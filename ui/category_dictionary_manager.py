"""
Category Dictionary Manager Dialog
Allows users to browse, edit, and bulk update dictionary mappings.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional

import tkinter as tk
from tkinter import messagebox, ttk

from core.category_dictionary import CategoryDictionary, CategoryMapping

logger = logging.getLogger(__name__)


@dataclass
class MappingRow:
    """Convenience structure for managing mapping rows inside the dialog."""

    description: str
    normalized_description: str
    category: str
    confidence: float
    notes: Optional[str]
    created_date: Optional[str]
    usage_count: int
    original_description: Optional[str] = None

    @classmethod
    def from_dict(cls, data: Dict[str, object]) -> "MappingRow":
        return cls(
            description=str(data.get("description", "")),
            normalized_description=str(data.get("normalized_description", "")).lower().strip(),
            category=str(data.get("category", "")),
            confidence=float(data.get("confidence", 1.0)),
            notes=data.get("notes"),
            created_date=data.get("created_date"),
            usage_count=int(data.get("usage_count", 0)),
            original_description=data.get("original_description"),
        )

    def to_payload(self) -> Dict[str, object]:
        return {
            "description": self.description,
            "category": self.category,
            "confidence": self.confidence,
            "notes": self.notes,
            "created_date": self.created_date,
            "usage_count": self.usage_count,
            "original_description": self.original_description or self.description,
        }


class CategoryDictionaryManager:
    """Dialog that enables in-depth management of the category dictionary."""

    def __init__(self, parent: tk.Misc, category_dictionary: CategoryDictionary):
        self.parent = parent
        self.category_dictionary = category_dictionary

        self.dialog: Optional[tk.Toplevel] = None
        self.tree: Optional[ttk.Treeview] = None
        self.save_button: Optional[ttk.Button] = None
        self.delete_button: Optional[ttk.Button] = None
        self.bulk_category_button: Optional[ttk.Button] = None
        self.edit_button: Optional[ttk.Button] = None

        self.filter_var = tk.StringVar()
        self.status_var = tk.StringVar()

        self._dirty = False
        self._mappings: Dict[str, MappingRow] = {}
        self._deleted_descriptions: set[str] = set()

        self._load_snapshot()
        self._create_dialog()

    # ------------------------------------------------------------------
    # Initialization helpers
    # ------------------------------------------------------------------
    def _load_snapshot(self) -> None:
        """Load current dictionary snapshot into local structures."""
        self._mappings.clear()
        try:
            snapshot = self.category_dictionary.list_mappings()
        except Exception as exc:
            logger.error("Failed to load dictionary mappings: %s", exc)
            messagebox.showerror("Error", f"Unable to load category dictionary:\n{exc}")
            snapshot = []

        for row in snapshot:
            mapping_row = MappingRow.from_dict(row)
            if not mapping_row.normalized_description:
                continue
            self._mappings[mapping_row.normalized_description] = mapping_row

    def _create_dialog(self) -> None:
        """Create and display the modal dialog."""
        dialog = tk.Toplevel(self.parent)
        dialog.title("Manage Category Dictionary")
        dialog.geometry("900x600")
        dialog.minsize(780, 520)
        dialog.transient(self.parent)
        dialog.grab_set()
        dialog.configure(padx=12, pady=10)

        self.dialog = dialog

        # Configure grid
        dialog.grid_rowconfigure(2, weight=1)
        dialog.grid_columnconfigure(0, weight=1)

        self._create_header(dialog)
        self._create_toolbar(dialog)
        self._create_tree(dialog)
        self._create_footer(dialog)

        dialog.protocol("WM_DELETE_WINDOW", self._on_close)

        self._populate_tree()
        self._update_status()

    def _create_header(self, container: tk.Toplevel) -> None:
        header = ttk.Frame(container)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 8))
        header.columnconfigure(1, weight=1)

        ttk.Label(header, text="Filter:", font=("TkDefaultFont", 10, "bold")).grid(
            row=0, column=0, padx=(0, 8)
        )

        filter_entry = ttk.Entry(header, textvariable=self.filter_var)
        filter_entry.grid(row=0, column=1, sticky=tk.EW)
        filter_entry.bind("<KeyRelease>", lambda _event: self._populate_tree())

        ttk.Label(
            header,
            text="Type to filter descriptions or categories (case-insensitive).",
        ).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(4, 0))

    def _create_toolbar(self, container: tk.Toplevel) -> None:
        toolbar = ttk.Frame(container)
        toolbar.grid(row=1, column=0, sticky=tk.EW, pady=(0, 8))
        toolbar.columnconfigure(4, weight=1)

        add_button = ttk.Button(toolbar, text="Add Mapping", command=self._add_mapping)
        add_button.grid(row=0, column=0, padx=(0, 6))

        edit_button = ttk.Button(toolbar, text="Edit Selected", command=self._edit_selected)
        edit_button.grid(row=0, column=1, padx=(0, 6))
        self.edit_button = edit_button

        bulk_button = ttk.Button(
            toolbar, text="Set Category…", command=self._bulk_set_category
        )
        bulk_button.grid(row=0, column=2, padx=(0, 6))
        self.bulk_category_button = bulk_button

        delete_button = ttk.Button(
            toolbar, text="Delete Selected", command=self._delete_selected
        )
        delete_button.grid(row=0, column=3)
        self.delete_button = delete_button

        for widget in (edit_button, bulk_button, delete_button):
            widget.state(["disabled"])

    def _create_tree(self, container: tk.Toplevel) -> None:
        tree_frame = ttk.Frame(container)
        tree_frame.grid(row=2, column=0, sticky=tk.NSEW)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        columns = ("Description", "Category", "Confidence", "Notes", "Usage Count")
        tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            selectmode="extended",
        )
        self.tree = tree

        tree.heading("Description", text="Description")
        tree.heading("Category", text="Category")
        tree.heading("Confidence", text="Confidence")
        tree.heading("Notes", text="Notes")
        tree.heading("Usage Count", text="Usage Count")

        tree.column("Description", width=360, anchor=tk.W)
        tree.column("Category", width=160, anchor=tk.W)
        tree.column("Confidence", width=100, anchor=tk.CENTER)
        tree.column("Notes", width=200, anchor=tk.W)
        tree.column("Usage Count", width=100, anchor=tk.CENTER)

        y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        tree.grid(row=0, column=0, sticky=tk.NSEW)
        y_scroll.grid(row=0, column=1, sticky=tk.NS)
        x_scroll.grid(row=1, column=0, sticky=tk.EW)

        tree.bind("<<TreeviewSelect>>", lambda _event: self._refresh_action_states())
        tree.bind("<Double-1>", lambda _event: self._edit_selected())

    def _create_footer(self, container: tk.Toplevel) -> None:
        footer = ttk.Frame(container)
        footer.grid(row=3, column=0, sticky=tk.EW, pady=(8, 0))
        footer.columnconfigure(0, weight=1)

        status_label = ttk.Label(footer, textvariable=self.status_var)
        status_label.grid(row=0, column=0, sticky=tk.W)

        button_frame = ttk.Frame(footer)
        button_frame.grid(row=0, column=1, sticky=tk.E)

        cancel_button = ttk.Button(button_frame, text="Cancel", command=self._on_close)
        cancel_button.grid(row=0, column=0, padx=(0, 6))

        save_button = ttk.Button(button_frame, text="Save Changes", command=self._save_changes)
        save_button.grid(row=0, column=1)
        save_button.state(["disabled"])
        self.save_button = save_button

    # ------------------------------------------------------------------
    # Tree population & filtering
    # ------------------------------------------------------------------
    def _filtered_rows(self) -> List[MappingRow]:
        raw_filter = self.filter_var.get().strip().lower()
        rows = list(self._mappings.values())
        if not raw_filter:
            return sorted(rows, key=lambda row: row.description.lower())

        return sorted(
            [
                row
                for row in rows
                if raw_filter in row.description.lower()
                or raw_filter in row.category.lower()
            ],
            key=lambda row: row.description.lower(),
        )

    def _populate_tree(self) -> None:
        if not self.tree:
            return

        selection = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())

        for row in self._filtered_rows():
            iid = row.normalized_description
            notes = row.notes or ""
            confidence_str = f"{row.confidence:.2f}"
            self.tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    row.description,
                    row.category,
                    confidence_str,
                    notes,
                    row.usage_count,
                ),
            )
            if iid in selection:
                self.tree.selection_add(iid)

        self._refresh_action_states()

    # ------------------------------------------------------------------
    # Action helpers
    # ------------------------------------------------------------------
    def _add_mapping(self) -> None:
        categories = self.category_dictionary.get_all_categories()
        editor = MappingEditDialog(self.dialog, categories=categories)
        result = editor.show()
        if result is None:
            return
        self._upsert_local_mapping(None, result)

    def _edit_selected(self) -> None:
        if not self.tree:
            return
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("Edit Mapping", "Please select a mapping to edit.")
            return
        if len(selected) > 1:
            messagebox.showinfo(
                "Edit Mapping", "Please select a single mapping to edit at a time."
            )
            return

        key = selected[0]
        mapping = self._mappings.get(key)
        if not mapping:
            messagebox.showerror("Edit Mapping", "Selected mapping could not be located.")
            return

        categories = self.category_dictionary.get_all_categories()
        editor = MappingEditDialog(
            self.dialog,
            categories=categories,
            mapping=mapping,
        )
        result = editor.show()
        if result is None:
            return

        self._upsert_local_mapping(mapping, result)

    def _bulk_set_category(self) -> None:
        if not self.tree:
            return
        selection = self.tree.selection()
        if not selection:
            return

        categories = self.category_dictionary.get_all_categories()
        dialog = BulkCategoryDialog(self.dialog, categories)
        new_category = dialog.show()
        if not new_category:
            return

        updated = 0
        for key in selection:
            mapping = self._mappings.get(key)
            if not mapping:
                continue
            if mapping.category == new_category:
                continue
            mapping.category = new_category
            updated += 1

        if updated:
            self._mark_dirty()
            self._populate_tree()
            self._update_status(
                f"Applied category '{new_category}' to {updated} mapping(s)."
            )

    def _delete_selected(self) -> None:
        if not self.tree:
            return
        selection = self.tree.selection()
        if not selection:
            return

        confirm = messagebox.askyesno(
            "Delete Mappings",
            f"Delete {len(selection)} selected mapping(s)? This cannot be undone.",
            parent=self.dialog,
        )
        if not confirm:
            return

        for key in selection:
            if key in self._mappings:
                del self._mappings[key]
                self._deleted_descriptions.add(key)

        self._mark_dirty()
        self._populate_tree()
        self._update_status(f"Deleted {len(selection)} mapping(s).")

    def _refresh_action_states(self) -> None:
        if not self.tree:
            return

        selection = self.tree.selection()
        has_selection = bool(selection)
        has_single_selection = len(selection) == 1

        for widget, enabled in (
            (self.delete_button, has_selection),
            (self.bulk_category_button, has_selection),
            (self.edit_button, has_single_selection),
        ):
            if widget is None:
                continue
            if enabled:
                widget.state(["!disabled"])
            else:
                widget.state(["disabled"])

    # ------------------------------------------------------------------
    # Persistence helpers
    # ------------------------------------------------------------------
    def _upsert_local_mapping(
        self,
        original: Optional[MappingRow],
        new_values: Dict[str, object],
    ) -> None:
        raw_description = str(new_values.get("description", "")).strip()
        if not raw_description:
            messagebox.showerror("Invalid Mapping", "Description cannot be empty.")
            return

        normalized = raw_description.lower()

        if original and normalized != original.normalized_description:
            self._deleted_descriptions.add(original.normalized_description)
            self._mappings.pop(original.normalized_description, None)

        self._deleted_descriptions.discard(normalized)

        if not original and normalized in self._mappings:
            messagebox.showerror(
                "Duplicate Description",
                "A mapping with this description already exists.",
            )
            return

        mapping_row = MappingRow(
            description=raw_description,
            normalized_description=normalized,
            category=str(new_values.get("category", "")).strip(),
            confidence=float(new_values.get("confidence", 1.0)),
            notes=new_values.get("notes"),
            created_date=new_values.get("created_date"),
            usage_count=int(new_values.get("usage_count", 0)),
            original_description=new_values.get("original_description") or raw_description,
        )

        self._mappings[normalized] = mapping_row
        self._mark_dirty()
        self._populate_tree()
        self._update_status("Mapping saved locally. Remember to save changes.")

    def _save_changes(self) -> None:
        if not self._dirty:
            return

        if not messagebox.askyesno(
            "Save Changes",
            "Save all pending dictionary changes?",
            parent=self.dialog,
        ):
            return

        payloads = [row.to_payload() for row in self._mappings.values()]
        backup_path = self.category_dictionary.backup_current_file()
        deleted = 0
        added = 0
        updated = 0

        try:
            if self._deleted_descriptions:
                deleted = self.category_dictionary.delete_mappings(
                    self._deleted_descriptions
                )

            added, updated = self.category_dictionary.upsert_mappings(payloads)
            if not self.category_dictionary.save_dictionary():
                raise RuntimeError("Failed to persist dictionary file.")
        except Exception as exc:
            logger.exception("Failed to persist category dictionary changes.")
            messagebox.showerror(
                "Save Failed",
                f"An error occurred while saving changes:\n{exc}",
                parent=self.dialog,
            )
            return

        # If we reach here, persistence succeeded
        self._dirty = False
        self._deleted_descriptions.clear()
        self._load_snapshot()
        self._populate_tree()

        summary = (
            f"Changes saved successfully.\n\n"
            f"Added: {added}\n"
            f"Updated: {updated}\n"
            f"Deleted: {deleted}"
        )
        if backup_path:
            summary += f"\nBackup: {backup_path.name}"

        messagebox.showinfo("Save Complete", summary, parent=self.dialog)
        self._update_status("All changes saved.")

    # ------------------------------------------------------------------
    # Dialog lifecycle helpers
    # ------------------------------------------------------------------
    def _mark_dirty(self) -> None:
        self._dirty = True
        if self.save_button:
            self.save_button.state(["!disabled"])
        if self.dialog and not self.dialog.title().endswith("*"):
            self.dialog.title("Manage Category Dictionary*")

    def _update_status(self, message: Optional[str] = None) -> None:
        if message:
            self.status_var.set(message)
        else:
            self.status_var.set(
                f"Loaded {len(self._mappings)} mapping(s). "
                f"{len(self._deleted_descriptions)} pending deletion."
            )
        if self._dirty and self.save_button:
            self.save_button.state(["!disabled"])
        elif not self._dirty and self.save_button:
            self.save_button.state(["disabled"])
        if self.dialog:
            title = "Manage Category Dictionary"
            if self._dirty:
                title += "*"
            self.dialog.title(title)

    def _on_close(self) -> None:
        if self._dirty:
            choice = messagebox.askyesnocancel(
                "Unsaved Changes",
                "You have unsaved changes.\n"
                "Save before closing?",
                parent=self.dialog,
            )
            if choice is None:
                return
            if choice:
                self._save_changes()
                if self._dirty:
                    return
            else:
                # User chose to discard changes
                pass

        if self.dialog:
            self.dialog.destroy()
            self.dialog = None


class MappingEditDialog:
    """Modal dialog for editing or creating a single mapping."""

    def __init__(
        self,
        parent: tk.Misc,
        categories: Iterable[str],
        mapping: Optional[MappingRow] = None,
    ):
        self.parent = parent
        self.categories = sorted(set(categories))
        self.mapping = mapping
        self.result: Optional[Dict[str, object]] = None

        self.dialog: Optional[tk.Toplevel] = None

        self.description_var = tk.StringVar(value=mapping.description if mapping else "")
        self.category_var = tk.StringVar(value=mapping.category if mapping else "")
        self.confidence_var = tk.StringVar(
            value=f"{mapping.confidence:.2f}" if mapping else "1.0"
        )
        initial_notes = ""
        if mapping and mapping.notes:
            initial_notes = str(mapping.notes)
        self.notes_var = tk.StringVar(value=initial_notes)

    def show(self) -> Optional[Dict[str, object]]:
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Edit Mapping" if self.mapping else "Add Mapping")
        self.dialog.geometry("420x260")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()

        frame = ttk.Frame(self.dialog, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Description:").grid(row=0, column=0, sticky=tk.W, pady=4)
        description_entry = ttk.Entry(frame, textvariable=self.description_var)
        description_entry.grid(row=0, column=1, sticky=tk.EW, pady=4)

        ttk.Label(frame, text="Category:").grid(row=1, column=0, sticky=tk.W, pady=4)
        category_combo = ttk.Combobox(
            frame,
            textvariable=self.category_var,
            values=self.categories,
            state="normal",
        )
        category_combo.grid(row=1, column=1, sticky=tk.EW, pady=4)

        ttk.Label(frame, text="Confidence (0-1):").grid(
            row=2, column=0, sticky=tk.W, pady=4
        )
        confidence_entry = ttk.Entry(frame, textvariable=self.confidence_var)
        confidence_entry.grid(row=2, column=1, sticky=tk.EW, pady=4)

        ttk.Label(frame, text="Notes:").grid(row=3, column=0, sticky=tk.W, pady=4)
        notes_entry = ttk.Entry(frame, textvariable=self.notes_var)
        notes_entry.grid(row=3, column=1, sticky=tk.EW, pady=4)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(12, 0), sticky=tk.E)

        ttk.Button(button_frame, text="Cancel", command=self._cancel).pack(
            side=tk.RIGHT, padx=(0, 6)
        )
        ttk.Button(button_frame, text="Save", command=self._submit).pack(side=tk.RIGHT)

        description_entry.focus_set()
        self.dialog.protocol("WM_DELETE_WINDOW", self._cancel)

        self.dialog.wait_window()
        return self.result

    def _submit(self) -> None:
        description = self.description_var.get().strip()
        category = self.category_var.get().strip()
        confidence_raw = self.confidence_var.get().strip() or "1.0"
        notes = self.notes_var.get().strip() or None

        if not description:
            messagebox.showerror("Invalid Input", "Description cannot be empty.", parent=self.dialog)
            return
        if not category:
            messagebox.showerror("Invalid Input", "Category cannot be empty.", parent=self.dialog)
            return

        try:
            confidence = float(confidence_raw)
        except ValueError:
            messagebox.showerror(
                "Invalid Input",
                "Confidence must be a number between 0 and 1.",
                parent=self.dialog,
            )
            return

        if not (0.0 <= confidence <= 1.0):
            messagebox.showerror(
                "Invalid Input",
                "Confidence must be between 0 and 1.",
                parent=self.dialog,
            )
            return

        self.result = {
            "description": description,
            "category": category,
            "confidence": confidence,
            "notes": notes,
            "usage_count": self.mapping.usage_count if self.mapping else 0,
            "created_date": self.mapping.created_date if self.mapping else None,
            "original_description": description,
        }
        self.dialog.destroy()

    def _cancel(self) -> None:
        self.result = None
        if self.dialog:
            self.dialog.destroy()


class BulkCategoryDialog:
    """Dialog used to select a category for bulk updates."""

    def __init__(self, parent: tk.Misc, categories: Iterable[str]):
        self.parent = parent
        self.categories = sorted(set(categories))
        self.dialog: Optional[tk.Toplevel] = None
        self.category_var = tk.StringVar()
        self.result: Optional[str] = None

    def show(self) -> Optional[str]:
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Set Category for Selection")
        self.dialog.geometry("360x180")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()

        frame = ttk.Frame(self.dialog, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(0, weight=1)

        ttk.Label(
            frame,
            text="Select an existing category or type a new one:",
            wraplength=320,
        ).pack(anchor=tk.W, pady=(0, 8))

        category_combo = ttk.Combobox(
            frame,
            textvariable=self.category_var,
            values=self.categories,
            state="normal",
        )
        category_combo.pack(fill=tk.X, pady=(0, 12))

        button_frame = ttk.Frame(frame)
        button_frame.pack(anchor=tk.E)

        ttk.Button(button_frame, text="Cancel", command=self._cancel).pack(
            side=tk.RIGHT, padx=(0, 6)
        )
        ttk.Button(button_frame, text="Apply", command=self._submit).pack(side=tk.RIGHT)

        category_combo.focus_set()
        self.dialog.protocol("WM_DELETE_WINDOW", self._cancel)
        self.dialog.wait_window()
        return self.result

    def _submit(self) -> None:
        value = self.category_var.get().strip()
        if not value:
            messagebox.showerror(
                "Invalid Category",
                "Category name cannot be empty.",
                parent=self.dialog,
            )
            return

        self.result = value
        self.dialog.destroy()

    def _cancel(self) -> None:
        self.result = None
        if self.dialog:
            self.dialog.destroy()


__all__ = ["CategoryDictionaryManager"]

