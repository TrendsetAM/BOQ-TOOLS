#!/usr/bin/env python3
"""
Build script for creating BOQ-Tools executable
This script automates the entire process of building a standalone executable.
"""

import os
import sys
import subprocess
import shutil
import tempfile
from pathlib import Path
from typing import List, Optional
import argparse
import logging

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('build.log')
    ]
)
logger = logging.getLogger(__name__)

class ExecutableBuilder:
    """Handles the creation of the executable"""
    
    def __init__(self, project_root: Path):
        self.project_root = project_root
        self.build_dir = project_root / 'build'
        self.dist_dir = project_root / 'dist'
        self.spec_file = project_root / 'boq_tools.spec'
        
    def clean_build_dirs(self):
        """Clean previous build directories"""
        logger.info("Cleaning previous build directories...")
        
        for dir_path in [self.build_dir, self.dist_dir]:
            if dir_path.exists():
                shutil.rmtree(dir_path)
                logger.info(f"Removed {dir_path}")
        
        # Also clean __pycache__ directories
        for pycache_dir in self.project_root.rglob('__pycache__'):
            shutil.rmtree(pycache_dir)
            logger.info(f"Removed {pycache_dir}")
    
    def check_python_version(self):
        """Check if Python version is compatible"""
        logger.info(f"Python version: {sys.version}")
        
        if sys.version_info < (3, 8):
            logger.error("Python 3.8 or higher is required")
            return False
        return True
    
    def install_dependencies(self):
        """Install required dependencies"""
        logger.info("Installing dependencies...")
        
        requirements_file = self.project_root / 'requirements.txt'
        if not requirements_file.exists():
            logger.error("requirements.txt not found")
            return False
        
        try:
            # Upgrade pip first
            subprocess.run([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'], 
                         check=True, capture_output=True, text=True)
            
            # Install requirements
            result = subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', str(requirements_file)], 
                                  check=True, capture_output=True, text=True)
            logger.info("Dependencies installed successfully")
            return True
            
        except subprocess.CalledProcessError as e:
            logger.error(f"Failed to install dependencies: {e}")
            logger.error(f"stdout: {e.stdout}")
            logger.error(f"stderr: {e.stderr}")
            return False
    
    def create_icon(self):
        """Create a default icon if none exists"""
        icon_path = self.project_root / 'resources' / 'icon.ico'
        
        if icon_path.exists():
            logger.info(f"Using existing icon: {icon_path}")
            return True
        
        # Create resources directory if it doesn't exist
        icon_path.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            # Create a simple icon using PIL
            from PIL import Image, ImageDraw, ImageFont
            
            # Create a 256x256 image
            img = Image.new('RGBA', (256, 256), (70, 130, 180, 255))  # Steel blue background
            draw = ImageDraw.Draw(img)
            
            # Draw a simple BOQ text
            try:
                # Try to use a larger font
                font = ImageFont.truetype("arial.ttf", 48)
            except:
                # Fall back to default font
                font = ImageFont.load_default()
            
            # Draw text
            text = "BOQ"
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            x = (256 - text_width) // 2
            y = (256 - text_height) // 2
            
            draw.text((x, y), text, fill=(255, 255, 255, 255), font=font)
            
            # Save as ICO
            img.save(str(icon_path), format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
            logger.info(f"Created default icon: {icon_path}")
            return True
            
        except Exception as e:
            logger.warning(f"Failed to create icon: {e}")
            return False
    
    def copy_user_dictionaries(self):
        """Copy user dictionaries to local config folder for bundling"""
        logger.info("Copying user dictionaries to local config folder...")
        
        try:
            # Add project root to Python path temporarily to import utils
            sys.path.insert(0, str(self.project_root))
            
            # Import the config utility
            from utils.config import get_user_config_path
            
            # Files to copy from user directory to local config
            config_files = [
                "category_dictionary.json",
                "canonical_mappings.json", 
                "boq_settings.json"
            ]
            
            local_config_dir = self.project_root / 'config'
            copied_files = []
            
            for filename in config_files:
                try:
                    # Get user config file path
                    user_file_path = Path(get_user_config_path(filename))
                    local_file_path = local_config_dir / filename
                    
                    # Copy if user file exists
                    if user_file_path.exists():
                        # Backup existing local file if it exists
                        if local_file_path.exists():
                            backup_path = local_file_path.with_suffix(f'.backup_{filename}')
                            shutil.copy2(local_file_path, backup_path)
                            logger.info(f"Backed up existing {filename} to {backup_path.name}")
                        
                        # Copy user file to local config
                        shutil.copy2(user_file_path, local_file_path)
                        copied_files.append(filename)
                        logger.info(f"Copied user {filename} from {user_file_path}")
                    else:
                        logger.warning(f"User config file not found: {user_file_path}")
                        
                except Exception as e:
                    logger.warning(f"Failed to copy {filename}: {e}")
            
            if copied_files:
                logger.info(f"Successfully copied {len(copied_files)} user config files: {', '.join(copied_files)}")
            else:
                logger.warning("No user config files were copied")
                
            return True
            
        except ImportError as e:
            logger.error(f"Failed to import utils.config: {e}")
            return False
        except Exception as e:
            logger.error(f"Error copying user dictionaries: {e}")
            return False
        finally:
            # Remove project root from Python path
            if str(self.project_root) in sys.path:
                sys.path.remove(str(self.project_root))
    
    def run_pyinstaller(self, debug: bool = False):
        """Run PyInstaller to create the executable"""
        logger.info("Running PyInstaller...")
        
        if not self.spec_file.exists():
            logger.error(f"Spec file not found: {self.spec_file}")
            return False
        
        cmd = [sys.executable, '-m', 'PyInstaller']
        
        if debug:
            cmd.append('--debug=all')
        
        cmd.extend([
            '--clean',
            '--noconfirm',
            str(self.spec_file)
        ])
        
        try:
            logger.info(f"Running command: {' '.join(cmd)}")
            result = subprocess.run(cmd, check=True, capture_output=True, text=True, cwd=str(self.project_root))
            logger.info("PyInstaller completed successfully")
            
            if result.stdout:
                logger.debug(f"PyInstaller stdout: {result.stdout}")
            if result.stderr:
                logger.debug(f"PyInstaller stderr: {result.stderr}")
            
            return True
            
        except subprocess.CalledProcessError as e:
            logger.error(f"PyInstaller failed: {e}")
            logger.error(f"stdout: {e.stdout}")
            logger.error(f"stderr: {e.stderr}")
            return False
    
    def create_installer_script(self):
        """Create a simple installer script"""
        installer_script = self.project_root / 'create_installer.iss'
        
        exe_path = self.dist_dir / 'BOQ-Tools.exe'
        if not exe_path.exists():
            logger.warning("Executable not found, skipping installer script creation")
            return False
        
        iss_content = f'''
[Setup]
AppName=BOQ Tools
AppVersion=1.0.0
AppPublisher=Your Company
AppPublisherURL=https://your-website.com
DefaultDirName={{pf}}\\BOQ Tools
DefaultGroupName=BOQ Tools
OutputDir={self.project_root / 'installer'}
OutputBaseFilename=BOQ-Tools-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{{cm:CreateDesktopIcon}}"; GroupDescription: "{{cm:AdditionalIcons}}"; Flags: unchecked

[Files]
Source: "{exe_path}"; DestDir: "{{app}}"; Flags: ignoreversion

[Icons]
Name: "{{group}}\\BOQ Tools"; Filename: "{{app}}\\BOQ-Tools.exe"
Name: "{{commondesktop}}\\BOQ Tools"; Filename: "{{app}}\\BOQ-Tools.exe"; Tasks: desktopicon

[Run]
Filename: "{{app}}\\BOQ-Tools.exe"; Description: "{{cm:LaunchProgram,BOQ Tools}}"; Flags: nowait postinstall skipifsilent
'''
        
        try:
            with open(installer_script, 'w') as f:
                f.write(iss_content)
            
            logger.info(f"Created installer script: {installer_script}")
            logger.info("To create an installer, install Inno Setup and compile the .iss file")
            return True
            
        except Exception as e:
            logger.error(f"Failed to create installer script: {e}")
            return False
    
    def verify_executable(self):
        """Verify that the executable was created and works"""
        exe_path = self.dist_dir / 'BOQ-Tools.exe'
        
        if not exe_path.exists():
            logger.error(f"Executable not found: {exe_path}")
            return False
        
        logger.info(f"Executable created: {exe_path}")
        logger.info(f"Executable size: {exe_path.stat().st_size / (1024*1024):.1f} MB")
        
        # For GUI applications, we can't easily test with command line
        # Just verify the file exists and has reasonable size
        file_size_mb = exe_path.stat().st_size / (1024*1024)
        if file_size_mb < 10:
            logger.warning(f"Executable seems too small ({file_size_mb:.1f} MB), may be missing dependencies")
        elif file_size_mb > 500:
            logger.warning(f"Executable seems very large ({file_size_mb:.1f} MB), consider optimizing")
        
        logger.info("Executable verification completed")
        return True
    
    def build(self, debug: bool = False, clean: bool = True):
        """Main build process"""
        logger.info("Starting BOQ-Tools executable build process...")
        
        # Check Python version
        if not self.check_python_version():
            return False
        
        # Clean build directories
        if clean:
            self.clean_build_dirs()
        
        # Install dependencies
        if not self.install_dependencies():
            return False
        
        # Create icon
        self.create_icon()

        # Copy user dictionaries to local config for bundling
        if not self.copy_user_dictionaries():
            logger.warning("Failed to copy user dictionaries, proceeding with existing config files")
        
        # Run PyInstaller
        if not self.run_pyinstaller(debug):
            return False
        
        # Verify executable
        if not self.verify_executable():
            return False
        
        # Create installer script
        self.create_installer_script()
        
        # Run standalone test
        self.run_standalone_test()
        
        logger.info("Build process completed successfully!")
        logger.info(f"Executable location: {self.dist_dir / 'BOQ-Tools.exe'}")
        
        return True
    
    def run_standalone_test(self):
        """Run the standalone test suite"""
        logger.info("Running standalone test suite...")
        
        exe_path = self.dist_dir / 'BOQ-Tools.exe'
        test_script = self.project_root / 'test_standalone.py'
        
        if not test_script.exists():
            logger.warning("Test script not found, skipping tests")
            return
        
        try:
            result = subprocess.run([
                sys.executable, str(test_script), str(exe_path)
            ], capture_output=True, text=True, timeout=60)
            
            if result.returncode == 0:
                logger.info("Standalone tests passed!")
            else:
                logger.warning("Some standalone tests failed")
            
            # Log test output
            if result.stdout:
                logger.info("Test output:")
                for line in result.stdout.split('\n'):
                    if line.strip():
                        logger.info(f"  {line}")
            
            if result.stderr:
                logger.warning("Test errors:")
                for line in result.stderr.split('\n'):
                    if line.strip():
                        logger.warning(f"  {line}")
            
        except subprocess.TimeoutExpired:
            logger.warning("Standalone tests timed out")
        except Exception as e:
            logger.warning(f"Could not run standalone tests: {e}")

def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(description='Build BOQ-Tools executable')
    parser.add_argument('--debug', action='store_true', help='Enable debug mode')
    parser.add_argument('--no-clean', action='store_true', help='Skip cleaning build directories')
    
    args = parser.parse_args()
    
    # Get project root
    project_root = Path(__file__).parent
    
    # Create builder and run
    builder = ExecutableBuilder(project_root)
    
    success = builder.build(
        debug=args.debug,
        clean=not args.no_clean
    )
    
    if success:
        print("\n" + "="*50)
        print("BUILD SUCCESSFUL!")
        print("="*50)
        print(f"Executable: {builder.dist_dir / 'BOQ-Tools.exe'}")
        print(f"Size: {(builder.dist_dir / 'BOQ-Tools.exe').stat().st_size / (1024*1024):.1f} MB")
        print("\nTo create an installer:")
        print("1. Install Inno Setup (https://jrsoftware.org/isinfo.php)")
        print("2. Compile the create_installer.iss file")
        sys.exit(0)
    else:
        print("\n" + "="*50)
        print("BUILD FAILED!")
        print("="*50)
        print("Check the build.log file for details")
        sys.exit(1)

if __name__ == '__main__':
    main() 