# BOQ Tools Comparison User Guide

## Overview

The BOQ Tools comparison feature allows you to analyze multiple Bill of Quantities (BOQ) files and merge offer-specific data into a comprehensive dataset. This is particularly useful for comparing different contractor offers or analyzing variations in BOQ items across different projects.

## Prerequisites

Before using the comparison feature, ensure that:

1. **Master BOQ is Categorized**: Your primary BOQ file must be fully processed and categorized
2. **File Compatibility**: Comparison files should have similar structure to the master file
3. **Column Mappings**: Both files should have compatible column mappings

## Comparison Workflow

### Step 1: Prepare Master Dataset

1. **Load Master BOQ**: Open your primary BOQ file using the main interface
2. **Process File**: Complete the file processing workflow including:
   - Sheet classification
   - Column mapping
   - Row validation
   - Categorization
3. **Verify Categorization**: Ensure all items are properly categorized
4. **Save Work**: Optionally save your work to resume later

### Step 2: Start Comparison

1. **Select "Compare Full"**: From the main interface, click the "Compare Full" button
2. **Choose Comparison File**: Select the BOQ file you want to compare against the master
3. **Provide Offer Information**: Enter details about the comparison offer:
   - **Offer Name**: A unique identifier for this offer (e.g., "Contractor A", "Bid 2024")
   - **Offer Date**: When the offer was submitted
   - **Additional Notes**: Any relevant information about the offer

### Step 3: File Validation

The system automatically validates the comparison file:

- **Sheet Compatibility**: Ensures both files have the same sheet structure
- **Column Mapping**: Validates that column mappings are compatible
- **Data Structure**: Checks for required columns and data format

If validation fails, you'll see an error message explaining the issue.

### Step 4: Row Review and Validation

The comparison row review dialog allows you to manually validate each row:

#### Understanding the Interface

- **Row List**: Shows all comparison rows with their key information
- **Status Column**: Indicates whether each row is valid (green) or invalid (red)
- **Reason Column**: Explains why a row was marked as invalid
- **Summary**: Real-time statistics showing valid/invalid row counts

#### Manual Validation

1. **Review Each Row**: Examine the description, quantity, and pricing information
2. **Toggle Validity**: Click on any row to toggle its validity status
3. **Check Reasons**: Review the reason for any automatically invalidated rows
4. **Confirm Review**: Click "Confirm" when you're satisfied with the validation

#### Common Validation Scenarios

- **Valid Rows**: Rows that match the master dataset structure and contain valid data
- **Invalid Rows**: Rows with missing data, formatting issues, or structural problems
- **Manual Override**: You can manually validate rows that were auto-invalidated or vice versa

### Step 5: Data Processing

Once you confirm the row review, the system processes the comparison data:

#### MERGE Operations
- **Matched Items**: For items that exist in both master and comparison datasets
- **Offer-Specific Columns**: Creates new columns like `quantity[OfferName]`, `unit_price[OfferName]`
- **Data Updates**: Updates the master dataset with comparison values

#### ADD Operations
- **New Items**: For items that exist only in the comparison dataset
- **Row Addition**: Adds new rows to the master dataset
- **Position Assignment**: Assigns appropriate position numbers to new items

#### Instance Management
- **Multiple Instances**: Handles cases where the same item appears multiple times
- **Instance Matching**: Matches instances based on description and other criteria
- **Data Consolidation**: Consolidates data from multiple instances

### Step 6: Results and Export

After processing, you can:

1. **View Results**: See a summary of the comparison processing
2. **Review Statistics**: Check how many rows were merged, added, or skipped
3. **Export Data**: Export the final dataset with offer-specific columns
4. **Save Work**: Save the comparison results for future reference

## Understanding the Results

### Offer-Specific Columns

The comparison process creates new columns for each offer:

- `quantity[OfferName]`: Quantity values from the comparison offer
- `unit_price[OfferName]`: Unit price values from the comparison offer
- `total_price[OfferName]`: Total price values from the comparison offer
- `manhours[OfferName]`: Manhours values from the comparison offer
- `wage[OfferName]`: Wage values from the comparison offer

### Data Structure

The final dataset contains:

- **Base Columns**: Original master dataset columns (description, code, unit, etc.)
- **Master Values**: Values from the original master dataset
- **Offer Columns**: New columns containing comparison offer data
- **Categorization**: Category information for all items

### Processing Statistics

The system provides detailed statistics:

- **Total Rows Processed**: Number of rows in the comparison file
- **Valid Rows**: Number of rows that passed validation
- **Invalid Rows**: Number of rows that failed validation
- **Merged Rows**: Number of rows that were merged with existing master rows
- **Added Rows**: Number of new rows added to the master dataset
- **Errors**: Any processing errors or warnings

## Best Practices

### File Preparation

1. **Consistent Format**: Ensure comparison files follow the same format as the master
2. **Data Quality**: Clean and validate data before comparison
3. **Column Headers**: Use consistent column headers across files
4. **Data Types**: Ensure numeric data is properly formatted

### Comparison Process

1. **Review Carefully**: Take time to review each comparison row
2. **Validate Manually**: Don't rely solely on automatic validation
3. **Document Decisions**: Keep notes on manual validation decisions
4. **Test with Small Files**: Practice with small files before large comparisons

### Data Management

1. **Backup Master**: Always backup your master dataset before comparison
2. **Save Incrementally**: Save work at each step of the process
3. **Version Control**: Keep track of different comparison versions
4. **Export Results**: Export final results for external analysis

## Troubleshooting

### Common Issues

#### Validation Errors
- **Missing Columns**: Ensure comparison file has required columns
- **Data Type Mismatches**: Check that numeric data is properly formatted
- **Sheet Structure**: Verify that sheet names and structure match

#### Processing Errors
- **Memory Issues**: For large files, consider processing in smaller chunks
- **Data Corruption**: Check for corrupted or malformed data
- **Encoding Issues**: Ensure files use consistent character encoding

#### UI Issues
- **Dialog Not Appearing**: Check that the master file is properly categorized
- **Row Review Problems**: Restart the comparison process if the dialog becomes unresponsive
- **Progress Stuck**: Cancel and restart the process if progress appears stuck

### Error Messages

#### "Master BoQ must be categorized before comparison"
- **Solution**: Complete the categorization process for your master file
- **Prevention**: Always categorize files before starting comparison

#### "Could not find master file mapping"
- **Solution**: Close and reopen the master file
- **Prevention**: Ensure the master file is properly loaded

#### "Comparison data validation failed"
- **Solution**: Check the comparison file structure and data quality
- **Prevention**: Validate comparison files before starting the process

## Advanced Features

### Multiple Comparisons

You can perform multiple comparisons against the same master dataset:

1. **Sequential Processing**: Process one comparison file at a time
2. **Cumulative Results**: Each comparison adds new offer columns
3. **Data Integrity**: The system maintains data integrity across multiple comparisons

### Data Export Options

- **Excel Export**: Export to Excel with formatting and styling
- **CSV Export**: Export to CSV for external analysis
- **Custom Formats**: Configure export settings in the application

### Configuration Options

- **Validation Thresholds**: Adjust validation sensitivity
- **Processing Limits**: Configure memory and performance limits
- **Column Mappings**: Customize column mapping behavior

## Support and Resources

### Documentation
- **README.md**: General application documentation
- **IMPLEMENTATION_SUMMARY.md**: Technical implementation details
- **Configuration Guide**: Settings and configuration options

### Testing
- **Demo Files**: Use provided demo files to practice the comparison workflow
- **Test Scenarios**: Various test scenarios are available for validation

### Troubleshooting
- **Error Logs**: Check application logs for detailed error information
- **Validation Reports**: Review validation reports for data quality issues
- **User Feedback**: Report issues through the application's feedback system

## Conclusion

The BOQ Tools comparison feature provides powerful capabilities for analyzing multiple BOQ files and merging offer-specific data. By following this guide and best practices, you can effectively use the comparison workflow to enhance your BOQ analysis and decision-making processes.

For additional support or questions, refer to the main application documentation or contact the development team. 