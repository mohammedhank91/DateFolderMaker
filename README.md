# Folder Creation Script

## Description

This PowerShell script is designed to automate the creation of a series of folders based on a specified date range. Each date in the range will have its own folder, which will contain two predefined subfolders. The script also provides visual feedback through a progress bar to indicate its progress as it creates the folder structure.

## How It Works

1. **Input Retrieval**:

   - The script retrieves values from the user interface, specifically from two text boxes: `$textBoxSubFolder2` and `$textBoxDirectory`.
   - These values are used to define the subfolder names (`$subfolder1`, `$subfolder2`) and the output directory (`$output_dir`), where the folders will be created.
2. **Directory Validation**:

   - The script checks if the specified output directory exists using the `Test-Path` cmdlet.
   - If the directory does not exist, the script throws an error ("Invalid output directory") and stops further execution to prevent errors.
3. **Progress Calculation**:

   - The total number of days in the specified date range (from start to end date) is calculated.
   - This number is used to set the maximum value of a progress bar, providing the user with visual feedback as the script runs.
4. **Folder Creation Loop**:

   - The script iterates through each day in the specified date range.
   - For each date, a main folder is created with the date formatted as `dd-MM-yy`.
   - Inside each main folder, two subfolders are created with names defined by `$subfolder1` and `$subfolder2`.
5. **Progress Update**:

   - As each folder is created, the script increments the progress bar to reflect the progress.

## Example Output

The folder structure created by this script will look like this:
```
Output Directory/
├── 10-11-24/
│ ├── Subfolder1/
│ └── Subfolder2/
├── 11-11-24/
│ ├── Subfolder1/
│ └── Subfolder2/
...
```

## Requirements

- PowerShell 5.1 or later.
- A user interface with the following components:
  - **TextBox for Start Date** (`dateTimePickerStart`)
  - **TextBox for End Date** (`dateTimePickerEnd`)
  - **TextBox for Subfolder Names** (`$textBoxSubFolder1`, `$textBoxSubFolder2`)
  - **Progress Bar** for tracking progress.

## Usage

1. Specify the start and end dates in the date text boxes.
2. Enter the names of the subfolders.
3. Select an output directory where the folder structure will be created.
4. Run the script. The progress bar will show the current progress as folders are created.

## Error Handling

If the specified output directory does not exist, the script will display an error message and stop.

## License

This script is provided as-is without warranty. Feel free to modify it for your own use.
