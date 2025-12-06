# Batch Renaming Tool - User Guide

This tool allows you to batch rename files using a CSV mapping. It is built using native Windows tools (PowerShell and Batch), so no extra installations like Python are required.

## Installation

No installation needed! The tool consists of two files in your directory:
1.  `rename_tool.bat` (The wrapper you run)
2.  `AdvancedRenamer.ps1` (The logic script)

Ensure both are in the same folder.

## Usage

Open `d:/kadam/Documents/batch_files_rename_cmd_tool` in File Explorer, type `cmd` in the address bar to open a terminal, or just drag `rename_tool.bat` into a terminal window.

### Basic Syntax
```cmd
rename_tool.bat [YourCSVFile.csv] [options]
```

### Options
- `[CSV File]`: The path to your CSV file. Must have at least two columns (Old Name, New Name).
- `-dryrun`: **Always recommended first!** Shows what *will* happen without changing anything.
- `-undo`: Undoes the last batch of renames.
- `-subfolders`: Searches subdirectories for the "Old Name" files.
- `-conflict [skip|overwrite|autonumber]`: Decides what to do if the new filename already exists. (Default: skip)

### Examples

**1. Preview changes (Safe Mode)**
```cmd
rename_tool.bat my_list.csv -dryrun
```

**2. Rename files**
```cmd
rename_tool.bat my_list.csv
```

**3. Rename recursively (find files in subfolders)**
```cmd
rename_tool.bat my_list.csv -subfolders
```

### Organizing Your Files (Best Practice)
You don't need to put the script, the CSV, and the files all in the same folder!
*   **Keep the tool clean**: Put `rename_tool.bat` and `AdvancedRenamer.ps1` in a dedicated folder (e.g., `C:\Apps\BatchRenamer`).
*   **Your Data**: Keep your CSV and files in your working folder (e.g., `D:\Photos`).
*   **How to run**:
    1.  Open CMD in your working folder (`D:\Photos`).
    2.  Run the tool by its full path: `C:\Apps\BatchRenamer\rename_tool.bat my_list.csv`
    3.  Alternatively, use the `-targetdir` option:
        ```cmd
        rename_tool.bat D:\Photos\my_list.csv -targetdir D:\Photos
        ```

**4. Undo the last mistake**
> [!IMPORTANT]
> If you specified a `-targetdir` when renaming (or if your files are in a different folder than the script), you **MUST** specify where to look for the undo log!

```cmd
rename_tool.bat -undo -targetdir D:\Photos
```

## CSV Format
Your CSV should look like this (headers are optional but recommended, the tool uses the first two columns by default):
```csv
Old Name,New Name
photo001.jpg,vacation_001.jpg
photo002.jpg,vacation_002.jpg
```
