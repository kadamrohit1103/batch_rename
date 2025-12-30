# Batch File Management Tools

This toolkit provides two powerful command-line utilities for managing files in bulk using CSV mappings:
1.  **Batch Rename Tool** (`rename_tool.bat`)
2.  **Batch Copy Tool** (`copy_tool.bat`)

Both tools are built with native PowerShell and Batch, requiring **no installation**.

---

## 1. Batch Rename Tool (`rename_tool.bat`)

Renames files in place based on your CSV list.

### Usage
```cmd
rename_tool.bat [Mapping.csv] [options]
```

### CSV Formats
The tool automatically detects the mode based on your CSV columns.

**Option A: 3-Column Mode (recommended)**
Gives you specific control. The tool validates the folder path once and processes all files within it (High Performance).
```csv
Folder Path,Old Name,New Name
D:\Photos,img_001.jpg,Summer_01.jpg
D:\Docs,report_draft.pdf,Final_Report.pdf
```

**Option B: 2-Column "Scan" Mode**
Searches for the "Old Name" in the current directory (or `-targetdir`).
```csv
Old Name,New Name
img_001.jpg,Summer_01.jpg
```

### Key Options
*   `-dryrun`: **Highly Recommended.** Previews changes without renaming anything.
*   `-undo`: Reverts the last operation.
*   `-conflict [skip|overwrite|autonumber]`: default is `skip`.
*   `-subfolders`: (Only for 2-Column Scan mode) Looks in subdirectories.

---

## 2. Batch Copy Tool (`copy_tool.bat`)

Copies files from a Source path to a Destination folder, with optional renaming.

### Usage
```cmd
copy_tool.bat [Mapping.csv] [options]
```

### CSV Format
Requires 2 or 3 columns.
```csv
Source Path,Destination Folder,New Name (Optional)
C:\Data\file.txt,D:\Backup,
C:\Data\image.png,D:\Backup\Images,renamed_image.png
D:\Archive.zip\data.txt,D:\Extracted,
```

### Features
*   **Auto-Conflict Handling**: If `file.txt` exists in the destination, it copies as `file_1.txt` automatically.
*   **Zip Support**: You can extract files directly from inside `.zip` files by treating the zip as a folder path (e.g., `C:\Archive.zip\folder\file.txt`).
*   **Dry Run**: Use `-dryrun` to see what will be copied.
*   **Undo**: Use `-undo` to delete the files copied in the last session.

---

## General Tips
*   **Performance**: The tools are optimized to group operations by folder, minimizing disk access.
*   **Undo**: If you used a specific folder for your files, you might need to specify it for undoing (e.g., `rename_tool.bat -undo -targetdir D:\MyFiles`).
