# DST-Edit

**[Open DST-Edit →](https://abrman.github.io/dst-edit/)**

DST-Edit is a web-based tool designed to simplify the process of editing AutoCAD Sheet Set (\*.dst) files. It allows users to modify sheet names, numbers, and various other fields, including custom fields, directly through a user-friendly web interface. Whether you need to update information for multiple sheets or customize specific fields, DST-Edit has got you covered.

## Features

- **Web-Based Editing:** Access and edit AutoCAD Sheet Set files directly through your web browser. Files never leave your computer.
- **Folders and Order:** See sheets in their Sheet Set Manager folders (subsets). Drag to reorder, move sheets between folders, create folders, and remove folders while keeping their sheets.
- **Spreadsheet-Style Editing:** Edit numbers, titles, descriptions and custom properties in a grid, paste straight from Excel, and undo any change.
- **Bulk Edits:** Number series (e.g. `C1.01`, `C1.02`… or restarting per folder), set value, find & replace, and prefix/suffix across selected sheets.
- **Sheet Set Settings:** Edit the sheet set name shown in AutoCAD, project details, the default publish/plot folder, and add, rename or remove custom properties.
- **CSV Round-Trip:** Export the sheet list to CSV, edit it in Excel, and import it back; rows are matched to sheets by ID.
- **AutoCAD Text Rules:** A `/` in a sheet number, title or description becomes a look-alike slash (AutoCAD doesn't allow the real one), and titles over 64 characters are flagged and block saving.
- **Drawing Path Check:** Flags sheets whose drawings are stored in a user profile folder and offers to repoint them to a shared path, with the matching `mklink` command.
- **Safe Round-Trip:** Edits are written back into the original XML, so anything DST-Edit doesn't know about is kept. Works with sheet sets from AutoCAD and ARES Commander, including non-ASCII characters.
- **GitHub Pages Integration:** Utilize the tool seamlessly via GitHub Pages at [https://abrman.github.io/dst-edit/](https://abrman.github.io/dst-edit/).

Sheets can be removed from a sheet set but not added, because each sheet links to a layout in a drawing (DWG); add new sheets in AutoCAD.

## Usage

1. Visit the [DST-Edit GitHub Pages](https://abrman.github.io/dst-edit/).
2. Upload your AutoCAD Sheet Set (\*.dst) file.
3. Edit sheet information and custom fields as required.
4. Save the changes and download the updated Sheet Set file.

## How to Contribute

If you have any suggestions, find bugs, or want to contribute to DST-Edit, feel free to open an issue or submit a pull request. Your contributions are highly appreciated!

## Credits

Thanks to [Bedz01](https://github.com/Bedz01) for rebuilding the .dst decoder ([#4](https://github.com/abrman/dst-edit/pull/4)). He replaced the byte lookup table with the general cipher formula, so characters such as em dashes and `{ | } ~` are no longer corrupted. He also fixed saves from Firefox so ARES Commander accepts them.

---

Happy editing with DST-Edit!
