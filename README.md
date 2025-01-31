# SlideMasterRemover

This is a simple C# console application that removes any **unused slide masters** and **custom layouts** from a specified PowerPoint presentation file. It helps to clean up your `.pptx` or `.pptm` files by deleting unnecessary designs or layouts, potentially reducing file size and preventing clutter.

## Features

- **Opens a PowerPoint file** (supports `.pptx`, `.pptm`, and `.ppt` by default, with a user warning for unknown extensions).
- **Scans each slide** to determine which layouts are actively used.
- **Removes unused layouts** and **deletes designs** if they are completely unused.
- **Saves changes** directly to the specified PowerPoint file.

## Prerequisites

- **Windows OS** (the application uses COM interop with Microsoft PowerPoint).
- **Microsoft Office / PowerPoint** installed (the application relies on Microsoft.Office.Interop.PowerPoint).
- **.NET Framework / .NET runtime** to execute the C# application (depending on the version you compile against).

## How to Build

1. Clone or download this repository.
2. Open the solution (`.sln`) or project (`.csproj`) in Visual Studio (or your preferred IDE).
3. Restore any required NuGet packages if prompted.
4. Build the project.

Alternatively, from the command line using the .NET SDK:
```bash
dotnet build
```

## How to Use

1. **Open a command prompt or terminal** where the compiled `SlideMasterRemover.exe` (or equivalent) is located.
2. **Run the executable** with the full path to your PowerPoint file as an argument. For example:

   ```bash
   SlideMasterRemover.exe "C:\Path\To\Presentation.pptx"
   ```

3. If the presentation file has an extension other than `.pptx`, `.pptm`, or `.ppt`, the application will prompt you for confirmation before proceeding.
4. If successful, the program will **report which layouts are used**, **delete those that are unused**, and **remove any completely unused designs**.  
5. The PowerPoint file will be saved automatically after cleanup.

## Usage Notes

- **Backup your original file** before running this tool, just to be safe.
- The **presentation is saved in-place** by default (the same file path is used).
- **Microsoft PowerPoint must be installed** on the machine running this program because the tool relies on Microsoft Office Interop.

## Contributing

Contributions are welcome! If you encounter any issues or have suggestions for improvements:

- Create a [Pull Request](../../pulls) if you have changes or enhancements to contribute.

## License

This project is released under the [MIT License](LICENSE).  

Feel free to modify and distribute in accordance with the license terms.  
