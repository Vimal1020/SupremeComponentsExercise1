# ConvertCell Excel Add-In

## 😀 Overview

**ConvertCell** is an Excel add-in built with C# and the Office Ribbon XML framework. It provides two primary functions:

1. **Convert to Alphanumeric**: Strips any non-alphanumeric characters from the selected cells, storing original values for possible restoration.
2. **Revert to Original**: Restores any converted cells back to their original values.

This tool is ideal for users who need to clean up data by removing symbols, spaces, or special characters while preserving the ability to revert changes.

## 🛠️ Prerequisites

* **Development Environment**: Visual Studio 2017 or later with **Office Developer Tools** installed.
* **.NET Framework**: Target framework **.NET Framework 4.7.2** (or higher)
* **Excel Version**: Microsoft Excel 2010 or later
* **NuGet Packages**: Microsoft.Office.Interop.Excel

## 📥 Installation & Setup

1. **Clone or Download the Repository**

   ```bash
   git clone https://github.com/your-org/ConvertCellAddin.git
   ```

2. **Open the Solution**

   * Launch Visual Studio and open `ConvertCellAddin.sln`.

3. **Embed the Ribbon XML**

   * Ensure `ConvertCell.xml` is added under the project with **Build Action** set to `Embedded Resource`.
   * Confirm the resource name matches `SupremeComponentsExercise1.ConvertCell.xml` (check `GetResourceText`).

4. **Wire up the Ribbon**

   * In `ThisAddIn.cs` (or `ThisWorkbook.cs` / `ThisDocument.cs`), override the Ribbon creation method:

     ```csharp
     protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
     {
         return new ConvertCell();
     }
     ```

5. **Build & Debug**

   * Set the add-in project as the startup project.
   * Press **F5** to build and launch Excel with the add-in loaded.

## 📖 Usage

1. In Excel, navigate to the **Convert Cell** tab on the Ribbon.
2. **Convert to Alphanumeric**: Select a range of cells and click the button—non-alphanumeric characters will be removed, and the status bar will indicate how many cells were processed.
3. **Revert to Original**: Click to restore cells to their original values.

> **Note**: The add-in limits processing to **10,000 cells** at a time (configurable via `MAX_CELLS`).

## 🧩 File Structure

```
ConvertCellAddin/
├─ ConvertCell.cs          # Main Ribbon class and logic
├─ ConvertCell.xml         # Ribbon XML definition (embedded resource)
├─ ThisAddIn.cs            # VSTO startup, override Ribbon loader
├─ app.config              # Configuration (e.g., versioning)
└─ README.md               # This documentation
```

## 🤝 Contributing

Contributions, bug reports, and feature requests are welcome! Please follow these steps:

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/YourFeatureName`.
3. Commit your changes and push to your fork.
4. Open a Pull Request describing your changes.

## 📄 License

This project is licensed under the **MIT License**. See [LICENSE](./LICENSE) for details.
