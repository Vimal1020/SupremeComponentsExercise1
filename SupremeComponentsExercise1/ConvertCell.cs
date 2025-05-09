using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ConvertCell();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

namespace SupremeComponentsExercise1
{
    [ComVisible(true)]
    public class ConvertCell : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private Dictionary<string, string> originalValues = new Dictionary<string, string>();
        private const int MAX_CELLS = 10000;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SupremeComponentsExercise1.ConvertCell.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        /// <summary>
        /// Converts selected cell values to alphanumeric only
        /// </summary>
        /// <param name="control">Ribbon control that triggered the event</param>
        public void ConvertToAlphanumeric_Click(Office.IRibbonControl control)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            app.ScreenUpdating = false;

            try
            {
                Excel.Range selection = app.Selection as Excel.Range;

                // Check if anything is selected
                if (selection == null)
                {
                    MessageBox.Show("Please select a range of cells first.", "No Selection",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check if selection is too large - handle potential exception in one place
                try
                {
                    int cellCount = selection.Cells.Count;
                    if (cellCount > MAX_CELLS)
                    {
                        ShowTooManyCellsMessage(app);
                        return;
                    }
                }
                catch (Exception)
                {
                    // This catch handles OutOfMemoryException or other exceptions when counting cells
                    ShowTooManyCellsMessage(app);
                    return;
                }

                // Clear previous dictionary to avoid memory leaks
                originalValues.Clear();

                // Process each cell in the selection
                foreach (Excel.Range cell in selection.Cells)
                {
                    if (cell.Value != null)
                    {
                        string cellAddress = GetCellAddress(cell);
                        string originalValue = cell.Value.ToString();

                        // Store original value for later restoration
                        originalValues[cellAddress] = originalValue;

                        // Convert to alphanumeric only
                        string sanitizedValue = Regex.Replace(originalValue, @"[^a-zA-Z0-9]", "");
                        cell.Value = sanitizedValue;
                    }
                }

                app.StatusBar = $"Converted {originalValues.Count} cells to alphanumeric only.";
            }
            catch (Exception ex)
            {
                app.StatusBar = "Error occurred during conversion.";
                MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// Reverts cells back to their original values before alphanumeric conversion
        /// </summary>
        /// <param name="control">Ribbon control that triggered the event</param>
        public void RevertToOriginal_Click(Office.IRibbonControl control)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            app.ScreenUpdating = false;

            try
            {
                // Check if we have any original values to restore
                if (originalValues.Count == 0)
                {
                    app.StatusBar = "No previous conversion to revert.";
                    MessageBox.Show("There are no previous conversions to revert.",
                        "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int restoredCount = 0;

                // Restore each cell to its original value
                foreach (var entry in originalValues)
                {
                    Excel.Range cell = GetCellFromAddress(app.ActiveSheet, entry.Key);
                    if (cell != null)
                    {
                        cell.Value = entry.Value;
                        restoredCount++;
                    }
                }

                app.StatusBar = $"Restored {restoredCount} cells to their original values.";

                // Clear the dictionary after restoration
                originalValues.Clear();
            }
            catch (Exception ex)
            {
                app.StatusBar = "Error occurred during reversion.";
                MessageBox.Show($"An error occurred: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        #region Helper Methods

        /// <summary>
        /// Shows a message about too many cells being selected
        /// </summary>
        /// <param name="app">Excel application instance</param>
        private void ShowTooManyCellsMessage(Excel.Application app)
        {
            string message = $"Too many cells selected. Please select fewer than {MAX_CELLS} cells.";
            app.StatusBar = message;
            MessageBox.Show(message, "Selection Too Large", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Gets the address of a cell in A1 notation with worksheet name
        /// </summary>
        /// <param name="cell">The Excel cell</param>
        /// <returns>String representation of cell address</returns>
        private string GetCellAddress(Excel.Range cell)
        {
            return $"{cell.Worksheet.Name}!{cell.Address}";
        }

        /// <summary>
        /// Gets a cell from its address string
        /// </summary>
        /// <param name="sheet">The active worksheet</param>
        /// <param name="address">The cell address string</param>
        /// <returns>Excel Range object representing the cell</returns>
        private Excel.Range GetCellFromAddress(Excel.Worksheet sheet, string address)
        {
            try
            {
                // Split address to get worksheet name and cell reference
                string[] parts = address.Split('!');
                if (parts.Length != 2)
                    return null;

                string worksheetName = parts[0];
                string cellReference = parts[1];

                // Get the worksheet
                Excel.Worksheet worksheet = sheet.Parent.Worksheets[worksheetName];
                if (worksheet == null)
                    return null;

                // Return the cell
                return worksheet.Range[cellReference];
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the resource text from assembly
        /// </summary>
        /// <param name="resourceName">Name of the resource</param>
        /// <returns>Content of the resource</returns>
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            foreach (string name in resourceNames)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }

            return null;
        }

        #endregion
    }
}