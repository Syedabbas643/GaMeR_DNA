using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Label = System.Windows.Forms.Label;
using CheckBox = System.Windows.Forms.CheckBox;
using Button = System.Windows.Forms.Button;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace GaMeR
{
    [ComVisible(true)]
    public class CustomRibbon : ExcelRibbon
    {
        private Formfind formfind;
        private Form1 form1;
        public override string GetCustomUI(string RibbonID)
        {
            return @"
              <customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnLoad"">
          <ribbon>
            <tabs>
              <tab id=""tab2"" label=""GaMeR2"">
                <group id=""group2"" label=""Automate"">
                  <button id=""data1"" label=""Create New Costing"" getImage=""GetCustomImage"" size=""large"" onAction=""OnDataClick""/>
                  <button id=""server"" label=""database folder"" getImage=""GetCustomImage"" onAction=""OndatabasefolderClick""/>
                  <separator id=""separator1""/>
                  <button id=""find"" label=""Find in Database"" getImage=""GetCustomImage"" size=""large"" onAction=""OnfindClick""/>
                  <separator id=""separator3""/>
                  <button id=""bom"" label=""Make Bill of OLD Materials"" getImage=""GetCustomImage"" size=""large"" onAction=""OnbomClick""/>
                  <button id=""bom2"" label=""Make Bill of NEW Materials"" getImage=""GetCustomImage"" size=""large"" onAction=""OnbomnewClick""/>
                  <button id=""cad"" label=""Analyse Costing"" getImage=""GetCustomImage"" size=""large"" onAction=""OnanalyseClick""/>
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
        }
        public Bitmap GetCustomImage(IRibbonControl control)
        {
            // Adjust the image loading mechanism according to your actual structure
            string imageName = control.Id; // Assuming image names match button IDs
            string resourceName = $"GaMeR.Images.{imageName}.png"; // Adjust namespace and file extension

            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    return new Bitmap(stream);
                }
                else
                {
                    MessageBox.Show($"Resource {resourceName} not found.");
                    return null; // Or return a default image if not found
                }
            }
        }
        private string GetResourceText(string resourceName)
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return null;
                using (var reader = new System.IO.StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void OnbomClick(IRibbonControl control)
        {
            try
            {
                var excelApp = ExcelDnaUtil.Application as Excel.Application;
                Excel.Worksheet currentSheet = excelApp.ActiveSheet;
                Excel.Workbook currentWorkbook = excelApp.ActiveWorkbook;

                string savedPath = GetDatabaseFilePath();

                if (string.IsNullOrEmpty(savedPath))
                {
                    MessageBox.Show("No folder path selected. Please select a folder first.");
                    return;
                }

                string extFilePath = System.IO.Path.Combine(savedPath, "templates.xlsx");
                string extFilePath2 = System.IO.Path.Combine(savedPath, "bom_database.xlsx");

                Excel.Workbook extWorkbook2 = excelApp.Workbooks.Open(
                    extFilePath2,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Editable: false,
                    IgnoreReadOnlyRecommended: true
                );

                Dictionary<string, string> catalogNumberToSheetMap = new Dictionary<string, string>();

                foreach (Excel.Worksheet sheet in extWorkbook2.Sheets)
                {
                    Excel.Range usedRange1 = sheet.UsedRange;
                    Excel.Range columnB1 = usedRange1.Columns["B"];

                    foreach (Excel.Range cell in columnB1.Cells)
                    {
                        if (cell.Value2 != null)
                        {
                            string catalogNumber = cell.Value2.ToString();
                            if (!catalogNumberToSheetMap.ContainsKey(catalogNumber))
                            {
                                catalogNumberToSheetMap[catalogNumber] = sheet.Name;
                            }
                        }
                    }
                }

                extWorkbook2.Close(false);

                Excel.Workbook extWorkbook = excelApp.Workbooks.Open(
                    extFilePath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Editable: false,
                    IgnoreReadOnlyRecommended: true
                );

                Excel.Worksheet tempSheet = null;
                foreach (Excel.Worksheet sheet in extWorkbook.Sheets)
                {
                    if (sheet.Name == "BOM")
                    {
                        tempSheet = sheet;
                        break;
                    }
                }

                if (tempSheet == null)
                {
                    MessageBox.Show("The 'temp' sheet was not found in the template_bom.xlsx file.");
                    extWorkbook.Close(false);
                    return;
                }

                tempSheet.Copy(After: currentWorkbook.Sheets[currentWorkbook.Sheets.Count]);

                int sheetNumber = 1;
                string newSheetName;
                bool sheetNameExists;

                do
                {
                    newSheetName = "Sheet" + sheetNumber.ToString();
                    sheetNameExists = false;

                    foreach (Excel.Worksheet sheet in currentWorkbook.Sheets)
                    {
                        if (sheet.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            sheetNameExists = true;
                            sheetNumber++;
                            break;
                        }
                    }
                } while (sheetNameExists);

                Excel.Worksheet newSheet = currentWorkbook.Sheets[currentWorkbook.Sheets.Count];
                newSheet.Name = newSheetName;

                extWorkbook.Close(false);

                currentSheet.Activate();
                Excel.Range usedRange = currentSheet.UsedRange;
                Excel.Range columnB = usedRange.Columns["B"];
                int insertingColumn = 6;
                int countColumn = 6;
                Dictionary<string, int> productNextRow = new Dictionary<string, int>
                        {
                            { "ACB", 8 },
                            { "MCCB", 10 },
                            { "METER", 12 },
                            { "MCB", 14 },
                            { "LAMP", 16 },
                            { "REA-CAP", 18 },
                            { "CONTACTOR", 20 },
                            { "CT", 22 },
                            { "NULL", 31 },
                            { "MAKE", 34 }
                        };

                List<Excel.Range> headings = new List<Excel.Range>();
                foreach (Excel.Range cell in columnB.Cells)
                {
                    if (cell.Interior.Color == 15773696)
                    {
                        headings.Add(cell);
                    }
                }

                for (int i = 0; i < headings.Count; i++)
                {
                    Excel.Range heading = headings[i];
                    Excel.Range nextHeading = (i < headings.Count - 1) ? headings[i + 1] : null;

                    Excel.Range dataRange;
                    if (nextHeading != null)
                    {
                        dataRange = currentSheet.Range[heading.Offset[1, 0], nextHeading.Offset[-1, 0]];
                    }
                    else
                    {
                        dataRange = currentSheet.Range[heading.Offset[1, 0], currentSheet.Cells[currentSheet.Rows.Count, "B"].End(Excel.XlDirection.xlUp)];
                    }

                    string cellValue = heading.Value2?.ToString();
                    newSheet.Activate();

                    Excel.Range insertColumn = newSheet.Columns[insertingColumn];
                    insertingColumn++;

                    insertColumn.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    Excel.Range pasteRange = newSheet.Cells[4, insertColumn.Column - 1];
                    pasteRange.Value = cellValue;
                    pasteRange.WrapText = true;
                    pasteRange.Columns.AutoFit();

                    Excel.Range columnCDataRange = dataRange.Resize[dataRange.Rows.Count, 1].Offset[0, 1];
                    Excel.Range columnDDataRange = dataRange.Resize[dataRange.Rows.Count, 1].Offset[0, 2];

                    foreach (Excel.Range cell in columnCDataRange)
                    {
                        string catNumber = cell.Value2?.ToString();
                        if (catNumber != null && catalogNumberToSheetMap.TryGetValue(catNumber, out string productName))
                        {
                            Excel.Range dataRow = currentSheet.Rows[cell.Row];
                            string columnAValue = currentSheet.Cells[cell.Row, "A"].Value2?.ToString();
                            string columnBValue = currentSheet.Cells[cell.Row, "B"].Value2?.ToString();
                            string columnCValue = currentSheet.Cells[cell.Row, "C"].Value2?.ToString();
                            string columnDValue = currentSheet.Cells[cell.Row, "D"].Value2?.ToString();
                            

                            double columnEValue = Convert.ToDouble(currentSheet.Cells[cell.Row, "E"].Value2);
                            double columnFValue = Convert.ToDouble(currentSheet.Cells[cell.Row, "F"].Value2);
                            double product = columnEValue * columnFValue;

                            if (productNextRow.TryGetValue(productName, out int targetRow))
                            {
                                bool matchFound = false;

                                for (int row = 1; row <= newSheet.UsedRange.Rows.Count; row++)
                                {
                                    if (newSheet.Cells[row, "C"].Value2?.ToString() == columnCValue)
                                    {
                                        double currentEValue = Convert.ToDouble(newSheet.Cells[row, countColumn + i].Value2);
                                        newSheet.Cells[row, countColumn + i].Value2 = currentEValue + product;
                                        
                                        matchFound = true;
                                        break;
                                    }
                                }

                                if (!matchFound)
                                {
                                    newSheet.Rows[targetRow].Insert();

                                    newSheet.Cells[targetRow, "A"].Value2 = columnAValue;
                                    newSheet.Cells[targetRow, "B"].Value2 = columnBValue;
                                    newSheet.Cells[targetRow, "C"].Value2 = columnCValue;
                                    newSheet.Cells[targetRow, "D"].Value2 = columnDValue;
                                    newSheet.Cells[targetRow, countColumn + i].Value2 = product;

                                    productNextRow[productName] = targetRow + 1;

                                    bool startUpdating = false;
                                    foreach (var key in productNextRow.Keys.ToList())
                                    {
                                        if (startUpdating)
                                        {
                                            productNextRow[key]++;
                                        }
                                        if (key == productName)
                                        {
                                            startUpdating = true;
                                        }
                                    }
                                }
                            }
                        }
                        else {
                            
                            Excel.Range dataRow = currentSheet.Rows[cell.Row];
                            string columnAValue = currentSheet.Cells[cell.Row, "A"].Value2?.ToString();
                            string columnBValue = currentSheet.Cells[cell.Row, "B"].Value2?.ToString();
                            Excel.Range columnBCell = currentSheet.Cells[cell.Row, "B"];

                            // The color code for Excel's "Orange" color
                            int orangeColorCode = 49407;
                            if (columnBValue == "BUSBAR FABRICATION COST" ||  columnBValue == "CONSUMABLES" || columnBValue== "LABOUR WIRING"|| columnBCell.Interior.Color == orangeColorCode)
                            {
                                continue;
                            }
                            string columnCValue = currentSheet.Cells[cell.Row, "C"].Value2?.ToString();
                            string columnDValue = currentSheet.Cells[cell.Row, "D"].Value2?.ToString();
                            if (columnDValue == "ALUMINIUM")
                            {
                                continue;
                            }

                            double columnEValue = Convert.ToDouble(currentSheet.Cells[cell.Row, "E"].Value2);
                            double columnFValue = Convert.ToDouble(currentSheet.Cells[cell.Row, "F"].Value2);
                            double product = columnEValue * columnFValue;
                            if (productNextRow.TryGetValue("NULL", out int targetRow))
                            {
                                bool matchFound = false;

                                for (int row = 1; row <= newSheet.UsedRange.Rows.Count; row++)
                                {
                                    if (newSheet.Cells[row, "B"].Value2?.ToString() == columnBValue)
                                    {
                                        double currentEValue = Convert.ToDouble(newSheet.Cells[row, countColumn + i].Value2);
                                        newSheet.Cells[row, countColumn + i].Value2 = currentEValue + product;
                                        matchFound = true;
                                        break;
                                    }
                                }

                                if (!matchFound)
                                {
                                    newSheet.Rows[targetRow].Insert();

                                    newSheet.Cells[targetRow, "A"].Value2 = columnAValue;
                                    newSheet.Cells[targetRow, "B"].Value2 = columnBValue;
                                    newSheet.Cells[targetRow, "C"].Value2 = columnCValue;
                                    newSheet.Cells[targetRow, "D"].Value2 = columnDValue;
                                    newSheet.Cells[targetRow, countColumn + i].Value2 = product;

                                    productNextRow["NULL"] = targetRow + 1;

                                    bool startUpdating = false;
                                    foreach (var key in productNextRow.Keys.ToList())
                                    {
                                        if (startUpdating)
                                        {
                                            productNextRow[key]++;
                                        }
                                        if (key == "NULL")
                                        {
                                            startUpdating = true;
                                        }
                                    }
                                }
                             }

                            }
                    }

                    
                    }
                

                newSheet.Activate();
                Excel.Range usedRangeNew = newSheet.UsedRange;
                Excel.Range columnBNew = usedRangeNew.Columns["B"];
                int lastColumn = usedRangeNew.Columns.Count - 3;
                string lastColumnLetter = GetExcelColumnName(lastColumn - 1);

                for (int row = 7; row <= usedRangeNew.Rows.Count; row++)
                {
                    Excel.Range formulaCell = newSheet.Cells[row, lastColumn];
                    string formula = $"=SUM(F{row}:{lastColumnLetter}{row})";
                    formulaCell.Formula = formula;
                }

                foreach (Excel.Range cell in columnBNew.Cells)
                {
                    if (cell.Interior.Color == 15773696)
                    {
                        int rowToDelete = cell.Row + 1;

                        if (rowToDelete <= usedRangeNew.Rows.Count)
                        {
                            Excel.Range rowToDeleteRange = newSheet.Rows[rowToDelete];
                            rowToDeleteRange.Delete();
                        }

                        newSheet.Cells[cell.Row, lastColumn] = "";
                        Excel.Range targetCell = newSheet.Cells[cell.Row, lastColumn];
                        targetCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    }
                }

                for (int col = 6; col <= lastColumn; col++)
                {
                    Excel.Range columnRange = usedRangeNew.Columns[col];

                    // Set font to bold
                    columnRange.Font.Bold = true;

                    // Set text to center
                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        public void OnbomnewClick(IRibbonControl control)
        {
            try
            {
                var excelApp = ExcelDnaUtil.Application as Excel.Application;
                Excel.Worksheet currentSheet = excelApp.ActiveSheet;
                Excel.Workbook currentWorkbook = excelApp.ActiveWorkbook;

                string savedPath = GetDatabaseFilePath();

                if (string.IsNullOrEmpty(savedPath))
                {
                    MessageBox.Show("No folder path selected. Please select a folder first.");
                    return;
                }

                string extFilePath = System.IO.Path.Combine(savedPath, "templates.xlsx");
                string extFilePath2 = System.IO.Path.Combine(savedPath, "bom_database.xlsx");

                Excel.Workbook extWorkbook2 = excelApp.Workbooks.Open(
                    extFilePath2,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Editable: false,
                    IgnoreReadOnlyRecommended: true
                );

                Dictionary<string, string> catalogNumberToSheetMap = new Dictionary<string, string>();

                foreach (Excel.Worksheet sheet in extWorkbook2.Sheets)
                {
                    Excel.Range usedRange1 = sheet.UsedRange;
                    Excel.Range columnB1 = usedRange1.Columns["B"];

                    foreach (Excel.Range cell in columnB1.Cells)
                    {
                        if (cell.Value2 != null)
                        {
                            string catalogNumber = cell.Value2.ToString();
                            if (!catalogNumberToSheetMap.ContainsKey(catalogNumber))
                            {
                                catalogNumberToSheetMap[catalogNumber] = sheet.Name;
                            }
                        }
                    }
                }

                extWorkbook2.Close(false);

                Excel.Workbook extWorkbook = excelApp.Workbooks.Open(
                    extFilePath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Editable: false,
                    IgnoreReadOnlyRecommended: true
                );

                Excel.Worksheet tempSheet = null;
                foreach (Excel.Worksheet sheet in extWorkbook.Sheets)
                {
                    if (sheet.Name == "BOM")
                    {
                        tempSheet = sheet;
                        break;
                    }
                }

                if (tempSheet == null)
                {
                    MessageBox.Show("The 'temp' sheet was not found in the template_bom.xlsx file.");
                    extWorkbook.Close(false);
                    return;
                }

                tempSheet.Copy(After: currentWorkbook.Sheets[currentWorkbook.Sheets.Count]);

                int sheetNumber = 1;
                string newSheetName;
                bool sheetNameExists;

                do
                {
                    newSheetName = "Sheet" + sheetNumber.ToString();
                    sheetNameExists = false;

                    foreach (Excel.Worksheet sheet in currentWorkbook.Sheets)
                    {
                        if (sheet.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            sheetNameExists = true;
                            sheetNumber++;
                            break;
                        }
                    }
                } while (sheetNameExists);

                Excel.Worksheet newSheet = currentWorkbook.Sheets[currentWorkbook.Sheets.Count];
                newSheet.Name = newSheetName;

                extWorkbook.Close(false);

                currentSheet.Activate();
                Excel.Range usedRange = currentSheet.UsedRange;
                Excel.Range columnB = usedRange.Columns["B"];
                int insertingColumn = 6;
                int countColumn = 6;
                Dictionary<string, int> productNextRow = new Dictionary<string, int>
                        {
                            { "ACB", 8 },
                            { "MCCB", 10 },
                            { "METER", 12 },
                            { "MCB", 14 },
                            { "LAMP", 16 },
                            { "REA-CAP", 18 },
                            { "CONTACTOR", 20 },
                            { "CT", 22 },
                            { "NULL", 31 }
                        };

                List<Excel.Range> headings = new List<Excel.Range>();
                List<Excel.Range> headcounts = new List<Excel.Range>();
                foreach (Excel.Range cell in columnB.Cells)
                {
                    if (cell.Interior.Color == 15773696)
                    {
                        headings.Add(cell);
                        headcounts.Add(cell.Offset[0,1]);
                    }
                }

                for (int i = 0; i < headings.Count; i++)
                {
                    Excel.Range heading = headings[i];
                    Excel.Range headcount = headcounts[i];
                    Excel.Range nextHeading = (i < headings.Count - 1) ? headings[i + 1] : null;

                    Excel.Range dataRange;
                    if (nextHeading != null)
                    {
                        dataRange = currentSheet.Range[heading.Offset[1, 0], nextHeading.Offset[-1, 0]];
                    }
                    else
                    {
                        dataRange = currentSheet.Range[heading.Offset[1, 0], currentSheet.Cells[currentSheet.Rows.Count, "B"].End(Excel.XlDirection.xlUp)];
                    }

                    string cellValue = heading.Value2?.ToString();
                    string cellcountvalue =headcount.Value2?.ToString();    
                    newSheet.Activate();

                    Excel.Range insertColumn = newSheet.Columns[insertingColumn];
                    insertingColumn++;

                    insertColumn.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    Excel.Range pasteRange = newSheet.Cells[4, insertColumn.Column - 1];
                    pasteRange.Value = cellValue;
                    pasteRange.WrapText = true;
                    pasteRange.Columns.AutoFit();

                    Excel.Range pasteRange2 = newSheet.Cells[5, insertColumn.Column - 1];
                    pasteRange2.Value = cellcountvalue;
                    

                    Excel.Range columnCDataRange = dataRange.Resize[dataRange.Rows.Count, 1].Offset[0, 1];

                    foreach (Excel.Range cell in columnCDataRange)
                    {
                        string catNumber = cell.Value2?.ToString();
                        if (catNumber != null && catalogNumberToSheetMap.TryGetValue(catNumber, out string productName))
                        {
                            Excel.Range dataRow = currentSheet.Rows[cell.Row];
                            string columnAValue = currentSheet.Cells[cell.Row, "A"].Value2?.ToString();
                            string columnBValue = currentSheet.Cells[cell.Row, "B"].Value2?.ToString();
                            string columnCValue = currentSheet.Cells[cell.Row, "C"].Value2?.ToString();
                            string columnDValue = currentSheet.Cells[cell.Row, "D"].Value2?.ToString();

                            //double columnEValue = Convert.ToDouble(currentSheet.Cells[cell.Row, "E"].Value2);
                           //double columnFValue = Convert.ToDouble(currentSheet.Cells[cell.Row, "F"].Value2);
                            double product = Convert.ToDouble(currentSheet.Cells[cell.Row, "L"].Value2);

                            if (productNextRow.TryGetValue(productName, out int targetRow))
                            {
                                bool matchFound = false;

                                for (int row = 1; row <= newSheet.UsedRange.Rows.Count; row++)
                                {
                                    if (newSheet.Cells[row, "C"].Value2?.ToString() == columnCValue)
                                    {
                                        double currentEValue = Convert.ToDouble(newSheet.Cells[row, countColumn + i].Value2);
                                        newSheet.Cells[row, countColumn + i].Value2 = currentEValue + product;
                                        matchFound = true;
                                        break;
                                    }
                                }

                                if (!matchFound)
                                {
                                    newSheet.Rows[targetRow].Insert();

                                    newSheet.Cells[targetRow, "A"].Value2 = columnAValue;
                                    newSheet.Cells[targetRow, "B"].Value2 = columnBValue;
                                    newSheet.Cells[targetRow, "C"].Value2 = columnCValue;
                                    newSheet.Cells[targetRow, "D"].Value2 = columnDValue;
                                    newSheet.Cells[targetRow, countColumn + i].Value2 = product;

                                    productNextRow[productName] = targetRow + 1;

                                    bool startUpdating = false;
                                    foreach (var key in productNextRow.Keys.ToList())
                                    {
                                        if (startUpdating)
                                        {
                                            productNextRow[key]++;
                                        }
                                        if (key == productName)
                                        {
                                            startUpdating = true;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {

                            Excel.Range dataRow = currentSheet.Rows[cell.Row];
                            string columnAValue = currentSheet.Cells[cell.Row, "A"].Value2?.ToString();
                            string columnBValue = currentSheet.Cells[cell.Row, "B"].Value2?.ToString();
                            Excel.Range columnBCell = currentSheet.Cells[cell.Row, "B"];

                            // The color code for Excel's "Orange" color
                            int orangeColorCode = 49407;
                            if (columnBValue == "BUSBAR FABRICATION COST" || columnBValue == "CONSUMABLES" || columnBValue == "LABOUR WIRING" || columnBCell.Interior.Color == orangeColorCode)
                            {
                                continue;
                            }
                            string columnCValue = currentSheet.Cells[cell.Row, "C"].Value2?.ToString();
                            string columnDValue = currentSheet.Cells[cell.Row, "D"].Value2?.ToString();
                            if (columnDValue == "ALUMINIUM")
                            {
                                continue;
                            }

                            //double columnEValue = Convert.ToDouble(currentSheet.Cells[cell.Row, "E"].Value2);
                            //double columnFValue = Convert.ToDouble(currentSheet.Cells[cell.Row, "F"].Value2);
                            double product = Convert.ToDouble(currentSheet.Cells[cell.Row, "L"].Value2);
                            if (productNextRow.TryGetValue("NULL", out int targetRow))
                            {
                                bool matchFound = false;

                                for (int row = 1; row <= newSheet.UsedRange.Rows.Count; row++)
                                {
                                    if (newSheet.Cells[row, "B"].Value2?.ToString() == columnBValue)
                                    {
                                        double currentEValue = Convert.ToDouble(newSheet.Cells[row, countColumn + i].Value2);
                                        newSheet.Cells[row, countColumn + i].Value2 = currentEValue + product;
                                        matchFound = true;
                                        break;
                                    }
                                }

                                if (!matchFound)
                                {
                                    newSheet.Rows[targetRow].Insert();

                                    newSheet.Cells[targetRow, "A"].Value2 = columnAValue;
                                    newSheet.Cells[targetRow, "B"].Value2 = columnBValue;
                                    newSheet.Cells[targetRow, "C"].Value2 = columnCValue;
                                    newSheet.Cells[targetRow, "D"].Value2 = columnDValue;
                                    newSheet.Cells[targetRow, countColumn + i].Value2 = product;

                                    productNextRow["NULL"] = targetRow + 1;

                                    bool startUpdating = false;
                                    foreach (var key in productNextRow.Keys.ToList())
                                    {
                                        if (startUpdating)
                                        {
                                            productNextRow[key]++;
                                        }
                                        if (key == "NULL")
                                        {
                                            startUpdating = true;
                                        }
                                    }
                                }
                            }

                        }
                    }
                }

                newSheet.Activate();
                Excel.Range usedRangeNew = newSheet.UsedRange;
                Excel.Range columnBNew = usedRangeNew.Columns["B"];
                int lastColumn = usedRangeNew.Columns.Count;
                string lastColumnLetter = GetExcelColumnName(lastColumn - 1);

                for (int row = 7; row <= usedRangeNew.Rows.Count; row++)
                {
                    Excel.Range formulaCell = newSheet.Cells[row, lastColumn];
                    string formula = $"=SUM(F{row}:{lastColumnLetter}{row})";
                    formulaCell.Formula = formula;
                }

                foreach (Excel.Range cell in columnBNew.Cells)
                {
                    if (cell.Interior.Color == 15773696)
                    {
                        int rowToDelete = cell.Row + 1;

                        if (rowToDelete <= usedRangeNew.Rows.Count)
                        {
                            Excel.Range rowToDeleteRange = newSheet.Rows[rowToDelete];
                            rowToDeleteRange.Delete();
                        }

                        newSheet.Cells[cell.Row, lastColumn] = "";
                        Excel.Range targetCell = newSheet.Cells[cell.Row, lastColumn];
                        targetCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    }
                }

                for (int col = 6; col <= lastColumn; col++)
                {
                    Excel.Range columnRange = usedRangeNew.Columns[col];

                    // Set font to bold
                    columnRange.Font.Bold = true;

                    // Set text to center
                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        public void OnanalyseClick(IRibbonControl control)
        {
            try
            {
                var excelApp = ExcelDnaUtil.Application as Excel.Application;
                Excel.Worksheet currentSheet = excelApp.ActiveSheet;
                Excel.Workbook currentWorkbook = excelApp.ActiveWorkbook;

                string savedPath = GetDatabaseFilePath();

                if (string.IsNullOrEmpty(savedPath))
                {
                    MessageBox.Show("No folder path selected. Please select a folder first.");
                    return;
                }

                string extFilePath = System.IO.Path.Combine(savedPath, "templates.xlsx");
                string extFilePath2 = System.IO.Path.Combine(savedPath, "bom_database.xlsx");

                Excel.Workbook extWorkbook2 = excelApp.Workbooks.Open(
                    extFilePath2,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Editable: false,
                    IgnoreReadOnlyRecommended: true
                );

                Dictionary<string, string> catalogNumberToSheetMap = new Dictionary<string, string>();

                foreach (Excel.Worksheet sheet in extWorkbook2.Sheets)
                {
                    Excel.Range usedRange1 = sheet.UsedRange;
                    Excel.Range columnB1 = usedRange1.Columns["B"];

                    foreach (Excel.Range cell in columnB1.Cells)
                    {
                        if (cell.Value2 != null)
                        {
                            string catalogNumber = cell.Value2.ToString();
                            if (!catalogNumberToSheetMap.ContainsKey(catalogNumber))
                            {
                                catalogNumberToSheetMap[catalogNumber] = sheet.Name;
                            }
                        }
                    }
                }

                extWorkbook2.Close(false);

                Excel.Workbook extWorkbook = excelApp.Workbooks.Open(
                    extFilePath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Editable: false,
                    IgnoreReadOnlyRecommended: true
                );

                Excel.Worksheet tempSheet = null;
                foreach (Excel.Worksheet sheet in extWorkbook.Sheets)
                {
                    if (sheet.Name == "ANALYSE")
                    {
                        tempSheet = sheet;
                        break;
                    }
                }

                if (tempSheet == null)
                {
                    MessageBox.Show("The 'temp' sheet was not found in the template_bom.xlsx file.");
                    extWorkbook.Close(false);
                    return;
                }

                tempSheet.Copy(After: currentWorkbook.Sheets[currentWorkbook.Sheets.Count]);

                int sheetNumber = 1;
                string newSheetName;
                bool sheetNameExists;

                do
                {
                    newSheetName = "Sheet" + sheetNumber.ToString();
                    sheetNameExists = false;

                    foreach (Excel.Worksheet sheet in currentWorkbook.Sheets)
                    {
                        if (sheet.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            sheetNameExists = true;
                            sheetNumber++;
                            break;
                        }
                    }
                } while (sheetNameExists);

                Excel.Worksheet newSheet = currentWorkbook.Sheets[currentWorkbook.Sheets.Count];
                newSheet.Name = newSheetName;

                extWorkbook.Close(false);

                currentSheet.Activate();
                Excel.Range usedRange = currentSheet.UsedRange;
                Excel.Range columnB = usedRange.Columns["B"];
                int insertingColumn = 6;
                int countColumn = 6;
                Dictionary<string, int> productNextRow = new Dictionary<string, int>
                        {
                            { "ACB", 60 },
                            { "MCCB", 62 },
                            { "METER", 64 },
                            { "MCB", 66 },
                            { "LAMP", 68 },
                            { "REA-CAP", 70 },
                            { "CONTACTOR", 72 },
                            { "CT", 74 },
                            { "NULL", 83 }
                        };

                List<Excel.Range> headings = new List<Excel.Range>();
                List<Excel.Range> headcounts = new List<Excel.Range>();
                foreach (Excel.Range cell in columnB.Cells)
                {
                    if (cell.Interior.Color == 15773696)
                    {
                        headings.Add(cell);
                        headcounts.Add(cell.Offset[0, 1]);
                    }
                }

                for (int i = 0; i < headings.Count; i++)
                {
                    Excel.Range heading = headings[i];
                    Excel.Range headcount = headcounts[i];
                    Excel.Range nextHeading = (i < headings.Count - 1) ? headings[i + 1] : null;

                    Excel.Range dataRange;
                    if (nextHeading != null)
                    {
                        dataRange = currentSheet.Range[heading.Offset[1, 0], nextHeading.Offset[-1, 0]];
                    }
                    else
                    {
                        dataRange = currentSheet.Range[heading.Offset[1, 0], currentSheet.Cells[currentSheet.Rows.Count, "B"].End(Excel.XlDirection.xlUp)];
                    }

                    string cellValue = heading.Value2?.ToString();
                    string cellcountvalue = headcount.Value2?.ToString();
                    newSheet.Activate();

                    Excel.Range insertColumn = newSheet.Columns[insertingColumn];
                    insertingColumn++;

                    insertColumn.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    Excel.Range pasteRange = newSheet.Cells[56, insertColumn.Column - 1];
                    pasteRange.Value = cellValue;
                    pasteRange.WrapText = true;
                    pasteRange.Columns.AutoFit();

                    Excel.Range pasteRange2 = newSheet.Cells[57, insertColumn.Column - 1];
                    pasteRange2.Value = cellcountvalue;


                    Excel.Range columnCDataRange = dataRange.Resize[dataRange.Rows.Count, 1].Offset[0, 1];

                    foreach (Excel.Range cell in columnCDataRange)
                    {
                        string catNumber = cell.Value2?.ToString();
                        if (catNumber != null && catalogNumberToSheetMap.TryGetValue(catNumber, out string productName))
                        {
                            Excel.Range dataRow = currentSheet.Rows[cell.Row];
                            string columnAValue = currentSheet.Cells[cell.Row, "A"].Value2?.ToString();
                            string columnBValue = currentSheet.Cells[cell.Row, "B"].Value2?.ToString();
                            string columnCValue = currentSheet.Cells[cell.Row, "C"].Value2?.ToString();
                            string columnDValue = currentSheet.Cells[cell.Row, "D"].Value2?.ToString();
                            string make = currentSheet.Cells[cell.Row, "D"].Value2?.ToString();
                            string disc = currentSheet.Cells[cell.Row, "G"].Formula;

                            //if (make == "ALUMINIUM" || make == null || make == "" || make == "KCIPL")
                            //{
                                //continue;

                            //}
                            //else if (make.Contains("L&T"))
                            //{
                                //make = "L&T";
                            //}

                            double price = Convert.ToDouble(currentSheet.Cells[cell.Row, "F"].Value2);
                            double product = Convert.ToDouble(currentSheet.Cells[cell.Row, "L"].Value2);

                            if (productNextRow.TryGetValue(productName, out int targetRow))
                            {
                                bool matchFound = false;

                                for (int row = 1; row <= newSheet.UsedRange.Rows.Count; row++)
                                {
                                    if (newSheet.Cells[row, "C"].Value2?.ToString() == columnCValue)
                                    {
                                        double currentEValue = Convert.ToDouble(newSheet.Cells[row, countColumn + i].Value2);
                                        newSheet.Cells[row, countColumn + i].Value2 = currentEValue + product;
                                        matchFound = true;
                                        break;
                                    }
                                }

                                if (!matchFound)
                                {
                                    newSheet.Rows[targetRow].Insert();

                                    newSheet.Cells[targetRow, "A"].Value2 = columnAValue;
                                    newSheet.Cells[targetRow, "B"].Value2 = columnBValue;
                                    newSheet.Cells[targetRow, "C"].Value2 = columnCValue;
                                    newSheet.Cells[targetRow, "D"].Value2 = columnDValue;
                                    newSheet.Cells[targetRow, countColumn + i].Value2 = product;
                                    newSheet.Cells[targetRow, countColumn + i + 2].Value2 = price.ToString();
                                    newSheet.Cells[targetRow, countColumn + i + 4].Formula = disc;
                                    newSheet.Cells[targetRow, countColumn + i + 6].Value2 = make;

                                    productNextRow[productName] = targetRow + 1;

                                    bool startUpdating = false;
                                    foreach (var key in productNextRow.Keys.ToList())
                                    {
                                        if (startUpdating)
                                        {
                                            productNextRow[key]++;
                                        }
                                        if (key == productName)
                                        {
                                            startUpdating = true;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {

                            Excel.Range dataRow = currentSheet.Rows[cell.Row];
                            string columnAValue = currentSheet.Cells[cell.Row, "A"].Value2?.ToString();
                            string columnBValue = currentSheet.Cells[cell.Row, "B"].Value2?.ToString();
                            Excel.Range columnBCell = currentSheet.Cells[cell.Row, "B"];

                            // The color code for Excel's "Orange" color
                            int orangeColorCode = 49407;
                            if (columnBValue == "BUSBAR FABRICATION COST" || columnBValue == "CONSUMABLES" || columnBValue == "LABOUR WIRING" || columnBCell.Interior.Color == orangeColorCode)
                            {
                                continue;
                            }
                            string columnCValue = currentSheet.Cells[cell.Row, "C"].Value2?.ToString();
                            string columnDValue = currentSheet.Cells[cell.Row, "D"].Value2?.ToString();
                            if (columnDValue == "ALUMINIUM")
                            {
                                continue;
                            }

                            double product = Convert.ToDouble(currentSheet.Cells[cell.Row, "L"].Value2);
                            if (productNextRow.TryGetValue("NULL", out int targetRow))
                            {
                                bool matchFound = false;

                                for (int row = 1; row <= newSheet.UsedRange.Rows.Count; row++)
                                {
                                    if (newSheet.Cells[row, "B"].Value2?.ToString() == columnBValue)
                                    {
                                        double currentEValue = Convert.ToDouble(newSheet.Cells[row, countColumn + i].Value2);
                                        newSheet.Cells[row, countColumn + i].Value2 = currentEValue + product;
                                        matchFound = true;
                                        break;
                                    }
                                }

                                if (!matchFound)
                                {
                                    newSheet.Rows[targetRow].Insert();

                                    newSheet.Cells[targetRow, "A"].Value2 = columnAValue;
                                    newSheet.Cells[targetRow, "B"].Value2 = columnBValue;
                                    newSheet.Cells[targetRow, "C"].Value2 = columnCValue;
                                    newSheet.Cells[targetRow, "D"].Value2 = columnDValue;
                                    newSheet.Cells[targetRow, countColumn + i].Value2 = product;

                                    productNextRow["NULL"] = targetRow + 1;

                                    bool startUpdating = false;
                                    foreach (var key in productNextRow.Keys.ToList())
                                    {
                                        if (startUpdating)
                                        {
                                            productNextRow[key]++;
                                        }
                                        if (key == "NULL")
                                        {
                                            startUpdating = true;
                                        }
                                    }
                                }
                            }

                        }
                    }
                }

                newSheet.Activate();
                Excel.Range usedRangeNew = newSheet.UsedRange;
                Excel.Range columnBNew = usedRangeNew.Columns["B"];
                int lastColumn = usedRangeNew.Columns.Count - 5;
                string lastColumnLetter = GetExcelColumnName(lastColumn - 1);

                int disccolumn = usedRangeNew.Columns.Count - 2;
                string discLetter = GetExcelColumnName(disccolumn);
                Excel.Range columndisc = usedRangeNew.Columns[discLetter];
                columndisc.NumberFormat = "0%";
                for (int row = 59; row <= usedRangeNew.Rows.Count; row++)
                {
                    Excel.Range checkcell = newSheet.Cells[row, disccolumn];
                    if (checkcell != null && checkcell.Value != null)
                    {
                        Excel.Range formulaCell = newSheet.Cells[row, lastColumn + 2];
                        Excel.Range formulaCell2 = newSheet.Cells[row, lastColumn + 4];
                        string priceColumnLetter = GetExcelColumnName(lastColumn + 1);
                        string formula = $"={priceColumnLetter}{row.ToString()}*{GetExcelColumnName(lastColumn)}{row.ToString()}";
                        string formula2 = $"={GetExcelColumnName(lastColumn + 2)}{row}-{GetExcelColumnName(lastColumn + 2)}{row}*{discLetter}{row}";
                        formulaCell.Formula = formula;
                        formulaCell2.Formula = formula2;
                    }

                }

                for (int row = 59; row <= usedRangeNew.Rows.Count; row++)
                {
                    Excel.Range formulaCell = newSheet.Cells[row, lastColumn];
                    string formula = $"=SUM(F{row}:{lastColumnLetter}{row})";
                    formulaCell.Formula = formula;
                }

                foreach (Excel.Range cell in columnBNew.Cells)
                {
                    if (cell.Interior.Color == 15773696)
                    {
                        int rowToDelete = cell.Row + 1;

                        if (rowToDelete <= usedRangeNew.Rows.Count)
                        {
                            Excel.Range rowToDeleteRange = newSheet.Rows[rowToDelete];
                            rowToDeleteRange.Delete();
                        }

                        newSheet.Cells[cell.Row, lastColumn] = "";
                        Excel.Range targetCell = newSheet.Cells[cell.Row, lastColumn];
                        targetCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    }
                }

                for (int col = 6; col <= lastColumn; col++)
                {
                    Excel.Range columnRange = usedRangeNew.Columns[col];

                    // Set font to bold
                    columnRange.Font.Bold = true;

                    // Set text to center
                    columnRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
        public void OndatabasefolderClick(IRibbonControl control)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select the folder containing the database files";
                dialog.ShowNewFolderButton = false;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    config.AppSettings.Settings.Remove("DatabaseFolderPath");
                    config.AppSettings.Settings.Add("DatabaseFolderPath", dialog.SelectedPath);
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                }
            }
        }
        private string GetDatabaseFilePath()
        {
            string savedPath = ConfigurationManager.AppSettings["DatabaseFolderPath"];

            if (string.IsNullOrEmpty(savedPath))
            {
                using (FolderBrowserDialog dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "Select the folder containing the database files";
                    dialog.ShowNewFolderButton = false;

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        savedPath = dialog.SelectedPath;
                        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                        config.AppSettings.Settings.Remove("DatabaseFolderPath");
                        config.AppSettings.Settings.Add("DatabaseFolderPath", savedPath);
                        config.Save(ConfigurationSaveMode.Modified);
                        ConfigurationManager.RefreshSection("appSettings");
                    }
                }
            }

            return savedPath;
        }
        void ApplyBorders(Excel.Range cell)
        {
            cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        public void OnfindClick(IRibbonControl control)
        {
            var excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            // Check if the selected range is valid
            if (selectedRange == null || selectedRange.Cells.Count != 1)
            {
                MessageBox.Show("Please select a single cell.");
                return;
            }

            // Check if the selected cell is in column B
            if (selectedRange.Column != 2) // Column B is 2
            {
                MessageBox.Show("Please select a cell in column B.");
                return;
            }

            if (selectedRange.Interior.Color == 49407 && !string.IsNullOrEmpty(selectedRange.Value2?.ToString()))
            {
                if (formfind != null && !formfind.IsDisposed)
                {
                    formfind.Close();
                    formfind.Dispose();
                }
                string auto = selectedRange.Value2?.ToString();

                // Create a new instance of the form and show it
                formfind = new Formfind(auto,selectedRange);
                
            }
            else
            {
                if (formfind != null && !formfind.IsDisposed)
                {
                    formfind.Close();
                    formfind.Dispose();
                }
                string auto = "";
                // Create a new instance of the form and show it
                formfind = new Formfind(auto,selectedRange);
                formfind.OnFeederDataEntered += (feederData) =>
                {
                    string heading = $"{feederData.FeederName} -";
                    if (feederData.containsRYB) { heading = $"{heading} RYB"; };
                    if (feederData.containsRGA) { heading = $"{heading} RGA"; };
                    if (feederData.containsMFM) { heading = $"{heading} MFM"; };
                    if (feederData.containsELR) { heading = $"{heading} ELR"; };
                    if (feederData.containsSPD) { heading = $"{heading} SPD"; };
                    if (feederData.containsVM) { heading = $"{heading} VM"; };
                    if (feederData.containsAM) { heading = $"{heading} AM"; };
                    if (feederData.containsTEST1) { heading = $"{heading} TEST1"; };
                    if (feederData.containsTEST2) { heading = $"{heading} TEST2"; };
                    selectedRange.Value2 = heading;
                    selectedRange.Interior.Color = 49407;
                    selectedRange.Font.Bold = true;
                };
                formfind.Show();
            }

            
            
        }

        public void OnDataClick(IRibbonControl control)
        {
            // Close the existing form if it is open
            if (form1 != null && !form1.IsDisposed)
            {
                form1.Close();
                form1.Dispose();
            }

            // Create a new instance of the form and show it
            form1 = new Form1();
            form1.Show();
        }


    }
    

}
