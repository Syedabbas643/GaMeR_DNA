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
using Microsoft.Office.Core;
using IRibbonControl = ExcelDna.Integration.CustomUI.IRibbonControl;
using IRibbonUI = ExcelDna.Integration.CustomUI.IRibbonUI;
using System.Net.Http;
using System.Threading.Tasks;
using System.Runtime.Serialization.Json;

namespace GaMeR
{
    [ComVisible(true)]
    public class CustomRibbon : ExcelRibbon
    {
        private Formfind formfind;
        private Form1 form1;
        private Find_Data Find_Data;
        private Timer _authorizationCheckTimer;
        
        public override string GetCustomUI(string RibbonID)
        {
            return @"
              <customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""OnLoad"">
          <ribbon>
            <tabs>
              <tab id=""tab2"" label=""GaMeR"">
                <group id=""group2"" label=""Automate"">
                  <button id=""data1"" label=""Create New Costing"" getImage=""GetCustomImage"" size=""large"" onAction=""OnDataClick""/>
                  <separator id=""separator2""/>
                  <button id=""data2"" label=""GET from Feeder"" getImage=""GetCustomImage"" size=""large"" onAction=""OnfeederClick""/>
                  <button id=""find2"" label=""Search feeder"" getImage=""GetCustomImage"" size=""large"" onAction=""OnsearchClick""/>
                  <separator id=""separator1""/>
                  <button id=""data3"" label=""Automate"" getImage=""GetCustomImage"" size=""large"" onAction=""OnfindClick""/>
                  <button id=""find"" label=""Automate all Feeders"" getImage=""GetCustomImage"" size=""large"" onAction=""OnallClick""/>
                  <separator id=""separator3""/>
                  <button id=""bom2"" label=""Make Bill of Materials"" getImage=""GetCustomImage"" size=""large"" onAction=""OnbomnewClick""/>
                  <button id=""cad"" label=""Analyse Costing"" getImage=""GetCustomImage"" size=""large"" onAction=""OnanalyseClick""/>
                  <button id=""bom"" label=""Make Bill of OLD Materials"" getImage=""GetCustomImage"" onAction=""OnbomClick""/>
                  <button id=""server"" label=""database folder"" getImage=""GetCustomImage"" onAction=""OndatabasefolderClick""/>
                  <separator id=""separator4""/>
                  <button id=""auto2"" label=""Copy to below FEEDERS"" getImage=""GetCustomImage"" size=""large"" onAction=""OnbelowClick""/>
                  <button id=""auto"" label=""Read COSTING sheet"" getImage=""GetCustomImage"" size=""large"" onAction=""OnreadClick""/>
                  <button id=""layout"" label=""Automate GA sheet"" getImage=""GetCustomImage"" size=""large"" onAction=""OngaClick""/>
                  <separator id=""separator5""/>
                  <button id=""help"" label=""HELP"" getImage=""GetCustomImage"" size=""large"" onAction=""OnhelpClick""/>
                  <button id=""about"" label=""About ME"" getImage=""GetCustomImage"" size=""large"" onAction=""OnaboutClick""/>
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
        public void CheckAuthorizationAsync()
        {
            string clientName = ConfigurationManager.AppSettings["name"];
            if (clientName == null) 
            {
                clientName = "test";
            }
            string apiUrl = "https://syedabbas.up.railway.app/check";

            try
            {
                _authorizationCheckTimer = new Timer();
                _authorizationCheckTimer.Interval = 1200000; // 20 minutes in milliseconds
                _authorizationCheckTimer.Tick += async (sender, args) =>
                {
                    using (HttpClient client = new HttpClient())
                    {
                        try 
                        {
                            HttpResponseMessage response = await client.GetAsync($"{apiUrl}/{clientName}");
                            if (response.IsSuccessStatusCode)
                            {
                                string responseBody = await response.Content.ReadAsStringAsync();
                                using (var ms = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(responseBody)))
                                {
                                    var serializer = new DataContractJsonSerializer(typeof(bool));
                                    bool isAuthorized = (bool)serializer.ReadObject(ms);

                                    if (isAuthorized)
                                    {
                                        _authorizationCheckTimer.Stop();
                                    }
                                    else
                                    {
                                        ExcelAsyncUtil.QueueAsMacro(() =>
                                        {
                                            try
                                            {
                                                Excel.Application excelApp = ExcelDnaUtil.Application as Excel.Application;
                                                excelApp.DisplayAlerts = false; // Suppress any save changes dialogs
                                                excelApp.Quit(); // Quit the application
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine("Error closing Excel: " + ex.Message);
                                            }
                                        });

                                        Environment.Exit(0);
                                    }
                                }

                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }

                };
                _authorizationCheckTimer.Start();
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        
        public void OnLoad(IRibbonUI ribbonUI)
        {
            
           CheckAuthorizationAsync();

        }

        public void OnhelpClick(IRibbonControl control)
        {
            Excel.Application excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Workbook extWorkbook = null;

            string savedPath = GetDatabaseFilePath();

            if (string.IsNullOrEmpty(savedPath))
            {
                System.Windows.Forms.MessageBox.Show("No folder path selected. Please select a folder first.");
                return;
            }

            string extFilePath = System.IO.Path.Combine(savedPath, "database.xlsx");
            string workbookName = System.IO.Path.GetFileName(extFilePath);

            excelApp.DisplayAlerts = false;  // Disable alerts
            excelApp.ScreenUpdating = false;

            foreach (Excel.Workbook wb in excelApp.Workbooks)
            {
                if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                {
                    wb.Close(true);
                    break;
                    
                }
            }
            if (extWorkbook == null)
            {
                extWorkbook = excelApp.Workbooks.Open(extFilePath, false, false);

            }

            Excel.Worksheet tempSheet = null;
            foreach (Excel.Worksheet sheet in extWorkbook.Sheets)
            {
                if (sheet.Name == "HELP")
                {
                    tempSheet = sheet;
                    break;
                }
            }
            tempSheet.Activate();

            excelApp.DisplayAlerts = true;  // Disable alerts
            excelApp.ScreenUpdating = true;

            Marshal.ReleaseComObject(tempSheet);
            Marshal.ReleaseComObject(extWorkbook);

        }

        public void OnaboutClick(IRibbonControl control)
        {
            MessageBox.Show(
                "Welcome to the Add-In!\n\n" +
                "Thank you for choosing my Excel add-in to enhance your productivity and streamline your workflows. \n\n" +
                "My Mission is to simplify your tasks and unlock new possibilities within Excel, helping you turn challenges into opportunities.\n\n" +
                ">>Nothing is impossible<<\n\n" +
                "Developed by --- GaMeR " +
                "",
                "About GaMeR Add-In",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        public void OnreadClick(IRibbonControl control)
        {
            var excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Workbook costingWorkbook = null;
            Excel.Workbook databaseWorkbook = null;
            Excel.Worksheet costingSheet = null;
            List<string> errorMessages = new List<string>(); // List to store error messages

            try
            {
                string savedPath = GetDatabaseFilePath();

                if (string.IsNullOrEmpty(savedPath))
                {
                    MessageBox.Show("No folder path selected. Please select a folder first.");
                    return;
                }

                string extFilePath = System.IO.Path.Combine(savedPath, "database.xlsx");

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
                    openFileDialog.Title = "Select Excel Files";
                    openFileDialog.Multiselect = true; // Allow multiple file selection

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        databaseWorkbook = excelApp.Workbooks.Open(
                            extFilePath,
                            UpdateLinks: 0, // 0 to not update external links
                            ReadOnly: false,
                            Editable: true
                        );

                        HashSet<string> databaseValues = new HashSet<string>();

                        // Scan the database workbook first
                        foreach (Excel.Worksheet sheet in databaseWorkbook.Sheets)
                        {
                            Excel.Range usedRange = sheet.UsedRange;
                            Excel.Range columnB = usedRange.Columns["B"];

                            foreach (Excel.Range cell in columnB.Cells)
                            {
                                if (cell.Row > 52) // Exclude first 52 rows
                                {
                                    string cellValue = cell.Value2?.ToString();
                                    if (cell.Interior.Color != 49407 && !string.IsNullOrEmpty(cellValue))
                                    {
                                        databaseValues.Add(cellValue);
                                    }
                                }
                            }
                        }
                        Excel.Worksheet newSheet = (Excel.Worksheet)databaseWorkbook.Sheets.Add();
                        int newSheetRow = 2;


                        // Process each selected file
                        foreach (string fileName in openFileDialog.FileNames)
                        {
                            try
                            {
                                // Open the current workbook
                                excelApp.DisplayAlerts = false;
                                costingWorkbook = excelApp.Workbooks.Open(fileName, false);

                                // Find the COSTING sheet
                                costingSheet = null;
                                foreach (Excel.Worksheet sheet in costingWorkbook.Sheets)
                                {
                                    if (sheet.Name.Equals("COSTING", StringComparison.OrdinalIgnoreCase))
                                    {
                                        costingSheet = sheet;
                                        break;
                                    }
                                }

                                if (costingSheet == null)
                                {
                                    errorMessages.Add($"The workbook '{fileName}' does not contain a COSTING sheet.");
                                    continue; // Skip this file
                                }

                                Excel.Range costingUsedRange = costingSheet.UsedRange;
                                Excel.Range costingColumnB = costingUsedRange.Columns["B"];

                                foreach (Excel.Range cell in costingColumnB.Cells)
                                {
                                    if (cell.Row > 52) // Exclude first 52 rows
                                    {
                                        string cellValue = cell.Value2?.ToString();
                                        
                                        if (cell.Interior.Color != 49407 && cell.Interior.Color != 15773696 && cell.Interior.Color != 65535 && !string.IsNullOrEmpty(cellValue) && !databaseValues.Contains(cellValue))
                                        {
                                            databaseValues.Add(cellValue);
                                            Excel.Range entireRow = cell.EntireRow;
                                            entireRow.Copy(newSheet.Rows[newSheetRow]);
                                            newSheetRow++;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                errorMessages.Add($"Error processing '{fileName}': {ex.Message}");
                            }
                            finally
                            {
                                if (costingWorkbook != null)
                                {
                                    costingWorkbook.Close(false);
                                    costingWorkbook = null; // Reset the reference
                                }
                            }


                        }

                        OrganizeDataByMake(newSheet);
                    }
                    
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"General error: {ex.Message}");
            }
            finally
            {
                excelApp.DisplayAlerts = true;

                if (costingWorkbook != null)
                {
                    costingWorkbook.Close(false);
                }
            }

            if (errorMessages.Any())
            {
                MessageBox.Show(string.Join("\n", errorMessages), "Processing Errors", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void OrganizeDataByMake(Excel.Worksheet sheet)
        {
            Dictionary<string, List<Excel.Range>> makeRows = new Dictionary<string, List<Excel.Range>>();
            Excel.Range usedRange = sheet.UsedRange;

            int lastRow = usedRange.Rows.Count;
           
            for (int rowIndex = 2; rowIndex <= lastRow; rowIndex++) // Assuming data starts from row 2
            {
                Excel.Range row = sheet.Rows[rowIndex];
                string make = row.Cells[1, 4].Value2?.ToString() ?? ""; // Assuming make is in column C

                if (!makeRows.ContainsKey(make))
                {
                    makeRows[make] = new List<Excel.Range>();
                }

                makeRows[make].Add(row);
            }

            if (makeRows.Count == 0)
            {
                return;
            }

            // Reinsert data grouped by make
            int currentRow = lastRow + 2; // Start from the first row

            foreach (var make in makeRows.Keys)
            {
                foreach (var row in makeRows[make])
                {
                    row.Copy(sheet.Rows[currentRow]);
                    currentRow++;
                }
                // Add a blank row between different makes
                currentRow++;
            }
            if (lastRow > 2) // Avoid deleting all rows if there's no data
            {
                Excel.Range rowsToDelete = sheet.Rows["2:" + (lastRow + 1)];
                rowsToDelete.Delete(); // Delete all rows from row 2 to the last used row
            }

        }

        public void OnbelowClick(IRibbonControl control)
        {
            var excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;
            Excel.Worksheet currentsheet = selectedRange.Worksheet;
            Excel.Range usedrange = currentsheet.UsedRange;

            try
            {
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
                excelApp.DisplayAlerts = false; 
                string feederHeading = selectedRange.Value2?.ToString() ?? "";
                Excel.Range cell = selectedRange.Offset[1, 0];
                List<(string Description, string CatalogNumber, string Price, Excel.Range Row)> dataBelowSelectedRange = new List<(string, string, string, Excel.Range)>();

                while (cell.Interior.Color != 49407 && cell.Interior.Color != 15773696 && cell.Interior.Color != 65535 && cell.Row <= excelApp.ActiveSheet.UsedRange.Rows.Count)
                {
                    string description = cell.Value2?.ToString() ?? "";
                    string catalogNumber = cell.Offset[0, 1].Value2?.ToString() ?? "";
                    string price = cell.Offset[0, 4].Value2?.ToString() ?? "";

                    dataBelowSelectedRange.Add((description, catalogNumber, price, cell.EntireRow));
                    cell = cell.Offset[1, 0];
                }

                Excel.Range ColumnBrange = usedrange.Columns["B"];
                List<Excel.Range> columnBCells = new List<Excel.Range>();

                // Collect column B cells in reverse order
                foreach (Excel.Range cell2 in ColumnBrange.Cells)
                {
                    if (cell2.Row > selectedRange.Row)
                    {
                        columnBCells.Add(cell2);
                    }
                }
                string feederqty = "1";
                string panelqty = "1";
                int lastRow = usedrange.Rows.Count;
                // Sort columnBCells by row number in descending order
                columnBCells = columnBCells.OrderByDescending(c => c.Row).ToList();
                foreach (Excel.Range cells in columnBCells)
                {
                    if (cells.Row > selectedRange.Row)
                    {
                        Excel.Range cell2 = cells;
                        string cellValue = cells.Value2?.ToString();

                        if (cells.Interior.Color == 49407 && !string.IsNullOrEmpty(cellValue) && cellValue == feederHeading)
                        {
                            if (cells.Row == lastRow ||cells.Offset[1, 0].Interior.Color == 49407 || cells.Offset[1, 0].Interior.Color == 15773696 || cells.Offset[1, 0].Interior.Color == 65535)
                            {
                                List<Excel.Range> rowsToCopy = new List<Excel.Range>();

                                try
                                {
                                    feederqty = cells.Offset[0, 1].Value2.ToString();
                                    for (int row = cells.Row - 1; row >= 1; row--)
                                    {
                                        Excel.Range temp = usedrange.Cells[row, 2];
                                        if (temp.Interior.Color == 15773696)
                                        {
                                            string sum = temp.Offset[0, 1].Value2.ToString();
                                            panelqty = sum;
                                            break;
                                        }

                                    }
                                }
                                catch
                                {
                                    MessageBox.Show("No Panel or Feeder Quantity found. So keeping the default Value");
                                }

                                if (dataBelowSelectedRange.Any())
                                {
                                    foreach (var data in dataBelowSelectedRange)
                                    {
                                        rowsToCopy.Add(data.Row);
                                    }
                                    for (int i = rowsToCopy.Count - 1; i >= 0; i--)
                                    {
                                        Excel.Range row = rowsToCopy[i];
                                        row.Copy();
                                        cells.Offset[1, -1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                                    }

                                    for (int i = 1; i <= rowsToCopy.Count; i++)
                                    {
                                        if (feederqty != null)
                                        {
                                            cell2.Offset[i, 7].Value2 = feederqty;
                                        }
                                        if (panelqty != null)
                                        {
                                            cell2.Offset[i, 9].Value2 = panelqty;
                                        }
                                    }

                                }
                            }
                            else if ((cells.Offset[1, 0].Interior.Color != 49407 || cells.Offset[1, 0].Interior.Color != 15773696 || cells.Offset[1, 0].Interior.Color != 65535) && !string.IsNullOrEmpty(cells.Offset[1, 0].Value2.ToString()))
                            {
                                //while (cell.Interior.Color != 49407 && cell.Interior.Color != 15773696 && cell.Interior.Color != 65535 && cell.Row <= excelApp.ActiveSheet.UsedRange.Rows.Count)
                                //{
                                    //string description = cell.Value2?.ToString() ?? "";
                                    //string catalogNumber = cell.Offset[0, 1].Value2?.ToString() ?? "";
                                   // string price = cell.Offset[0, 4].Value2?.ToString() ?? "";

                                    //cell = cell.Offset[1, 0];
                               // }
                            }
                         }
                    }
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
            finally 
            {
                excelApp.DisplayAlerts = true;
            }

        }
        public void OnfeederClick(IRibbonControl control)
        {
            var excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Range selectedCell = excelApp.Selection as Excel.Range;
            if (selectedCell != null && selectedCell.Value2 != null)
            {
                
                CopyFromExternalWorkbook(selectedCell);
            }
        }
        private void CopyFromExternalWorkbook(Excel.Range selectedCell)
        {
            var excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Workbook currentWorkbook = excelApp.ActiveWorkbook;
            Excel.Worksheet currentSheet = excelApp.ActiveSheet;
            Excel.Workbook extWorkbook = null;
            Excel.Worksheet extSheet = null;


            try
            {
                string savedPath = GetDatabaseFilePath();

                if (string.IsNullOrEmpty(savedPath))
                {
                    MessageBox.Show("No folder path selected. Please select a folder first.");
                    return;
                }
                string extFilePath = System.IO.Path.Combine(savedPath, "feeder_database.xlsx");

                extWorkbook = excelApp.Workbooks.Open(
                    extFilePath,
                    UpdateLinks: 0, // 0 to not update external links
                    ReadOnly: true,
                    Editable: false,
                    IgnoreReadOnlyRecommended: true
                );

                int matchCount = 0;

                foreach (Excel.Worksheet sheet in extWorkbook.Sheets)
                {
                    // Check if cell A1 in the sheet contains the desired value (partial, case-insensitive match)
                    Excel.Range cellA1 = sheet.Cells[1, 2];
                    if (cellA1.Value2 != null && cellA1.Value2.ToString().IndexOf(selectedCell.Value2.ToString(), StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        extSheet = sheet;
                        matchCount++;
                    }
                }

                if (extSheet != null && matchCount == 1)
                {
                    Excel.Range copyRange = extSheet.UsedRange;
                    int numberOfRows = copyRange.Rows.Count;

                    // Calculate the paste range (one cell below the selected cell)
                    int pasteRow = selectedCell.Row;
                    int pasteColumn = selectedCell.Column -1;

                    // Insert the required number of rows below the selected cell
                    for (int i = 1; i < numberOfRows; i++)
                    {
                        Excel.Range insertRange = currentSheet.Rows[selectedCell.Row + i];
                        insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);
                    }

                    copyRange.Copy();
                    Excel.Range pasteRange = currentSheet.Cells[selectedCell.Row, pasteColumn];
                    pasteRange.PasteSpecial(Excel.XlPasteType.xlPasteAll,
                                            Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                            false, false);


                    excelApp.CutCopyMode = 0;

                    extWorkbook.Close(false);
                }
                else if (matchCount > 1)
                {
                    extWorkbook.Close(false);
                    System.Windows.Forms.MessageBox.Show("Multiple sheets contain the SAME value. Please refine your Data.");
                }
                else
                {
                    extWorkbook.Close(false);
                    System.Windows.Forms.MessageBox.Show("No match found in the external database.");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {

                if (extWorkbook != null)
                {
                    Marshal.ReleaseComObject(extWorkbook);
                }
            }

        }
        public void OnsearchClick(IRibbonControl control)
        {
            // Close the existing form if it is open
            if (Find_Data != null && !Find_Data.IsDisposed)
            {
                Find_Data.Close();
                Find_Data.Dispose();
            }

            // Create a new instance of the form and show it
            Find_Data = new Find_Data();
            Find_Data.Show();
        }
        public void OngaClick(IRibbonControl control)
        {
            try 
            {
                var excelApp = ExcelDnaUtil.Application as Excel.Application;
                Excel.Worksheet currentSheet = excelApp.ActiveSheet;
                Excel.Workbook currentWorkbook = excelApp.ActiveWorkbook;

                if (currentSheet.Name != "COSTING") 
                {
                    MessageBox.Show("PLZ RUN THE SCRIPT ON COSTING SHEET ONLY");
                    return;
                }

                Excel.Range usedRange = currentSheet.UsedRange;
                Excel.Range columnB = usedRange.Columns["B"];

                var panelData = new List<(Excel.Range panelHeading, List<(Excel.Range feederHeading, Excel.Range feederQuantity)>)>();

                List<(Excel.Range feederHeading, Excel.Range feederQuantity)> currentFeeders = null;
                Excel.Range currentPanelHeading = null;

                foreach (Excel.Range cell in columnB.Cells)
                {
                    if (cell.Interior.Color == 15773696)
                    {
                        if (currentPanelHeading != null && currentFeeders != null)
                        {
                            panelData.Add((currentPanelHeading, currentFeeders));
                        }

                        currentPanelHeading = cell;
                        currentFeeders = new List<(Excel.Range feederHeading, Excel.Range feederQuantity)>();
                    }
                    else if (cell.Interior.Color == 49407 && !string.IsNullOrEmpty(cell.Value2?.ToString()) && cell.Value2 != "PANEL UTILITY" && cell.Value2 != "ENCLOSURE AND BUSBAR + EARTH")
                    {
                        currentFeeders?.Add((cell, cell.Offset[0, 1])); // Add feeder heading and its quantity (one cell to the right)
                    }
                }

                if (currentPanelHeading != null && currentFeeders != null)
                {
                    panelData.Add((currentPanelHeading, currentFeeders));
                }

                // Create a new sheet
                Excel.Worksheet newSheet = (Excel.Worksheet)currentWorkbook.Sheets.Add();
                

                // Start copying headings to the new sheet
                int startRow = 1; // Start at row 1
                foreach (var (panelHeading, feederHeadings) in panelData)
                {
                    if (panelHeading != null)
                    {
                        Excel.Range widthCell = newSheet.Cells[startRow, 1];
                        widthCell.Value2 = "Width";
                        widthCell.Font.Bold = true;
                        ApplyBorders(widthCell);
                        Excel.Range heightCell = newSheet.Cells[startRow + 1, 1];
                        heightCell.Value2 = "Height";
                        heightCell.Font.Bold = true;
                        ApplyBorders(heightCell);
                        Excel.Range depthCell = newSheet.Cells[startRow + 2, 1];
                        depthCell.Value2 = "Depth";
                        depthCell.Font.Bold = true;
                        ApplyBorders(depthCell);

                        // Insert 0 in column B (second column)
                        Excel.Range zeroCell = newSheet.Cells[startRow, 2];
                        zeroCell.Value2 = 0;
                        ApplyBorders(zeroCell);
                        Excel.Range zeroCell1 = newSheet.Cells[startRow + 1, 2];
                        zeroCell1.Value2 = 0;
                        ApplyBorders(zeroCell1);
                        Excel.Range zeroCell2 = newSheet.Cells[startRow + 2, 2];
                        zeroCell2.Value2 = 0;
                        ApplyBorders(zeroCell2);

                        Excel.Range targetRange = newSheet.Range[newSheet.Cells[startRow, 4], newSheet.Cells[startRow, 8]];
                        targetRange.Merge();
                        targetRange.Value2 = panelHeading.Value2;
                        targetRange.Interior.Color = panelHeading.Interior.Color;
                        targetRange.Font.Bold = panelHeading.Font.Bold;
                        targetRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                        // Apply borders (if needed, you can modify this part to match your desired border settings)
                        ApplyBorders(targetRange);
                    }

                    
                    int feederStartColumn = 15;
                    int feedercolumn = startRow + 1;
                    foreach (var (feederHeading, feederQuantity) in feederHeadings)
                    {
                        Excel.Range newFeederCell = newSheet.Cells[feedercolumn, feederStartColumn];
                        newFeederCell.Value2 = feederHeading.Value2;
                        newFeederCell.Font.Bold = feederHeading.Font.Bold;
                        ApplyBorders(newFeederCell);

                        Excel.Range newFeederQuantityCell = newSheet.Cells[feedercolumn, feederStartColumn + 1];
                        newFeederQuantityCell.Value2 = feederQuantity.Value2;
                        newFeederQuantityCell.Font.Bold = feederQuantity.Font.Bold;
                        newFeederQuantityCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        ApplyBorders(newFeederQuantityCell);

                        feedercolumn++;
                    }

                    startRow += 40;
                    
                }
                newSheet.Columns[15].AutoFit();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}");
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

        public void OnallClick(IRibbonControl control)
        {
            try
            {
                var excelApp = ExcelDnaUtil.Application as Excel.Application;
                Excel.Worksheet currentSheet = excelApp.ActiveSheet;
                Excel.Workbook currentWorkbook = excelApp.ActiveWorkbook;

                if (currentSheet.Name != "COSTING") 
                {
                    MessageBox.Show("PLZ RUN THE SCRIPT ON COSTING SHEET ONLY");
                    return;
                }
                Excel.Range usedRange = currentSheet.UsedRange;

                // List to hold the orange cells
                List<Excel.Range> orangeCells = new List<Excel.Range>();

                // Collect all orange cells with values
                for (int row = usedRange.Rows.Count; row >= 1; row--)
                {
                    Excel.Range cell = currentSheet.Cells[row, 2]; // Column B is index 2
                    if (cell.Interior.Color == 49407 && !string.IsNullOrEmpty(cell.Value2?.ToString()) && (cell.Offset[1,0].Interior.Color == 49407 || cell.Offset[1, 0].Interior.Color == 15773696 || string.IsNullOrEmpty(cell.Offset[1,0].Value2?.ToString())))
                    {
                        orangeCells.Add(cell);
                    }
                }

                foreach (var orangeCell in orangeCells)
                {
                    Formfind form = new Formfind("automate643", orangeCell);
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
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

        [ExcelFunction(Description = "XLOOKUP UDF for Excel 2016")]
        public static object XLookup(
        [ExcelArgument(Name = "lookup_value", Description = "Value to search for")] object lookupValue,
        [ExcelArgument(Name = "lookup_array", Description = "Array to search within")] object[] lookupArray,
        [ExcelArgument(Name = "return_array", Description = "Array to return values from")] object[] returnArray,
        [ExcelArgument(Name = "if_not_found", Description = "Value to return if not found")] object ifNotFound = null)
        {
            // Ensure that lookupArray and returnArray are of the same length
            if (lookupArray.Length != returnArray.Length)
            {
                return ExcelError.ExcelErrorValue;
            }

            // Convert lookupArray and returnArray to strings
            string lookupValueStr = Convert.ToString(lookupValue);
            for (int i = 0; i < lookupArray.Length; i++)
            {
                string lookupArrayStr = Convert.ToString(lookupArray[i]);
                if (lookupValueStr.Equals(lookupArrayStr, StringComparison.OrdinalIgnoreCase))
                {
                    return returnArray[i];
                }
            }

            // Return the ifNotFound value if no match is found
            return ifNotFound ?? ExcelError.ExcelErrorNA;
        }


    }
    

}
