using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static GaMeR.Formfind;
using System.Runtime.InteropServices;

namespace GaMeR
{
    public partial class Form1 : Form
    {
        private List<TabPage> _hiddenTabs;
        
        public Form1()
        {
            InitializeComponent();
            InitializeTabs();

        }
        private void InitializeTabs()
        {
            _hiddenTabs = new List<TabPage>();
            for (int i = 1; i < tabControl1.TabPages.Count; i++)
            {
                var tabPage = tabControl1.TabPages[i];
                _hiddenTabs.Add(tabPage);
            }
            for (int i = tabControl1.TabPages.Count - 1; i > 0; i--)
            {
                tabControl1.TabPages.RemoveAt(i);
            }
        }
        
        private void button3_Click(object sender, EventArgs e)
        {
            if (_hiddenTabs.Count > 0)
            {
                var tabPage = _hiddenTabs[0];
                tabControl1.TabPages.Insert(tabControl1.TabPages.Count, tabPage);
                _hiddenTabs.RemoveAt(0);
                UpdateTabCountLabel();
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (tabControl1.TabPages.Count > 1)
            {
                var tabPage = tabControl1.TabPages[tabControl1.TabPages.Count - 1];
                _hiddenTabs.Insert(0, tabPage);
                tabControl1.TabPages.RemoveAt(tabControl1.TabPages.Count - 1);
                UpdateTabCountLabel();
            }
        }
        private void UpdateTabCountLabel()
        {
            label26.Text = (tabControl1.TabPages.Count - 1).ToString(); // Subtracting 1 because the first tab is always visible
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

        private void runbutton_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Workbook templateWorkbook = null;
            Excel.Workbook newWorkbook = null;
            Excel.Worksheet costingSheet = null;
            Excel.Worksheet titleSheet = null;
            try
            {
                string projectname = projectnamebox.Text.ToString().ToUpper();
                string customername = customernamebox.Text.ToString().ToUpper();

                string paneltype = "";
                string panelmodel = "";

                foreach (Control control in groupBox1.Controls)
                {
                    if (control is RadioButton radioButton && radioButton.Checked)
                    {
                        paneltype = radioButton.Text;
                        break;
                    }
                }

                foreach (Control control in groupBox2.Controls)
                {
                    if (control is RadioButton radioButton && radioButton.Checked)
                    {
                        panelmodel = radioButton.Text;
                        break;
                    }
                }
                int panelCount;
                if (!int.TryParse(label26.Text, out panelCount))
                {
                    panelCount = 0; // Default to 0 if parsing fails
                }

                // Retrieve and store values from text fields on visible tabs
                var panelNames = new List<string>();

                for (int i = 0; i < panelCount; i++)
                {
                    // Assuming each visible tab is named like "Panel 1", "Panel 2", etc.
                    var tabName = $"Panel {i + 1}";
                    var tab = tabControl1.TabPages.Cast<TabPage>().FirstOrDefault(t => t.Text.Equals(tabName, StringComparison.OrdinalIgnoreCase));

                    if (tab != null)
                    {
                        var textBox = tab.Controls.OfType<TextBox>().FirstOrDefault(c => c.Name.Equals($"p{i + 1}name", StringComparison.OrdinalIgnoreCase));

                        if (textBox == null || textBox.Text == "")
                        {
                            label25.Visible = true;
                            label25.Text = "Enter all Fields First!!";
                            return;
                        }
                        else 
                        {
                            panelNames.Add(textBox.Text.ToUpper());
                        }
                    }
                }

                if (projectname == "" || customername == "" || paneltype == "" || panelmodel == "" || label26.Text == "0")
                {
                    label25.Visible = true;
                    label25.Text = "Enter all Fields First!!";
                    return;
                }

                string savedPath = GetDatabaseFilePath();

                if (string.IsNullOrEmpty(savedPath))
                {
                    System.Windows.Forms.MessageBox.Show("No folder path selected. Please select a folder first.");
                    return;
                }

                string extFilePath = System.IO.Path.Combine(savedPath, "templates.xlsx");
                

                templateWorkbook = excelApp.Workbooks.Open(extFilePath, ReadOnly: true);

                // Create a new workbook
                newWorkbook = excelApp.Workbooks.Add();

                // Copy each sheet from the template workbook to the new workbook
                foreach (Excel.Worksheet sheet in templateWorkbook.Sheets)
                {
                    sheet.Copy(After: newWorkbook.Sheets[newWorkbook.Sheets.Count]);
                }

                // Remove the default empty sheet created in the new workbook
                newWorkbook.Sheets[1].Delete();

                // Set the title of the new workbook window
                newWorkbook.Windows[1].Caption = projectname;
                newWorkbook.Application.DisplayAlerts = false;
                newWorkbook.Activate();

                foreach (Excel.Worksheet sheet in newWorkbook.Sheets)
                {
                    if (sheet.Name.Equals("ANALYSE", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.Delete();
                    }
                    else if (sheet.Name.Equals("BOM", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.Delete();
                    }
                    else if (sheet.Name.Equals("COSTING", StringComparison.OrdinalIgnoreCase))
                    {
                        costingSheet = sheet;
                    }
                    else if (sheet.Name.Equals("TITLE", StringComparison.OrdinalIgnoreCase))
                    {
                        titleSheet = sheet;
                    }
                }
                int startRowPrice = 10;
                int slno = 1;
                foreach (var panelName in panelNames)
                {
                    Excel.Range row9 = titleSheet.Rows[startRowPrice]; // Row 9 because Excel is 1-based
                    row9.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);

                    Excel.Range newRow = titleSheet.Rows[startRowPrice]; // New row inserted at the same index
                    newRow.Cells[1, 1].Value2 = slno.ToString();
                    newRow.Cells[1, 2].Value2 = panelName;
                    newRow.Cells[1, 8].Value2 = "1";

                    startRowPrice++;
                    slno++;
                }
                titleSheet.Cells[5, 1].Value2 = $"PROJECT: {projectname}";
                titleSheet.Cells[6, 1].Value2 = $"CUSTOMER: {customername}";
                titleSheet.Rows[9].Delete();
                titleSheet.Rows[startRowPrice - 1].Delete();

                int startRow = 54;
                foreach (var panelName in panelNames)
                {
                    Excel.Range cell = costingSheet.Cells[startRow, 2]; // Column B is index 2
                    cell.Value2 = panelName;
                    cell.Offset[0, -1].Value2 = "AL";
                    cell.Font.Bold = true; // Make the text bold
                    cell.Interior.Color = 15773696;

                    Excel.Range cellright = costingSheet.Cells[startRow, 3]; // Column B is index 2
                    cellright.Value2 = "1";
                    cellright.Font.Bold = true;
                    cellright.Interior.Color = 15773696;
                    cellright.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; // Center align text

                    // Apply borders
                    ApplyBorders(cell);
                    ApplyBorders(cellright);

                    Excel.Range cellBelow2 = costingSheet.Cells[startRow + 1, 2];
                    cellBelow2.Interior.Color = 49407; // Orange color

                    Excel.Range cellright2 = costingSheet.Cells[startRow + 1, 3]; // Column B is index 2
                    cellright2.Value2 = "1";
                    cellright2.Font.Bold = true; // Make the text bold
                    cellright2.Interior.Color = 49407;
                    cellright2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    ApplyBorders(cellBelow2);
                    ApplyBorders(cellright2);

                    Excel.Range cellBelow3 = costingSheet.Cells[startRow + 2, 2];
                    cellBelow3.Value2 = "PANEL UTILITY";
                    cellBelow3.Font.Bold = true;
                    cellBelow3.Interior.Color = 49407; // Orange color

                    Excel.Range cellright3 = costingSheet.Cells[startRow + 2, 3]; // Column B is index 2
                    cellright3.Value2 = "1";
                    cellright3.Font.Bold = true; // Make the text bold
                    cellright3.Interior.Color = 49407;
                    cellright3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    ApplyBorders(cellBelow3);
                    ApplyBorders(cellright3);

                    Excel.Range cellBelow4 = costingSheet.Cells[startRow + 3, 2];
                    cellBelow4.Value2 = "ENCLOSURE AND BUSBAR + EARTH";
                    cellBelow4.Font.Bold = true;
                    cellBelow4.Interior.Color = 49407; // Orange color

                    Excel.Range cellright4 = costingSheet.Cells[startRow + 3, 3]; // Column B is index 2
                    cellright4.Value2 = "1";
                    cellright4.Font.Bold = true; // Make the text bold
                    cellright4.Interior.Color = 49407;
                    cellright4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    ApplyBorders(cellBelow4);
                    ApplyBorders(cellright4);

                    // Move startRow down by 5 to account for the four extra rows
                    startRow += 4;
                }
                templateWorkbook.Close(false);
                costingSheet.Activate();
                Excel.Range usedRange = costingSheet.UsedRange;

                // List to hold the orange cells
                List<Excel.Range> orangeCells = new List<Excel.Range>();

                // Collect all orange cells with values
                for (int row = usedRange.Rows.Count; row >= 1; row--)
                {
                    Excel.Range cell = costingSheet.Cells[row, 2]; // Column B is index 2
                    if (cell.Interior.Color == 49407 && !string.IsNullOrEmpty(cell.Value2?.ToString()))
                    {
                        orangeCells.Add(cell);
                    }
                }
                
                foreach (var orangeCell in orangeCells)
                {
                    Formfind form = new Formfind(orangeCell.Value2.ToString(), orangeCell);
                }

                label25.Visible = false;
            }
            catch (Exception ex) 
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                Marshal.ReleaseComObject( templateWorkbook );
                newWorkbook.Application.DisplayAlerts = true;
                this.Close();
            }

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

    }

    
    
}
