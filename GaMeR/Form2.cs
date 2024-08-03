using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using GaMeR;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Configuration;


namespace GaMeR
{
    public partial class Find_Data : Form
    {

        private List<string> allItems = new List<string>();
        private System.Windows.Forms.Timer filterTimer;
        bool check11;
        bool check12;
        public Find_Data()
        {
            InitializeComponent();
            textBox1.TextChanged += TextBox1_TextChanged;
            filterTimer = new System.Windows.Forms.Timer();
            filterTimer.Interval = 350; // 500 milliseconds delay
            filterTimer.Tick += FilterTimer_Tick;

            // Initialize context menu for listView
            InitializeContextMenu();
            PopulateListView();

            // Handle KeyDown event for Ctrl+C copy functionality
            listView1.KeyDown += ListView1_KeyDown;
        }
        private void InitializeContextMenu()
        {
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem copyMenuItem = new ToolStripMenuItem("Copy");
            copyMenuItem.Click += CopyMenuItem_Click;
            contextMenu.Items.Add(copyMenuItem);
            listView1.ContextMenuStrip = contextMenu;
        }
        
        private void FilterTimer_Tick(object sender, EventArgs e)
        {
            // Timer elapsed, stop the timer and run the filter function
            filterTimer.Stop();
            FilterListView(textBox1.Text);
        }
        private void CopyMenuItem_Click(object sender, EventArgs e)
        {
            CopySelectedListViewItem();
        }
        private void ListView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectedListViewItem();
            }
        }
        private void CopySelectedListViewItem()
        {
            if (listView1.SelectedItems.Count > 0)
            {
                string selectedValue = listView1.SelectedItems[0].Text;
                Clipboard.SetText(selectedValue);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            RePopulateListView();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                string selectedValue = listView1.SelectedItems[0].Text;
                OpenCellForEditing(selectedValue);
            }
            else
            {
                MessageBox.Show("Please select an item from the list.");
            }

        }
        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            // Reset the timer when text changes
            filterTimer.Stop();
            filterTimer.Start();
        }

        private void RePopulateListView()
        {
            textBox1.Clear();
            FilterListView(textBox1.Text);
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


        private void PopulateListView()
        {
            string savedPath = GetDatabaseFilePath();
            if (string.IsNullOrEmpty(savedPath))
            {
                MessageBox.Show("Database file path is not set.");
                return;
            }
            Excel.Application excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Workbook extWorkbook = null;
            try
            {
                excelApp.ScreenUpdating = false;
                
                    string extFilePath = System.IO.Path.Combine(savedPath, "feeder_database.xlsx");
                    extWorkbook = excelApp.Workbooks.Open(
                        extFilePath,
                        UpdateLinks: 0, // 0 to not update external links
                        ReadOnly: true,
                        Editable: false,
                        IgnoreReadOnlyRecommended: true
                    );
                    foreach (Excel.Worksheet sheet in extWorkbook.Sheets)
                    {
                        Excel.Range cellA1 = sheet.Cells[1, 2];
                        string cellValue = cellA1.Value2?.ToString() ?? "";
                        if (!allItems.Contains(cellValue)) // Prevent duplicates
                        {
                            allItems.Add(cellValue);
                        }
                    }
                    extWorkbook.Close(false);
                    Marshal.ReleaseComObject(extWorkbook); ;

                FilterListView(textBox1.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }

        }

        private void FilterListView(string filter)
        {
            listView1.Items.Clear();
            foreach (string item in allItems)
            {
                if (item.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    listView1.Items.Add(item);
                }
            }
        }

        private void OpenCellForEditing(string selectedValue)
        {
            string savedPath = GetDatabaseFilePath();
            if (string.IsNullOrEmpty(savedPath))
            {
                MessageBox.Show("Database file path is not set.");
                return;
            }

            string extFilePath = System.IO.Path.Combine(savedPath, "feeder_database.xlsx");
            //string extFilePath2 = System.IO.Path.Combine(savedPath, "abb_database.xlsm");
            Excel.Application excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Workbook extWorkbook = null;
            //Excel.Workbook extWorkbook2 = null;
            try
            {
                extWorkbook = excelApp.Workbooks.Open(extFilePath);
                foreach (Excel.Worksheet sheet in extWorkbook.Sheets)
                {
                    Excel.Range cellA1 = sheet.Cells[1, 1];
                    if (cellA1.Value2 != null && cellA1.Value2.ToString() == selectedValue)
                    {
                        sheet.Activate(); // Activate the worksheet
                        extWorkbook.Windows[1].Visible = true;
                        cellA1.Select();
                        excelApp.ActiveWindow.Activate();
                        return;
                    }
                }

                Marshal.ReleaseComObject(extWorkbook);

                this.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }




    }
}
