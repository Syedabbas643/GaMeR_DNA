using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Drawing;
using Rectangle = System.Drawing.Rectangle;
using Point = System.Drawing.Point;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TextBox = System.Windows.Forms.TextBox;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

namespace GaMeR
{
    public partial class Formfind : Form
    {
        TextBox feeder = new TextBox();
        public Formfind(string auto,Excel.Range SelectedRange)
        {
            InitializeComponent();
            
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            int x = workingArea.Width - this.Width;
            int y = workingArea.Height / 2 - this.Height / 2;
            this.Location = new Point(x, y);

            if (auto == "") 
            {
                openasnewform();
            }
            else
            {
                openasdatabase(auto,SelectedRange);
            }
        }
        public class FormData
        {
            public string FeederName { get; set; }
            public bool containsELR { get; set; }
            public bool containsSPD { get; set; }
            public bool containsMFM { get; set; }
            public bool containsRYB { get; set; }
            public bool containsRGA { get; set; }
            public bool containsVM { get; set; }
            public bool containsAM { get; set; }
            public bool containsTEST1 { get; set; }
            public bool containsTEST2 { get; set; }
            public string MFMcatno { get; set; }
            public List<string> SelectedCheckboxes { get; set; }

        }
        public event Action<FormData> OnFeederDataEntered;

        public void openasdatabase(string auto,Excel.Range SelectedRange)
        {
            try
            {
                Excel.Application excelApp = ExcelDnaUtil.Application as Excel.Application;


                string cellValue = SelectedRange.Value2.ToString();

                string name = cellValue.Replace(",", " ").ToLower();

                bool containsRYB = name.Contains("ryb");
                bool containsRGA = name.Contains("rga");
                bool containsMFM = name.Contains("mfm");
                bool containsMCCB = name.Contains("mccb");
                bool containsMCB = name.Contains("mcb");
                bool containsACB = name.Contains("acb");
                bool containsUTIL = name.Contains("utility");
                bool containsAM = name.Contains("am");
                bool containsVM = name.Contains("vm");
                bool containsELR = name.Contains("elr");
                bool containsSPD = name.Contains("spd");
                bool containsCOS = name.Contains("cos");
                bool containsSDFU = name.Contains("sdfu");
                bool containsELC = name.Contains("enclosure");
                bool containsAMSS = name.Contains("amss");
                bool containsVMSS = name.Contains("vmss");
                bool containsHRC = name.Contains("hrc");
                bool containsKVAR = name.Contains("kvar");

                int amps1 = 0;

                Match ampsMatch1 = Regex.Match(name, @"(\d+) ?a", RegexOptions.IgnoreCase);
                if (ampsMatch1.Success)
                {
                    amps1 = int.Parse(ampsMatch1.Groups[1].Value);
                }

                label1.Text = "NO MATCH:";
                if (containsUTIL)
                {
                    label1.Text = "PANEL UTILITY";
                }
                else if (containsELC)
                {
                    label1.Text = "ENCLOSURE AND BUSBAR + EARTH";
                }
                else if (containsCOS)
                {
                    string[] parts = name.Split(new string[] { "cos" }, StringSplitOptions.None);
                    string beforeMCB = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                    string result = beforeMCB.ToUpper();
                    label1.Text = $"{result} COS:";
                }
                else if (containsSDFU)
                {
                    string[] parts = name.Split(new string[] { "sdfu" }, StringSplitOptions.None);
                    string beforeMCB = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                    string result = beforeMCB.ToUpper();
                    label1.Text = $"{result} SDFU:";
                }
                else if (containsMCB)
                {
                    string[] parts = name.Split(new string[] { "mcb" }, StringSplitOptions.None);
                    string beforeMCB = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                    string result = beforeMCB.ToUpper();
                    label1.Text = $"{result} MCB:";
                }
                else if (containsMCCB)
                {
                    string[] parts = name.Split(new string[] { "mccb" }, StringSplitOptions.None);
                    string beforeMCCB = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                    string result = beforeMCCB.ToUpper();
                    label1.Text = $"{result} MCCB:";
                    comboBox1.Items.AddRange(new string[] { "AUTO", "DZ1", "DU", "DN0", "DN1", "DN2", "DN3" });
                    comboBox1.SelectedIndex = 0;
                    checkedListBox1.Items.Add("Spreaders");
                    checkedListBox1.Items.Add("Extended ROM");
                    checkedListBox1.Items.Add("Aux Contacts");
                    checkedListBox1.Items.Add("Shunt Release");
                    checkedListBox1.Items.Add("Additional item");
                    checkedListBox1.SetItemChecked(1,true);
                    if (amps1 > 64)
                    {
                        checkedListBox1.SetItemChecked(0, true);
                    }
                    if (containsELR)
                    {
                        checkedListBox1.SetItemChecked(3, true);
                    }
                    if (containsRGA)
                    {
                        checkedListBox1.SetItemChecked(2, true);
                    }
                }
                else if (containsACB)
                {
                    string[] parts = name.Split(new string[] { "acb" }, StringSplitOptions.None);
                    string beforeMCCB = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                    string result = beforeMCCB.ToUpper();
                    label1.Text = $"{result} ACB:";
                    comboBox1.Items.AddRange(new string[] { "AUTO", "50KA", "65KA", "80KA", "100KA" });
                    comboBox1.SelectedIndex = 0;
                    checkedListBox1.Items.Add("Door interlock");
                    checkedListBox1.Items.Add("C- Lock");
                    checkedListBox1.Items.Add("UV release");
                    checkedListBox1.Items.Add("Closing release");
                    checkedListBox1.Items.Add("Shunt release");
                    checkedListBox1.Items.Add("Additional item");
                }

                rybcheckbox.Checked = containsRYB;
                rgacheckbox.Checked = containsRGA;
                ambox.Checked = containsAM;
                vmbox.Checked = containsVM;
                elrbox.Checked = containsELR;
                spdbox.Checked = containsSPD;
                mfmcheckbox.Checked = containsMFM;
                hrccheckbox.Checked = containsHRC;
                if (containsKVAR)
                {
                    string[] parts = name.Split(new string[] { "kvar" }, StringSplitOptions.None);
                    string beforeMCB = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                    string result = beforeMCB.ToUpper();
                    label1.Text = $"{result}KVAR:";
                    reactorcheckbox.Visible = true;
                    reactorcheckbox.Checked = true;
                }
                if (containsAMSS)
                {
                    ambox.Checked = containsAMSS;
                    asscheckbox.Checked = containsAMSS;
                }
                if (containsVMSS) 
                {
                    vmbox.Checked = containsVMSS;
                    vsscheckbox.Checked = containsVMSS;
                }

                string getmfmcatno = ConfigurationManager.AppSettings["MfmCatno"];
                if (getmfmcatno != null)
                {
                    mfmcat.Text = getmfmcatno;
                }

                ctamps.Items.AddRange(new string[] { "CLASS 1", "CLASS 0.5" });

                ctva.Items.AddRange(new string[] { "5VA", "15VA" });

                fusecombobox.Items.AddRange(new string[] { "Control","6","8","10","16","20","25","32","40","50","63","80","100","125","160","200", "250","315","400","500","630","800" });

                Match hrcTypeMatch = Regex.Match(name, @"hrc(\d+)", RegexOptions.IgnoreCase);
                if (hrcTypeMatch.Success)
                {
                    fusecombobox.SelectedItem = hrcTypeMatch.Groups[1].Value;
                    if (fusecombobox.SelectedItem == null)
                    {
                        fusecombobox.SelectedIndex = 0;
                    }
                }
                else
                {
                    fusecombobox.SelectedIndex = 0;
                }
                                

                string getctamps = ConfigurationManager.AppSettings["ctamps"];
                if (getctamps != null)
                {
                    ctamps.SelectedItem = getctamps;
                }
                else
                {
                    ctamps.SelectedIndex = 0;
                }
                string getctva = ConfigurationManager.AppSettings["ctva"];
                if (getctva != null)
                {
                    ctva.SelectedItem = getctva;
                }
                else
                {
                    ctva.SelectedIndex = 0;
                }

                if (label1.Text == "PANEL UTILITY" || label1.Text == "ENCLOSURE AND BUSBAR + EARTH" || label1.Text.Contains("MCB"))
                {

                    Run_Clickalt(SelectedRange, EventArgs.Empty);
                }
                else if(auto == "automate643")
                {
                    Run_Clickalt(SelectedRange, EventArgs.Empty);
                }
                else
                {
                    this.Show();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error on initialize: {ex.Message}");
            }
        }
        public void openasnewform()
        {
            Run.Visible = false;
            Update.Visible = false;
            button1.Visible = true;
            label1.Text = "Enter the Feeder Items:";
            this.Controls.Remove(comboBox1);

            // Create a new TextBox
            
            feeder.Name = "feedhead";
            feeder.Width = 300;
            feeder.Location = comboBox1.Location;

            feeder.TextChanged += new EventHandler(textBox1_TextChanged);

            this.Controls.Add(feeder);
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null)
            {
                string text = textBox.Text.ToUpper();

                // Clear existing items
                checkedListBox1.Items.Clear();

                if (text.Contains("MCCB"))
                {
                    checkedListBox1.Items.Add("Spreaders");
                    checkedListBox1.Items.Add("Extended ROM");
                    checkedListBox1.Items.Add("Aux Contacts");
                    checkedListBox1.Items.Add("Shunt Release");
                }
                else if (text.Contains("ACB"))
                {
                    checkedListBox1.Items.Add("Door interlock");
                    checkedListBox1.Items.Add("C- Lock");
                    checkedListBox1.Items.Add("UV release");
                    checkedListBox1.Items.Add("Closing release");
                    checkedListBox1.Items.Add("Shunt release");
                }
            }
        }
        private void mfmcheckbox_CheckedChanged_1(object sender, EventArgs e)
        {
            if (mfmcheckbox.Checked)
            {
                label4.Visible = true;
                mfmcat.Visible = true;
                label6.Visible = true;
                label7.Visible = true;
                ctamps.Visible = true;
                ctva.Visible = true;
                label8.Visible = true;
            }
            else
            {
                label4.Visible = false;
                mfmcat.Visible = false;
                label6.Visible = false;
                label7.Visible = false;
                ctamps.Visible = false;
                ctva.Visible = false;
                label8.Visible = false;
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

        Excel.Range GetUsedRange(Excel.Worksheet sheet, ref Excel.Range range)
        {
            if (range == null)
            {
                range = sheet.UsedRange;
            }
            return range;
        }

        private void Run_Click(object sender, EventArgs e)
        {
            var excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            // Check if the selected range is valid
            if (selectedRange == null || selectedRange.Cells.Count != 1)
            {
                MessageBox.Show("Please select a single cell.");
                return;
            }

            Run_Clickalt(selectedRange, EventArgs.Empty);
        }

        private void Run_Clickalt(object sender, EventArgs e)
        {
            Excel.Application excelApp = ExcelDnaUtil.Application as Excel.Application;
            Excel.Workbook currentWorkbook = excelApp.ActiveWorkbook;
            Excel.Workbook extWorkbook = null;
            Excel.Range selectedRange = sender as Excel.Range;
            Excel.Worksheet currentSheet = selectedRange.Worksheet;
            try
            {
                bool containsRYB = rybcheckbox.Checked;
                bool containsRGA = rgacheckbox.Checked;
                bool containsMFM = mfmcheckbox.Checked;
                bool containsAM = ambox.Checked;
                bool containsVM = vmbox.Checked;
                bool containsAMSS = asscheckbox.Checked;
                bool containsVMSS = vsscheckbox.Checked;
                bool containsELR = elrbox.Checked;
                bool containsSPD = spdbox.Checked;
                bool containsHRC = hrccheckbox.Checked;
                bool containsREAC = reactorcheckbox.Checked;
                bool checkBoxSpreaders = false;
                bool checkBoxExtendedROM = false;
                bool checkBoxAuxilaryContacts = false;
                bool checkBoxShuntRelease = false;
                bool checkBoxdoorint = false;
                bool checkBoxckey = false;
                bool checkBoxshuntR = false;
                bool checkBoxuvR = false;
                bool checkBoxclosingR = false;
                bool containsAddmccb = false;
                bool containsAddacb = false;
                bool testbox1 = test1.Checked;
                bool testbox2 = test2.Checked;
                string acctype = comboBox1.SelectedItem?.ToString().ToLower();
                string header = selectedRange.Value2.ToString().ToUpper();
                string mfmcatno = null;
                string ctampsvalue = null;
                string ctvavalue = null;
                string hrcamps = null;


                Excel.Range accUsedRange = null;
                Excel.Range onekUsedRange = null;
                Excel.Range twokUsedRange = null;
                Excel.Range threekUsedRange = null;
                Excel.Range fivekUsedRange = null;
                Excel.Range sevenkUsedRange = null;
                Excel.Range mcbUsedRange = null;
                Excel.Range mfmUsedRange = null;
                Excel.Range cosUsedRange = null;
                Excel.Range sdfuUsedRange = null;
                Excel.Range apfcUsedRange = null;

                if (containsMFM)
                {
                    mfmcatno = mfmcat.Text;
                    ctampsvalue = ctamps.SelectedItem?.ToString();
                    ctvavalue = ctva.SelectedItem?.ToString();
                    Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    config.AppSettings.Settings.Remove("MfmCatno");
                    config.AppSettings.Settings.Remove("ctamps");
                    config.AppSettings.Settings.Remove("ctva");
                    config.AppSettings.Settings.Add("MfmCatno", mfmcatno);
                    config.AppSettings.Settings.Add("ctamps", ctampsvalue);
                    config.AppSettings.Settings.Add("ctva", ctvavalue);
                    config.Save(ConfigurationSaveMode.Modified);
                    ConfigurationManager.RefreshSection("appSettings");
                }

                if (containsHRC)
                {
                    hrcamps = fusecombobox.SelectedItem?.ToString();
                }

                if (header.Contains("ACB"))
                {
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (checkedListBox1.Items[i].ToString() == "Door interlock")
                        {
                            checkBoxdoorint = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "C- Lock")
                        {
                            checkBoxckey = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "UV release")
                        {
                            checkBoxshuntR = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "Closing release")
                        {
                            checkBoxuvR = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "Shunt release")
                        {
                            checkBoxclosingR = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "Additional item")
                        {
                            containsAddacb = checkedListBox1.GetItemChecked(i);
                        }
                    }
                }
                if (header.Contains("MCCB"))
                {
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (checkedListBox1.Items[i].ToString() == "Spreaders")
                        {
                            checkBoxSpreaders = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "Extended ROM")
                        {
                            checkBoxExtendedROM = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "Aux Contacts")
                        {
                            checkBoxAuxilaryContacts = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "Shunt Release")
                        {
                            checkBoxShuntRelease = checkedListBox1.GetItemChecked(i);
                        }
                        else if (checkedListBox1.Items[i].ToString() == "Additional item")
                        {
                            containsAddmccb = checkedListBox1.GetItemChecked(i);
                        }
                    }
                }


                // Define the path to the database file
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
                excelApp.ErrorCheckingOptions.BackgroundChecking = false;  // Disable background error checking
                

                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    if (wb.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase))
                    {
                        extWorkbook = wb;
                        break;
                    }
                }
                if (extWorkbook == null)
                {
                    extWorkbook = excelApp.Workbooks.Open(extFilePath, false, false);

                }

                Excel.Worksheet mainSheet = extWorkbook.Sheets["MAIN"];
                Excel.Range mainUsedRange = mainSheet.UsedRange;
                Excel.Worksheet accSheet = extWorkbook.Sheets["MCCB ACC"];
                accUsedRange = GetUsedRange(accSheet, ref accUsedRange);

                string amps = "";

                Match ampsMatch = Regex.Match(header, @"\d+ ?A", RegexOptions.IgnoreCase);
                if (ampsMatch.Success)
                {
                    amps = ampsMatch.Value;
                }

                string panelqty = "1";
                string feederqty = "1";
                bool foundqty = false;
                string bartype = "AL";
                Excel.Range fullRange = selectedRange.Worksheet.UsedRange;
                for (int row = selectedRange.Row - 1; row >= 1; row--)
                {
                    Excel.Range cell = fullRange.Cells[row, 2];
                    if (cell.Interior.Color == 15773696)
                    {
                        string sum = cell.Offset[0, 1].Value2.ToString();
                        panelqty = sum;
                        foundqty = true;
                        try
                        {
                            string bar = cell.Offset[0, -1].Value2.ToString().ToLower();
                            if(bar == "al")
                            {
                                bartype = "AL";
                            }
                            else if (bar == "cu")
                            {
                                bartype = "CU";
                            }
                            else
                            {
                                MessageBox.Show("Can't find BUSBAR material. So keeping the default Value as ALUMINIUM");
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Can't find BUSBAR material. So keeping the default Value as ALUMINIUM");
                        }
                        break;
                    }

                }

                List<Excel.Range> rowsToCopy = new List<Excel.Range>();

                //inserting main componentas and accessories

                if (header.Contains("UTILITY"))
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                        if (cellValue == "PU")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }
                else if (header.Contains("ENCLOSURE"))
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                        if (cellValue == "ELC" || cellValue == "TOTAL")
                        {
                            rowsToCopy.Add(row);
                        }

                    }
                }
                else if (header.Contains("SDFU"))
                {
                    Excel.Worksheet sdfuSheet = extWorkbook.Sheets["SDFU"];
                    sdfuUsedRange = GetUsedRange(sdfuSheet, ref sdfuUsedRange);
                    string poleType = "";
                    string hrc = "";
                    Match poleTypeMatch = Regex.Match(header, @"[SDTF1234]P[N]?", RegexOptions.IgnoreCase);
                    if (poleTypeMatch.Success)
                    {
                        switch (poleTypeMatch.Value)
                        {
                            case "TPN":
                                poleType = "TPN";
                                break;
                            case "TP":
                                poleType = "TPN";
                                break;
                            case "3P":
                                poleType = "TPN";
                                break;
                            case "FP":
                                poleType = "4P";
                                break;
                            case "4P":
                                poleType = "4P";
                                break;
                            case "2P":
                                poleType = "2P";
                                break;
                            case "DP":
                                poleType = "2P";
                                break;
                        }
                    }
                    foreach (Excel.Range row in sdfuUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                        string cellValue2 = row.Cells[1, 28].Value2?.ToString().ToUpper() ?? "";
                        string cellValue3 = row.Cells[1, 29].Value2?.ToString().ToUpper() ?? "";

                        if (cellValue == "SDFU" && cellValue2 == amps && cellValue3 == poleType)
                        {
                            hrc = row.Cells[1, 30].Value2.ToString();
                            rowsToCopy.Add(row);
                        }

                    }

                    if (hrc != "")
                    {
                        foreach (Excel.Range row in sdfuUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                            string cellValue2 = row.Cells[1, 28].Value2?.ToString().ToUpper() ?? "";
                            string cellValue3 = row.Cells[1, 29].Value2?.ToString() ?? "";

                            if (cellValue == "HRC" && cellValue2 == amps && cellValue3 == hrc)
                            {
                                rowsToCopy.Add(row);
                            }

                        }
                    }
                }
                else if (header.Contains("COS"))
                {
                    Excel.Worksheet cosSheet = extWorkbook.Sheets["COS"];
                    cosUsedRange = GetUsedRange(cosSheet, ref cosUsedRange);

                    string type = "COSM";
                    if (header.Contains("MOT"))
                    {
                        type = "COSA";
                    }
                    int amps1 = 0;

                    Match ampsMatch1 = Regex.Match(header, @"(\d+) ?A", RegexOptions.IgnoreCase);
                    if (ampsMatch1.Success)
                    {
                        amps1 = int.Parse(ampsMatch1.Groups[1].Value);
                    }
                    foreach (Excel.Range row in cosUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        string rating = row.Cells[1, 28].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == type.ToLower() && rating == amps1.ToString())
                        {
                            rowsToCopy.Add(row);
                        }
                    }

                }
                else if (header.Contains("ACB"))
                {
                    Excel.Worksheet acbSheet = extWorkbook.Sheets["ACB ACC"];
                    Excel.Range acbUsedRange = acbSheet.UsedRange;

                    bool edraw = false;
                    bool efix = false;
                    bool mdraw = false;
                    bool mfix = false;
                    string poleType = null;
                    string frame = null;
                    string breaking = null;
                    string generatedcat = null;
                    string rating = null;
                    string mtxtype = "9";

                    Match poleTypeMatch = Regex.Match(header, @"[SDTF1234]P[N]?", RegexOptions.IgnoreCase);
                    if (poleTypeMatch.Success)
                    {
                        switch (poleTypeMatch.Value)
                        {
                            case "TPN":
                                poleType = "X";
                                break;
                            case "TP":
                                poleType = "X";
                                break;
                            case "3P":
                                poleType = "X";
                                break;
                            case "FP":
                                poleType = "F";
                                break;
                            case "4P":
                                poleType = "F";
                                break;
                        }
                    }

                    if (header.Contains("EDO")) 
                    {
                        edraw = true;
                    }
                    else if (header.Contains("MDO"))
                    {
                        mdraw = true;
                    }
                    else if (header.Contains("MF"))
                    {
                        mfix = true;
                    }
                    else if (header.Contains("EF"))
                    {
                        efix= true;
                    }
                    string mtx = null;
                    Match mtxMatch = Regex.Match(header, @"(MTX \d+(\.\d+)?[A-Za-z]*)", RegexOptions.IgnoreCase);
                    if (mtxMatch.Success)
                    {
                        mtx = mtxMatch.Groups[1].Value.ToString().ToUpper(); // Capture the value after MTX
                                                                    // Store or process mtxValue as needed
                    }
                    if(mtx != null)
                    {
                        switch (mtx)
                        {
                            case "MTX 1.0":
                                mtxtype = "7";
                                break;
                            case "MTX 1G":
                                mtxtype = "8";
                                break;
                            case "MTX 1GI":
                                mtxtype = "B";
                                break;
                            case "MTX 1.5G":
                                mtxtype = "9";
                                break;
                            case "MTX 1.5GI":
                                mtxtype = "C";
                                break;
                            case "MTX 3.5":
                                mtxtype = "3";
                                break;
                            case "MTX 3.5EC":
                                mtxtype = "4";
                                break;
                            case "MTX 4.5":
                                mtxtype = "5";
                                break;
                            case "MTX 3.5H":
                                mtxtype = "6";
                                break;
                        }
                    }
                    

                    if (acctype != null && acctype == "auto")
                    {
                        Match breakingCapacityMatch = Regex.Match(header, @"\d+ ?KA", RegexOptions.IgnoreCase);
                        if (breakingCapacityMatch.Success)
                        {
                            acctype = breakingCapacityMatch.Value.ToLower();
                        }
                    }
 
                    if (acctype != null && acctype == "50ka")
                    {
                        breaking = "N";
                    }
                    else if (acctype != null && acctype == "65ka")
                    {
                        breaking = "S";
                    }
                    else if (acctype != null && acctype == "80ka")
                    {
                        breaking = "H";
                    }
                    else if (acctype != null && acctype == "100ka")
                    {
                        breaking = "V";
                    }

                    if(amps == "630A")
                    {
                        frame = "1";
                        rating = "06";
                    }
                    else if(amps == "400A")
                    {
                        frame = "1";
                        rating = "04";
                    }
                    else if (amps == "800A")
                    {
                        frame = "1";
                        rating = "08";
                    }
                    else if (amps == "1000A")
                    {
                        frame = "1";
                        rating = "10";

                    }
                    else if (amps == "1250A")
                    {
                        frame = "1";
                        rating = "12";

                    }
                    else if (amps == "1600A")
                    {
                        frame = "1";
                        rating = "16";

                    }
                    else if (amps == "2000A")
                    {
                        frame = "1";
                        rating = "20";

                    }
                    else if (amps == "2500A")
                    {
                        frame = "1";
                        rating = "25";

                    }
                    else if (amps == "3200A")
                    {
                        frame = "3";
                        rating = "32";

                    }
                    else if (amps == "4000A")
                    {
                        frame = "3";
                        rating = "40";

                    }
                    else if (amps == "5000A")
                    {
                        frame = "3";
                        rating = "50";

                    }
                    else if (amps == "6300A")
                    {
                        frame = "3";
                        rating = "63";

                    }
                    string D = null;
                    if (edraw || mdraw)
                    {
                        D = "D";
                    }
                    else
                    {
                        D = "F";
                    }
                    string type = null;
                    string controlvolt = null;
                    if (edraw || efix)
                    {
                        type = "2";
                        controlvolt = "1";
                    }
                    else 
                    {
                        type = "1";
                        controlvolt = "0";
                    }
                    try 
                    {
                        generatedcat = $"UW{frame}{rating}{breaking}{poleType}{D}{controlvolt}{type}{mtxtype}00";
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                            if (cellValue == "acb")
                            {
                                row.Cells[1, 3].Value2 = generatedcat;
                                rowsToCopy.Add(row);
                            }
                        }
                    }
                    catch
                    {
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                            if (cellValue == "acb")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                    if (checkBoxdoorint)
                    {
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            if (cellValue == "Door interlock")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                    if (checkBoxckey)
                    {
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            if (cellValue == "C- Lock")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                    if (checkBoxuvR)
                    {
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            if (cellValue == "UV release")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                    if (checkBoxshuntR)
                    {
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            if (cellValue == "Shunt release")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                    if (checkBoxclosingR)
                    {
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            if (cellValue == "Closing release")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                    if (containsAddacb)
                    {
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            if (cellValue == "Additonal item")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                    if (edraw || efix)
                    {
                        foreach (Excel.Range row in acbUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            if (cellValue == "TNC")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                }
                else if (header.Contains("MCCB"))
                {

                    string breakingCapacity = "";
                    string poleType = "";
                    string mtxType = "";

                    // Extract Breaking Capacity
                    Match breakingCapacityMatch = Regex.Match(header, @"\d+ ?KA", RegexOptions.IgnoreCase);
                    if (breakingCapacityMatch.Success)
                    {
                        breakingCapacity = breakingCapacityMatch.Value;
                    }

                    // Extract Pole Type (FP or TP)
                    Match poleTypeMatch = Regex.Match(header, @"[SDTF1234]P[N]?", RegexOptions.IgnoreCase);
                    if (poleTypeMatch.Success)
                    {
                        switch (poleTypeMatch.Value)
                        {
                            case "TPN":
                                poleType = "TPN";
                                break;
                            case "TP":
                                poleType = "TPN";
                                break;
                            case "3P":
                                poleType = "TPN";
                                break;
                            case "FP":
                                poleType = "4P";
                                break;
                            case "4P":
                                poleType = "4P";
                                break;
                        }
                    }

                    // Extract MTX Type
                    Match mtxTypeMatch = Regex.Match(header, @"MTX \d+\.\d", RegexOptions.IgnoreCase);
                    if (mtxTypeMatch.Success)
                    {
                        mtxType = mtxTypeMatch.Value.Trim();
                    }

                    if (mtxType != "")
                    {
                        if (breakingCapacity == "18KA")
                        {
                            breakingCapacity = "36KA";
                        }
                        else if (breakingCapacity == "25KA")
                        {
                            breakingCapacity = "36KA";
                        }
                    }

                    if (breakingCapacity == "18KA")
                    {
                        Excel.Worksheet onekSheet = extWorkbook.Sheets["MCCB 18KA"];
                        onekUsedRange = GetUsedRange(onekSheet, ref onekUsedRange);

                        foreach (Excel.Range row in onekUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                            string cellValue2 = row.Cells[1, 28].Value2?.ToString().ToUpper() ?? "";
                            if (cellValue == amps && cellValue2 == poleType)
                            {
                                if (acctype == "auto")
                                {
                                    acctype = row.Cells[1, 29].Value2.ToString();
                                }

                                rowsToCopy.Add(row);
                            }
                        }
                    }
                    else if (breakingCapacity == "25KA")
                    {
                        Excel.Worksheet twokSheet = extWorkbook.Sheets["MCCB 25KA"];
                        twokUsedRange = GetUsedRange(twokSheet, ref twokUsedRange);

                        foreach (Excel.Range row in twokUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                            if (cellValue == $"{amps} {poleType}")
                            {
                                if (acctype == "auto")
                                {
                                    acctype = row.Cells[1, 28].Value2.ToString();
                                }

                                rowsToCopy.Add(row);
                            }
                        }
                    }
                    else if (breakingCapacity == "36KA")
                    {
                        Excel.Worksheet threekSheet = extWorkbook.Sheets["MCCB 36KA"];
                        threekUsedRange = GetUsedRange(threekSheet, ref threekUsedRange);

                        foreach (Excel.Range row in threekUsedRange.Rows)
                        {
                            string cellValue1 = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                            string cellValue2 = row.Cells[1, 28].Value2?.ToString().ToUpper() ?? "";
                            string cellValue3 = row.Cells[1, 29].Value2?.ToString().ToUpper() ?? "";

                            if (mtxType != "")
                            {
                                if (cellValue1 == amps && cellValue2 == poleType && cellValue3 == mtxType)
                                {
                                    if (acctype == "auto")
                                    {
                                        acctype = row.Cells[1, 30].Value2.ToString();
                                    }

                                    rowsToCopy.Add(row);
                                }
                            }
                            else
                            {
                                if (cellValue1 == amps && cellValue2 == poleType && cellValue3 == "")
                                {
                                    if (acctype == "auto")
                                    {
                                        acctype = row.Cells[1, 30].Value2.ToString();
                                    }

                                    rowsToCopy.Add(row);
                                }
                            }
                        }
                    }
                    else if (breakingCapacity == "50KA")
                    {
                        Excel.Worksheet fivekSheet = extWorkbook.Sheets["MCCB 50KA"];
                        fivekUsedRange = GetUsedRange(fivekSheet, ref fivekUsedRange);

                        foreach (Excel.Range row in fivekUsedRange.Rows)
                        {
                            string cellValue1 = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                            string cellValue2 = row.Cells[1, 28].Value2?.ToString().ToUpper() ?? "";
                            string cellValue3 = row.Cells[1, 29].Value2?.ToString().ToUpper() ?? "";

                            if (mtxType != "")
                            {
                                if (cellValue1 == amps && cellValue2 == poleType && cellValue3 == mtxType)
                                {
                                    if (acctype == "auto")
                                    {
                                        acctype = row.Cells[1, 30].Value2.ToString();
                                    }

                                    rowsToCopy.Add(row);
                                }
                            }
                            else
                            {
                                if (cellValue1 == amps && cellValue2 == poleType && cellValue3 == "")
                                {
                                    if (acctype == "auto")
                                    {
                                        acctype = row.Cells[1, 30].Value2.ToString();
                                    }

                                    rowsToCopy.Add(row);
                                }
                            }

                        }
                    }
                    else if (breakingCapacity == "70KA")
                    {
                        Excel.Worksheet sevenkSheet = extWorkbook.Sheets["MCCB 70KA"];
                        sevenkUsedRange = GetUsedRange(sevenkSheet, ref sevenkUsedRange);

                        foreach (Excel.Range row in sevenkUsedRange.Rows)
                        {
                            string cellValue1 = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                            string cellValue2 = row.Cells[1, 28].Value2?.ToString().ToUpper() ?? "";
                            string cellValue3 = row.Cells[1, 29].Value2?.ToString().ToUpper() ?? "";

                            if (mtxType != "")
                            {
                                if (cellValue1 == amps && cellValue2 == poleType && cellValue3 == mtxType)
                                {
                                    if (acctype == "auto")
                                    {
                                        acctype = row.Cells[1, 30].Value2.ToString();
                                    }

                                    rowsToCopy.Add(row);
                                }
                            }
                            else
                            {
                                if (cellValue1 == amps && cellValue2 == poleType && cellValue3 == "")
                                {
                                    if (acctype == "auto")
                                    {
                                        acctype = row.Cells[1, 30].Value2.ToString();
                                    }

                                    rowsToCopy.Add(row);
                                }
                            }

                        }
                    }

                    if (acctype != null)
                    {
                        if (checkBoxSpreaders)
                        {
                            foreach (Excel.Range row in accUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";

                                if (amps == "400A")
                                {
                                    if (cellValue == $"{acctype.ToLower()} spr400 {poleType.ToLower()}")
                                    {
                                        rowsToCopy.Add(row);
                                    }
                                }
                                else
                                {
                                    if (cellValue == $"{acctype.ToLower()} spr {poleType.ToLower()}")
                                    {
                                        rowsToCopy.Add(row);
                                    }
                                }
                                
                            }
                        }

                        if (checkBoxExtendedROM)
                        {
                            foreach (Excel.Range row in accUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                                if (cellValue == $"{acctype.ToLower()} rom")
                                {
                                    rowsToCopy.Add(row);
                                }
                            }
                        }

                        if (checkBoxAuxilaryContacts)
                        {
                            foreach (Excel.Range row in accUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                                if (cellValue == $"{acctype.ToLower()} aux")
                                {
                                    rowsToCopy.Add(row);
                                }
                            }
                        }

                        if (checkBoxShuntRelease)
                        {
                            foreach (Excel.Range row in accUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                                if (cellValue == $"{acctype.ToLower()} shunt")
                                {
                                    rowsToCopy.Add(row);
                                }
                            }
                        }

                        if (containsAddmccb)
                        {
                            foreach (Excel.Range row in accUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                                if (cellValue == $"{acctype.ToLower()} add")
                                {
                                    rowsToCopy.Add(row);
                                }
                            }
                        }
                    }


                }
                else if (header.Contains("MCB"))
                {

                    Excel.Worksheet mcbSheet = extWorkbook.Sheets["MCB"];
                    mcbUsedRange = GetUsedRange(mcbSheet, ref mcbUsedRange);

                    string poleType = "";
                    Match poleTypeMatch = Regex.Match(header, @"[SDTF1234]P[N]?", RegexOptions.IgnoreCase);
                    if (poleTypeMatch.Success)
                    {
                        switch (poleTypeMatch.Value)
                        {
                            case "SP":
                                poleType = "SP";
                                break;
                            case "1P":
                                poleType = "SP";
                                break;
                            case "DP":
                                poleType = "DP";
                                break;
                            case "2P":
                                poleType = "DP";
                                break;
                            case "TPN":
                                poleType = "TP";
                                break;
                            case "TP":
                                poleType = "TP";
                                break;
                            case "3P":
                                poleType = "TP";
                                break;
                            case "FP":
                                poleType = "FP";
                                break;
                            case "4P":
                                poleType = "FP";
                                break;
                        }
                    }

                    foreach (Excel.Range row in mcbUsedRange.Rows)
                    {
                        string cellValue1 = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                        string cellValue2 = row.Cells[1, 28].Value2?.ToString().ToUpper() ?? "";
                        if (cellValue1 == amps && cellValue2 == poleType)
                        {
                            rowsToCopy.Add(row);
                        }
                    }

                    foreach (Excel.Range row in mcbUsedRange.Rows)
                    {
                        string cellValue1 = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";

                        if (cellValue1 == "MCB ACC")
                        {
                            rowsToCopy.Add(row);
                        }
                    }



                }


                if (containsRGA && containsRYB)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "ryb,rga")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }
                else if (containsRYB)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "ryb")
                        {
                            rowsToCopy.Add(row);
                        }
                    }

                }
                else if (containsRGA)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "rga")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                if (header.Contains("KVAR"))
                {
                    Excel.Worksheet apfcSheet = extWorkbook.Sheets["APFC"];
                    apfcUsedRange = GetUsedRange(apfcSheet, ref apfcUsedRange);

                    string kvar = "";

                    Match kvarMatch = Regex.Match(header, @"\d+\s*KVAR", RegexOptions.IgnoreCase);
                    if (kvarMatch.Success)
                    {
                        kvar = kvarMatch.Value;
                    }

                    foreach (Excel.Range row in apfcUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                        string cellValue2 = row.Cells[1, 28].Value2?.ToString() ?? "";

                        if (cellValue == kvar && (cellValue2 == "CO" || cellValue2 == "CP"))
                        {
                            rowsToCopy.Add(row);

                        }
                    }

                    if (containsREAC)
                    {
                        if (bartype == "CU")
                        {
                            foreach (Excel.Range row in apfcUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                                string cellValue2 = row.Cells[1, 28].Value2?.ToString() ?? "";

                                if (cellValue == kvar && cellValue2 == "CUR")
                                {
                                    rowsToCopy.Add(row);

                                }
                            }
                        }
                        else if (bartype == "AL")
                        {
                            foreach (Excel.Range row in apfcUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                                string cellValue2 = row.Cells[1, 28].Value2?.ToString() ?? "";

                                if (cellValue == kvar && cellValue2 == "ALR")
                                {
                                    rowsToCopy.Add(row);

                                }
                            }
                        }
                    }

                    foreach (Excel.Range row in apfcUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";

                        if (cellValue == "APFC ACC")
                        {
                            rowsToCopy.Add(row);

                        }
                    }

                }

                if (containsMFM)
                {
                    Excel.Worksheet mfmSheet = extWorkbook.Sheets["MFM"];
                    mfmUsedRange = GetUsedRange(mfmSheet, ref mfmUsedRange);

                    if (mfmcatno != null)
                    {
                        bool found = false;
                        if(mfmcatno != "")
                        {
                            foreach (Excel.Range row in mfmUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 3].Value2?.ToString() ?? "";
                                if (cellValue == mfmcatno)
                                {
                                    rowsToCopy.Add(row);
                                    found = true;
                                }
                            }
                        }

                        if (!found) 
                        {
                            foreach (Excel.Range row in mfmUsedRange.Rows)
                            {
                                string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                                if (cellValue == "mfm")
                                {
                                    rowsToCopy.Add(row);
                                    
                                }
                            }
                        }
                    }

                    foreach (Excel.Range row in mfmUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToUpper() ?? "";
                        if (cellValue == "CT")
                        {
                            string cellamps = row.Cells[1, 28].Value2?.ToString().ToUpper() ?? "";
                            string cellclass = row.Cells[1, 29].Value2?.ToString().ToUpper() ?? "";
                            string cellva = row.Cells[1, 30].Value2?.ToString().ToUpper() ?? "";
                            if (cellamps == amps && cellclass == ctampsvalue && cellva == ctvavalue)
                            {
                                rowsToCopy.Add(row);
                            }

                        }
                    }

                    foreach (Excel.Range row in mfmUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "mfm acc")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                if (containsRYB || containsRGA || containsMFM)
                {
                    string count = "1";
                    if (containsRYB && containsRGA && containsMFM) { count = "4"; }
                    else if (containsRYB && containsRGA) { count = "4"; }
                    else if (containsRGA && containsMFM) { count = "4"; }
                    else if (containsRYB && containsMFM) { count = "3"; }
                    else if (containsRYB) { count = "3"; }
                    else if (containsMFM) { count = "3"; }
                    else if (containsRGA) { count = "1"; }

                    if (containsHRC && hrcamps == "Control")
                    {
                        foreach (Excel.Range row in mainUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                            if (cellValue == "hrc6")
                            {
                                row.Cells[1, 5].Value2 = count;
                                rowsToCopy.Add(row);
                            }
                        }
                    }
                    else
                    {
                        foreach (Excel.Range row in mainUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                            if (cellValue == "mcb")
                            {
                                row.Cells[1, 5].Value2 = count;
                                rowsToCopy.Add(row);
                            }
                        }
                    }
                }

                if (containsHRC)
                {
                    if (hrcamps != "Control")
                    {
                        Excel.Worksheet sdfuSheet = extWorkbook.Sheets["SDFU"];
                        sdfuUsedRange = GetUsedRange(sdfuSheet, ref sdfuUsedRange);

                        string HRCamps = $"{hrcamps}A";
                        foreach (Excel.Range row in sdfuUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            string cellValue2 = row.Cells[1, 29].Value2?.ToString() ?? "";
                            if (cellValue == "HRC" && cellValue2 == HRCamps)
                            {
                                rowsToCopy.Add(row);
                            }
                        }

                        int HRCampsint = int.Parse(hrcamps);
                        string basetype = null;

                        if (HRCampsint > 250)
                        {
                            basetype = "HB 800";
                        }
                        else if (HRCampsint > 160)
                        {
                            basetype = "HB 250";
                        }
                        else if (HRCampsint > 63)
                        {
                            basetype = "HB 160";
                        }
                        else
                        {
                            basetype = "HB 63";
                        }

                        foreach (Excel.Range row in sdfuUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString() ?? "";
                            string cellValue2 = row.Cells[1, 29].Value2?.ToString() ?? "";
                            if (cellValue == "HRC" && cellValue2 == basetype)
                            {
                                rowsToCopy.Add(row);
                            }
                        }

                    }
                }

                if (containsAM)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "am")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }
                if (containsAMSS)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "ass")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                if (containsVM)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "vm")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }
                if (containsVMSS)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "vss")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                if (containsELR)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "elr")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                if (containsSPD)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "spd")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                if (testbox1)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "test1")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                if (testbox2)
                {
                    foreach (Excel.Range row in mainUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (cellValue == "test2")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                //inserting busbar interconnection and consumebles

                if (header.Contains("MCCB") || header.Contains("COS") || header.Contains("SDFU"))
                {
                    int amps1 = 0;

                    Match ampsMatch1 = Regex.Match(header, @"(\d+) ?A", RegexOptions.IgnoreCase);
                    if (ampsMatch1.Success)
                    {
                        amps1 = int.Parse(ampsMatch1.Groups[1].Value);
                    }

                    if (amps1 != 0 && amps1 <= 63)
                    {
                        foreach (Excel.Range row in accUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";

                            if(bartype == "CU")
                            {
                                if (cellValue == "63cu")
                                {
                                    rowsToCopy.Add(row);
                                }
                            }
                            else if (bartype == "AL")
                            {
                                if (cellValue == "63")
                                {
                                    rowsToCopy.Add(row);
                                }
                            }
                            
                        }
                    }
                    else if (amps1 > 66)
                    {
                        foreach (Excel.Range row in accUsedRange.Rows)
                        {
                            string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                            

                            if (bartype == "CU")
                            {
                                if (cellValue == $"{amps1.ToString()}cu")
                                {
                                    rowsToCopy.Add(row);
                                }
                            }
                            else if (bartype == "AL")
                            {
                                if (cellValue == amps1.ToString())
                                {
                                    rowsToCopy.Add(row);
                                }
                            }
                        }
                    }

                }else if (header.Contains("ACB"))
                {
                    Excel.Worksheet acbSheet = extWorkbook.Sheets["ACB ACC"];
                    Excel.Range acbUsedRange = acbSheet.UsedRange;

                    int amps1 = 0;

                    Match ampsMatch1 = Regex.Match(header, @"(\d+) ?A", RegexOptions.IgnoreCase);
                    if (ampsMatch1.Success)
                    {
                        amps1 = int.Parse(ampsMatch1.Groups[1].Value);
                    }

                    foreach (Excel.Range row in acbUsedRange.Rows)
                    {
                        string cellValue = row.Cells[1, 27].Value2?.ToString().ToLower() ?? "";
                        if (bartype == "CU")
                        {
                            if (cellValue == $"{amps1.ToString()}cu")
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                        else if (bartype == "AL")
                        {
                            if (cellValue == amps1.ToString())
                            {
                                rowsToCopy.Add(row);
                            }
                        }
                    }

                }


                if (rowsToCopy.Count == 0)
                {
                    return;
                }

                // Copy and insert the rows below the selected cell
                for (int i = rowsToCopy.Count - 1; i >= 0; i--)
                {
                    Excel.Range row = rowsToCopy[i];
                    row.Copy();
                    selectedRange.Offset[1, -1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                }
                
                if (header.Contains("ENCLOSURE"))
                {
                    selectedRange.Value2 = "ENCLOSURE AND BUSBAR + EARTH";
                    selectedRange.Interior.Color = 49407;
                    selectedRange.Font.Bold = true;

                    Excel.Range totalprice = selectedRange.Offset[3, 8];
                    string totalrpiceref = totalprice.Address[false, false];
                    int totalpriceRow = totalprice.Row;
                    string foundedcell = null;
                    string foundedhead = null;
                    Excel.Range usedRange = selectedRange.Worksheet.UsedRange;

                    for (int row = totalpriceRow - 1; row >= 1; row--)
                    {
                        Excel.Range cell = usedRange.Cells[row, 2];
                        if (cell.Interior.Color == 15773696) 
                        {
                            foundedhead = cell.Value2.ToString();
                            foundedcell = cell.Offset[0, 8].Address[false, false]; // Append the cell address to the formula
                            break;
                        }
                        
                    }

                    if(foundedcell == null)
                    {
                        MessageBox.Show("NO PANEL HEADING FOUND, CHECK FOR COLOR MISMATCH");
                        return;
                    } 

                    string sumFormula = $"=SUM({foundedcell}:J{(totalpriceRow - 1).ToString()})";
                    totalprice.Formula = sumFormula;
                    totalprice.NumberFormat = "0";
                    Excel.Worksheet titlesheet = null;
                    foreach (Excel.Worksheet wb in currentWorkbook.Worksheets)
                    {
                        if (wb.Name.Equals("TITLE", StringComparison.OrdinalIgnoreCase))
                        {
                            titlesheet = wb;
                            break;
                        }
                    }

                    if (titlesheet != null)
                    {
                        Excel.Range titlerange = titlesheet.UsedRange;
                        foreach (Excel.Range row in titlerange.Rows)
                        {
                            string cellValue = row.Cells[1, 2].Value2?.ToString() ?? "";
                            if (cellValue == foundedhead)
                            {
                                row.Cells[1, 4].Formula = $"=COSTING!{totalrpiceref}";
                                row.Cells[1, 5].Formula = $"={row.Cells[1,4].Address[false, false]}*$F$7";
                                row.Cells[1, 6].Formula = $"={row.Cells[1,4].Address[false, false]}+{row.Cells[1,5].Address[false, false]}";
                                row.Cells[1, 7].Formula = $"=ROUNDUP({row.Cells[1, 6].Address[false, false]},-3)";
                                row.Cells[1, 9].Formula = $"={row.Cells[1, 7].Address[false, false]}*{row.Cells[1, 8].Address[false, false]}";

                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("TITLE sheet not found.");
                    }
                }

                //changing the names in feeder heading
                if (header.Contains("MCCB"))
                {
                    selectedRange.Interior.Color = 49407;
                    selectedRange.Font.Bold = true;
                    string feederhead = selectedRange.Value2.ToString();
                    string pattern = @"MTX \d+\.\d";
                    string newhead = Regex.Replace(feederhead, pattern, "MP", RegexOptions.IgnoreCase);
                    selectedRange.Value2 = newhead;
                    selectedRange.Offset[0,30].Value2 = feederhead;
                }else if (header.Contains("ACB"))
                {
                    selectedRange.Interior.Color = 49407;
                    selectedRange.Font.Bold = true;
                    string feederhead = selectedRange.Value2.ToString();
                    string pattern = @"MTX \d+(\.\d+)?\w*";
                    string newhead = Regex.Replace(feederhead, pattern, "", RegexOptions.IgnoreCase);
                    selectedRange.Value2 = newhead;
                    selectedRange.Offset[0, 30].Value2 = feederhead;
                }

                
                try
                {
                    feederqty = selectedRange.Offset[0, 1].Value2.ToString();
                }
                catch
                {
                    MessageBox.Show("No Feeder Quantity found. So keeping the default Value");
                }
                if (!foundqty) 
                {
                    MessageBox.Show("No Panel Quantity found. So keeping the default Value");
                }

                for (int i = 1; i <= rowsToCopy.Count; i++)
                {
                    if (selectedRange.Offset[i, 5].HasFormula) 
                    {
                        string discformula = selectedRange.Offset[i, 5].Formula.ToString();
                        string changedformula = discformula.Replace("$C", "$B");
                        selectedRange.Offset[i,23].Formula = changedformula;
                    }

                    if (feederqty != null) 
                    {
                        selectedRange.Offset[i,7].Value2 = feederqty;
                    }
                    if (panelqty != null)
                    {
                        selectedRange.Offset[i, 9].Value2 = panelqty;
                    }
                }
                Marshal.ReleaseComObject(extWorkbook);
            }
            catch (Exception ex)
            {
                
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {

                excelApp.DisplayAlerts = true;  // Disable alerts
                excelApp.ErrorCheckingOptions.BackgroundChecking = true;
                excelApp.ScreenUpdating = true;
                currentWorkbook.Activate();
                currentSheet.Activate();

                Marshal.ReleaseComObject(excelApp);
                Marshal.ReleaseComObject(currentWorkbook);
                
                this.Close();

            }


        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            if (feeder.Text != "") 
            {
                FormData formData = new FormData
                {
                    FeederName = feeder.Text.ToString(),
                    containsELR = elrbox.Checked,
                    containsSPD = spdbox.Checked,
                    containsMFM = mfmcheckbox.Checked,
                    containsRYB = rybcheckbox.Checked,
                    containsRGA = rgacheckbox.Checked,
                    containsVM = vmbox.Checked,
                    containsAM = ambox.Checked,
                    containsTEST1 = test1.Checked,
                    containsTEST2 = test2.Checked,
                    SelectedCheckboxes = new List<string>()
                };

                foreach (var item in checkedListBox1.CheckedItems)
                {
                    formData.SelectedCheckboxes.Add(item.ToString());
                }

                // Trigger the event and pass the formData object
                OnFeederDataEntered?.Invoke(formData);

                // Optionally, close the form after sending the value
                this.Close();
            }
            else
            {
                labelred.Visible = true;
            }
            

        }
    }
}
