using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using FlaUI.UIA3.Patterns;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
//using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Ink;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = FlaUI.Core.Application;
using Label = System.Windows.Forms.Label;
using Menu = FlaUI.Core.AutomationElements.Menu;
using MenuItem = FlaUI.Core.AutomationElements.MenuItem;


namespace BurnInTestValidate
{
   
    public partial class FrmBurnIntest : Form
    {
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.TextBox txtCustomer;
        private System.Windows.Forms.Timer barcodeTimer = new System.Windows.Forms.Timer();
        public System.Windows.Forms.Label lblCount;
        public System.Windows.Forms.Label lblFGValue;
        public System.Windows.Forms.Label lblPCBValue;
        public System.Windows.Forms.Label lblFailValue;
        public RichTextBox _rtbLog;
        public System.Windows.Forms.Label lblFG;
        public System.Windows.Forms.Label lblserialNoPCBA;
        public IUserService _userService;
        PassmarkHistory payhistory = new PassmarkHistory();
        string exePath = string.Empty;
        public string[] stages = { "", "" };
        int PartitionCount = 0;
        popupform ppfrm = new popupform();
        Blink frmblink = new Blink();
        //  Autopopclose appfrm = new Autopopclose();
        // BlinkPopupForm popup = new BlinkPopupForm(6000);

        //private readonly string _customer;
        //private readonly int _productTypeId;
        //private readonly string _productTypeName;
        //private readonly string _fgName;
        public string _customer;
        public int _productTypeId;
        public string  _productTypeName;
        public string _fgName;
        public string _pcbSno;
        public UserSession _session;
        public FgDetails _fgDetails;
        public string filenameNow=string.Empty;
        public int stageid = 0;
        public string stageName = string.Empty;
        //public FrmBurnIntest(IUserService userService,UserSession userSession, string customer,
        //string productTypeName, int productTypeId,
        //string fgName, FgDetails fgDetails)
        //{
        //    InitializeComponent();
        //    _userService = userService;
        //    _customer = customer;
        //    _productTypeName = productTypeName;
        //    _productTypeId = productTypeId;
        //    _fgName = fgName;
        //    _session = userSession;
        //    SetupUI();
        //    _fgDetails = fgDetails;
        //}
        public FrmBurnIntest(IUserService userService)        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
            _userService = userService;
            SetupUI();
        }
        private void SetupUI()
        {
            // this.Text = "Burn In Test Automator";
            // this.Size = new System.Drawing.Size(700, 500);
            // this.StartPosition = FormStartPosition.CenterParent;

            // Panel infoPanel = new Panel
            // {
            //     Dock = DockStyle.Top,
            //     Height = 450,
            //     Padding = new Padding(15),
            //     BackColor = Color.WhiteSmoke
            // };
            // //Label lblUser = CreateInfoLabel($"User : {_session.UserName}", 0);
            // lblCount = CreateInfoLabel("FailCount", 5);
            // txtCustomer = CreateInfoTextBox("", 15);
            // Label lblProduct = CreateInfoLabel(
            //     $"Product Type :  M.2 ", 40);
            //  lblFG = CreateInfoLabel($"FG Name : ", 60);
            //  lblserialNoPCBA = CreateInfoLabel($"PCBA ID : ", 80);

            //// infoPanel.Controls.Add(lblUser);
            // infoPanel.Controls.Add(txtCustomer);
            // infoPanel.Controls.Add(lblProduct);
            // infoPanel.Controls.Add(lblFG);
            // infoPanel.Controls.Add(lblserialNoPCBA);

            // this.Controls.Add(infoPanel);
            // // Start Button
            // btnStart = new System.Windows.Forms.Button
            // {

            //     Text = "Start Automation",
            //     Size = new System.Drawing.Size(150, 40),
            //     Location = new System.Drawing.Point(500, 15),
            //     BackColor = Color.FromArgb(0, 120, 215),
            //     ForeColor = Color.White,
            //     FlatStyle = FlatStyle.Flat
            // };
            // btnStart.FlatAppearance.BorderSize = 0;
            // btnStart.Click += btnStart_Click;
            // //this.Controls.Add(btnStart);
            // infoPanel.Controls.Add(btnStart);

            // // Log Box
            // _rtbLog = new RichTextBox
            // {
            //     Location = new System.Drawing.Point(15, 110),
            //     Size = new System.Drawing.Size(this.ClientSize.Width - 30,
            //         this.ClientSize.Height - 130),
            //     Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            //     ReadOnly = true,
            //     Font = new Font("Consolas", 9)
            // };
            // infoPanel.Controls.Add(_rtbLog);

            // // Store for use
            // this.Tag = _rtbLog;


            this.Text = "Burn In Test Automator";
            this.WindowState = FormWindowState.Maximized;
            this.BackColor = Color.WhiteSmoke;

            Panel topPanel = new Panel();
            topPanel.Dock = DockStyle.Top;
            topPanel.Height = 200;
            topPanel.BackColor = Color.White;
            this.Controls.Add(topPanel);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 3;
            layout.RowCount = 5;

           
            layout.RowStyles.Clear();

            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // Customer
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // Product
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // FG + Button
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // PCBA + FailCount

            topPanel.Controls.Add(layout);

            Label lblCustomer = new Label();
            lblCustomer.Text = "Customer Serial No";
            lblCustomer.Font = new Font("Segoe UI", 11);
            lblCustomer.Anchor = AnchorStyles.Left;

            txtCustomer = new System.Windows.Forms.TextBox();
            txtCustomer.Dock = DockStyle.Fill;
            txtCustomer.Font = new Font("Segoe UI", 11);
            txtCustomer.Width = 300;

            Label lblProduct = new Label();
            lblProduct.Text = "Product Type : M.2";
            lblProduct.Anchor = AnchorStyles.Left;

            Label lblFG = new Label();
            lblFG.Text = "FG Name :";
            lblFG.Anchor = AnchorStyles.Left;

            Label lblPCBA = new Label();
            lblPCBA.Text = "PCBA ID :";
            lblPCBA.Anchor = AnchorStyles.Left;

            Label lblCount = new Label();
            lblCount.Text = "Fail Count :";
            lblCount.Anchor = AnchorStyles.Left;

            lblFGValue = new Label();
            lblFGValue.Text = "-";
            lblFGValue.Font = new Font("Segoe UI", 11, System.Drawing.FontStyle.Bold);
            lblFGValue.ForeColor = Color.DarkBlue;
            lblFGValue.AutoSize = true;

            lblPCBValue = new Label();
            lblPCBValue.Text = "-";
            lblPCBValue.Font = new Font("Segoe UI", 11, System.Drawing.FontStyle.Bold);
            lblPCBValue.ForeColor = Color.DarkGreen;
            lblPCBValue.AutoSize = true;

            lblFailValue = new Label();
            lblFailValue.Text = "0";
            lblFailValue.Font = new Font("Segoe UI", 11, System.Drawing.FontStyle.Bold);
            lblFailValue.BackColor = Color.Red;
            lblFailValue.ForeColor = Color.White;
            lblFailValue.Padding = new Padding(5);
            lblFailValue.AutoSize = true;



            layout.Controls.Add(lblCustomer, 0, 0);
            layout.Controls.Add(txtCustomer, 1, 0);

            layout.Controls.Add(lblProduct, 0, 1);

            layout.Controls.Add(lblFG, 0, 2);
           

            layout.Controls.Add(lblPCBA, 0, 3);
            layout.Controls.Add(lblCount, 1, 3);

         

            layout.Controls.Add(lblFGValue, 1, 2);

           
            layout.Controls.Add(lblPCBValue, 1, 3);

            layout.Controls.Add(lblCount, 0, 4);
            layout.Controls.Add(lblFailValue, 1, 4);


            btnStart = new System.Windows.Forms.Button();
            btnStart.Text = "START AUTOMATION";
            btnStart.Width = 180;
            btnStart.Height = 40;
            btnStart.BackColor = Color.FromArgb(0, 120, 215);
            btnStart.ForeColor = Color.White;
            btnStart.FlatStyle = FlatStyle.Flat;
            btnStart.FlatAppearance.BorderSize = 0;

            btnStart.Click += btnStart_Click;

            layout.Controls.Add(btnStart, 2, 4);

         
            PictureBox gifBox = new PictureBox();
            gifBox.Image = Properties.Resources.Animation;
            gifBox.SizeMode = PictureBoxSizeMode.Zoom;
            gifBox.Dock = DockStyle.Fill;

            layout.Controls.Add(gifBox, 2, 1);
            layout.SetRowSpan(gifBox, 3);


            _rtbLog = new RichTextBox();
            _rtbLog.Dock = DockStyle.Fill;
            _rtbLog.BackColor = Color.Black;
            _rtbLog.ForeColor = Color.Lime;
            _rtbLog.Font = new Font("Consolas", 13);
            

            this.Controls.Add(_rtbLog);
            this.Tag = _rtbLog;
            _rtbLog.Text= "=== Burn In Test Automator ===\n\nPlease start the automation process.\n\n";
            _rtbLog.ReadOnly = true;
            _rtbLog.BringToFront();
            
        }
     
        protected override void OnPaint(PaintEventArgs e)
        {
            using (LinearGradientBrush brush = new LinearGradientBrush(
                this.ClientRectangle,
                Color.LightBlue,
                Color.White,
                90F))
            {
                e.Graphics.FillRectangle(brush, this.ClientRectangle);
            }
        }
    
        private void txtCustomer_TextChanged(object sender, EventArgs e)
        {
            barcodeTimer.Stop();
            barcodeTimer.Start();
        }
        //private void BarcodeTimer_Tick(object sender, EventArgs e)
        //{
        //    barcodeTimer.Stop();
        //    BarcodeScannerFunction();
        //}

        //private void BarcodeScannerFunction()
        //{
        //    string barcode = txtCustomer.Text.Trim();

        //    if (!string.IsNullOrEmpty(barcode))
        //    {
        //        btnStart.PerformClick();
        //    }
        //}

        private void UpdateUI(string fgName,string pcbaId,string customerserialno)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateUI(fgName,pcbaId, customerserialno)));
                return;
            }

            lblFGValue.Text =fgName;
            lblPCBValue.Text =pcbaId;
            txtCustomer.Text = customerserialno;
        }

        public Dictionary<bool,int> checkCustomerSerialNo(string customerSerialNo, RichTextBox log)
        {
            Dictionary<bool, int> checkSNN = new Dictionary<bool, int>();
            try
            {
               
              var  fgDetails = _userService.GetFgDetails(1, customerSerialNo);
                if (fgDetails != null)
                {
                    _customer = customerSerialNo;
                    _pcbSno = fgDetails.ProductType;
                    _fgName = fgDetails.FgName;
                    payhistory.FgNumber = _fgName;
                    payhistory.PCBAID = _pcbSno;
                    payhistory.CustomerSerialNumber = customerSerialNo;
                    UpdateUI(_fgName, _pcbSno, _customer);
                  
                  
                    var stage = _userService.startchecksfcs(fgDetails.FgName);

                    if (stage != null)
                    {
                        stages = stage;
                        stageid = Convert.ToInt32(stage[0]);
                        stageName = stage[1];
                      
                    }
                    //payhistory.StageId = stageid;
                    //payhistory.StageName = stageName;
                    checkSNN = _userService.Check_Curr_Stage(_pcbSno, stageid.ToString(), "Performance Test", true);
                    Log(log, "FG Name :" + _fgName + " -PCBAID " + _pcbSno);
                }
                else
                {
                    Log(log, "Serial No not found in DB" + "SN Not Found", Color.Red);
                    checkSNN.Add(false, 0);
                    SafeUI(() => txtCustomer.Clear());
                    SafeUI(() => txtCustomer.Focus());
                }

                return checkSNN;
            }
            catch (Exception ex)
            {
                writeErrorMessage(ex.Message.ToString(), "checkCustomerSerialNo");
                Log(log, $"ERROR: {ex.Message}", System.Drawing.Color.Red);
                checkSNN.Add(false , 0);
                return checkSNN;
            }
        }
        private void SafeUI(Action action)
        {
            if (InvokeRequired)
                Invoke(action);
            else
                action();
        }
        private async void btnStart_Click(object sender, EventArgs e)
        {
           
            var btn = (System.Windows.Forms.Button)sender;
            btn.Enabled = false;
            var log = (RichTextBox)this.Tag;
            Log(log, " Customerb Serial No :" + txtCustomer.Text.ToString(), Color.Lime);
            string serialFilePath = ConfigurationManager.AppSettings["HWPath"].ToString();
            //string serialFilePath = @"C:\project\hwi_822\hwi_822";
            var chkserialNo = await getSerialNo(serialFilePath, log);
            string customerSerialNo = txtCustomer.Text.ToString();
            //  var chkserialNo = await Task.Run(() => checkCustomerSerialNo(customerSerialNo, log));
            if (chkserialNo != null && chkserialNo.Any())
            {
                int value = chkserialNo.Values.FirstOrDefault();

                if (lblFailValue != null)
                {
                    if (this.InvokeRequired)
                    {
                        this.Invoke(new Action(() =>
                        {
                            lblFailValue.Text = "Fail Count : " + value;
                        }));
                    }
                    else
                    {
                        lblFailValue.Text = "Fail Count : " + value;
                    }
                }
            }
            if (chkserialNo.ContainsKey(false))
            {
                writeErrorMessage("Stage MissMatch/Serial No Not Found", serialFilePath);
                Log(log,"Stage MissMatch/Serial No Not Found",Color.Red);
              
                _userService.SQL_Upload(_pcbSno, _customer, true, "Stage MissMatch/Serial No Not Found",stages);
                return;
            }
          
            try
            {
               
                await Task.Run(() => RunAutomation(log));
            }
            catch (Exception ex)
            {
                writeErrorMessage(ex.Message.ToString(), "btnStart_Click");
                Log(log, $"ERROR: {ex.Message}", System.Drawing.Color.Red);
                _userService.SQL_Upload(_pcbSno, _customer, true, ex.Message.ToString(),stages);
            }
            finally
            {
                btn.Enabled = true;
            }
        }
        //static string Safe(string s)
        //{
        //    if (string.IsNullOrWhiteSpace(s))
        //        return "(empty)";
        //    return s;
        //}

   

        public async Task<string> DiskPartitionDynamic_NoWMI(RichTextBox log)
        {
            try
            {
                Log(log, "Starting auto disk partitioning (Disk 1 → D:, Disk 2 → E:, etc.)\r\n", Color.Lime);

                char driveLetter = 'D';
                int diskIndex = 1;

                //while (driveLetter <= 'Z')
                //{
                    string scriptPath = Path.Combine(Path.GetTempPath(), $"diskpart_auto_{diskIndex}.txt");

                    string script = $@"select disk {diskIndex}
attributes disk clear readonly
online disk noerr
clean
create partition primary
format fs=ntfs label=""Data"" quick
assign letter={driveLetter}
exit";

                    File.WriteAllText(scriptPath, script);

                    Log(log, $"Trying Disk {diskIndex} → {driveLetter}: ... ",Color.Lime);

                    var psi = new ProcessStartInfo
                    {
                        FileName = "diskpart.exe",
                        Arguments = "/s \"" + scriptPath + "\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        Verb = "runas"
                    };

                    string fullOutput = "";
                    string fullError = "";

                    using (var process = new Process { StartInfo = psi })
                    {
                        var outputBuilder = new StringBuilder();
                        var errorBuilder = new StringBuilder();

                        process.OutputDataReceived += (s, e) => { if (e.Data != null) outputBuilder.AppendLine(e.Data); };
                        process.ErrorDataReceived += (s, e) => { if (e.Data != null) errorBuilder.AppendLine(e.Data); };

                        process.Start();
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();
                        await Task.Run(() => process.WaitForExit());

                        fullOutput = outputBuilder.ToString().ToLowerInvariant();
                        fullError = errorBuilder.ToString().ToLowerInvariant();
                    }

                    // --- Compatible logic for .NET 4.7.2 ---
                    bool diskExists = true;
                    bool success = false;

                if (fullError.Contains("no disk") ||
                    fullError.Contains("there is no disk") ||
                    fullError.Contains("the specified disk does not exist") ||
                    fullOutput.Contains("no disk selected"))
                {
                    diskExists = false;
                    payhistory.DiskPartition = "Fail";
                }
                else if (fullOutput.Contains("diskpart succeeded") ||
                         fullOutput.Contains("100 percent completed") ||
                         fullOutput.Contains("successfully formatted") ||
                         fullOutput.Contains("assigned the drive letter"))
                {
                    success = true;

                }

                //Testing
                // --- Now simple if/else (no tuples!) ---
                //if (!diskExists)
                //{
                //    if (driveLetter == 'E')
                //        PartitionCount = 1;

                //    Log(log, "No more disks found.\r\n", Color.Orange);
                //    //break;
                //}

                if (success)
                {
                    Log(log, $"SUCCESS → {driveLetter}:\r\n", Color.LimeGreen);
                    writeErrorMessage("Drive Partition", $"SUCCESS → {driveLetter}:\r\n");
                    driveLetter++;
                    diskIndex++;
                    payhistory.DiskPartition = "Pass";
                }
                else
                {
                    Log(log, $"Skipped (already partitioned or failed)\r\n", Color.Yellow);
                    writeErrorMessage("Drive Partition", $"Skipped (already partitioned or failed)\r\n");
                    diskIndex++; // Try next disk anyway
                    payhistory.DiskPartition = "Pass";
                }

                    await Task.Delay(3000); // Let disk settle
             //   }

                Log(log, $"Finished. Last assigned letter: {(char)(driveLetter - 1)}:\r\n", Color.Cyan);
                Log(log,"Disk partitioning completed!Success",System.Drawing.Color.Green);
                writeErrorMessage("Drive Partition", "Disk partitioning completed!Success");
                return "true";
            }
            catch (Exception ex)
            {
                Log(log, $"Error: {ex.Message}\r\n", Color.Red);
                writeErrorMessage($"Error: {ex.Message}\r\n","Error");
                payhistory.DiskPartition = "Fail";
                _userService.SQL_Upload(_pcbSno, _customer, true, "Disk Partion Fail",stages);  
                return "false";
            }
        }
        private async void RunAutomation(RichTextBox log)
        {
            Log(log, "Starting automation...", Color.Lime);
            Log(log, "Disk Partition Start", Color.Lime);
            // bool eStatus = false;
            //Testing
            var status = await DiskPartitionDynamic_NoWMI(log);
            if (status == "false") return;

            //BurnIn Test Start


            exePath = ConfigurationManager.AppSettings["BurnInTest"];
            if (!File.Exists(exePath))
            {
                writeErrorMessage("File path Not Exists -", exePath.ToString());
                Log(log, $"EXE not found: {exePath}", System.Drawing.Color.Red);
                payhistory.burnintest ="Fail";
                _userService.SQL_Upload(_pcbSno, _customer, true, "File path Not Exists", stages);
                return;
            }
            writeErrorMessage("File path Exists -", exePath.ToString());
         


            Application app = null;
            try
            {
                //this.WindowState = FormWindowState.Maximized;
                //this.Activate();
                //this.BringToFront();
                using (var automation = new UIA3Automation())
                {
                    Console.WriteLine("=== All Open Window Names ===");

                    var desktop = automation.GetDesktop();
                    var allWindows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));

                    string drivemsg = drivecheck(log, allWindows);

                    Log(log, drivemsg, System.Drawing.Color.DarkBlue);

                    Thread.Sleep(2000);
                    Log(log, "Start Crystal DiskMark", Color.Lime);
                    //===========================================
                    writeErrorMessage("Message", "Start Crystal DiskMark");

                    var Crystalpath = ConfigurationManager.AppSettings["Crystal"];
                    if (!File.Exists(Crystalpath))
                    {
                        Log(log, "Crystal DiskMark Not Found", Color.Lime);
                        writeErrorMessage("Error -", "Crystal DiskMark Not Found");
                        payhistory.CrystalReport = "Fail";
                        _userService.SQL_Upload(_pcbSno, _customer, true, "Crystal DiskMark Not Found", stages);
                        return;
                    }
                    app = LaunchWithAdmin(Crystalpath);
                    System.Threading.Thread.Sleep(2000);

                    var mainWindowCrystal = desktop.FindFirstDescendant(cf =>
cf.ByControlType(ControlType.Window)
.And(cf.ByName("CrystalDiskMark 8.0.1 x86 [Admin]")))
?.AsWindow();

                    if (mainWindowCrystal == null)
                    {
                        Log(log, "CrystalDiskMark Window Not Found", Color.Lime);
                        writeErrorMessage("Error -", "CrystalDiskMark Window Not Found");
                        payhistory.CrystalReport = "Fail";
                        //popup.SetMessage("CrystalDiskMark Window Not Found!");
                        //popup.Show();
                        _userService.SQL_Upload(_pcbSno, _customer, true, "CrystalDiskMark Window Not Found", stages);
                        ppfrm.AddText("ERROR: ", Color.Red, true);
                        ppfrm.AddText("CrystalDiskMark Window Not Found!", Color.Black);
                        ppfrm.ShowDialog();
                        ppfrm.Focus();
                        return;
                    }
                    mainWindowCrystal.Focus();

                    //Comobox value
                    System.Threading.Thread.Sleep(1000);
                    string comboAutomationId = "1027";

                    var comboElement = mainWindowCrystal.FindFirstDescendant(cf =>
                cf.ByAutomationId(comboAutomationId)
                  .And(cf.ByControlType(ControlType.ComboBox)));
                    if (comboElement == null)
                    {
                        Log(log, $"❌ ComboBox with AutomationId '{comboAutomationId}' not found.", Color.Lime);
                        payhistory.CrystalReport = "Fail";
                        _userService.SQL_Upload(_pcbSno, _customer, true, "ComboBox with AutomationId", stages);
                        return;
                    }


                    var combo = comboElement.AsComboBox();
                    combo?.Expand();
                    Thread.Sleep(500); // allow items to appear

                    Log(log, $"✅ ComboBox Found: {comboAutomationId}", Color.Lime);
                    Log(log, "----------------------------------------------------", Color.Lime);

                    // 🔹 List all dropdown values
                    if (combo?.Items != null && combo.Items.Length > 0)
                    {
                        Log(log, "Available Items:");
                        foreach (var item in combo.Items)
                        {
                            if (item.Name == "D: 18% (42/232GiB)" || item.Name.Contains("D:"))
                            {
                                combo.Select(item.Name);
                                Log(log, "  • " + item.Text, Color.Lime);
                                break;
                            }
                        }
                        if (combo.Items.Length >= 3)
                            PartitionCount = 1;

                        Log(log, "Partition Count - " + PartitionCount.ToString(), Color.Lime);
                    }
                    else
                    {
                        Log(log, "⚠️ No items found or combo not expandable.", Color.Red);

                    }

                    // 🔹 Show currently selected item
                    if (combo?.SelectedItem != null)
                        Log(log, $"Selected: {combo.SelectedItem.Text}", Color.Lime);
                    else
                        Log(log, "No item currently selected.", Color.Lime);

                    combo?.Collapse();

                    //All Ok Button

                    if (combo.SelectedItem.Text.Contains("C:"))
                    {
                        Log(log, "D: Drive Not showing in crystal report.", Color.Lime);
                        writeErrorMessage("Error", "D: Drive Issue");
                        _userService.SQL_Upload(_pcbSno, _customer, true, "D: Drive Not showing in crystal report.", stages);
                        return;
                    }

                    var btnAll = mainWindowCrystal.FindFirstDescendant(cf => cf.ByName("All"))?.AsButton();
                    if (btnAll == null)
                    {
                        Log(log, "Button All Not Found", Color.Red);
                        _userService.SQL_Upload(_pcbSno, _customer, true, "Button All Not Found", stages);
                        return;
                    }
                    btnAll.Invoke();

                    Log(log, "D: crystal Report Started.",Color.Lime);
                    writeErrorMessage("Message", "D: crystal Report Started");
                    payhistory.CrystalReport = "Pass";



                    //Testing
                    //Application appnew = null;

                    //if (PartitionCount > 0)
                    //{
                    //                        Log(log, "Check Next crystal Report Entry.");
                    //                        appnew = LaunchWithAdmin(Crystalpath);
                    //                        string comboAutomationId_1 = "1027";
                    //                        System.Threading.Thread.Sleep(2000);
                    //                        var emainWindowCrystal = desktop.FindFirstDescendant(cf =>
                    //cf.ByControlType(ControlType.Window)
                    //.And(cf.ByName("CrystalDiskMark 8.0.1 x86 [Admin]")))
                    //?.AsWindow();
                    //                        if (emainWindowCrystal == null)
                    //                        {
                    //                            Log(log, $"❌ Second window not found.");
                    //                            return;
                    //                        }
                    //                        emainWindowCrystal.Focus();

                    //                        var comboElement_1 = emainWindowCrystal.FindFirstDescendant(cf =>
                    //                    cf.ByAutomationId(comboAutomationId_1)
                    //                      .And(cf.ByControlType(ControlType.ComboBox)));
                    //                        if (comboElement_1 == null)
                    //                        {
                    //                            Log(log, $"❌ ComboBox with AutomationId '{comboAutomationId_1}' not found.");
                    //                            appnew.Close();
                    //                            return;
                    //                        }


                    //                        var combo1 = comboElement_1.AsComboBox();

                    //                        combo1?.Expand();
                    //                        //Thread.Sleep(300); // allow items to appear

                    //                        //Log(log, $"✅ ComboBox Found: {comboAutomationId_1}");
                    //                        ////Log(log, "----------------------------------------------------");

                    //                        //Log(log,"Second list --" + combo1.Items.Count().ToString());
                    //                        Thread.Sleep(300);
                    //                        writeErrorMessage("Message", "E: crystal Report ");
                    //                        // 🔹 List all dropdown values
                    //                        if (combo1?.Items != null && combo1.Items.Length > 0)
                    //                        {

                    //                            foreach (var item in combo1.Items)
                    //                            {
                    //                                Log(log, "Available Items:" + item.Name);
                    //                                if (item.Name.Contains("E:"))
                    //                                {
                    //                                    Log(log, "Found E: Drive ");
                    //                                    combo1.Select(item.Name);
                    //                                    Log(log, "  • " + item.Text);
                    //                                    eStatus = true;
                    //                                    writeErrorMessage("Message", "E: crystal Report Found ");
                    //                                    break;
                    //                                }

                    //                            }

                    //                            if (!eStatus)
                    //                            {
                    //                                Log(log, "E: Drive not found.");
                    //                                writeErrorMessage("Error", "E: Drive not found ");
                    //                                appnew.Close(true);

                    //                            }
                    //                            else if (eStatus)
                    //                            {
                    //                                if (combo1?.SelectedItem != null)
                    //                                    Log(log, $"Selected: {combo1.SelectedItem.Text}");
                    //                                else
                    //                                    Log(log, "No item currently selected.");


                    //                                combo1?.Collapse();

                    //                                Log(log, "Collapse");

                    //                                //All Ok Button
                    //                                var btnAll_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByName("All"))?.AsButton();
                    //                                if (btnAll_1 == null)
                    //                                {
                    //                                    Log(log, "Button All Not Found");
                    //                                    return;
                    //                                }
                    //                                btnAll_1.Invoke();
                    //                                Log(log, "All Button clicked");
                    //                                writeErrorMessage("Message", "E: crystal Report Started ");
                    //                                //if (combo1.Items.Length > 4)
                    //                                //    PartitionCount = 2;
                    //                            }


                    //}
                    //else
                    //{
                    //    Log(log, "⚠️ No items found or combo not expandable.");
                    //}

                    // 🔹 Show currently selected item



                    // Thread.Sleep(100000); Testing
                    //check crystal result strat

                    Thread.Sleep(1000);

                    //check crystal result End

                    //var elements = mainWindowCrystal.FindAllDescendants();
                    //int count = 0;
                    //foreach (var el in elements)
                    //{
                    //    string name = SafeGet(() => el.Name);
                    //    string automationId = SafeGet(() => el.AutomationId);
                    //    string controlType = SafeGet(() => el.ControlType.ToString());


                    //    if (automationId.ToString() == "1009" || automationId.ToString() == "1010" || automationId.ToString() == "1011" || automationId.ToString() == "1012" || automationId.ToString() == "1014" || automationId.ToString() == "1015" || automationId.ToString() == "1016" || automationId.ToString() == "1017")
                    //    {

                    //        var ftxt = mainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId(automationId.ToString()));
                    //        string readfinal = getcrystalvalue(ftxt);
                    //       if (readfinal.ToString() == "0")
                    //        {
                    //            Log(log, "Crystal DiskMark Fail " + readfinal);
                    //            writeErrorMessage("D :Crystal DiskMark Fail - Read", "Error");
                    //            ppfrm.AddText("ERROR: ", Color.Red, true);
                    //            ppfrm.AddText("D:Crystal DiskMark Fail - Read!", Color.Black);
                    //            ppfrm.ShowDialog();
                    //            ppfrm.Focus();
                    //            return;
                    //        }
                    //        count++;
                    //        if (count == 8)
                    //        {
                    //            if (!(readfinal.ToString() == "0"))
                    //            {
                    //                Log(log, "Crystal DiskMark Pass -" + readfinal);

                    //                this.BeginInvoke(new Action(() =>
                    //                {
                    //                    BlinkPopupForm popup = new BlinkPopupForm(6000);
                    //                    popup.SetMessage("Crystal DiskMark Pass -Read");
                    //                    popup.Show();
                    //                    popup.Activate();
                    //                }));
                    //            }
                    //            else
                    //                break;
                    //        }
                    //        Log(log, $"Type: {controlType} | Name: {name} | AutomationId: {automationId}");

                    //    }


                    //}

                    //var Stxt = mainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1014"));
                    //if (Stxt != null)
                    //{
                    //    var Sval = Stxt.ToString().Split('.');
                    //    if (Sval.Length > 0)
                    //    {
                    //        if (Sval[0].Length < 3 || Sval[0].ToString() == "0")
                    //        {
                    //            Log(log, "Crystal DiskMark Fail - write");
                    //            writeErrorMessage("D :Crystal DiskMark Fail - write", "Error");
                    //            ppfrm.AddText("ERROR: ", Color.Red, true);
                    //            ppfrm.AddText("D:Crystal DiskMark Fail - write!", Color.Black);
                    //            ppfrm.ShowDialog();
                    //            ppfrm.Focus();
                    //            //popup.SetMessage("D:Crystal DiskMark Fail - write");
                    //            //popup.Show();
                    //            return;
                    //        }
                    //        else
                    //        {
                    //            var writeval = Sval[0].Split(',');
                    //            if (writeval.Length > 0)
                    //            {
                    //                Log(log, "Write Value Split --" + writeval[1]);
                    //                var writefinal = writeval[1].ToString();
                    //                if (!(writefinal.Length < 3 || writefinal.ToString() == "0"))
                    //                {
                    //                    Log(log, "Crystal DiskMark Pass -write -" + writefinal);
                    //                    this.BeginInvoke(new Action(() =>
                    //                    {
                    //                        BlinkPopupForm popup = new BlinkPopupForm(6000);
                    //                        popup.SetMessage("Crystal DiskMark Pass -Write");
                    //                        popup.Show();
                    //                        popup.Activate();
                    //                    }));
                    //                }
                    //                else if (writefinal.Length < 3 || writefinal.ToString() == "0")
                    //                {
                    //                    Log(log, "Crystal DiskMark Fail - write" + writefinal);
                    //                    writeErrorMessage("D :Crystal DiskMark Fail - write", "Error");
                    //                    ppfrm.AddText("ERROR: ", Color.Red, true);
                    //                    ppfrm.AddText("D:Crystal DiskMark Fail - write!", Color.Black);
                    //                    ppfrm.ShowDialog();
                    //                    ppfrm.Focus();
                    //                    //popup.SetMessage("D:Crystal DiskMark Fail - write");
                    //                    //popup.Show();
                    //                    return;



                    //                }
                    //            }
                    //            else
                    //            {
                    //                Log(log, "Write Value Split Issue --" + Sval[0]);
                    //                return;
                    //            }

                    //        }

                    //    }
                    //}
                    //Log(log, "Second Text Box-" + Stxt.Name);
                    //Testing
                    //                    if (eStatus)
                    //                    {
                    //                        var ecryCheck = mainWindowCrystal
                    //.FindFirstDescendant(cf => cf.ByName("Stop"))
                    //?.AsButton();

                    //                        if (ecryCheck != null)
                    //                        {
                    //                            do
                    //                            {
                    //                                Thread.Sleep(1000);
                    //                                Log(log, "E:Waiting for Crystal Test to complete... --" + ecryCheck.Name);
                    //                            } while (ecryCheck.Name != "All");
                    //                        }
                    //                        Log(log, " E:Crystal Test to completed");

                    //                    }

                    //this.Invoke(new Action(() =>
                    //{
                    //    MessageBox.Show(
                    //        this,
                    //        "Process finished",
                    //        "Done",
                    //        MessageBoxButtons.OK,
                    //        MessageBoxIcon.Information
                    //    );
                    //}));
                    writeErrorMessage("Crystal Disk Test completed Successfully", "Message");


                    //return; //Testing




                    //==============================
                    writeErrorMessage("Burn In Test Run Started", "Message");
                    Log(log, "Burn In Test Run Started", Color.Lime);
                    app = LaunchWithAdmin(exePath);
                    // Log(log, "Burn In Test Run Exe Launched");
                    Thread.Sleep(7000);

                    var mainWindow = desktop.FindFirstDescendant(cf =>
             cf.ByControlType(ControlType.Window)
               .And(cf.ByName("BurnInTest V8.1 Pro (1006)")))
             ?.AsWindow();

                    Thread.Sleep(2000);
                    if (mainWindow == null)
                    {
                        Log(log, "Main window not found", System.Drawing.Color.Red);
                        writeErrorMessage("Main window not found", "Error");
                        payhistory.burnintest = "Fail";
                        _userService.SQL_Upload(_pcbSno, _customer, true, "Main window not found", stages);
                        return;
                    }

                    mainWindow.Focus();

                    Thread.Sleep(1000);

                    var menuBar = mainWindow.FindFirstDescendant(cf => cf.ByControlType(ControlType.MenuBar))?.AsMenu();
                    if (menuBar == null)
                    {
                        Log(log, "Menu bar not found", System.Drawing.Color.Red);
                        payhistory.burnintest = "Fail";
                        _userService.SQL_Upload(_pcbSno, _customer, true, "Menu bar not found", stages);
                        return;
                    }
                    System.Threading.Thread.Sleep(1500);
                    var configMenu = mainWindow.FindFirstDescendant(cf => cf.ByName("Configuration"))?.AsMenuItem();
                    if (configMenu == null)
                    {
                        Log(log, "Configuration menu not found", System.Drawing.Color.Red);
                        payhistory.burnintest = "Fail";
                        _userService.SQL_Upload(_pcbSno, _customer, true, "Configuration menu not found", stages);
                        return;
                    }
                    configMenu?.Click();
                    Log(log, "Configuration menu Clicked", System.Drawing.Color.Green);

                    System.Threading.Thread.Sleep(1000);
                    var allWindowsnew = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));

                    foreach (var w in allWindowsnew)
                    {
                        try
                        {
                            System.Threading.Thread.Sleep(2500);

                            var testPref = w.FindFirstDescendant(cf => cf.ByName("Test Preferences..."));

                            if (testPref != null)
                            {
                                testPref.Click();

                                Log(log, "Test Preferences found-", System.Drawing.Color.Green);

                            }
                            System.Threading.Thread.Sleep(4000);



                            var prefWindow = mainWindow.FindFirstDescendant(cf => cf.ByName("BurnInTest Preferences"))?.AsWindow()
                             ?? mainWindow.FindFirstDescendant(cf => cf.ByControlType(ControlType.Window).And(cf.ByName("Preferences")))?.AsWindow();

                            if (prefWindow == null)
                            {
                                Log(log, "BurnInTest Preferences window not found.", Color.Red);
                                payhistory.burnintest = "Fail";
                                _userService.SQL_Upload(_pcbSno, _customer, true, "BurnInTest Preferences window not found.", stages);
                                return;
                            }
                            prefWindow.Focus();
                            Log(log, "BurnInTest Preferences window Focus.",Color.Lime);

                            Thread.Sleep(500);

                            var checkbox = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1224"))?.AsCheckBox();

                            if (checkbox == null)
                            {
                                Log(log, "Checkbox not found (check AutomationId-)" + checkbox.AutomationId + "-" + checkbox.Name, Color.Red);
                                payhistory.burnintest = "Fail";
                                _userService.SQL_Upload(_pcbSno, _customer, true, "Checkbox not found (check AutomationId-).", stages);
                                return;
                            }
                            checkbox.IsChecked = true;
                            Thread.Sleep(500);
                            checkbox.IsChecked = false;


                            for (int i = 0; i < 10; i++)
                            {
                                var lstC = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("ListViewItem-" + i))?.AsListBoxItem();
                                if (lstC == null)
                                {
                                    Log(log, "List Item Not Fount-" + i.ToString(), Color.Red);
                                    payhistory.burnintest = "Fail";
                                    break;
                                }
                                string lstName = lstC.Name.ToString();
                                string checkCDrive = lstName.Substring(0, 2);
                                if (checkCDrive == "C:")
                                {
                                    lstC.Select();

                                    var checkC = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1223"))?.AsCheckBox();

                                    if (checkC == null)
                                    {
                                        Log(log, "ListView C: not found .", Color.Red);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "ListView C: not found .", stages);
                                        return;
                                    }
                                    Log(log, "ListView C: found");
                                    checkC.IsChecked = false;


                                    // break;
                                }
                                if (checkCDrive == "D:")
                                {
                                    lstC.Select();
                                    System.Threading.Thread.Sleep(500);
                                    var checkD = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1223"))?.AsCheckBox();

                                    if (checkD == null)
                                    {
                                        Log(log, "ListView D: not found .", Color.Red);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "ListView D: not found .", stages);
                                        return;
                                    }
                                    Log(log, "ListView D: found", Color.Red);
                                    checkD.IsChecked = true;
                                    var fileSizeD = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1011"))?.AsTextBox();
                                    if (fileSizeD == null)
                                    {
                                        Log(log, "File Size Box Not Found", System.Drawing.Color.Red);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "File Size Box Not Found", stages);
                                        return;
                                    }
                                    System.Threading.Thread.Sleep(500);
                                    var filesized = fileSizeD.Patterns.Value.Pattern;
                                    if (filesized.ToString() != "3.00")
                                    {
                                        System.Threading.Thread.Sleep(500);
                                        filesized.SetValue("3.00");
                                        Log(log, "D: File Size Box value set", System.Drawing.Color.Green);
                                    }



                                    var comboElement_D = prefWindow.FindFirstDescendant(cf =>
                        cf.ByAutomationId("1148")
                          .And(cf.ByControlType(ControlType.ComboBox)));
                                    if (comboElement_D == null)
                                    {
                                        Log(log, "D: Block Size Box Not Found", System.Drawing.Color.Red);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "D: Block Size Box Not Found", stages);
                                        return;
                                    }
                                    var combod = comboElement_D.AsComboBox();
                                    combod.Focus();
                                    System.Threading.Thread.Sleep(200);

                                    Keyboard.Type("2");
                                    System.Threading.Thread.Sleep(250);

                                    Keyboard.Type("{ENTER}");

                                    if (combod.Value.ToString() != "256")
                                    {
                                        Log(log, "D:Block Size value 256 Not selectd" + combod.Value.ToString(), Color.Red);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "D:Block Size value 256 Not selectd", stages);
                                        return;
                                    }
                                    Log(log, " D: Block Size value Selected" + combod.Value.ToString(), Color.Lime);
                                    //break;
                                }
                                if (checkCDrive == "E:")
                                {
                                    lstC.Select();
                                    System.Threading.Thread.Sleep(500);
                                    var checkE = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1223"))?.AsCheckBox();

                                    if (checkE == null)
                                    {
                                        Log(log, "ListView E: not found .", Color.Red);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "ListView E: not found .", stages);
                                        return;
                                    }
                                    Log(log, "ListView E: found", Color.Lime);
                                    checkE.IsChecked = true;
                                    var fileSizeE = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1011"))?.AsTextBox();
                                    System.Threading.Thread.Sleep(500);
                                    if (fileSizeE == null)
                                    {
                                        Log(log, "E:File Size Box Not Found", System.Drawing.Color.Red);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "E:File Size Box Not Found", stages);
                                        return;
                                    }

                                    var filesizee = fileSizeE.Patterns.Value.Pattern;

                                    if (filesizee.ToString() != "3.00")
                                    {
                                        System.Threading.Thread.Sleep(500);
                                        filesizee.SetValue("3.00");
                                        Log(log, "E: File Size Box value set", System.Drawing.Color.Green);
                                    }


                                    var comboElement_E = prefWindow.FindFirstDescendant(cf =>
                          cf.ByAutomationId("1148")
                            .And(cf.ByControlType(ControlType.ComboBox)));
                                    if (comboElement_E == null)
                                    {
                                        Log(log, "E: Block Size Box Not Found", System.Drawing.Color.Red);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "E: Block Size Box Not Found", stages);
                                        return;
                                    }


                                    var comboe = comboElement_E.AsComboBox();
                                    comboe.Focus();
                                    System.Threading.Thread.Sleep(200);

                                    Keyboard.Type("2");
                                    System.Threading.Thread.Sleep(250);

                                    Keyboard.Type("{ENTER}");

                                    if (comboe.Value.ToString() != "256")
                                    {
                                        Log(log, "E Block Size value 256 Not selectd" + comboe.Value.ToString(), Color.Lime);
                                        payhistory.burnintest = "Fail";
                                        _userService.SQL_Upload(_pcbSno, _customer, true, "E Block Size value 256 Not selectd", stages);
                                        return;
                                    }
                                    Log(log, " E: Block Size value Selected" + comboe.Value.ToString());

                                    writeErrorMessage("Burn In Test 1st Stage Completed", "Message");

                                    break;
                                }

                            }





                            var okButton = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1"))?.AsButton();

                            if (okButton != null)
                            {
                                okButton.Invoke();
                                Log(log, "Clicked OK button.");
                            }
                            else
                            {
                                Log(log, "OK button not found (check AutomationId).");
                                payhistory.burnintest = "Fail";
                            }

                            configMenu?.Click();

                            Log(log, "Configuration menu Clicked", System.Drawing.Color.Green);
                            System.Threading.Thread.Sleep(2000);


                            var testPrefnext = w.FindFirstDescendant(cf => cf.ByName("Test Selection && Duty Cycles..."));

                            if (testPrefnext == null)
                            {

                                Log(log, "Test Selection & Duty Cycles not found-", System.Drawing.Color.Green);
                                payhistory.burnintest = "Fail";

                            }
                            else
                            {
                                Log(log, "Test Selection & Duty Cycles found-", System.Drawing.Color.Green);
                                testPrefnext.Click();
                            }

                            System.Threading.Thread.Sleep(2000);
                            var prefWindowcycles = mainWindow.FindFirstDescendant(cf =>
  cf.ByControlType(ControlType.Window)
    .And(cf.ByName("Test selection and duty cycles")))
  ?.AsWindow();

                            if (prefWindowcycles == null)
                            {
                                //Console.WriteLine("Test selection and duty cycles window not found.");
                                payhistory.burnintest = "Fail";
                                return;
                            }

                            prefWindowcycles.Focus();
                            //Console.WriteLine("Test selection and duty cycles window found.");

                            var checkBoxes1 = prefWindowcycles.FindAllDescendants(cf => cf.ByControlType(ControlType.CheckBox));
                            foreach (var cTest in checkBoxes1)
                            {
                                var checkboxnew = prefWindowcycles.FindFirstDescendant(cf => cf.ByAutomationId(cTest.AutomationId.ToString()))?.AsCheckBox();
                                if (cTest.AutomationId != "1048")
                                    checkboxnew.IsChecked = false;

                                if (cTest.AutomationId == "1048")
                                {
                                    if (!checkboxnew.IsChecked.HasValue || !checkboxnew.IsChecked.Value)
                                    {
                                        checkboxnew.IsChecked = true;
                                        Log(log, "Disk Checkbox checked successfully ✅", Color.Lime);
                                    }
                                    else
                                    {
                                        Log(log, "Disk Checkbox was already checked ✅", Color.Lime);
                                    }

                                }

                            }

                            Log(log, "Disk Checkbox checked completed", Color.Lime);
                            //M.2
                            var txtMinutes = prefWindowcycles.FindFirstDescendant(cf => cf.ByAutomationId("1074"))?.AsTextBox();
                            var txtMinvalue = txtMinutes.Patterns.Value.Pattern;
                            if (txtMinvalue != null)
                            {
                                if (txtMinvalue.ToString() != "0")
                                    txtMinvalue.SetValue("0");
                            }

                            var txtCycles = prefWindowcycles.FindFirstDescendant(cf => cf.ByAutomationId("1087"))?.AsTextBox();
                            var txtCylvalue = txtCycles.Patterns.Value.Pattern;
                            if (txtCylvalue != null)
                            {
                                if (txtCylvalue.ToString() != "11")
                                    txtCylvalue.SetValue("11");
                            }

                            var txtRow = prefWindowcycles.FindFirstDescendant(cf => cf.ByAutomationId("1067"))?.AsTextBox();
                            var txtRowvalue = txtRow.Patterns.Value.Pattern;
                            if (txtRowvalue != null)
                            {
                                if (txtRowvalue.ToString() != "50")
                                    txtRowvalue.SetValue("50");
                            }

                            var txtDisk = prefWindowcycles.FindFirstDescendant(cf => cf.ByAutomationId("1061"))?.AsTextBox();
                            var txtDiskvalue = txtDisk.Patterns.Value.Pattern;
                            if (txtDiskvalue != null)
                            {
                                if (txtDiskvalue.ToString() != "100")
                                    txtDiskvalue.SetValue("100");
                            }

                            //Testing
                            var btnOk = prefWindowcycles.FindFirstDescendant(cf => cf.ByName("OK"))?.AsButton();
                            if (btnOk != null)
                                btnOk.Invoke();
                            writeErrorMessage("Burn In Test 2nd Stage Completed", "Message");

                            System.Threading.Thread.Sleep(2000);

                            var TestMenu = mainWindow.FindFirstDescendant(cf => cf.ByName("Test"))?.AsMenuItem();
                            if (TestMenu != null)
                            {
                                TestMenu.Invoke();
                                System.Threading.Thread.Sleep(500);
                                Log(log, "Test menu Clicked", System.Drawing.Color.Green);
                            }
                            var testStart = w.FindFirstDescendant(cf => cf.ByName("Start Test Run"));
                            if (testStart != null)
                            {
                                testStart.Click();
                                Log(log, "Start Test Run Found", System.Drawing.Color.Green);
                            }

                            var prefWindowcyclesWarning = mainWindow.FindFirstDescendant(cf =>
cf.ByControlType(ControlType.Window)
 .And(cf.ByName("Getting ready to run Burn in tests")))
?.AsWindow();

                            if (prefWindowcyclesWarning == null)
                            {
                                //Console.WriteLine("warning Getting ready to run Burn in tests window not found.");
                                payhistory.burnintest = "Fail";
                                _userService.SQL_Upload(_pcbSno, _customer, true, "warning Getting ready to run Burn in tests window not found.", stages);
                                return;
                            }


                            prefWindowcyclesWarning.Focus();
                            // Testing
                            var btnOkwarning = prefWindowcyclesWarning.FindFirstDescendant(cf => cf.ByName("OK"))?.AsButton();
                            if (btnOkwarning != null)
                                btnOkwarning.Invoke();


                            Log(log, "Task Running", Color.Lime);
                            Thread.Sleep(146000);
                            //                        var windows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));

                            //                        foreach (var win in windows)
                            //                        {
                            //                            Log(log, "==================================");
                            //                            Log(log, $"Window Name       : {win.Name}");
                            //                            string automationId = win.Properties.AutomationId.IsSupported
                            //? win.Properties.AutomationId.Value
                            //: "(Not Supported)";

                            //                            Log(log, $"AutomationId : {automationId}");
                            //                            // Console.WriteLine($"AutomationId : {win.AutomationId ?? "(Not Available)"}");
                            //                            //Console.WriteLine($"ClassName         : {win.ClassName}");
                            //                            //Console.WriteLine($"ProcessId         : {win.Properties.ProcessId.Value}");
                            //                        }


                            //                        var childWindows = mainWindow.FindAllChildren(cf =>
                            //cf.ByControlType(ControlType.Window));

                            //                        foreach (var win in childWindows)
                            //                        {
                            //                            Console.WriteLine("-------- Child Window --------");
                            //                            Log(log, "==================================");
                            //                            Log(log, $"Window Name       : {win.Name}");
                            //                            string automationId = win.Properties.AutomationId.IsSupported
                            //? win.Properties.AutomationId.Value
                            //: "(Not Supported)";

                            //                            Log(log, $"AutomationId : {automationId}");
                            //                        }

                            //crystal Report check - Start
                            var cryCheck = mainWindowCrystal
.FindFirstDescendant(cf => cf.ByName("Stop"))
?.AsButton();

                            if (cryCheck != null)
                            {
                                do
                                {
                                    Thread.Sleep(1000);
                                    // Log(log, " D:Waiting for Crystal Test to complete... --" + cryCheck.Name);
                                } while (cryCheck.Name != "All");
                            }
                            Log(log, " D:Crystal Test to completed", Color.Lime);
                            var validIds = new HashSet<string>
{
    "1009","1010","1011","1012",
    "1014","1015","1016","1017"
};

                            int count = 0;

                            foreach (var el in mainWindowCrystal.FindAllDescendants())
                            {
                                string automationId = SafeGet(() => el.AutomationId);

                                if (!validIds.Contains(automationId))
                                    continue;

                                string name = SafeGet(() => el.Name);
                                string controlType = SafeGet(() => el.ControlType.ToString());

                                var element = mainWindowCrystal.FindFirstDescendant(cf =>
                                    cf.ByAutomationId(automationId));

                                if (element == null)
                                    continue;

                                string readValue = getcrystalvalue(element);
                                if (automationId == "1009")
                                    payhistory.read_one = readValue;
                                else if (automationId == "1010")
                                    payhistory.read_two = readValue;
                                else if (automationId == "1011")
                                    payhistory.read_three = readValue;
                                else if (automationId == "1012")
                                    payhistory.read_four = readValue;
                                else if (automationId == "1014")
                                    payhistory.write_one = readValue;
                                else if (automationId == "1015")
                                    payhistory.write_two = readValue;
                                else if (automationId == "1016")
                                    payhistory.write_three = readValue;
                                else if (automationId == "1017")
                                    payhistory.write_four = readValue;
                                // ❌ FAIL case
                                if (readValue == "0")
                                {
                                    _userService.SQL_Upload(_pcbSno, _customer, true, "D :Crystal DiskMark Fail - Read", stages);
                                    Log(log, "Crystal DiskMark Fail " + readValue, Color.Red);
                                    writeErrorMessage("D :Crystal DiskMark Fail - Read", "Error");
                                    payhistory.CrystalReport = "Fail";
                                    ppfrm.AddText("ERROR: ", Color.Red, true);
                                    ppfrm.AddText("D: Crystal DiskMark Fail - Read!", Color.Black);
                                    ppfrm.ShowDialog();
                                    ppfrm.Focus();

                                    return;
                                }
                                if ((automationId == "1009" || automationId == "1014") && Convert.ToDecimal(readValue) < 1000)
                                {
                                    _userService.SQL_Upload(_pcbSno, _customer, true, "D :Crystal DiskMark Fail - Read", stages);
                                    Log(log, "Crystal DiskMark Fail " + readValue, Color.Red);
                                    writeErrorMessage("D :Crystal DiskMark Fail - Read", "Error");
                                    payhistory.CrystalReport = "Fail";
                                    ppfrm.AddText("ERROR: ", Color.Red, true);
                                    ppfrm.AddText("D: Crystal DiskMark Fail - Read!", Color.Black);
                                    ppfrm.ShowDialog();
                                    ppfrm.Focus();
                                    return;
                                }
                                count++;

                                // ✅ PASS case (all 8 values processed)
                                if (count == validIds.Count)
                                {
                                    Log(log, "Crystal DiskMark Pass - " + readValue, Color.Lime);

                                    this.BeginInvoke(new Action(() =>
                                    {
                                        var popup = new BlinkPopupForm(6000);
                                        popup.SetMessage("Crystal DiskMark Pass - Read");
                                        popup.Show();
                                        popup.Activate();
                                    }));
                                    payhistory.CrystalReport = "Pass";
                                    Log(log, "Crystal DiskMark Pass", Color.Green);
                                    break;
                                }

                                Log(log, $"Type: {controlType} | Name: {name} | AutomationId: {automationId}",Color.Lime);
                            }

                            //Crystal Report check - end

                            //BurmInTest Result Check

                            string resultname = string.Empty;

                            do
                            {
                                var burnintestresult = mainWindow.FindFirstDescendant(cf =>
         cf.ByControlType(ControlType.Window)
           .And(cf.ByName("BurnInTest test result")));

                                resultname = burnintestresult == null ? string.Empty : burnintestresult.Name.ToString();

                                Thread.Sleep(1000);
                                string name = burnintestresult == null ? "N/A" : burnintestresult.Name;
                                Log(log, "Waiting for Burn In Test to complete... " + name, Color.Lime);

                            } while (resultname == string.Empty);
                            Thread.Sleep(2000);
                        


                            //Testing

                            //if (eStatus)
                            //{
                            //    // Thread.Sleep(100000);
                            //    var ftxt_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1009"));
                            //    if (ftxt_1 != null)
                            //    {
                            //        var fval_1 = ftxt_1.ToString().Split('.');
                            //        if (fval_1.Length > 0)
                            //        {
                            //            if (fval_1[0].Length > 3)
                            //            {
                            //                Log(log, "Crystal DiskMark Pass -Read");
                            //                MessageBox.Show("Crystal DiskMark Pass -Read", "Message");
                            //            }
                            //            else
                            //            {
                            //                Log(log, "Crystal DiskMark Fail - Read");
                            //                MessageBox.Show(" E :Crystal DiskMark Fail - Read", "Error");
                            //                writeErrorMessage("E :Crystal DiskMark Fail - Read", "Error");
                            //            }
                            //        }

                            //    }
                            //    Log(log, "First Text Box-" + ftxt_1.Name);

                            //    var Stxt_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1014"));
                            //    if (Stxt_1 != null)
                            //    {
                            //        var Sval_1 = Stxt_1.ToString().Split('.');
                            //        if (Sval_1.Length > 0)
                            //        {
                            //            if (Sval_1[0].Length > 3)
                            //            {
                            //                Log(log, "Crystal DiskMark Pass -write");
                            //                MessageBox.Show("Crystal DiskMark Pass -write", "Message");
                            //            }
                            //            else
                            //            {
                            //                Log(log, "Crystal DiskMark Fail - write");
                            //                MessageBox.Show("E:Crystal DiskMark Fail - write", "Error");
                            //                writeErrorMessage("E :Crystal DiskMark Fail - write", "Error");
                            //            }
                            //        }

                            //    }
                            //    Log(log, "Second Text Box-" + Stxt_1.Name);
                            //}

                            //first Crystal Report stage



                            //                            var Crystalpath = ConfigurationManager.AppSettings["Crystal"];
                            //                            if (!File.Exists(Crystalpath))
                            //                            { 
                            //                                Log(log, "Crystal DiskMark Not Found");
                            //                                return;
                            //                            }
                            //                            app = LaunchWithAdmin(Crystalpath);
                            //                            System.Threading.Thread.Sleep(2000);

                            //                            var mainWindowCrystal = desktop.FindFirstDescendant(cf =>
                            //cf.ByControlType(ControlType.Window)
                            //.And(cf.ByName("CrystalDiskMark 8.0.1 x86 [Admin]")))
                            //?.AsWindow();

                            //                            if (mainWindowCrystal == null)
                            //                            {
                            //                                Log(log, "CrystalDiskMark Window Not Found");
                            //                                return;
                            //                            }
                            //                            mainWindowCrystal.Focus();

                            //                            //Comobox value
                            //                            System.Threading.Thread.Sleep(1000);
                            //                            string comboAutomationId = "1027";

                            //                            var comboElement = mainWindowCrystal.FindFirstDescendant(cf =>
                            //                        cf.ByAutomationId(comboAutomationId)
                            //                          .And(cf.ByControlType(ControlType.ComboBox)));
                            //                            if (comboElement == null)
                            //                            {
                            //                                Log(log, $"❌ ComboBox with AutomationId '{comboAutomationId}' not found.");
                            //                                return;
                            //                            }


                            //                            var combo = comboElement.AsComboBox();
                            //                            combo?.Expand();
                            //                            Thread.Sleep(500); // allow items to appear

                            //                            Log(log, $"✅ ComboBox Found: {comboAutomationId}");
                            //                            Log(log, "----------------------------------------------------");

                            //                            // 🔹 List all dropdown values
                            //                            if (combo?.Items != null && combo.Items.Length > 0)
                            //                            {
                            //                                Log(log, "Available Items:");
                            //                                foreach (var item in combo.Items)
                            //                                {
                            //                                    if (item.Name == "D: 18% (42/232GiB)" || item.Name.Contains("D:"))
                            //                                    {
                            //                                        combo.Select(item.Name);
                            //                                        Log(log, "  • " + item.Text);
                            //                                        break;
                            //                                    }
                            //                                }
                            //                                if (combo.Items.Length >= 3)
                            //                                    PartitionCount = 1;
                            //                            }
                            //                            else
                            //                            {
                            //                                Log(log, "⚠️ No items found or combo not expandable.");
                            //                            }

                            //                            // 🔹 Show currently selected item
                            //                            if (combo?.SelectedItem != null)
                            //                                Log(log, $"Selected: {combo.SelectedItem.Text}");
                            //                            else
                            //                                Log(log, "No item currently selected.");

                            //                            combo?.Collapse();

                            //                            //All Ok Button
                            //                            var btnAll = mainWindowCrystal.FindFirstDescendant(cf => cf.ByName("All"))?.AsButton();
                            //                            if (btnAll == null)
                            //                            {
                            //                                Log(log, "Button All Not Found");
                            //                                return;
                            //                            }
                            //                            btnAll.Invoke();

                            //                            Log(log, "D: crystal Report Started.");

                            //                            Application appnew = null;

                            //                            if (PartitionCount > 0 )
                            //                            {
                            //                                Log(log, "Check Next crystal Report Entry.");
                            //                                appnew = LaunchWithAdmin(Crystalpath);
                            //                                string comboAutomationId_1 = "1027";
                            //                                System.Threading.Thread.Sleep(2000);
                            //                                var emainWindowCrystal = desktop.FindFirstDescendant(cf =>
                            //cf.ByControlType(ControlType.Window)
                            //.And(cf.ByName("CrystalDiskMark 8.0.1 x86 [Admin]")))
                            //?.AsWindow();
                            //                                if (emainWindowCrystal == null)
                            //                                {
                            //                                    Log(log, $"❌ Second window not found.");
                            //                                    return;
                            //                                }
                            //                                emainWindowCrystal.Focus();

                            //                                var comboElement_1 = emainWindowCrystal.FindFirstDescendant(cf =>
                            //                            cf.ByAutomationId(comboAutomationId_1)
                            //                              .And(cf.ByControlType(ControlType.ComboBox)));
                            //                                if (comboElement_1 == null)
                            //                                {
                            //                                    Log(log, $"❌ ComboBox with AutomationId '{comboAutomationId_1}' not found.");
                            //                                    appnew.Close();
                            //                                    return;
                            //                                }


                            //                                var combo1 = comboElement_1.AsComboBox();

                            //                                combo1?.Expand();
                            //                                //Thread.Sleep(300); // allow items to appear

                            //                                //Log(log, $"✅ ComboBox Found: {comboAutomationId_1}");
                            //                                ////Log(log, "----------------------------------------------------");

                            //                                //Log(log,"Second list --" + combo1.Items.Count().ToString());
                            //                                Thread.Sleep(300);
                            //                                // 🔹 List all dropdown values
                            //                                if (combo1?.Items != null && combo1.Items.Length > 0)
                            //                                {
                            //                                    bool eStatus = false;
                            //                                    foreach (var item in combo1.Items)
                            //                                    {
                            //                                        Log(log, "Available Items:" + item.Name);
                            //                                        if (item.Name.Contains("E:"))
                            //                                        {
                            //                                            Log(log, "Found E: Drive ");
                            //                                            combo1.Select(item.Name);
                            //                                            Log(log, "  • " + item.Text);
                            //                                            eStatus = true;
                            //                                            break;
                            //                                        }
                            //                                    }

                            //                                    if(!eStatus)
                            //                                    {
                            //                                        Log(log, "E: Drive found.");
                            //                                        appnew.Close(true);
                            //                                        return;
                            //                                    }

                            //                                    if (combo1.Items.Length > 4)
                            //                                        PartitionCount = 2;
                            //                                }
                            //                                else
                            //                                {
                            //                                    Log(log, "⚠️ No items found or combo not expandable.");
                            //                                }

                            //                                // 🔹 Show currently selected item
                            //                                if (combo1?.SelectedItem != null)
                            //                                    Log(log, $"Selected: {combo1.SelectedItem.Text}");
                            //                                else
                            //                                    Log(log, "No item currently selected.");


                            //                                combo1?.Collapse();



                            //                                //All Ok Button
                            //                                var btnAll_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByName("All"))?.AsButton();
                            //                                if (btnAll_1 == null)
                            //                                {
                            //                                    Log(log, "Button All Not Found");
                            //                                    return;
                            //                                }
                            //                                btnAll_1.Invoke();

                            //                                Thread.Sleep(100000);
                            //                                var ftxt_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1009"));
                            //                                if (ftxt_1 != null)
                            //                                {
                            //                                    var fval_1 = ftxt_1.ToString().Split('.');
                            //                                    if (fval_1.Length > 0)
                            //                                    {
                            //                                        if (fval_1[0].Length > 3)
                            //                                            Log(log, "Crystal DiskMark Pass -Read");
                            //                                        else
                            //                                            Log(log, "Crystal DiskMark Fail - Read");
                            //                                    }

                            //                                }
                            //                                Log(log, "First Text Box-" + ftxt_1.Name);

                            //                                var Stxt_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1014"));
                            //                                if (Stxt_1 != null)
                            //                                {
                            //                                    var Sval_1 = Stxt_1.ToString().Split('.');
                            //                                    if (Sval_1.Length > 0)
                            //                                    {
                            //                                        if (Sval_1[0].Length > 3)
                            //                                            Log(log, "Crystal DiskMark Pass -write");
                            //                                        else
                            //                                            Log(log, "Crystal DiskMark Fail - write");
                            //                                    }

                            //                                }
                            //                                Log(log, "Second Text Box-" + Stxt_1.Name);

                            //                                //first Crystal Report stage

                            //                                var ftxt = mainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1009"));
                            //                                if (ftxt != null)
                            //                                {
                            //                                    var fval = ftxt.ToString().Split('.');
                            //                                    if (fval.Length > 0)
                            //                                    {
                            //                                        if (fval[0].Length > 3)
                            //                                            Log(log, "Crystal DiskMark Pass -Read");
                            //                                        else
                            //                                            Log(log, "Crystal DiskMark Fail - Read");
                            //                                    }

                            //                                }
                            //                                Log(log, "First Text Box-" + ftxt.Name);

                            //                                var Stxt = mainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1014"));
                            //                                if (Stxt != null)
                            //                                {
                            //                                    var Sval = Stxt.ToString().Split('.');
                            //                                    if (Sval.Length > 0)
                            //                                    {
                            //                                        if (Sval[0].Length > 3)
                            //                                            Log(log, "Crystal DiskMark Pass -write");
                            //                                        else
                            //                                            Log(log, "Crystal DiskMark Fail - write");
                            //                                    }

                            //                                }
                            //                                Log(log, "Second Text Box-" + Stxt.Name);

                            // appnew.Close();
                            //    }

                            //Testing
                            payhistory.overall_result = "Pass";
                            payhistory.burnintest = "Pass";
                           var nextstage = _userService.Nextstartchecksfcs(_fgName);
                           _userService.SQL_Upload(_pcbSno, _customer, false, "Passmark Test Completed.", nextstage);
                            payhistory.CreatedBy = "Admin";
                            int resultHistory = _userService.inserthistory(payhistory);
                            if (resultHistory > 0)
                            {
                                Log(log, "Test history saved to DB", System.Drawing.Color.Green);
                            }
                            else
                            {
                                Log(log, "Failed to save test history to DB", System.Drawing.Color.Red);

                            }
                            // Burn In Test Check End

                            //ppfrm.TopMost = true;
                            //ppfrm.AddText("SUCCESS: ", Color.Green, true);
                            //ppfrm.AddText("Passmark Test Completed.\n", Color.Green);
                            //ppfrm.ShowDialog();
                            //ppfrm.Activate();
                            this.Invoke((MethodInvoker)delegate
                            {
                                frmblink = new Blink();
                                frmblink.TopMost = true;
                                frmblink.Show();

                                frmblink.AddText("SUCCESS:", Color.Green, true);
                                frmblink.AddText("Passmark Test Completed.", Color.Green);
                            });


                            if (mainWindowCrystal == null)
                            {
                                if (mainWindowCrystal.Patterns.Window.Pattern.WindowVisualState.Value
     == FlaUI.Core.Definitions.WindowVisualState.Minimized)
                                {
                                    mainWindowCrystal.Patterns.Window.Pattern.SetWindowVisualState(
                                        FlaUI.Core.Definitions.WindowVisualState.Normal);
                                }
                                mainWindowCrystal.Close();
                               
                            }
                            else if (mainWindowCrystal != null)
                            {
                                mainWindowCrystal.Close();
                              
                            }
                            else
                                Log(log, "Crystal DiskMark - Page not found", Color.Red);
                            // PartitionCount = 0;


                            //  app.Close();




                            break;
                        }
                        catch (Exception ex)
                        {

                            payhistory.overall_result = "Fail";
                            Log(log, "error-" + ex.Message.ToString(), System.Drawing.Color.Red);
                            writeErrorMessage(ex.Message.ToString(), "Crystal DiskMark");
                            ppfrm.AddText("ERROR: ", Color.Red, true);
                            ppfrm.AddText(ex.Message.ToString(), Color.Black);
                            _userService.SQL_Upload(_pcbSno, _customer, true, ex.Message.ToString(),stages);
                            ppfrm.ShowDialog();
                            ppfrm.Focus();
                            ppfrm.Activate();
                            break;
                        }
                        finally
                        {

                            mainWindow.Close();
                            app.Dispose();
                           
                            

                        }

                        // }

                    }

                }
            }


            catch (Exception ex)
            {
                writeErrorMessage(ex.Message.ToString(), "RunAutomation");
                Log(log, $"FAILED: {ex.Message}", System.Drawing.Color.Red);
            }
            finally
            {
                // app?.Dispose();
            }
        }
        
        private Application LaunchWithAdmin(string path)
        {
            var psi = new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true,
                Verb = "runas"
            };
            return Application.Launch(psi);
        }

        public string drivecheck(RichTextBox log, AutomationElement[] desktopAll)
        {
            string result = string.Empty;
            try
            {
                var dwindow = desktopAll
    .Where(w => w.ClassName == "CabinetWClass" &&
                !string.IsNullOrEmpty(w.Name) &&
                w.Name.Contains("D:"))
    .ToList();

                if (dwindow.Count <= 0)
                {
                    Log(log, "D:\\ window not found",Color.Lime);
                    result = "D:\\ window not found";
                }
                else
                {

                    foreach (var win in dwindow)
                    {
                        var titleBar = win.FindFirstDescendant(cf =>
    cf.ByControlType(ControlType.TitleBar));

                        if (titleBar != null)
                        {
                            var buttons = titleBar.FindAllDescendants(cf =>
                                cf.ByControlType(ControlType.Button));

                            // Close button is always the LAST button in titlebar
                            var closeBtn = buttons.LastOrDefault();

                            closeBtn?.AsButton()?.Invoke();

                            result = "D:\\ window Closed";
                        }
                    }
                    
                }


                        var ewindow = desktopAll
            .Where(w => w != null && w.ClassName == "CabinetWClass" &&
                        !string.IsNullOrEmpty(w.Name) &&
                        w.Name.Contains("E:")).ToList();

                if (ewindow.Count <= 0)
                {
                    Log(log, "E:\\ window not found", Color.Lime);
                    result = "E:\\ window not found -" + result;
                }
                else
                {

                    foreach (var ewin in ewindow)
                    {
                        var etitleBar = ewin.FindFirstDescendant(cf =>
    cf.ByControlType(ControlType.TitleBar));

                        if (etitleBar != null)
                        {
                            var ebuttons = etitleBar.FindAllDescendants(cf =>
                                cf.ByControlType(ControlType.Button));

                            // Close button is always the LAST button in titlebar
                            var ecloseBtn = ebuttons.LastOrDefault();

                            ecloseBtn?.AsButton()?.Invoke();

                            result = "E:\\ window Closed" + "--" + result;
                        }
                    }

                }
                    
            }
            catch (Exception ex)
            {
                result = ex.Message.ToString() + "-" + result;
            }

            return result;
        }

        //public void Log(RichTextBox rtb, string message, System.Drawing.Color? color = null)
        //{
        //    this.Invoke((MethodInvoker)delegate
        //    {
        //        rtb.SelectionColor = color ?? System.Drawing.Color.Black;
        //        rtb.AppendText($"{DateTime.Now:HH:mm:ss} | {message}\r\n");
        //        rtb.ScrollToCaret();
        //    });
        //}

        public void Log(RichTextBox rtb, string message, Color? color = null)
        {
            if (rtb == null) return;

            if (rtb.InvokeRequired)
            {
                rtb.Invoke(new Action(() => Log(rtb, message, color)));
                return;
            }

            rtb.SelectionStart = rtb.TextLength;
            rtb.SelectionLength = 0;
            rtb.SelectionColor = color ?? Color.Black;

            rtb.AppendText($"{DateTime.Now:HH:mm:ss} | {message}\r\n");
            rtb.SelectionColor = rtb.ForeColor;
        }

        public void writeErrorMessage(string errorMessage, string functionName)
        {
            var systemPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\BurnInTest" + "\\" + DateTime.Now.ToString("dd-MM-yyyy");
           // var systemPath =@"D:" + "\\BurnInTest" + "\\" + DateTime.Now.ToString("dd-MM-yyyy");

            if (!Directory.Exists(systemPath))
            {
                Directory.CreateDirectory(systemPath);
            }

            string WrErrorLog = String.Format(@"{0}\{1}.txt", systemPath, "BurnInTestLog");
            using (StreamWriter errLogs = new StreamWriter(WrErrorLog, true))
            {
                errLogs.WriteLine("--------------------------------------------------------------------------------------------------------------------" + Environment.NewLine);
                errLogs.WriteLine("---------------------------------------------------" + DateTime.Now + "----------------------------------------------" + Environment.NewLine);
                errLogs.WriteLine(errorMessage + Environment.NewLine + "-----" + functionName);
                errLogs.Close();
            }
        }

        

        public string getcrystalvalue(AutomationElement ftxt)
        {
            string readval = string.Empty;
          
            

                if (ftxt != null)
                {
                    var fval = ftxt.ToString().Split('.');
                    if (fval.Length > 0)
                    {
                        var readvalue = fval[0].Split(',');
                        if (readvalue.Length > 0)
                        {
                            var readvalfinal = readvalue[1].Split(':');
                            readval = readvalfinal[1].Trim();
                          
                            return readval.ToString();
                        }
                        else
                            return string.Empty;

                    }

                }
          
            return readval;
        }
        static string SafeGet(Func<string> getter)
        {
            try
            {
                return string.IsNullOrWhiteSpace(getter()) ? "N/A" : getter();
            }
            catch
            {
                return "Not Supported";
            }
        }
        private void FrmBurnIntest_Load(object sender, EventArgs e)
        {
            barcodeTimer.Interval = 300; // 300ms delay
           // barcodeTimer.Tick += BarcodeTimer_Tick;

          //  txtCustomer.TextChanged += txtCustomer_TextChanged;
        }
       

        public async Task<Dictionary<bool,int>> getSerialNo(string HW_PATH , RichTextBox log)
        {
            Dictionary<bool, int> checkSN = new Dictionary<bool, int>();
            try
            {
                string[] invalues = new string[7];
                string serialno = string.Empty;
                ProcessStartInfo proc = new ProcessStartInfo();
                proc.FileName = Path.Combine(HW_PATH, "HWiNFO64.exe");
                proc.WorkingDirectory = HW_PATH;
                proc.Verb = "runas";
                Process.Start(proc);
                this.SendToBack();
                Thread.Sleep(5000);

                // --- Ensure HWiNFO is running ---
                Process[] hwinfoApps = Process.GetProcessesByName("HWiNFO64");

                if (hwinfoApps.Length == 0)
                {
                    Log(log,"Unable to Open HWiNFO... Contact Test Dev Team" + "HWiNFO Not Opening",Color.Red);
                    //Application.Exit();
                    checkSN.Add(false, 0);
                    return checkSN;
                }

                // --- Retrieve main window AutomationElement ---
                //AutomationElement hwWindow = AutomationElement.FromHandle(hwinfoApps[0].MainWindowHandle);
                using (var automation = new UIA3Automation())
                {
                    Thread.Sleep(1500);
                    var hwWindow = automation.FromHandle(
                        hwinfoApps[0].MainWindowHandle
                    ).AsWindow();


                    if (hwWindow == null)
                    {
                        Log(log,"Automation Element is Not Working"+ "UI Element Error", Color.Red);
                        // Application.Exit();
                        checkSN.Add(false, 0);
                        return checkSN ;
                    }

                    // --- Click "Save Report" Button ---
                   await ClickButtonByName(hwWindow, "Save Report" , log);

                    Thread.Sleep(1000);


                    // --- Kill HWiNFO processes ---
                    try
                    {
                        foreach (Process p in hwinfoApps)
                        {
                            if (!p.HasExited)
                                p.Kill();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log(log,ex.Message,Color.Red);
                    }

                    this.BringToFront();
                    this.Refresh();


                    // string csvFilePath = Path.Combine(HW_PATH, Environment.MachineName + ".csv");
                    string csvFilePath = Path.Combine(HW_PATH, filenameNow + ".csv");

                    List<string[]> lines = new List<string[]>();

                    using (TextFieldParser parser = new TextFieldParser(csvFilePath))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");

                        while (!parser.EndOfData)
                        {
                            try
                            {
                                string[] fields = parser.ReadFields();
                                lines.Add(fields);
                            }
                            catch (Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show(ex.ToString());
                                checkSN.Add(false, 0);
                                return checkSN;
                            }
                        }
                    }



                    // -------- Loop through CSV --------
                    for (int i = 0; i < lines.Count; i++)
                    {
                        string[] rows = lines[i];

                        if (rows.Length == 2)
                        {
                            if (rows[0] == "Drive Serial Number:")
                            {
                                // string rpt_SN = rows[1];

                               
                                serialno = rows[1].ToString();
                                if (!serialno.Contains("SSD"))
                                {
                                    writeErrorMessage("Serial No-" + serialno, "getSerialNo");
                                    _productTypeId = 1;
                                    //serialno = "ESP3C1Q06CNA01656"; //Testing
                                    _fgDetails = _userService.GetFgDetails(_productTypeId, serialno);
                                    if (_fgDetails != null)
                                    {
                                        // _customer = _fgDetails.Customer;
                                        _customer = serialno;
                                        _pcbSno = _fgDetails.ProductType;
                                        _fgName = _fgDetails.FgName;
                                        UpdateUI(_fgName, _pcbSno, _customer);
                                        payhistory.FgNumber = _fgName;
                                        payhistory.PCBAID = _pcbSno;
                                        payhistory.CustomerSerialNumber = _customer;
                                        var stage = _userService.startchecksfcs(_fgName);

                                        if (stage != null)
                                        {
                                            stages = stage;
                                            stageid = Convert.ToInt32(stage[0]);
                                            stageName = stage[1];

                                        }
                                        checkSN = _userService.Check_Curr_Stage(_pcbSno, "262", "Performance Test", true);
                                        Log(log, "FG Name :" + _fgName + " -PCBAID " + _pcbSno, Color.Lime);
                                    }
                                    else
                                    {
                                        //Log(log, "Serial No not found in Data Base", Color.Red);
                                        //checkSN.Add(false, 0);
                                        Log(log, "Serial No not found in DB", Color.Red);
                                        checkSN.Add(false, 0);
                                        SafeUI(() => txtCustomer.Clear());
                                        SafeUI(() => txtCustomer.Focus());

                                    }
                                    return checkSN;
                                }
                                 
                                  
                            }
                        }
                    }





                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
                checkSN.Add(false, 0);
                return checkSN;
            }
            return checkSN;

        }


        private async Task ClickButtonByName(AutomationElement root, string containsText,RichTextBox log)
        {
            using (var automation = new UIA3Automation())
            {
                var buttonCondition = automation.ConditionFactory
          .ByControlType(FlaUI.Core.Definitions.ControlType.Button);

                //var buttons = root.FindAllDescendants();"AutomationId:2061, Name:, ControlType:edit, FrameworkId:Win32"
              
                var buttons = root.FindAllDescendants();
                var buttonElement = root.FindFirstDescendant(cf => cf.ByName("Create a Report File"));
                if (buttonElement != null)
                {
                    var btn = buttonElement.AsButton();
                    if (btn != null)
                    {
                        btn.Invoke();
                    }
                    else
                    {
                        Log(log,"Button Not Found"+ "UI Element Error", Color.Red);
                        return;
                    }
                }
                Thread.Sleep(500);
                var delomited = root.FindFirstDescendant(cf => cf.ByAutomationId("1219"));
                if (delomited != null)
                {
                    var delbtn = delomited.AsButton();
                    if (delbtn != null)
                    {
                        delbtn.Invoke();
                    }
                    else
                    {
                        Log(log,"Delimited Button Not Found"+ "UI Element Error", Color.Red);
                        return;
                    }
                }
                Thread.Sleep(500);
             
                var filename = root.FindFirstDescendant(cf => cf.ByAutomationId("2061"))?.AsTextBox();
                if (filename != null)
                {
                    filename.Patterns.Value.Pattern.SetValue("");
                    var fname = DateAndTime.Now.ToString();
                     filenameNow = Regex.Replace(fname, @"[^0-9]", "");
                    filename.Patterns.Value.Pattern.SetValue(filenameNow + ".CSV");
                }
                else
                {
                    Log(log,"Filename TextBox Not Found"+ "UI Element Error", Color.Red);
                    return;
                }


                var nextbtn = root.FindFirstDescendant(cf => cf.ByAutomationId("12324"));
                if (nextbtn != null)
                {
                    var nbtn = nextbtn.AsButton();
                    if (nbtn != null)
                    {
                        nbtn.Invoke();

                    }
                    else
                    {
                        Log(log,"Next Button Not Found"+ "UI Element Error", Color.Red);
                        return;
                    }
                }
                Thread.Sleep(500);
                var finishbtn = root.FindFirstDescendant(cf => cf.ByAutomationId("12325"));
                if(finishbtn != null)
                {
                    var fbtn = finishbtn.AsButton();
                    if (fbtn != null)
                    {
                        fbtn.Invoke();

                    }
                    else
                    {
                        Log(log,"Finish Button Not Found"+ "UI Element Error", Color.Red);
                        return;
                    }
                }

              
            }

        }

        private void SelectRadioButton(AutomationElement root, string containsText)
        {
            using (var automation = new UIA3Automation())
            {
                var rbCondition = automation.ConditionFactory
         .ByControlType(FlaUI.Core.Definitions.ControlType.RadioButton);

                var radioButtons = root.FindAll(TreeScope.Descendants, rbCondition);

                foreach (var rb in radioButtons)
                {
                    if (rb.Name.Contains(containsText))
                    {
                        var selectionItem = rb.Patterns.SelectionItem;
                        if (selectionItem.IsSupported)
                        {
                            selectionItem.Pattern.Select();
                        }
                        return;
                    }
                }
            }
        }

        private AutomationElement FindDialogWindow()
        {
            // Windows dialog = class #32770
            //Condition dlgCond = new PropertyCondition(AutomationElement.ClassNameProperty, "#32770");
            //var desktop = AutomationElement.RootElement;
            using (var automation = new UIA3Automation())
            {

                var dlgCond = automation.ConditionFactory
                    .ByClassName("#32770");

                var desktop = automation.GetDesktop();

                return desktop.FindFirst(TreeScope.Children, dlgCond);
            }
        }
    }
}
