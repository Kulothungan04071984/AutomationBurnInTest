using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.Http;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Ink;
using System.Xml.Linq;
using Application = FlaUI.Core.Application;
using Menu = FlaUI.Core.AutomationElements.Menu;
using MenuItem = FlaUI.Core.AutomationElements.MenuItem;
using System.Management;
using FlaUI.Core.Conditions;


namespace BurnInTestValidate
{
    public partial class FrmBurnIntest : Form
    {
        string exePath = string.Empty;
        int PartitionCount = 0;
        public FrmBurnIntest()
        {
            InitializeComponent();
            SetupUI();

        }


        private void SetupUI()
        {
            this.Text = "Burn In Test Automator";
            this.Size = new System.Drawing.Size(700, 500);
            this.StartPosition = FormStartPosition.CenterParent;

            // Start Button
            var btnStart = new System.Windows.Forms.Button
            {
                Text = "Start Automation",
                Size = new System.Drawing.Size(150, 40),
                Location = new System.Drawing.Point(450, 20)
            };
            btnStart.Click += btnStart_Click;
            this.Controls.Add(btnStart);

            // Log Box
            var rtbLog = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9),
                Location = new System.Drawing.Point(20, 80)
            };
            this.Controls.Add(rtbLog);

            // Store for use
            this.Tag = rtbLog;
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            var btn = (System.Windows.Forms.Button)sender;
            btn.Enabled = false;
            var log = (RichTextBox)this.Tag;

            try
            {
                await Task.Run(() => RunAutomation(log));
            }
            catch (Exception ex)
            {
                writeErrorMessage(ex.Message.ToString(), "btnStart_Click");
                Log(log, $"ERROR: {ex.Message}", System.Drawing.Color.Red);
            }
            finally
            {
                btn.Enabled = true;
            }
        }
        static string Safe(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return "(empty)";
            return s;
        }

   

        public async Task<string> DiskPartitionDynamic_NoWMI(RichTextBox log)
        {
            try
            {
                Log(log, "Starting auto disk partitioning (Disk 1 → D:, Disk 2 → E:, etc.)\r\n");

                char driveLetter = 'D';
                int diskIndex = 1;

                while (driveLetter <= 'Z')
                {
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

                    Log(log, $"Trying Disk {diskIndex} → {driveLetter}: ... ");

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
                    }
                    else if (fullOutput.Contains("diskpart succeeded") ||
                             fullOutput.Contains("100 percent completed") ||
                             fullOutput.Contains("successfully formatted") ||
                             fullOutput.Contains("assigned the drive letter"))
                    {
                        success = true;
                    }

                    // --- Now simple if/else (no tuples!) ---
                    if (!diskExists)
                    {
                        Log(log, "No more disks found.\r\n", Color.Orange);
                        break;
                    }

                    if (success)
                    {
                        Log(log, $"SUCCESS → {driveLetter}:\r\n", Color.LimeGreen);
                        driveLetter++;
                        diskIndex++;
                    }
                    else
                    {
                        Log(log, $"Skipped (already partitioned or failed)\r\n", Color.Yellow);
                        diskIndex++; // Try next disk anyway
                    }

                    await Task.Delay(3000); // Let disk settle
                }

                Log(log, $"Finished. Last assigned letter: {(char)(driveLetter - 1)}:\r\n", Color.Cyan);
                Log(log,"Disk partitioning completed!Success",System.Drawing.Color.Green);
                return "true";
            }
            catch (Exception ex)
            {
                Log(log, $"Error: {ex.Message}\r\n", Color.Red);
                //MessageBox.Show("Error: " + ex.Message, "Error");
                return "false";
            }
        }
        private async void RunAutomation(RichTextBox log)
        {
            Log(log, "Starting automation...");
            Log(log, "Disk Partition Start");

            //Testing
            var status = await DiskPartitionDynamic_NoWMI(log);
            if (status == "false") return;



            Log(log, "Starting automation BurnIn Test...");

            //BurnIn Test Start
        
       
           exePath= ConfigurationManager.AppSettings["BurnInTest"];
            if (!File.Exists(exePath))
            {
                writeErrorMessage("File path Not Exists -", exePath.ToString());
                Log(log, $"EXE not found: {exePath}", System.Drawing.Color.Red);
                return;
            }
            writeErrorMessage("File path Exists -", exePath.ToString());
         


            Application app = null;
            try
            {
               app = LaunchWithAdmin(exePath);
                Thread.Sleep(1000);
                using (var automation = new UIA3Automation())
                {
                    Console.WriteLine("=== All Open Window Names ===");

                    var desktop = automation.GetDesktop();
                    var allWindows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));
                  
                    string drivemsg = drivecheck(log, allWindows);

                    Log(log,drivemsg,System.Drawing.Color.DarkBlue);
                   
                    Thread.Sleep(2000);

                    var mainWindow = desktop.FindFirstDescendant(cf =>
                 cf.ByControlType(ControlType.Window)
                   .And(cf.ByName("BurnInTest V8.1 Pro (1006)")))
                 ?.AsWindow();

                    if (mainWindow == null)
                    {
                        Log(log, "Main window not found", System.Drawing.Color.Red);

                        return;
                    }

                    mainWindow.Focus();

                    Thread.Sleep(1000);

                    var menuBar =mainWindow.FindFirstDescendant(cf => cf.ByControlType(ControlType.MenuBar))?.AsMenu();
                    if (menuBar == null)
                    {
                        Log(log, "Menu bar not found", System.Drawing.Color.Red);

                        return;
                    }
                    System.Threading.Thread.Sleep(1500);
                    var configMenu = mainWindow.FindFirstDescendant(cf => cf.ByName("Configuration"))?.AsMenuItem();
                    if (configMenu == null)
                    {
                        Log(log, "Configuration menu not found", System.Drawing.Color.Red);

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
                                Log(log, "BurnInTest Preferences window not found.");
                                return;
                            }
                            prefWindow.Focus();
                            Log(log, "BurnInTest Preferences window Focus.");

                            Thread.Sleep(500);

                            var checkbox = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1224"))?.AsCheckBox();

                            if (checkbox == null)
                            {
                                Log(log, "Checkbox not found (check AutomationId-)"+ checkbox.AutomationId + "-" + checkbox.Name);
                                return;
                            }
                            checkbox.IsChecked = true;
                            Thread.Sleep(500);
                            checkbox.IsChecked = false;
                         

                            for (int i = 0; i < 5; i++)
                            {
                               var lstC = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("ListViewItem-" + i))?.AsListBoxItem();
                                if(lstC == null)
                                {
                                    Log(log, "List Item Not Fount-" + i.ToString());
                                }
                               string lstName = lstC.Name.ToString();
                                string checkCDrive = lstName.Substring(0, 2);
                                if(checkCDrive == "C:")
                                {
                                    lstC.Select();

                                    var checkC = prefWindow.FindFirstDescendant(cf => cf.ByAutomationId("1223"))?.AsCheckBox();

                                    if (checkC == null)
                                    {
                                        Log(log, "ListView C: not found .");
                                        return;
                                    }
                                    Log(log, "ListView C: found");
                                    checkC.IsChecked = false;
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
                            }

                            configMenu?.Click();

                            Log(log, "Configuration menu Clicked", System.Drawing.Color.Green);
                            System.Threading.Thread.Sleep(2000);


                            var testPrefnext = w.FindFirstDescendant(cf => cf.ByName("Test Selection && Duty Cycles..."));

                            if (testPrefnext == null)
                            {

                                Log(log, "Test Selection & Duty Cycles not found-", System.Drawing.Color.Green);

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
                                Console.WriteLine("Test selection and duty cycles window not found.");
                                return;
                            }

                            prefWindowcycles.Focus();
                            Console.WriteLine("Test selection and duty cycles window found.");

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
                                        Log(log, "Disk Checkbox checked successfully ✅");
                                    }
                                    else
                                    {
                                        Log(log, "Disk Checkbox was already checked ✅");
                                    }

                                }


                            }

                            Log(log, "Disk Checkbox checked completed");
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
                                Console.WriteLine("warning Getting ready to run Burn in tests window not found.");
                                return;
                            }


                            prefWindowcyclesWarning.Focus();
                            // Testing
                            var btnOkwarning = prefWindowcyclesWarning.FindFirstDescendant(cf => cf.ByName("OK"))?.AsButton();
                            if (btnOkwarning != null)
                                btnOkwarning.Invoke();


                            Log(log, "Task Completed");

                            Thread.Sleep(3000);

                            Log(log, "Start Crystal DiskMark");


                            var Crystalpath = ConfigurationManager.AppSettings["Crystal"];
                            if (!File.Exists(Crystalpath))
                            { 
                                Log(log, "Crystal DiskMark Not Found");
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
                                Log(log, "CrystalDiskMark Window Not Found");
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
                                Log(log, $"❌ ComboBox with AutomationId '{comboAutomationId}' not found.");
                                return;
                            }


                            var combo = comboElement.AsComboBox();
                            combo?.Expand();
                            Thread.Sleep(500); // allow items to appear

                            Log(log, $"✅ ComboBox Found: {comboAutomationId}");
                            Log(log, "----------------------------------------------------");

                            // 🔹 List all dropdown values
                            if (combo?.Items != null && combo.Items.Length > 0)
                            {
                                Log(log, "Available Items:");
                                foreach (var item in combo.Items)
                                {
                                    if (item.Name == "D: 18% (42/232GiB)" || item.Name.Contains("D:"))
                                    {
                                        combo.Select(item.Name);
                                        Log(log, "  • " + item.Text);
                                        break;
                                    }
                                }
                                if (combo.Items.Length >= 3)
                                    PartitionCount = 1;
                            }
                            else
                            {
                                Log(log, "⚠️ No items found or combo not expandable.");
                            }

                            // 🔹 Show currently selected item
                            if (combo?.SelectedItem != null)
                                Log(log, $"Selected: {combo.SelectedItem.Text}");
                            else
                                Log(log, "No item currently selected.");

                            combo?.Collapse();

                            //All Ok Button
                            var btnAll = mainWindowCrystal.FindFirstDescendant(cf => cf.ByName("All"))?.AsButton();
                            if (btnAll == null)
                            {
                                Log(log, "Button All Not Found");
                                return;
                            }
                            btnAll.Invoke();
                           
                            Log(log, "D: crystal Report Started.");
                          
                            Application appnew = null;
                          
                            if (PartitionCount > 0 )
                            {
                                Log(log, "Check Next crystal Report Entry.");
                                appnew = LaunchWithAdmin(Crystalpath);
                                string comboAutomationId_1 = "1027";
                                System.Threading.Thread.Sleep(2000);
                                var emainWindowCrystal = desktop.FindFirstDescendant(cf =>
cf.ByControlType(ControlType.Window)
.And(cf.ByName("CrystalDiskMark 8.0.1 x86 [Admin]")))
?.AsWindow();
                                if (emainWindowCrystal == null)
                                {
                                    Log(log, $"❌ Second window not found.");
                                    return;
                                }
                                emainWindowCrystal.Focus();

                                var comboElement_1 = emainWindowCrystal.FindFirstDescendant(cf =>
                            cf.ByAutomationId(comboAutomationId_1)
                              .And(cf.ByControlType(ControlType.ComboBox)));
                                if (comboElement_1 == null)
                                {
                                    Log(log, $"❌ ComboBox with AutomationId '{comboAutomationId_1}' not found.");
                                    appnew.Close();
                                    return;
                                }


                                var combo1 = comboElement_1.AsComboBox();

                                combo1?.Expand();
                                //Thread.Sleep(300); // allow items to appear

                                //Log(log, $"✅ ComboBox Found: {comboAutomationId_1}");
                                ////Log(log, "----------------------------------------------------");

                                //Log(log,"Second list --" + combo1.Items.Count().ToString());
                                Thread.Sleep(300);
                                // 🔹 List all dropdown values
                                if (combo1?.Items != null && combo1.Items.Length > 0)
                                {
                                    bool eStatus = false;
                                    foreach (var item in combo1.Items)
                                    {
                                        Log(log, "Available Items:" + item.Name);
                                        if (item.Name.Contains("E:"))
                                        {
                                            Log(log, "Found E: Drive ");
                                            combo1.Select(item.Name);
                                            Log(log, "  • " + item.Text);
                                            eStatus = true;
                                            break;
                                        }
                                    }

                                    if(!eStatus)
                                    {
                                        Log(log, "E: Drive found.");
                                        appnew.Close(true);
                                        return;
                                    }
                                    
                                    if (combo1.Items.Length > 4)
                                        PartitionCount = 2;
                                }
                                else
                                {
                                    Log(log, "⚠️ No items found or combo not expandable.");
                                }

                                // 🔹 Show currently selected item
                                if (combo1?.SelectedItem != null)
                                    Log(log, $"Selected: {combo1.SelectedItem.Text}");
                                else
                                    Log(log, "No item currently selected.");


                                combo1?.Collapse();



                                //All Ok Button
                                var btnAll_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByName("All"))?.AsButton();
                                if (btnAll_1 == null)
                                {
                                    Log(log, "Button All Not Found");
                                    return;
                                }
                                btnAll_1.Invoke();

                                Thread.Sleep(100000);
                                var ftxt_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1009"));
                                if (ftxt_1 != null)
                                {
                                    var fval_1 = ftxt_1.ToString().Split('.');
                                    if (fval_1.Length > 0)
                                    {
                                        if (fval_1[0].Length > 3)
                                            Log(log, "Crystal DiskMark Pass -Read");
                                        else
                                            Log(log, "Crystal DiskMark Fail - Read");
                                    }

                                }
                                Log(log, "First Text Box-" + ftxt_1.Name);

                                var Stxt_1 = emainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1014"));
                                if (Stxt_1 != null)
                                {
                                    var Sval_1 = Stxt_1.ToString().Split('.');
                                    if (Sval_1.Length > 0)
                                    {
                                        if (Sval_1[0].Length > 3)
                                            Log(log, "Crystal DiskMark Pass -write");
                                        else
                                            Log(log, "Crystal DiskMark Fail - write");
                                    }

                                }
                                Log(log, "Second Text Box-" + Stxt_1.Name);

                                //first Crystal Report stage

                                var ftxt = mainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1009"));
                                if (ftxt != null)
                                {
                                    var fval = ftxt.ToString().Split('.');
                                    if (fval.Length > 0)
                                    {
                                        if (fval[0].Length > 3)
                                            Log(log, "Crystal DiskMark Pass -Read");
                                        else
                                            Log(log, "Crystal DiskMark Fail - Read");
                                    }

                                }
                                Log(log, "First Text Box-" + ftxt.Name);

                                var Stxt = mainWindowCrystal.FindFirstDescendant(cf => cf.ByAutomationId("1014"));
                                if (Stxt != null)
                                {
                                    var Sval = Stxt.ToString().Split('.');
                                    if (Sval.Length > 0)
                                    {
                                        if (Sval[0].Length > 3)
                                            Log(log, "Crystal DiskMark Pass -write");
                                        else
                                            Log(log, "Crystal DiskMark Fail - write");
                                    }

                                }
                                Log(log, "Second Text Box-" + Stxt.Name);

                               // appnew.Close();
                            }
                           
                            PartitionCount = 0;

                          //  app.Close();
                            


                            break;
                        }
                        catch (Exception ex)
                        {
                            Log(log, "error-" + ex.Message.ToString(), System.Drawing.Color.Red);
                            writeErrorMessage(ex.Message.ToString(), "Crystal DiskMark");
                            break;
                        }
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
                    Log(log, "D:\\ window not found");
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
                    Log(log, "E:\\ window not found");
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

        public void Log(RichTextBox rtb, string message, System.Drawing.Color? color = null)
        {
            this.Invoke((MethodInvoker)delegate
            {
                rtb.SelectionColor = color ?? System.Drawing.Color.Black;
                rtb.AppendText($"{DateTime.Now:HH:mm:ss} | {message}\r\n");
                rtb.ScrollToCaret();
            });
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
        private void FrmBurnIntest_Load(object sender, EventArgs e)
        {

        }
    }
}
