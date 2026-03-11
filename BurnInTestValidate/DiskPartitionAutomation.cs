using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace BurnInTestValidate
{
    public partial class DiskPartitionAutomation : Form
    {
        private readonly FrmBurnIntest _formAuto;

        public DiskPartitionAutomation()
        {
            InitializeComponent();
            InitializeCustomControls();
          
        }

        public DiskPartitionAutomation(FrmBurnIntest formAuto)
        {
            _formAuto = formAuto;
        }


        private System.Windows.Forms.ComboBox cmbDisks;
        private System.Windows.Forms.ListBox lstPartitions;
        private System.Windows.Forms.TextBox txtSizeMB, txtLabel, txtLetter;
        private System.Windows.Forms.ComboBox cmbFileSystem;
        private System.Windows.Forms.Button btnRefresh, btnCreate, btnFormat, btnAssign, btnDelete, btnRunAutomation;
        private RichTextBox rtbOutput;

        
        private void InitializeCustomControls()
        {
            this.Text = "Disk Partition Tool + FlaUI Auto";
            this.Width = 950;
            this.Height = 650;
            this.StartPosition = FormStartPosition.CenterScreen;

            var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 9 };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            this.Controls.Add(table);

            int row = 0;

            // Disk
             table.Controls.Add(new System.Windows.Forms.Label { Text = "Disk:", TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, row);
          
            cmbDisks = new System.Windows.Forms.ComboBox { Dock = DockStyle.Fill, Name = "cmbDisks" };
            table.Controls.Add(cmbDisks, 1, row);
            btnRefresh = new System.Windows.Forms.Button { Text = "Refresh", Name = "btnRefresh" };
            btnRefresh.Click += async (s, e) => await LoadDisksAndPartitions();
            table.Controls.Add(btnRefresh, 2, row++);

            // Partitions
            //table.Controls.Add(new System.Windows.Forms.Label { Text = "Partitions:", TextAlign = ContentAlignment.MiddleLeft }, 0, row);
            table.Controls.Add(
    new System.Windows.Forms.Label
    {
        Text = "Partitions:",
        TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    },
    0,
    row
    
);
            lstPartitions = new System.Windows.Forms.ListBox { Dock = DockStyle.Fill, Name = "lstPartitions" };
            lstPartitions.SelectedIndexChanged += (s, e) => UpdateButtons();
            table.Controls.Add(lstPartitions, 1, row);
            table.SetRowSpan(lstPartitions, 5); row += 5;

            // Size
            table.Controls.Add(new System.Windows.Forms.Label { Text = "Size (MB):" }, 0, row);
            txtSizeMB = new System.Windows.Forms.TextBox { Dock = DockStyle.Fill, Text = "192400", Name = "txtSizeMB" };
            table.Controls.Add(txtSizeMB, 1, row);
            btnCreate = new System.Windows.Forms.Button { Text = "Create Primary", Name = "btnCreate" };
            btnCreate.Click += async (s, e) => await CreatePartition();
            table.Controls.Add(btnCreate, 2, row++);

            // FS + Label
            table.Controls.Add(new System.Windows.Forms.Label { Text = "File System:" }, 0, row);
            cmbFileSystem = new System.Windows.Forms.ComboBox { Dock = DockStyle.Fill, Name = "cmbFileSystem" };
            cmbFileSystem.Items.AddRange(new[] { "NTFS", "FAT32", "exFAT" });
            cmbFileSystem.SelectedIndex = 0;
            table.Controls.Add(cmbFileSystem, 1, row);
            btnFormat = new System.Windows.Forms.Button { Text = "Format", Name = "btnFormat" };
            btnFormat.Click += async (s, e) => await FormatPartition();
            table.Controls.Add(btnFormat, 2, row++);

            table.Controls.Add(new System.Windows.Forms.Label { Text = "Label:" }, 0, row);
            txtLabel = new System.Windows.Forms.TextBox { Dock = DockStyle.Fill, Name = "txtLabel" };
            table.Controls.Add(txtLabel, 1, row++);
            table.Controls.Add(new Panel(), 2, row - 1);

            // Letter
            table.Controls.Add(new System.Windows.Forms.Label { Text = "Letter:" }, 0, row);
            txtLetter = new System.Windows.Forms.TextBox { Dock = DockStyle.Fill, Text = "D", MaxLength = 1, Name = "txtLetter" };
            txtLetter.TextChanged += (s, e) => txtLetter.Text = txtLetter.Text.ToUpper();
            table.Controls.Add(txtLetter, 1, row);
            btnAssign = new System.Windows.Forms.Button { Text = "Assign", Name = "btnAssign" };
            btnAssign.Click += async (s, e) => await AssignLetter();
            table.Controls.Add(btnAssign, 2, row++);

            // Delete
            btnDelete = new System.Windows.Forms.Button { Text = "Delete Selected", BackColor = System.Drawing.Color.IndianRed, Name = "btnDelete" };
            btnDelete.Click += async (s, e) => await DeletePartition();
            table.SetColumnSpan(btnDelete, 3);
            table.Controls.Add(btnDelete, 0, row++);

            // === AUTOMATION BUTTON ===
            btnRunAutomation = new System.Windows.Forms.Button
            {
                Text = "RUN FLAUI AUTOMATION",
                BackColor = System.Drawing.Color.MediumSeaGreen,
                ForeColor = System.Drawing.Color.White,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                Dock = DockStyle.Fill,
                Name = "btnRunAutomation"
            };
            btnRunAutomation.Click += (s, e) => Task.Run(() => RunFlaUIAutomation());
            table.SetColumnSpan(btnRunAutomation, 3);
            table.Controls.Add(btnRunAutomation, 0, row++);

            // Output
            rtbOutput = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 9), Name = "rtbOutput" };
            table.SetRowSpan(rtbOutput, 3);
            table.Controls.Add(rtbOutput, 0, row);
            table.SetColumnSpan(rtbOutput, 3);

            this.Load += async (s, e) => await LoadDisksAndPartitions();
        }

        private void UpdateButtons()
        {
            bool hasSelection = lstPartitions.SelectedItem is PartitionInfo;
            btnFormat.Enabled = hasSelection;
            btnAssign.Enabled = hasSelection;
            btnDelete.Enabled = hasSelection;
        }

        private async Task<string> RunDiskpart(string script)
        {
            string temp = Path.Combine(Path.GetTempPath(), $"dp_{Guid.NewGuid()}.txt");
            await Task.Run(() => File.WriteAllText(temp, script, Encoding.ASCII));

            var psi = new ProcessStartInfo
            {
                FileName = "diskpart.exe",
                Arguments = $"/s \"{temp}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.GetEncoding(437)
            };

            var p = new Process { StartInfo = psi };
            var sb = new StringBuilder();
            p.OutputDataReceived += (s, e) => { if (e.Data != null) sb.AppendLine(e.Data); };
            p.ErrorDataReceived += (s, e) => { if (e.Data != null) sb.AppendLine("ERR: " + e.Data); };

            p.Start();
            p.BeginOutputReadLine();
            p.BeginErrorReadLine();
            await Task.Run(() => p.WaitForExit());

            try { File.Delete(temp); } catch { }
            return p.ExitCode == 0 ? sb.ToString() : $"[FAILED] Code {p.ExitCode}\n{sb}";
        }

        private async Task LoadDisksAndPartitions() { /* ... same as before ... */ }
        private async Task CreatePartition() { /* ... */ }
        private async Task FormatPartition() { /* ... */ }
        private async Task AssignLetter() { /* ... */ }
        private async Task DeletePartition() { /* ... */ }

        private void RunFlaUIAutomation()
        {
            try
            {
                _formAuto.writeErrorMessage("Start Disk Partion", "strat");
                AppendLog("Starting FlaUI Automation...");

                var automation = new UIA3Automation();
                var window = GetCurrentWindow(automation);

                // Step 1: Select Disk 1
                SelectDisk(window, 0);
                CreatePartition(window, 192400);
                Thread.Sleep(2000);

                // Step 2: Refresh
                ClickButton(window, "btnRefresh");
                Thread.Sleep(1500);

                // Step 3: Get new partition
                int partNum = GetLatestPartition(window);
                SelectPartition(window, partNum);

                // Step 4: Format
                SetComboBox(window, "cmbFileSystem", "NTFS");
                SetTextBox(window, "txtLabel", "AutoData");
                ClickButton(window, "btnFormat");
                Thread.Sleep(2000);

                // Step 5: Assign E:
                SetTextBox(window, "txtLetter", "D");
                ClickButton(window, "btnAssign");

                AppendLog("Automation Completed!");
            }
            catch (Exception ex)
            {
                AppendLog($"ERROR: {ex.Message}");
            }
        }

        private Window GetCurrentWindow(UIA3Automation automation)
        {
            var process = Process.GetCurrentProcess();
            return automation.GetDesktop().FindFirstDescendant(cf => cf.ByProcessId(process.Id))?.AsWindow()
                   ?? throw new Exception("Window not found");
        }

        private void SelectDisk(Window window, int diskId)
        {
            var combo = window.FindFirstDescendant(cf => cf.ByName("cmbDisks"))?.AsComboBox();
            combo?.Expand();
            Thread.Sleep(300);
            var item = combo?.FindFirstDescendant(cf => cf.ByText($"Disk {diskId}"));
            item?.Click();
            combo?.Collapse();
            AppendLog($"Selected Disk {diskId}");
        }

        private void CreatePartition(Window window, int sizeMB)
        {
            SetTextBox(window, "txtSizeMB", sizeMB.ToString());
            ClickButton(window, "btnCreate");
            AppendLog($"Creating {sizeMB} MB partition...");
        }

        private int GetLatestPartition(Window window)
        {
            var listBox = window.FindFirstDescendant(cf => cf.ByName("lstPartitions"))?.AsListBox();
            var items = listBox?.Items;
            if (items == null || items.Length == 0) throw new Exception("No partitions");
            var text = items[items.Length - 1].Name;
            var num = int.Parse(text.Split('|')[0].Replace("Partition", "").Trim());
            AppendLog($"Latest partition: {num}");
            return num;
        }

        private void SelectPartition(Window window, int num)
        {
            var list = window.FindFirstDescendant(cf => cf.ByName("lstPartitions"))?.AsListBox();
            var item = list?.FindFirstDescendant(cf => cf.ByText($"Partition {num}"));
            item?.Click();
        }

        private void SetTextBox(Window window, string name, string text)
        {
            var tb = window.FindFirstDescendant(cf => cf.ByName(name))?.AsTextBox();
            tb?.Enter(text);
        }

        private void SetComboBox(Window window, string name, string value)
        {
            var cb = window.FindFirstDescendant(cf => cf.ByName(name))?.AsComboBox();
            cb?.Select(value);
        }

        private void ClickButton(Window window, string name)
        {
            var btn = window.FindFirstDescendant(cf => cf.ByName(name))?.AsButton();
            btn?.Click();
        }

        private void AppendLog(string text)
        {
            if (rtbOutput.InvokeRequired)
            {
                rtbOutput.Invoke(new Action(() => rtbOutput.AppendText($"\n[{DateTime.Now:HH:mm:ss}] {text}")));
            }
            else
            {
                rtbOutput.AppendText($"\n[{DateTime.Now:HH:mm:ss}] {text}");
            }
        }

        private class DiskInfo { public int Id; public string Size; public override string ToString() => $"Disk {Id} ({Size})"; }
        private class PartitionInfo { public int DiskId, Number; public string Size, Type; public override string ToString() => $"Partition {Number} | {Size} | {Type}"; }
    }
}
