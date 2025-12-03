using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BurnInTestValidate
{
    public partial class DiskPartition : Form
    {
        public DiskPartition()
        {
            InitializeComponent();
            InitializeCustomControls();
        }

        private ComboBox cmbDisks;
        private ListBox lstPartitions;
        private TextBox txtSizeMB;
        private TextBox txtLabel;
        private ComboBox cmbFileSystem;
        private TextBox txtLetter;
        private Button btnRefresh;
        private Button btnCreate;
        private Button btnFormat;
        private Button btnAssign;
        private Button btnDelete;
        private RichTextBox rtbOutput;

        private void InitializeCustomControls()
        {
            this.Text = "Disk Partition Tool (diskpart via C#)";
            this.Width = 900;
            this.Height = 600;
            this.StartPosition = FormStartPosition.CenterScreen;

            // --- Layout ---
            var table = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 8 };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            this.Controls.Add(table);

            int row = 0;

            // Disk selection
            table.Controls.Add(new Label { Text = "Disk:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, row);
            cmbDisks = new ComboBox { Dock = DockStyle.Fill };
            table.Controls.Add(cmbDisks, 1, row);
            btnRefresh = new Button { Text = "Refresh", Width = 80 };
            btnRefresh.Click += async (s, e) => await LoadDisksAndPartitions();
            table.Controls.Add(btnRefresh, 2, row++);

            // Partitions list
            table.Controls.Add(new Label { Text = "Partitions:", Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft }, 0, row);
            lstPartitions = new ListBox { Dock = DockStyle.Fill };
            lstPartitions.SelectedIndexChanged += (s, e) => UpdateSelectedPartition();
            table.Controls.Add(lstPartitions, 1, row);
            table.SetRowSpan(lstPartitions, 4);
            table.Controls.Add(new Panel(), 2, row); row += 4;

            // Create partition
            table.Controls.Add(new Label { Text = "Size (MB):", Dock = DockStyle.Fill }, 0, row);
            txtSizeMB = new TextBox { Dock = DockStyle.Fill, Text = "500" };
            table.Controls.Add(txtSizeMB, 1, row);
            btnCreate = new Button { Text = "Create Primary", Dock = DockStyle.Fill };
            btnCreate.Click += async (s, e) => await CreatePartition();
            table.Controls.Add(btnCreate, 2, row++);

            // Format
            table.Controls.Add(new Label { Text = "FS:", Dock = DockStyle.Fill }, 0, row);
            cmbFileSystem = new ComboBox { Dock = DockStyle.Fill };
            cmbFileSystem.Items.AddRange(new[] { "NTFS", "FAT32", "exFAT" });
            cmbFileSystem.SelectedIndex = 0;
            table.Controls.Add(cmbFileSystem, 1, row);
            btnFormat = new Button { Text = "Format", Dock = DockStyle.Fill };
            btnFormat.Click += async (s, e) => await FormatPartition();
            table.Controls.Add(btnFormat, 2, row++);

            // Label
            table.Controls.Add(new Label { Text = "Label:", Dock = DockStyle.Fill }, 0, row);
            txtLabel = new TextBox { Dock = DockStyle.Fill };
            table.Controls.Add(txtLabel, 1, row);
            table.Controls.Add(new Panel(), 2, row++);

            // Assign letter
            table.Controls.Add(new Label { Text = "Letter:", Dock = DockStyle.Fill }, 0, row);
            txtLetter = new TextBox { Dock = DockStyle.Fill, Text = "E", MaxLength = 1 };
            txtLetter.TextChanged += (s, e) => txtLetter.Text = txtLetter.Text.ToUpper();
            table.Controls.Add(txtLetter, 1, row);
            btnAssign = new Button { Text = "Assign", Dock = DockStyle.Fill };
            btnAssign.Click += async (s, e) => await AssignLetter();
            table.Controls.Add(btnAssign, 2, row++);

            // Delete
            btnDelete = new Button { Text = "Delete Selected", Dock = DockStyle.Fill, BackColor = System.Drawing.Color.IndianRed };
            btnDelete.Click += async (s, e) => await DeletePartition();
            table.SetColumnSpan(btnDelete, 3);
            table.Controls.Add(btnDelete, 0, row++);

            // Output
            rtbOutput = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 9) };
            table.SetRowSpan(rtbOutput, 3);
            table.Controls.Add(rtbOutput, 0, row);
            table.SetColumnSpan(rtbOutput, 3);

            // Load disks on start
            this.Load += async (s, e) => await LoadDisksAndPartitions();
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

        private async Task LoadDisksAndPartitions()
        {
            rtbOutput.Clear();
            cmbDisks.Items.Clear();
            lstPartitions.Items.Clear();

            string output = await RunDiskpart("list disk\nlist partition\n");
            rtbOutput.Text = output;

            // Parse disks
            var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                if (line.StartsWith("  Disk "))
                {
                    var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (int.TryParse(parts[1], out int diskId))
                    {
                        string size = parts.Length > 5 ? $"{parts[5]} {parts[6]}" : "";
                        cmbDisks.Items.Add(new DiskInfo(diskId, size));
                    }
                }
            }
            if (cmbDisks.Items.Count > 0) cmbDisks.SelectedIndex = 0;
        }

        private void UpdateSelectedPartition()
        {
            // No action needed – just enable/disable buttons
            btnFormat.Enabled = lstPartitions.SelectedItem is PartitionInfo;
            btnAssign.Enabled = lstPartitions.SelectedItem is PartitionInfo;
            btnDelete.Enabled = lstPartitions.SelectedItem is PartitionInfo;
        }

        private async Task CreatePartition()
        {
            var disk = cmbDisks.SelectedItem as DiskInfo;
            if (disk == null) return;

            if (!int.TryParse(txtSizeMB.Text, out int size) || size <= 0)
            {
                MessageBox.Show("Enter valid size in MB");
                return;
            }

            string script = $"select disk {disk.Id}\ncreate partition primary size={size}\n";
            string res = await RunDiskpart(script);
            rtbOutput.AppendText($"\n--- CREATE ---\n{res}\n");
            await LoadDisksAndPartitions();
        }

        private async Task FormatPartition()
        {
            var part = lstPartitions.SelectedItem as PartitionInfo;
            if (part == null) return;

            string fs = cmbFileSystem.SelectedItem?.ToString() ?? "NTFS";
            string label = string.IsNullOrWhiteSpace(txtLabel.Text) ? "" : $"label=\"{txtLabel.Text}\"";
            string script = $"select disk {part.DiskId}\nselect partition {part.Number}\nformat fs={fs} quick {label}\n";
            string res = await RunDiskpart(script);
            rtbOutput.AppendText($"\n--- FORMAT ---\n{res}\n");
        }

        private async Task AssignLetter()
        {
            var part = lstPartitions.SelectedItem as PartitionInfo;
            if (part == null) return;

            if (string.IsNullOrWhiteSpace(txtLetter.Text))
            {
                MessageBox.Show("Enter a letter");
                return;
            }

            char letter = txtLetter.Text.Trim().ToUpper()[0];
            if (letter < 'D' || letter > 'Z')
            {
                MessageBox.Show("Letter D-Z only");
                return;
            }

            string script = $"select disk {part.DiskId}\nselect partition {part.Number}\nassign letter={letter}\n";
            string res = await RunDiskpart(script);
            rtbOutput.AppendText($"\n--- ASSIGN {letter}: ---\n{res}\n");
        }

        private async Task DeletePartition()
        {
            var part = lstPartitions.SelectedItem as PartitionInfo;
            if (part == null) return;

            var confirm = MessageBox.Show(
                $"Delete partition {part.Number} on Disk {part.DiskId}?\nALL DATA WILL BE LOST!",
                "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (confirm != DialogResult.Yes) return;

            string script = $"select disk {part.DiskId}\nselect partition {part.Number}\ndelete partition override\n";
            string res = await RunDiskpart(script);
            rtbOutput.AppendText($"\n--- DELETE ---\n{res}\n");
            await LoadDisksAndPartitions();
        }

        // --------------------------------------------------------------
        // Helper classes
        // --------------------------------------------------------------
        private class DiskInfo
        {
            public int Id { get; }
            public string Size { get; }
            public DiskInfo(int id, string size) { Id = id; Size = size; }
            public override string ToString() => $"Disk {Id} ({Size})";
        }

        private class PartitionInfo
        {
            public int DiskId { get; }
            public int Number { get; }
            public string Size { get; }
            public string Type { get; }
            public PartitionInfo(int diskId, int number, string size, string type)
            {
                DiskId = diskId; Number = number; Size = size; Type = type;
            }
            public override string ToString() => $"Partition {Number} | {Size} | {Type}";
        }

        private async void cmbDisks_SelectedIndexChanged(object sender, EventArgs e)
        {
           var disk = cmbDisks.SelectedItem as DiskInfo;
    if (disk == null) return;

    lstPartitions.Items.Clear();

    string output = await RunDiskpart($"select disk {disk.Id}\nlist partition\n");
            var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            bool inPartition = false;
            foreach (var line in lines)
            {
                if (line.Contains("Partition ###")) { inPartition = true; continue; }
                if (!inPartition) continue;
                if (line.Trim() == "") continue;

                var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 5 && int.TryParse(parts[1], out int num))
                {
                    string size = parts[2] + " " + parts[3];
                    string type = string.Join(" ", parts, 5, parts.Length - 5);
                    lstPartitions.Items.Add(new PartitionInfo(disk.Id, num, size, type));
                }
            }


        }
    }
}
