using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BurnInTestValidate
{
    public partial class FailBlink : Form
    {
        private bool isGreenVisible = true;
        private System.Windows.Forms.Timer blinkTimer;
        private CancellationTokenSource _cts;
        public string MessageTest = string.Empty;
        public System.Windows.Forms.Timer autoCloseTimer;
        public FailBlink()
        {
            InitializeComponent();
            rtbFail.Text = string.Empty;
            rtbFail.ReadOnly = true;
            rtbFail.BorderStyle = BorderStyle.None;
            rtbFail.BackColor = Color.Black;
            rtbFail.Font = new Font("Segoe UI", 30, FontStyle.Bold);

            // Blink timer
            timer1 = new System.Windows.Forms.Timer();
            timer1.Interval = 500; // blink speed
            timer1.Tick += timer1_Tick;
            timer1.Start();

        //    autoCloseTimer = new System.Windows.Forms.Timer();
        //    autoCloseTimer.Interval = 1000; // 1 second
        //    autoCloseTimer.Tick += AutoCloseTimer_Tick;
        //    autoCloseTimer.Start(); // <-- missing line
        }
        //private void AutoCloseTimer_Tick(object sender, EventArgs e)
        //{
        //    autoCloseTimer.Stop();
        //    this.Close();
        //}
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (rtbFail.IsDisposed) return;

            try
            {
                rtbFail.SelectAll();
                rtbFail.SelectionColor = isGreenVisible ? Color.Red : Color.Black;
                isGreenVisible = !isGreenVisible;
            }
            catch (ObjectDisposedException)
            {

            }
        }

        public void SafeUI(Action action)
        {
            if (this.IsDisposed || !this.IsHandleCreated)
                return; // form is closed, skip updates

            try
            {
                if (InvokeRequired)
                    Invoke(action);
                else
                    action();
            }
            catch (ObjectDisposedException)
            {

            }
        }

        public void AddText(string text, Color color = default, bool bold = true)
        {
            SafeUI(() =>
            {
                if (rtbFail.IsDisposed) return;  // extra safety

                rtbFail.Text = string.Empty;
                MessageTest = text;
                rtbFail.Clear();
                rtbFail.BackColor = Color.Black;

                var style = bold ? FontStyle.Bold : FontStyle.Regular;
                rtbFail.SelectionFont = new Font("Segoe UI", 30, style);
                rtbFail.SelectionColor = color;
                rtbFail.BackColor = Color.Black;
                rtbFail.ForeColor = Color.Red;

                int lines = rtbFail.Height / rtbFail.Font.Height;
                for (int i = 0; i < lines / 2; i++)
                    rtbFail.AppendText(Environment.NewLine);

                rtbFail.SelectionAlignment = HorizontalAlignment.Center;
                rtbFail.AppendText(text);
            });
            
        }
        private void btnOk_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Dispose();
            this.Close();
        }
        private void FailBlink_FormClosing(object sender, FormClosingEventArgs e)
        {
            _cts.Cancel();
        }

        private void FailBlink_Load(object sender, EventArgs e)
        {
            _cts = new CancellationTokenSource();

            Task.Run(() =>
            {
                while (!_cts.Token.IsCancellationRequested)
                {
                    AddText("PassMark Test Fail....", Color.Red);
                    Thread.Sleep(500);
                }
            });
        }
    }
}
