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
using Timer = System.Windows.Forms.Timer;

namespace BurnInTestValidate
{
    public partial class Blink : Form
    {
        private bool isGreenVisible = true;
        private Timer blinkTimer;
        private CancellationTokenSource _cts;
        public Blink()
        {
            InitializeComponent();
            rtbBlinkw.ReadOnly = true;
            rtbBlinkw.BorderStyle = BorderStyle.None;
            rtbBlinkw.BackColor = Color.Black;
            rtbBlinkw.Font = new Font("Segoe UI", 30, FontStyle.Bold);

            // Blink timer
            blinkTimer = new Timer();
            blinkTimer.Interval = 500; // blink speed
            blinkTimer.Tick += timerBlink_Tick;
            blinkTimer.Start();

        }

        private void timerBlink_Tick(object sender, EventArgs e)
        {
            if (rtbBlinkw.IsDisposed) return;

            try
            {
                rtbBlinkw.SelectAll();
                rtbBlinkw.SelectionColor = isGreenVisible ? Color.Green : Color.Black;
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
                if (rtbBlinkw.IsDisposed) return;  // extra safety

                rtbBlinkw.Clear();
                rtbBlinkw.BackColor = Color.Black;

                var style = bold ? FontStyle.Bold : FontStyle.Regular;
                rtbBlinkw.SelectionFont = new Font("Segoe UI", 30, style);
                rtbBlinkw.SelectionColor = color;
                rtbBlinkw.BackColor = Color.Black;
                rtbBlinkw.ForeColor = Color.Green;

                int lines = rtbBlinkw.Height / rtbBlinkw.Font.Height;
                for (int i = 0; i < lines / 2; i++)
                    rtbBlinkw.AppendText(Environment.NewLine);

                rtbBlinkw.SelectionAlignment = HorizontalAlignment.Center;
                rtbBlinkw.AppendText(text);
            });
            //SafeUI(() =>
            //{
            //    if (rtbBlinkw.IsDisposed) return;

            //    rtbBlinkw.SelectionStart = rtbBlinkw.TextLength;
            //    rtbBlinkw.SelectionLength = 0;
            //    rtbBlinkw.SelectionColor = color;
            //    rtbBlinkw.SelectionFont = new Font("Segoe UI", 30, bold ? FontStyle.Bold : FontStyle.Regular);
            //    rtbBlinkw.SelectionAlignment = HorizontalAlignment.Center;
            //    rtbBlinkw.AppendText(text + Environment.NewLine);
            //    rtbBlinkw.ScrollToCaret();
            //});
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            blinkTimer.Stop();
            blinkTimer.Dispose();
            this.Close();
        }
        private void frmBlink_FormClosing(object sender, FormClosingEventArgs e)
        {
            _cts.Cancel();
        }
        private void Blink_Load(object sender, EventArgs e)
        {
            _cts = new CancellationTokenSource();

            Task.Run(() =>
            {
                while (!_cts.Token.IsCancellationRequested)
                {
                    AddText("Pass Mark Test Completed...", Color.Green);
                    Thread.Sleep(500);
                }
            });
        }
    }
}
