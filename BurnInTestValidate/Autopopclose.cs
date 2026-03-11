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
    public partial class Autopopclose : Form
    {
        private System.Windows.Forms.Timer closeTimer;
        private int closeAfterMs = 3000;
        public Autopopclose(int autoCloseMilliseconds = 3000)
        {
            InitializeComponent();
            closeAfterMs = autoCloseMilliseconds;

            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.StartPosition = FormStartPosition.Manual;
            this.ShowInTaskbar = false;
            this.TopMost = true;

            rtbMessage.ReadOnly = true;
            rtbMessage.BorderStyle = BorderStyle.None;

            closeTimer = new System.Windows.Forms.Timer();
            closeTimer.Interval = closeAfterMs;
            closeTimer.Tick += timer_Tick;
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            for (double i = 1.0; i >= 0; i -= 0.1)
            {
                this.Opacity = i;
                Application.DoEvents();
                Thread.Sleep(30);
            }
            closeTimer.Stop();
            this.Close();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            closeTimer.Start();
        }

        public void AddText(string text, Color color, bool bold = false)
        {
            rtbMessage.SelectionStart = rtbMessage.TextLength;
            rtbMessage.SelectionLength = 0;
            rtbMessage.SelectionColor = color;
            rtbMessage.SelectionFont = new Font(
                rtbMessage.Font,
                bold ? FontStyle.Bold : FontStyle.Regular);
            rtbMessage.AppendText(text);
        }
    }
}
