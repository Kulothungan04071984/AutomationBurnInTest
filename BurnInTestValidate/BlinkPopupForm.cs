using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BurnInTestValidate
{
    public partial class BlinkPopupForm : Form
    {

        private int colorIndex = 0;
        private Color[] blinkColors =
        {
        Color.Green,
        };


        private Timer blinkTimer;
        private Timer closeTimer;
        public BlinkPopupForm(int autoCloseMs = 5000)
        {
            InitializeComponent();
            // Form style
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;
            this.ShowInTaskbar = false;

            rtbMessagenew.Clear();
            rtbMessagenew.ReadOnly = true;
            rtbMessagenew.BorderStyle = BorderStyle.None;
            rtbMessagenew.Font = new Font("Segoe UI", 30, FontStyle.Bold);
            rtbMessagenew.BackColor= Color.Black;
            rtbMessagenew.SelectionAlignment = HorizontalAlignment.Center;

            // Blink Timer
            blinkTimer = new Timer();
            blinkTimer.Interval = 500; // blink speed
            blinkTimer.Tick += timersecond_Tick;

            // Auto-close Timer
            closeTimer = new Timer();
            closeTimer.Interval = autoCloseMs;
            closeTimer.Tick += (s, e) => this.Close();
           // SetMessage(string.Empty);
        }
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            blinkTimer.Start();
            closeTimer.Start();
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            StopTimers();
            base.OnFormClosing(e);
        }
        private void timersecond_Tick(object sender, EventArgs e)
        {
            if (rtbMessagenew == null || rtbMessagenew.IsDisposed || !rtbMessagenew.IsHandleCreated)
                return;
            rtbMessagenew.SelectAll();
            rtbMessagenew.SelectionColor = blinkColors[colorIndex];
            colorIndex = (colorIndex + 1) % blinkColors.Length;
        }

        public void SetMessage(string message)
        {
            //message="Success : BurnIn Test Completed Successfully!";
            rtbMessagenew.Text = string.Empty;
            rtbMessagenew.Text = message;
        }
        private void StopTimers()
        {
            if (timersecond != null)
            {
                timersecond.Stop();
                timersecond.Tick -= timersecond_Tick;
                timersecond.Dispose();
                timersecond = null;
            }
        }

    }
}
