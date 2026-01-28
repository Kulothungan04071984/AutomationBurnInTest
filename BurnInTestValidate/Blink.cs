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
    public partial class Blink : Form
    {
        private bool isGreenVisible = true;
        private Timer blinkTimer;

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
            rtbBlinkw.SelectAll();
            rtbBlinkw.SelectionColor = isGreenVisible ? Color.Green : Color.Black;
            isGreenVisible = !isGreenVisible;
        }

        public void AddText(string text, Color color = default, bool bold = true)
        {
            rtbBlinkw.Text = string.Empty;
            rtbBlinkw.SelectionFont = new Font("Segoe UI", 30, FontStyle.Bold);
            rtbBlinkw.SelectionColor = Color.Green; // force green
            rtbBlinkw.AppendText(text);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            blinkTimer.Stop();
            blinkTimer.Dispose();
            this.Close();
        }

    }
}
