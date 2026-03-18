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
    public partial class popupform : Form
    {
        
        bool isRed = true;
        public popupform()
        {
            InitializeComponent();
            richTextBox.ReadOnly = true;
            richTextBox.BorderStyle = BorderStyle.None;
            richTextBox.Dock = DockStyle.Fill;

            richTextBox.BackColor = Color.Black;
            richTextBox.SelectionAlignment = HorizontalAlignment.Center;

            timer1.Interval = 500; // blinking speed (milliseconds)
            timer1.Tick += timer1_Tick;
        }

        public void AddText(string text, Color color, bool bold = false)
        {
            if (richTextBox.InvokeRequired)
            {
                richTextBox.Invoke(new Action(() => AddText(text,color,bold)));
                return;
            }

            richTextBox.Clear();

            richTextBox.SelectionStart = 0;
            richTextBox.SelectionLength = 0;

            richTextBox.SelectionColor = Color.Red;

            richTextBox.SelectionFont = new Font("Segoe UI", 32,
                bold ? FontStyle.Bold : FontStyle.Regular);

            richTextBox.SelectionAlignment = HorizontalAlignment.Center;

            richTextBox.AppendText(text);

            timer1.Start();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            this.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (isRed)
                richTextBox.ForeColor = Color.Black;
            else
                richTextBox.ForeColor = Color.Red;

            isRed = !isRed;
        }
    }
}
