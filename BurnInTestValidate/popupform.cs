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
        public popupform()
        {
            InitializeComponent();
            rtbMessage.ReadOnly = true;
            rtbMessage.BorderStyle = BorderStyle.None;
        }

        public void AddText(string text, Color color, bool bold = false)
        {
            rtbMessage.Text = string.Empty;
            rtbMessage.SelectionStart = rtbMessage.TextLength;
            rtbMessage.SelectionLength = 0;

            rtbMessage.SelectionColor = color;
            rtbMessage.Font = new Font("Segoe UI", 30, FontStyle.Bold);
            rtbMessage.SelectionFont = new Font(
                rtbMessage.Font,
                bold ? FontStyle.Bold : FontStyle.Regular);

            rtbMessage.AppendText(text);
            rtbMessage.SelectionColor = rtbMessage.ForeColor;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
