namespace BurnInTestValidate
{
    partial class Blink
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.timerBlink = new System.Windows.Forms.Timer(this.components);
            this.rtbBlinkw = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // timerBlink
            // 
            this.timerBlink.Tick += new System.EventHandler(this.timerBlink_Tick);
            // 
            // rtbBlinkw
            // 
            this.rtbBlinkw.Location = new System.Drawing.Point(11, 11);
            this.rtbBlinkw.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.rtbBlinkw.Name = "rtbBlinkw";
            this.rtbBlinkw.Size = new System.Drawing.Size(728, 270);
            this.rtbBlinkw.TabIndex = 0;
            this.rtbBlinkw.Text = "";
            // 
            // Blink
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(750, 292);
            this.Controls.Add(this.rtbBlinkw);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "Blink";
            this.Text = "Blink";
            this.Load += new System.EventHandler(this.Blink_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer timerBlink;
        private System.Windows.Forms.RichTextBox rtbBlinkw;
    }
}