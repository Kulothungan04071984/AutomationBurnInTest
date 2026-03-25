namespace BurnInTestValidate
{
    partial class FailBlink
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
            this.rtbFail = new System.Windows.Forms.RichTextBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // rtbFail
            // 
            this.rtbFail.Location = new System.Drawing.Point(26, 12);
            this.rtbFail.Name = "rtbFail";
            this.rtbFail.Size = new System.Drawing.Size(762, 245);
            this.rtbFail.TabIndex = 0;
            this.rtbFail.Text = "";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // FailBlink
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 269);
            this.Controls.Add(this.rtbFail);
            this.Name = "FailBlink";
            this.Text = "FailBlink";
            this.Load += new System.EventHandler(this.FailBlink_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox rtbFail;
        private System.Windows.Forms.Timer timer1;
    }
}