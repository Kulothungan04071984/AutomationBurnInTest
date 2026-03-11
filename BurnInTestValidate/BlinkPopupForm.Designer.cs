namespace BurnInTestValidate
{
    partial class BlinkPopupForm
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
            this.rtbMessagenew = new System.Windows.Forms.RichTextBox();
            this.timersecond = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // rtbMessagenew
            // 
            this.rtbMessagenew.Location = new System.Drawing.Point(13, 13);
            this.rtbMessagenew.Name = "rtbMessagenew";
            this.rtbMessagenew.Size = new System.Drawing.Size(1069, 416);
            this.rtbMessagenew.TabIndex = 0;
            this.rtbMessagenew.Text = "";
            // 
            // timersecond
            // 
            this.timersecond.Tick += new System.EventHandler(this.timersecond_Tick);
            // 
            // BlinkPopupForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1094, 450);
            this.Controls.Add(this.rtbMessagenew);
            this.Name = "BlinkPopupForm";
            this.Text = "BlinkPopupForm";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox rtbMessagenew;
        private System.Windows.Forms.Timer timersecond;
    }
}