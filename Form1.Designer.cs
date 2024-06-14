namespace WeightMachinFormApp
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.RichTextBox rtbData;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.rtbData = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // rtbData
            // 
            this.rtbData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rtbData.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.rtbData.Location = new System.Drawing.Point(0, 0);
            this.rtbData.Name = "rtbData";
            this.rtbData.Size = new System.Drawing.Size(421, 409);
            this.rtbData.TabIndex = 0;
            this.rtbData.Text = "";
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(421, 409);
            this.Controls.Add(this.rtbData);
            this.Name = "Form1";
            this.Text = "Wight Machine Service";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.ResumeLayout(false);

        }
    }
}
