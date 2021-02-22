namespace SpertroApp
{
    partial class Form_Classify
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
            this.CCDImage = new System.Windows.Forms.PictureBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.CCDImage)).BeginInit();
            this.SuspendLayout();
            // 
            // CCDImage
            // 
            this.CCDImage.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.CCDImage.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.CCDImage.Cursor = System.Windows.Forms.Cursors.Cross;
            this.CCDImage.Location = new System.Drawing.Point(7, -15);
            this.CCDImage.Margin = new System.Windows.Forms.Padding(0);
            this.CCDImage.Name = "CCDImage";
            this.CCDImage.Size = new System.Drawing.Size(787, 480);
            this.CCDImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.CCDImage.TabIndex = 3;
            this.CCDImage.TabStop = false;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form_Classify
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(870, 575);
            this.Controls.Add(this.CCDImage);
            this.Name = "Form_Classify";
            this.Text = "Form_Classify";
            this.Load += new System.EventHandler(this.Form_Classify_Load);
            ((System.ComponentModel.ISupportInitialize)(this.CCDImage)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox CCDImage;
        private System.Windows.Forms.Timer timer1;
    }
}