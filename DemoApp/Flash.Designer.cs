
namespace DemoApp
{
    partial class Flash
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Flash));
            this.flash_img = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.loader_timer = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.flash_img)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // flash_img
            // 
            this.flash_img.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flash_img.Image = ((System.Drawing.Image)(resources.GetObject("flash_img.Image")));
            this.flash_img.Location = new System.Drawing.Point(0, 0);
            this.flash_img.Name = "flash_img";
            this.flash_img.Size = new System.Drawing.Size(800, 450);
            this.flash_img.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.flash_img.TabIndex = 0;
            this.flash_img.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.ForestGreen;
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.panel1.Location = new System.Drawing.Point(0, 587);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(800, 13);
            this.panel1.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.LimeGreen;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(50, 13);
            this.panel2.TabIndex = 2;
            // 
            // loader_timer
            // 
            this.loader_timer.Enabled = true;
            this.loader_timer.Interval = 15;
            this.loader_timer.Tick += new System.EventHandler(this.loader_timer_Tick);
            // 
            // Flash
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.flash_img);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Flash";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Flash";
            ((System.ComponentModel.ISupportInitialize)(this.flash_img)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox flash_img;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Timer loader_timer;
    }
}