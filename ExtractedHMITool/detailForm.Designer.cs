namespace ExtractedHMITool
{
    partial class detailForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(detailForm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.animationTextBox = new System.Windows.Forms.RichTextBox();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.animationTextBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1442, 152);
            this.panel1.TabIndex = 0;
            // 
            // animationTextBox
            // 
            this.animationTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.animationTextBox.BackColor = System.Drawing.Color.White;
            this.animationTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.animationTextBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.animationTextBox.Location = new System.Drawing.Point(3, 3);
            this.animationTextBox.Name = "animationTextBox";
            this.animationTextBox.ReadOnly = true;
            this.animationTextBox.Size = new System.Drawing.Size(1439, 149);
            this.animationTextBox.TabIndex = 0;
            this.animationTextBox.Text = "";
            // 
            // detailForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1442, 152);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "detailForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Details";
            this.Activated += new System.EventHandler(this.detailForm_Activated);
            this.Load += new System.EventHandler(this.detailForm_Load);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.RichTextBox animationTextBox;
        private System.Windows.Forms.BindingSource bindingSource1;
    }
}