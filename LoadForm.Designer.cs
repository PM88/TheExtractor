namespace Report_generator
{
    partial class LoadForm
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
            this.constantLabel = new System.Windows.Forms.Label();
            this.dynamicLabel1 = new System.Windows.Forms.Label();
            this.dynamicLabel2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // constantLabel
            // 
            this.constantLabel.AutoSize = true;
            this.constantLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.constantLabel.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.constantLabel.Location = new System.Drawing.Point(131, 66);
            this.constantLabel.Name = "constantLabel";
            this.constantLabel.Size = new System.Drawing.Size(169, 46);
            this.constantLabel.TabIndex = 0;
            this.constantLabel.Text = "Loading";
            // 
            // dynamicLabel1
            // 
            this.dynamicLabel1.AutoSize = true;
            this.dynamicLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dynamicLabel1.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.dynamicLabel1.Location = new System.Drawing.Point(321, 66);
            this.dynamicLabel1.Name = "dynamicLabel1";
            this.dynamicLabel1.Size = new System.Drawing.Size(48, 46);
            this.dynamicLabel1.TabIndex = 1;
            this.dynamicLabel1.Text = "--";
            // 
            // dynamicLabel2
            // 
            this.dynamicLabel2.AutoSize = true;
            this.dynamicLabel2.Font = new System.Drawing.Font("Microsoft Sans Serif", 30F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dynamicLabel2.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.dynamicLabel2.Location = new System.Drawing.Point(43, 66);
            this.dynamicLabel2.Name = "dynamicLabel2";
            this.dynamicLabel2.Size = new System.Drawing.Size(48, 46);
            this.dynamicLabel2.TabIndex = 2;
            this.dynamicLabel2.Text = "--";
            // 
            // LoadForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.ClientSize = new System.Drawing.Size(408, 175);
            this.ControlBox = false;
            this.Controls.Add(this.dynamicLabel2);
            this.Controls.Add(this.dynamicLabel1);
            this.Controls.Add(this.constantLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "LoadForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Load += new System.EventHandler(this.LoadForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label constantLabel;
        private System.Windows.Forms.Label dynamicLabel1;
        private System.Windows.Forms.Label dynamicLabel2;
    }
}