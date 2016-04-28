namespace Report_generator
{
    partial class mainForm
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
            this.excelFileSheetsComboBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.queryTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.excelQueryGridView = new System.Windows.Forms.DataGridView();
            this.excelSourceFilePathButton = new System.Windows.Forms.Button();
            this.excelFilePathTextbox = new System.Windows.Forms.TextBox();
            this.queryLoadButton = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.excelQueryGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // excelFileSheetsComboBox
            // 
            this.excelFileSheetsComboBox.FormattingEnabled = true;
            this.excelFileSheetsComboBox.Location = new System.Drawing.Point(92, 71);
            this.excelFileSheetsComboBox.Name = "excelFileSheetsComboBox";
            this.excelFileSheetsComboBox.Size = new System.Drawing.Size(193, 21);
            this.excelFileSheetsComboBox.TabIndex = 25;
            this.excelFileSheetsComboBox.SelectedIndexChanged += new System.EventHandler(this.excelFileSheetsComboBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(7, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(416, 25);
            this.label2.TabIndex = 24;
            this.label2.Text = "Select the source excel file and worksheet";
            // 
            // queryTextBox
            // 
            this.queryTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.queryTextBox.Location = new System.Drawing.Point(94, 124);
            this.queryTextBox.Name = "queryTextBox";
            this.queryTextBox.Size = new System.Drawing.Size(507, 20);
            this.queryTextBox.TabIndex = 23;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(9, 95);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 25);
            this.label1.TabIndex = 22;
            this.label1.Text = "Report preview";
            // 
            // excelQueryGridView
            // 
            this.excelQueryGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.excelQueryGridView.Location = new System.Drawing.Point(12, 151);
            this.excelQueryGridView.Name = "excelQueryGridView";
            this.excelQueryGridView.Size = new System.Drawing.Size(589, 219);
            this.excelQueryGridView.TabIndex = 21;
            // 
            // excelSourceFilePathButton
            // 
            this.excelSourceFilePathButton.Location = new System.Drawing.Point(11, 42);
            this.excelSourceFilePathButton.Name = "excelSourceFilePathButton";
            this.excelSourceFilePathButton.Size = new System.Drawing.Size(75, 23);
            this.excelSourceFilePathButton.TabIndex = 20;
            this.excelSourceFilePathButton.Text = "Browse";
            this.excelSourceFilePathButton.UseVisualStyleBackColor = true;
            this.excelSourceFilePathButton.Click += new System.EventHandler(this.excelFilePathButton_Click);
            // 
            // excelFilePathTextbox
            // 
            this.excelFilePathTextbox.Location = new System.Drawing.Point(92, 44);
            this.excelFilePathTextbox.Name = "excelFilePathTextbox";
            this.excelFilePathTextbox.Size = new System.Drawing.Size(508, 20);
            this.excelFilePathTextbox.TabIndex = 19;
            // 
            // queryLoadButton
            // 
            this.queryLoadButton.Location = new System.Drawing.Point(12, 122);
            this.queryLoadButton.Name = "queryLoadButton";
            this.queryLoadButton.Size = new System.Drawing.Size(75, 23);
            this.queryLoadButton.TabIndex = 18;
            this.queryLoadButton.Text = "Load query";
            this.queryLoadButton.UseVisualStyleBackColor = true;
            this.queryLoadButton.Click += new System.EventHandler(this.queryLoadButton_Click);
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(501, 99);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(99, 23);
            this.button1.TabIndex = 26;
            this.button1.Text = "Export to Excel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(612, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 25);
            this.label3.TabIndex = 27;
            this.label3.Text = "Fields";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(804, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(85, 25);
            this.label4.TabIndex = 28;
            this.label4.Text = "Presets";
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(955, 385);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.excelFileSheetsComboBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.queryTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.excelQueryGridView);
            this.Controls.Add(this.excelSourceFilePathButton);
            this.Controls.Add(this.excelFilePathTextbox);
            this.Controls.Add(this.queryLoadButton);
            this.Name = "mainForm";
            this.Text = "Report generator";
            ((System.ComponentModel.ISupportInitialize)(this.excelQueryGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox excelFileSheetsComboBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox queryTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView excelQueryGridView;
        private System.Windows.Forms.Button excelSourceFilePathButton;
        private System.Windows.Forms.TextBox excelFilePathTextbox;
        private System.Windows.Forms.Button queryLoadButton;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}

