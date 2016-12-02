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
            this.flexiLabel = new System.Windows.Forms.Label();
            this.queryTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.previewGridView = new System.Windows.Forms.DataGridView();
            this.excelFileBrowsePathButton = new System.Windows.Forms.Button();
            this.excelFilePathTextbox = new System.Windows.Forms.TextBox();
            this.queryLoadButton = new System.Windows.Forms.Button();
            this.exportToExcelButton = new System.Windows.Forms.Button();
            this.tableListView = new System.Windows.Forms.ListView();
            this.label3 = new System.Windows.Forms.Label();
            this.tableAddButton = new System.Windows.Forms.Button();
            this.tableEditButton = new System.Windows.Forms.Button();
            this.tableDeleteButton = new System.Windows.Forms.Button();
            this.masterButton = new System.Windows.Forms.Button();
            this.quitButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.masterQueryLoadButton = new System.Windows.Forms.Button();
            this.masterQueryExportToExcelButton = new System.Windows.Forms.Button();
            this.masterQueryTextBox = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.previewGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // excelFileSheetsComboBox
            // 
            this.excelFileSheetsComboBox.FormattingEnabled = true;
            this.excelFileSheetsComboBox.Location = new System.Drawing.Point(386, 98);
            this.excelFileSheetsComboBox.Name = "excelFileSheetsComboBox";
            this.excelFileSheetsComboBox.Size = new System.Drawing.Size(278, 21);
            this.excelFileSheetsComboBox.TabIndex = 25;
            this.excelFileSheetsComboBox.SelectedIndexChanged += new System.EventHandler(this.excelFileSheetsComboBox_SelectedIndexChanged);
            // 
            // flexiLabel
            // 
            this.flexiLabel.AutoSize = true;
            this.flexiLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.flexiLabel.Location = new System.Drawing.Point(300, 12);
            this.flexiLabel.Name = "flexiLabel";
            this.flexiLabel.Size = new System.Drawing.Size(318, 25);
            this.flexiLabel.TabIndex = 24;
            this.flexiLabel.Text = "Welcome to the Extractor v2.1.1";
            // 
            // queryTextBox
            // 
            this.queryTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.queryTextBox.Location = new System.Drawing.Point(386, 175);
            this.queryTextBox.Multiline = true;
            this.queryTextBox.Name = "queryTextBox";
            this.queryTextBox.Size = new System.Drawing.Size(554, 76);
            this.queryTextBox.TabIndex = 23;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(300, 145);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(177, 25);
            this.label1.TabIndex = 22;
            this.label1.Text = "Data source view";
            // 
            // previewGridView
            // 
            this.previewGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.previewGridView.Location = new System.Drawing.Point(12, 264);
            this.previewGridView.Name = "previewGridView";
            this.previewGridView.Size = new System.Drawing.Size(928, 342);
            this.previewGridView.TabIndex = 21;
            // 
            // excelFileBrowsePathButton
            // 
            this.excelFileBrowsePathButton.Location = new System.Drawing.Point(305, 69);
            this.excelFileBrowsePathButton.Name = "excelFileBrowsePathButton";
            this.excelFileBrowsePathButton.Size = new System.Drawing.Size(75, 23);
            this.excelFileBrowsePathButton.TabIndex = 20;
            this.excelFileBrowsePathButton.Text = "Browse";
            this.excelFileBrowsePathButton.UseVisualStyleBackColor = true;
            this.excelFileBrowsePathButton.Click += new System.EventHandler(this.excelFilePathButton_Click);
            // 
            // excelFilePathTextbox
            // 
            this.excelFilePathTextbox.Location = new System.Drawing.Point(386, 71);
            this.excelFilePathTextbox.Name = "excelFilePathTextbox";
            this.excelFilePathTextbox.ReadOnly = true;
            this.excelFilePathTextbox.Size = new System.Drawing.Size(554, 20);
            this.excelFilePathTextbox.TabIndex = 19;
            // 
            // queryLoadButton
            // 
            this.queryLoadButton.Enabled = false;
            this.queryLoadButton.Location = new System.Drawing.Point(305, 173);
            this.queryLoadButton.Name = "queryLoadButton";
            this.queryLoadButton.Size = new System.Drawing.Size(75, 23);
            this.queryLoadButton.TabIndex = 18;
            this.queryLoadButton.Text = "Load query";
            this.queryLoadButton.UseVisualStyleBackColor = true;
            this.queryLoadButton.Click += new System.EventHandler(this.queryLoadButton_Click);
            // 
            // exportToExcelButton
            // 
            this.exportToExcelButton.Enabled = false;
            this.exportToExcelButton.Location = new System.Drawing.Point(305, 214);
            this.exportToExcelButton.Name = "exportToExcelButton";
            this.exportToExcelButton.Size = new System.Drawing.Size(75, 37);
            this.exportToExcelButton.TabIndex = 26;
            this.exportToExcelButton.Text = "Export to CSV";
            this.exportToExcelButton.UseVisualStyleBackColor = true;
            this.exportToExcelButton.Click += new System.EventHandler(this.exportToExcelButton_Click);
            // 
            // tableListView
            // 
            this.tableListView.Location = new System.Drawing.Point(102, 170);
            this.tableListView.Name = "tableListView";
            this.tableListView.Size = new System.Drawing.Size(166, 81);
            this.tableListView.TabIndex = 27;
            this.tableListView.UseCompatibleStateImageBehavior = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(16, 142);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(139, 25);
            this.label3.TabIndex = 28;
            this.label3.Text = "Data sources";
            // 
            // tableAddButton
            // 
            this.tableAddButton.Location = new System.Drawing.Point(21, 170);
            this.tableAddButton.Name = "tableAddButton";
            this.tableAddButton.Size = new System.Drawing.Size(75, 23);
            this.tableAddButton.TabIndex = 29;
            this.tableAddButton.Text = "ADD";
            this.tableAddButton.UseVisualStyleBackColor = true;
            this.tableAddButton.Click += new System.EventHandler(this.tableAddButton_Click);
            // 
            // tableEditButton
            // 
            this.tableEditButton.Location = new System.Drawing.Point(21, 199);
            this.tableEditButton.Name = "tableEditButton";
            this.tableEditButton.Size = new System.Drawing.Size(75, 23);
            this.tableEditButton.TabIndex = 30;
            this.tableEditButton.Text = "EDIT";
            this.tableEditButton.UseVisualStyleBackColor = true;
            // 
            // tableDeleteButton
            // 
            this.tableDeleteButton.Location = new System.Drawing.Point(21, 228);
            this.tableDeleteButton.Name = "tableDeleteButton";
            this.tableDeleteButton.Size = new System.Drawing.Size(75, 23);
            this.tableDeleteButton.TabIndex = 31;
            this.tableDeleteButton.Text = "DELETE";
            this.tableDeleteButton.UseVisualStyleBackColor = true;
            // 
            // masterButton
            // 
            this.masterButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.masterButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.masterButton.Location = new System.Drawing.Point(42, 16);
            this.masterButton.Name = "masterButton";
            this.masterButton.Size = new System.Drawing.Size(195, 35);
            this.masterButton.TabIndex = 32;
            this.masterButton.Text = "Generate master report";
            this.masterButton.UseVisualStyleBackColor = false;
            this.masterButton.Click += new System.EventHandler(this.masterButton_Click);
            // 
            // quitButton
            // 
            this.quitButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.quitButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.quitButton.Location = new System.Drawing.Point(42, 66);
            this.quitButton.Name = "quitButton";
            this.quitButton.Size = new System.Drawing.Size(195, 35);
            this.quitButton.TabIndex = 33;
            this.quitButton.Text = "Quit";
            this.quitButton.UseVisualStyleBackColor = false;
            this.quitButton.Click += new System.EventHandler(this.quitButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Location = new System.Drawing.Point(274, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(10, 253);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Location = new System.Drawing.Point(12, 125);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(928, 14);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            // 
            // masterQueryLoadButton
            // 
            this.masterQueryLoadButton.Enabled = false;
            this.masterQueryLoadButton.Location = new System.Drawing.Point(623, 13);
            this.masterQueryLoadButton.Name = "masterQueryLoadButton";
            this.masterQueryLoadButton.Size = new System.Drawing.Size(75, 28);
            this.masterQueryLoadButton.TabIndex = 35;
            this.masterQueryLoadButton.Text = "Load query";
            this.masterQueryLoadButton.UseVisualStyleBackColor = true;
            // 
            // masterQueryExportToExcelButton
            // 
            this.masterQueryExportToExcelButton.Enabled = false;
            this.masterQueryExportToExcelButton.Location = new System.Drawing.Point(704, 14);
            this.masterQueryExportToExcelButton.Name = "masterQueryExportToExcelButton";
            this.masterQueryExportToExcelButton.Size = new System.Drawing.Size(116, 26);
            this.masterQueryExportToExcelButton.TabIndex = 36;
            this.masterQueryExportToExcelButton.Text = "Export to CSV";
            this.masterQueryExportToExcelButton.UseVisualStyleBackColor = true;
            // 
            // masterQueryTextBox
            // 
            this.masterQueryTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.masterQueryTextBox.Location = new System.Drawing.Point(290, 46);
            this.masterQueryTextBox.Multiline = true;
            this.masterQueryTextBox.Name = "masterQueryTextBox";
            this.masterQueryTextBox.Size = new System.Drawing.Size(650, 82);
            this.masterQueryTextBox.TabIndex = 37;
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(952, 618);
            this.ControlBox = false;
            this.Controls.Add(this.masterQueryExportToExcelButton);
            this.Controls.Add(this.masterQueryLoadButton);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.quitButton);
            this.Controls.Add(this.masterButton);
            this.Controls.Add(this.tableDeleteButton);
            this.Controls.Add(this.tableEditButton);
            this.Controls.Add(this.tableAddButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tableListView);
            this.Controls.Add(this.exportToExcelButton);
            this.Controls.Add(this.excelFileSheetsComboBox);
            this.Controls.Add(this.flexiLabel);
            this.Controls.Add(this.queryTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.previewGridView);
            this.Controls.Add(this.excelFileBrowsePathButton);
            this.Controls.Add(this.excelFilePathTextbox);
            this.Controls.Add(this.queryLoadButton);
            this.Controls.Add(this.masterQueryTextBox);
            this.Name = "mainForm";
            this.Text = "Report generator";
            this.Load += new System.EventHandler(this.mainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.previewGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox excelFileSheetsComboBox;
        private System.Windows.Forms.Label flexiLabel;
        private System.Windows.Forms.TextBox queryTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView previewGridView;
        private System.Windows.Forms.Button excelFileBrowsePathButton;
        private System.Windows.Forms.TextBox excelFilePathTextbox;
        private System.Windows.Forms.Button queryLoadButton;
        private System.Windows.Forms.Button exportToExcelButton;
        private System.Windows.Forms.ListView tableListView;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button tableAddButton;
        private System.Windows.Forms.Button tableEditButton;
        private System.Windows.Forms.Button tableDeleteButton;
        private System.Windows.Forms.Button masterButton;
        private System.Windows.Forms.Button quitButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button masterQueryLoadButton;
        private System.Windows.Forms.Button masterQueryExportToExcelButton;
        private System.Windows.Forms.TextBox masterQueryTextBox;
    }
}

