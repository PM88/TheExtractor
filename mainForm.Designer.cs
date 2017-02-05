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
            this.flexiLabel = new System.Windows.Forms.Label();
            this.queryTextBox = new System.Windows.Forms.TextBox();
            this.previewGridView = new System.Windows.Forms.DataGridView();
            this.queryLoadButton = new System.Windows.Forms.Button();
            this.dataObjectsListView = new System.Windows.Forms.ListView();
            this.columnHeaderDataObject = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.tableAddButton = new System.Windows.Forms.Button();
            this.tableDeleteButton = new System.Windows.Forms.Button();
            this.masterButton = new System.Windows.Forms.Button();
            this.quitButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.masterQueryLoadButton = new System.Windows.Forms.Button();
            this.masterQueryTextBox = new System.Windows.Forms.TextBox();
            this.tableEditButton = new System.Windows.Forms.Button();
            this.loadPresetsButton = new System.Windows.Forms.Button();
            this.saveDataObjectButton = new System.Windows.Forms.Button();
            this.totalRecordsLabel = new System.Windows.Forms.Label();
            this.exportFromGridViewButton = new System.Windows.Forms.Button();
            this.tableRenameButton = new System.Windows.Forms.Button();
            this.dataSourceGroupBox = new System.Windows.Forms.GroupBox();
            this.getSheetsButton = new System.Windows.Forms.Button();
            this.autoRunCheckBox = new System.Windows.Forms.CheckBox();
            this.sourceTypesGroupBox = new System.Windows.Forms.GroupBox();
            this.sourceSharePointRadioButton = new System.Windows.Forms.RadioButton();
            this.sourceExcelRadioButton = new System.Windows.Forms.RadioButton();
            this.persStorageCheckBox = new System.Windows.Forms.CheckBox();
            this.excelFileSheetsComboBox = new System.Windows.Forms.ComboBox();
            this.excelFileBrowsePathButton = new System.Windows.Forms.Button();
            this.excelFilePathTextBox = new System.Windows.Forms.TextBox();
            this.savePresetsButton = new System.Windows.Forms.Button();
            this.descriptionDOTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dataObjectGroupBox = new System.Windows.Forms.GroupBox();
            this.currentStatusLabel = new System.Windows.Forms.Label();
            this.statusLabel = new System.Windows.Forms.Label();
            this.resetButton = new System.Windows.Forms.Button();
            this.tableCloneButton = new System.Windows.Forms.Button();
            this.tableUpButton = new System.Windows.Forms.Button();
            this.tableDownButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.sourcePasswordTextBox = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.previewGridView)).BeginInit();
            this.dataSourceGroupBox.SuspendLayout();
            this.sourceTypesGroupBox.SuspendLayout();
            this.dataObjectGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // flexiLabel
            // 
            this.flexiLabel.AutoSize = true;
            this.flexiLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.flexiLabel.Location = new System.Drawing.Point(300, 12);
            this.flexiLabel.Name = "flexiLabel";
            this.flexiLabel.Size = new System.Drawing.Size(318, 25);
            this.flexiLabel.TabIndex = 24;
            this.flexiLabel.Text = "Welcome to the Extractor v2.0.0";
            // 
            // queryTextBox
            // 
            this.queryTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.queryTextBox.Location = new System.Drawing.Point(87, 30);
            this.queryTextBox.Multiline = true;
            this.queryTextBox.Name = "queryTextBox";
            this.queryTextBox.Size = new System.Drawing.Size(554, 45);
            this.queryTextBox.TabIndex = 23;
            // 
            // previewGridView
            // 
            this.previewGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.previewGridView.Location = new System.Drawing.Point(12, 277);
            this.previewGridView.Name = "previewGridView";
            this.previewGridView.Size = new System.Drawing.Size(928, 310);
            this.previewGridView.TabIndex = 21;
            this.previewGridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.previewGridView_CellContentClick);
            // 
            // queryLoadButton
            // 
            this.queryLoadButton.Location = new System.Drawing.Point(6, 19);
            this.queryLoadButton.Name = "queryLoadButton";
            this.queryLoadButton.Size = new System.Drawing.Size(75, 23);
            this.queryLoadButton.TabIndex = 18;
            this.queryLoadButton.Text = "Load query";
            this.queryLoadButton.UseVisualStyleBackColor = true;
            this.queryLoadButton.Click += new System.EventHandler(this.queryLoadButton_Click);
            // 
            // dataObjectsListView
            // 
            this.dataObjectsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderDataObject});
            this.dataObjectsListView.FullRowSelect = true;
            this.dataObjectsListView.HideSelection = false;
            this.dataObjectsListView.Location = new System.Drawing.Point(79, 145);
            this.dataObjectsListView.MultiSelect = false;
            this.dataObjectsListView.Name = "dataObjectsListView";
            this.dataObjectsListView.Size = new System.Drawing.Size(189, 126);
            this.dataObjectsListView.TabIndex = 27;
            this.dataObjectsListView.UseCompatibleStateImageBehavior = false;
            this.dataObjectsListView.View = System.Windows.Forms.View.Details;
            this.dataObjectsListView.SelectedIndexChanged += new System.EventHandler(this.dataObjectsListView_SelectedIndexChanged);
            // 
            // columnHeaderDataObject
            // 
            this.columnHeaderDataObject.Text = "Data objects";
            this.columnHeaderDataObject.Width = 185;
            // 
            // tableAddButton
            // 
            this.tableAddButton.Location = new System.Drawing.Point(12, 144);
            this.tableAddButton.Name = "tableAddButton";
            this.tableAddButton.Size = new System.Drawing.Size(61, 21);
            this.tableAddButton.TabIndex = 29;
            this.tableAddButton.Text = "ADD";
            this.tableAddButton.UseVisualStyleBackColor = true;
            this.tableAddButton.Click += new System.EventHandler(this.tableAddButton_Click);
            // 
            // tableDeleteButton
            // 
            this.tableDeleteButton.Location = new System.Drawing.Point(12, 207);
            this.tableDeleteButton.Name = "tableDeleteButton";
            this.tableDeleteButton.Size = new System.Drawing.Size(61, 21);
            this.tableDeleteButton.TabIndex = 31;
            this.tableDeleteButton.Text = "DELETE";
            this.tableDeleteButton.UseVisualStyleBackColor = true;
            this.tableDeleteButton.Click += new System.EventHandler(this.tableDeleteButton_Click);
            // 
            // masterButton
            // 
            this.masterButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.masterButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.masterButton.Location = new System.Drawing.Point(42, 7);
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
            this.quitButton.Location = new System.Drawing.Point(42, 89);
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
            this.groupBox1.Size = new System.Drawing.Size(10, 266);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // groupBox2
            // 
            this.groupBox2.ForeColor = System.Drawing.Color.Red;
            this.groupBox2.Location = new System.Drawing.Point(12, 125);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(928, 14);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Please make sure that Excel is closed before you run any functionality. You might" +
    " lose your work progress.";
            // 
            // masterQueryLoadButton
            // 
            this.masterQueryLoadButton.Location = new System.Drawing.Point(623, 13);
            this.masterQueryLoadButton.Name = "masterQueryLoadButton";
            this.masterQueryLoadButton.Size = new System.Drawing.Size(132, 28);
            this.masterQueryLoadButton.TabIndex = 35;
            this.masterQueryLoadButton.Text = "Load master query";
            this.masterQueryLoadButton.UseVisualStyleBackColor = true;
            this.masterQueryLoadButton.Visible = false;
            this.masterQueryLoadButton.Click += new System.EventHandler(this.masterQueryLoadButton_Click);
            // 
            // masterQueryTextBox
            // 
            this.masterQueryTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.masterQueryTextBox.Location = new System.Drawing.Point(290, 46);
            this.masterQueryTextBox.Multiline = true;
            this.masterQueryTextBox.Name = "masterQueryTextBox";
            this.masterQueryTextBox.Size = new System.Drawing.Size(650, 82);
            this.masterQueryTextBox.TabIndex = 37;
            this.masterQueryTextBox.Visible = false;
            // 
            // tableEditButton
            // 
            this.tableEditButton.Location = new System.Drawing.Point(12, 165);
            this.tableEditButton.Name = "tableEditButton";
            this.tableEditButton.Size = new System.Drawing.Size(61, 21);
            this.tableEditButton.TabIndex = 38;
            this.tableEditButton.Text = "EDIT";
            this.tableEditButton.UseVisualStyleBackColor = true;
            this.tableEditButton.Click += new System.EventHandler(this.tableEditButton_Click);
            // 
            // loadPresetsButton
            // 
            this.loadPresetsButton.Location = new System.Drawing.Point(12, 593);
            this.loadPresetsButton.Name = "loadPresetsButton";
            this.loadPresetsButton.Size = new System.Drawing.Size(89, 20);
            this.loadPresetsButton.TabIndex = 40;
            this.loadPresetsButton.Text = "Load presets";
            this.loadPresetsButton.UseVisualStyleBackColor = true;
            this.loadPresetsButton.Click += new System.EventHandler(this.loadPresetsButton_Click);
            // 
            // saveDataObjectButton
            // 
            this.saveDataObjectButton.Location = new System.Drawing.Point(6, 76);
            this.saveDataObjectButton.Name = "saveDataObjectButton";
            this.saveDataObjectButton.Size = new System.Drawing.Size(75, 37);
            this.saveDataObjectButton.TabIndex = 41;
            this.saveDataObjectButton.Text = "Save data object";
            this.saveDataObjectButton.UseVisualStyleBackColor = true;
            this.saveDataObjectButton.Click += new System.EventHandler(this.saveDataObjectButton_Click);
            // 
            // totalRecordsLabel
            // 
            this.totalRecordsLabel.AutoSize = true;
            this.totalRecordsLabel.Location = new System.Drawing.Point(745, 597);
            this.totalRecordsLabel.Name = "totalRecordsLabel";
            this.totalRecordsLabel.Size = new System.Drawing.Size(10, 13);
            this.totalRecordsLabel.TabIndex = 45;
            this.totalRecordsLabel.Text = " ";
            // 
            // exportFromGridViewButton
            // 
            this.exportFromGridViewButton.Location = new System.Drawing.Point(851, 593);
            this.exportFromGridViewButton.Name = "exportFromGridViewButton";
            this.exportFromGridViewButton.Size = new System.Drawing.Size(89, 20);
            this.exportFromGridViewButton.TabIndex = 46;
            this.exportFromGridViewButton.Text = "Export data";
            this.exportFromGridViewButton.UseVisualStyleBackColor = true;
            this.exportFromGridViewButton.Click += new System.EventHandler(this.exportFromGridViewButton_Click);
            // 
            // tableRenameButton
            // 
            this.tableRenameButton.Location = new System.Drawing.Point(12, 186);
            this.tableRenameButton.Name = "tableRenameButton";
            this.tableRenameButton.Size = new System.Drawing.Size(61, 21);
            this.tableRenameButton.TabIndex = 47;
            this.tableRenameButton.Text = "RENAME";
            this.tableRenameButton.UseVisualStyleBackColor = true;
            this.tableRenameButton.Click += new System.EventHandler(this.tableRenameButton_Click);
            // 
            // dataSourceGroupBox
            // 
            this.dataSourceGroupBox.Controls.Add(this.sourcePasswordTextBox);
            this.dataSourceGroupBox.Controls.Add(this.label1);
            this.dataSourceGroupBox.Controls.Add(this.getSheetsButton);
            this.dataSourceGroupBox.Controls.Add(this.autoRunCheckBox);
            this.dataSourceGroupBox.Controls.Add(this.sourceTypesGroupBox);
            this.dataSourceGroupBox.Controls.Add(this.persStorageCheckBox);
            this.dataSourceGroupBox.Controls.Add(this.excelFileSheetsComboBox);
            this.dataSourceGroupBox.Controls.Add(this.excelFileBrowsePathButton);
            this.dataSourceGroupBox.Controls.Add(this.excelFilePathTextBox);
            this.dataSourceGroupBox.Location = new System.Drawing.Point(291, 5);
            this.dataSourceGroupBox.Name = "dataSourceGroupBox";
            this.dataSourceGroupBox.Size = new System.Drawing.Size(649, 123);
            this.dataSourceGroupBox.TabIndex = 49;
            this.dataSourceGroupBox.TabStop = false;
            this.dataSourceGroupBox.Text = "Data source";
            this.dataSourceGroupBox.Visible = false;
            // 
            // getSheetsButton
            // 
            this.getSheetsButton.Location = new System.Drawing.Point(6, 96);
            this.getSheetsButton.Name = "getSheetsButton";
            this.getSheetsButton.Size = new System.Drawing.Size(75, 23);
            this.getSheetsButton.TabIndex = 58;
            this.getSheetsButton.Text = "Get sheets";
            this.getSheetsButton.UseVisualStyleBackColor = true;
            this.getSheetsButton.Visible = false;
            this.getSheetsButton.Click += new System.EventHandler(this.getSheetsButton_Click);
            // 
            // autoRunCheckBox
            // 
            this.autoRunCheckBox.AutoSize = true;
            this.autoRunCheckBox.Location = new System.Drawing.Point(211, 35);
            this.autoRunCheckBox.Name = "autoRunCheckBox";
            this.autoRunCheckBox.Size = new System.Drawing.Size(84, 17);
            this.autoRunCheckBox.TabIndex = 57;
            this.autoRunCheckBox.Text = "Run on load";
            this.autoRunCheckBox.UseVisualStyleBackColor = true;
            this.autoRunCheckBox.Visible = false;
            // 
            // sourceTypesGroupBox
            // 
            this.sourceTypesGroupBox.Controls.Add(this.sourceSharePointRadioButton);
            this.sourceTypesGroupBox.Controls.Add(this.sourceExcelRadioButton);
            this.sourceTypesGroupBox.Enabled = false;
            this.sourceTypesGroupBox.Location = new System.Drawing.Point(6, 19);
            this.sourceTypesGroupBox.Name = "sourceTypesGroupBox";
            this.sourceTypesGroupBox.Size = new System.Drawing.Size(180, 39);
            this.sourceTypesGroupBox.TabIndex = 56;
            this.sourceTypesGroupBox.TabStop = false;
            this.sourceTypesGroupBox.Text = "Source type";
            this.sourceTypesGroupBox.Visible = false;
            // 
            // sourceSharePointRadioButton
            // 
            this.sourceSharePointRadioButton.AutoSize = true;
            this.sourceSharePointRadioButton.Location = new System.Drawing.Point(64, 18);
            this.sourceSharePointRadioButton.Name = "sourceSharePointRadioButton";
            this.sourceSharePointRadioButton.Size = new System.Drawing.Size(106, 17);
            this.sourceSharePointRadioButton.TabIndex = 1;
            this.sourceSharePointRadioButton.Text = "SharePoint Excel";
            this.sourceSharePointRadioButton.UseVisualStyleBackColor = true;
            this.sourceSharePointRadioButton.CheckedChanged += new System.EventHandler(this.sourceSharePointRadioButton_CheckedChanged);
            // 
            // sourceExcelRadioButton
            // 
            this.sourceExcelRadioButton.AutoSize = true;
            this.sourceExcelRadioButton.Checked = true;
            this.sourceExcelRadioButton.Location = new System.Drawing.Point(7, 18);
            this.sourceExcelRadioButton.Name = "sourceExcelRadioButton";
            this.sourceExcelRadioButton.Size = new System.Drawing.Size(51, 17);
            this.sourceExcelRadioButton.TabIndex = 0;
            this.sourceExcelRadioButton.TabStop = true;
            this.sourceExcelRadioButton.Text = "Excel";
            this.sourceExcelRadioButton.UseVisualStyleBackColor = true;
            this.sourceExcelRadioButton.CheckedChanged += new System.EventHandler(this.sourceExcelRadioButton_CheckedChanged_1);
            // 
            // persStorageCheckBox
            // 
            this.persStorageCheckBox.AutoSize = true;
            this.persStorageCheckBox.Location = new System.Drawing.Point(531, 100);
            this.persStorageCheckBox.Name = "persStorageCheckBox";
            this.persStorageCheckBox.Size = new System.Drawing.Size(110, 17);
            this.persStorageCheckBox.TabIndex = 55;
            this.persStorageCheckBox.Text = "Persistent storage";
            this.persStorageCheckBox.UseVisualStyleBackColor = true;
            this.persStorageCheckBox.Visible = false;
            this.persStorageCheckBox.CheckedChanged += new System.EventHandler(this.persStorageCheckBox_CheckedChanged);
            // 
            // excelFileSheetsComboBox
            // 
            this.excelFileSheetsComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.excelFileSheetsComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.excelFileSheetsComboBox.FormattingEnabled = true;
            this.excelFileSheetsComboBox.Location = new System.Drawing.Point(87, 98);
            this.excelFileSheetsComboBox.Name = "excelFileSheetsComboBox";
            this.excelFileSheetsComboBox.Size = new System.Drawing.Size(438, 21);
            this.excelFileSheetsComboBox.TabIndex = 53;
            this.excelFileSheetsComboBox.Visible = false;
            // 
            // excelFileBrowsePathButton
            // 
            this.excelFileBrowsePathButton.Location = new System.Drawing.Point(6, 69);
            this.excelFileBrowsePathButton.Name = "excelFileBrowsePathButton";
            this.excelFileBrowsePathButton.Size = new System.Drawing.Size(75, 23);
            this.excelFileBrowsePathButton.TabIndex = 52;
            this.excelFileBrowsePathButton.Text = "Browse";
            this.excelFileBrowsePathButton.UseVisualStyleBackColor = true;
            this.excelFileBrowsePathButton.Visible = false;
            this.excelFileBrowsePathButton.Click += new System.EventHandler(this.excelFileBrowsePathButton_Click);
            // 
            // excelFilePathTextBox
            // 
            this.excelFilePathTextBox.Location = new System.Drawing.Point(87, 71);
            this.excelFilePathTextBox.Name = "excelFilePathTextBox";
            this.excelFilePathTextBox.Size = new System.Drawing.Size(554, 20);
            this.excelFilePathTextBox.TabIndex = 51;
            this.excelFilePathTextBox.Visible = false;
            // 
            // savePresetsButton
            // 
            this.savePresetsButton.Location = new System.Drawing.Point(107, 593);
            this.savePresetsButton.Name = "savePresetsButton";
            this.savePresetsButton.Size = new System.Drawing.Size(89, 20);
            this.savePresetsButton.TabIndex = 50;
            this.savePresetsButton.Text = "Save preset";
            this.savePresetsButton.UseVisualStyleBackColor = true;
            this.savePresetsButton.Click += new System.EventHandler(this.savePresetsButton_Click);
            // 
            // descriptionDOTextBox
            // 
            this.descriptionDOTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.descriptionDOTextBox.Location = new System.Drawing.Point(87, 93);
            this.descriptionDOTextBox.Name = "descriptionDOTextBox";
            this.descriptionDOTextBox.Size = new System.Drawing.Size(554, 20);
            this.descriptionDOTextBox.TabIndex = 58;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(119, 79);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 13);
            this.label2.TabIndex = 59;
            this.label2.Text = "Description";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(119, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(57, 13);
            this.label4.TabIndex = 60;
            this.label4.Text = "SQL query";
            // 
            // dataObjectGroupBox
            // 
            this.dataObjectGroupBox.Controls.Add(this.currentStatusLabel);
            this.dataObjectGroupBox.Controls.Add(this.statusLabel);
            this.dataObjectGroupBox.Controls.Add(this.queryLoadButton);
            this.dataObjectGroupBox.Controls.Add(this.label4);
            this.dataObjectGroupBox.Controls.Add(this.queryTextBox);
            this.dataObjectGroupBox.Controls.Add(this.descriptionDOTextBox);
            this.dataObjectGroupBox.Controls.Add(this.saveDataObjectButton);
            this.dataObjectGroupBox.Controls.Add(this.label2);
            this.dataObjectGroupBox.Location = new System.Drawing.Point(291, 145);
            this.dataObjectGroupBox.Name = "dataObjectGroupBox";
            this.dataObjectGroupBox.Size = new System.Drawing.Size(649, 122);
            this.dataObjectGroupBox.TabIndex = 61;
            this.dataObjectGroupBox.TabStop = false;
            this.dataObjectGroupBox.Text = "Data object view";
            // 
            // currentStatusLabel
            // 
            this.currentStatusLabel.AutoSize = true;
            this.currentStatusLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.currentStatusLabel.Location = new System.Drawing.Point(11, 60);
            this.currentStatusLabel.Name = "currentStatusLabel";
            this.currentStatusLabel.Size = new System.Drawing.Size(35, 13);
            this.currentStatusLabel.TabIndex = 62;
            this.currentStatusLabel.Text = "IDLE";
            this.currentStatusLabel.Visible = false;
            this.currentStatusLabel.Click += new System.EventHandler(this.currentStatusLabel_Click);
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(6, 45);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(40, 13);
            this.statusLabel.TabIndex = 61;
            this.statusLabel.Text = "Status:";
            this.statusLabel.Visible = false;
            // 
            // resetButton
            // 
            this.resetButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.resetButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.resetButton.Location = new System.Drawing.Point(42, 48);
            this.resetButton.Name = "resetButton";
            this.resetButton.Size = new System.Drawing.Size(195, 35);
            this.resetButton.TabIndex = 62;
            this.resetButton.Text = "Reset";
            this.resetButton.UseVisualStyleBackColor = false;
            this.resetButton.Click += new System.EventHandler(this.resetButton_Click);
            // 
            // tableCloneButton
            // 
            this.tableCloneButton.Location = new System.Drawing.Point(12, 229);
            this.tableCloneButton.Name = "tableCloneButton";
            this.tableCloneButton.Size = new System.Drawing.Size(61, 21);
            this.tableCloneButton.TabIndex = 63;
            this.tableCloneButton.Text = "CLONE";
            this.tableCloneButton.UseVisualStyleBackColor = true;
            this.tableCloneButton.Click += new System.EventHandler(this.tableCloneButton_Click);
            // 
            // tableUpButton
            // 
            this.tableUpButton.Location = new System.Drawing.Point(12, 250);
            this.tableUpButton.Name = "tableUpButton";
            this.tableUpButton.Size = new System.Drawing.Size(31, 21);
            this.tableUpButton.TabIndex = 64;
            this.tableUpButton.Text = "/\\";
            this.tableUpButton.UseVisualStyleBackColor = true;
            this.tableUpButton.Click += new System.EventHandler(this.tableUpButton_Click);
            // 
            // tableDownButton
            // 
            this.tableDownButton.Location = new System.Drawing.Point(42, 250);
            this.tableDownButton.Name = "tableDownButton";
            this.tableDownButton.Size = new System.Drawing.Size(31, 21);
            this.tableDownButton.TabIndex = 65;
            this.tableDownButton.Text = "\\/";
            this.tableDownButton.UseVisualStyleBackColor = true;
            this.tableDownButton.Click += new System.EventHandler(this.tableDownButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(329, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 60;
            this.label1.Text = "Password";
            // 
            // sourcePasswordTextBox
            // 
            this.sourcePasswordTextBox.BackColor = System.Drawing.SystemColors.Window;
            this.sourcePasswordTextBox.Location = new System.Drawing.Point(386, 32);
            this.sourcePasswordTextBox.Name = "sourcePasswordTextBox";
            this.sourcePasswordTextBox.Size = new System.Drawing.Size(188, 20);
            this.sourcePasswordTextBox.TabIndex = 63;
            // 
            // mainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(952, 618);
            this.ControlBox = false;
            this.Controls.Add(this.tableDownButton);
            this.Controls.Add(this.tableUpButton);
            this.Controls.Add(this.tableCloneButton);
            this.Controls.Add(this.resetButton);
            this.Controls.Add(this.dataSourceGroupBox);
            this.Controls.Add(this.dataObjectGroupBox);
            this.Controls.Add(this.savePresetsButton);
            this.Controls.Add(this.tableRenameButton);
            this.Controls.Add(this.exportFromGridViewButton);
            this.Controls.Add(this.totalRecordsLabel);
            this.Controls.Add(this.loadPresetsButton);
            this.Controls.Add(this.tableEditButton);
            this.Controls.Add(this.masterQueryLoadButton);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.quitButton);
            this.Controls.Add(this.masterButton);
            this.Controls.Add(this.tableDeleteButton);
            this.Controls.Add(this.tableAddButton);
            this.Controls.Add(this.dataObjectsListView);
            this.Controls.Add(this.flexiLabel);
            this.Controls.Add(this.previewGridView);
            this.Controls.Add(this.masterQueryTextBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "mainForm";
            this.Text = "Report generator";
            this.Load += new System.EventHandler(this.mainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.previewGridView)).EndInit();
            this.dataSourceGroupBox.ResumeLayout(false);
            this.dataSourceGroupBox.PerformLayout();
            this.sourceTypesGroupBox.ResumeLayout(false);
            this.sourceTypesGroupBox.PerformLayout();
            this.dataObjectGroupBox.ResumeLayout(false);
            this.dataObjectGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label flexiLabel;
        private System.Windows.Forms.TextBox queryTextBox;
        private System.Windows.Forms.DataGridView previewGridView;
        private System.Windows.Forms.Button queryLoadButton;
        private System.Windows.Forms.ListView dataObjectsListView;
        private System.Windows.Forms.Button tableAddButton;
        private System.Windows.Forms.Button tableDeleteButton;
        private System.Windows.Forms.Button masterButton;
        private System.Windows.Forms.Button quitButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button masterQueryLoadButton;
        private System.Windows.Forms.TextBox masterQueryTextBox;
        private System.Windows.Forms.ColumnHeader columnHeaderDataObject;
        private System.Windows.Forms.Button tableEditButton;
        private System.Windows.Forms.Button loadPresetsButton;
        private System.Windows.Forms.Button saveDataObjectButton;
        private System.Windows.Forms.Label totalRecordsLabel;
        private System.Windows.Forms.Button exportFromGridViewButton;
        private System.Windows.Forms.Button tableRenameButton;
        private System.Windows.Forms.GroupBox dataSourceGroupBox;
        private System.Windows.Forms.CheckBox persStorageCheckBox;
        private System.Windows.Forms.ComboBox excelFileSheetsComboBox;
        private System.Windows.Forms.Button excelFileBrowsePathButton;
        private System.Windows.Forms.TextBox excelFilePathTextBox;
        private System.Windows.Forms.GroupBox sourceTypesGroupBox;
        private System.Windows.Forms.RadioButton sourceSharePointRadioButton;
        private System.Windows.Forms.RadioButton sourceExcelRadioButton;
        private System.Windows.Forms.Button savePresetsButton;
        private System.Windows.Forms.TextBox descriptionDOTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox dataObjectGroupBox;
        private System.Windows.Forms.Label currentStatusLabel;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.CheckBox autoRunCheckBox;
        private System.Windows.Forms.Button resetButton;
        private System.Windows.Forms.Button getSheetsButton;
        private System.Windows.Forms.Button tableCloneButton;
        private System.Windows.Forms.Button tableUpButton;
        private System.Windows.Forms.Button tableDownButton;
        private System.Windows.Forms.TextBox sourcePasswordTextBox;
        private System.Windows.Forms.Label label1;
    }
}

