using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ADOX;
using System.Data.OleDb;

/*Developer Comments*/
//Rough notes (ignore it)

namespace Report_generator
{
    public partial class mainForm : Form
    {
        public const string appVersion = "2.0.0.0";
        //public ADOX.Catalog storageDbCatalog;
        public Dictionary<string, DataObject> dataObjectCollecion;
        public string currentFolder;
        public OleDbCommand cmdDraft;
        public string masterConnString;
        public string accessStorageDbPath;
        //public string masterQuery;
        //public List<string> tempDBList;

        //enum Days
        //{
        //    Sunday = 1,
        //    TuesDay = 2,
        //    wednesday = 3
        //}
        ////cast to enum
        //Days day = (Days)3;

        enum DOstatus
        {
            Idle,
            Extracting,
            Downloading
        }

        public mainForm() 
        { 
            InitializeComponent();
            this.Text = "The Extractor v " + appVersion; /* App label */
            
            dataObjectCollecion = new Dictionary<string, DataObject>();
            currentFolder = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\";
            accessStorageDbPath = currentFolder + "ExtractorStorage.accdb";
            masterConnString = FunRepository.GetConnectionString(accessStorageDbPath);
            CheckNecessaryFiles();
            ResetSettings();
        }
        private void ResetSettings()
        {
            foreach (Control ctrl in dataObjectGroupBox.Controls) { ctrl.Enabled = false; }
            flexiLabel.Text = "Welcome to the Extractor v " + appVersion;
            dataObjectCollecion.Clear();
            ClearTempDB();
            dataObjectsListView.Items.Clear();
            foreach (Control ctrl in this.Controls) { if (ctrl.Name.Length >= 7) { if (FunRepository.Right(ctrl.Name, 7) == "TextBox") { ctrl.Text = ""; } } }
        }
        private void CheckNecessaryFiles()
        {
            string adoDbLibrary = currentFolder + "adodb.dll";
            string vbLibrary = currentFolder + "Microsoft.VisualBasic.dll";
            string[] criticalFiles = { adoDbLibrary, vbLibrary, accessStorageDbPath };

            foreach (string filePath in criticalFiles)
            { if (!(System.IO.File.Exists(filePath))) { MessageBox.Show("Critical file was not found. The application will exit." + Environment.NewLine + filePath); Environment.Exit(1); } }
        }
        private void RefreshExcelSheets(string workbookPath)
        {
            var excelFileSheetsList = FunRepository.GetOleDbSchema(workbookPath);
            if (excelFileSheetsList == null) { MessageBox.Show("Could not load the sheets. Is the Excel file correct?"); return; }
            this.excelFileSheetsComboBox.DataSource = excelFileSheetsList;
        }
        private void excelFileBrowsePathButton_Click(object sender, EventArgs e)
        {
            string workbookPath = FunRepository.BrowseFilePath("Excel Files|*.xlsx;*.xls;*.xlsm");//;*.csv
            this.excelFilePathTextbox.Text = workbookPath;
            if (workbookPath == "") { return; }

            RefreshExcelSheets(workbookPath);

            this.queryLoadButton.Enabled = true;
        }
        //private void excelFileRefreshSheetsButton_Click_1(object sender, EventArgs e) { RefreshExcelSheets(this.excelFilePathTextbox.Text); }
        private void queryLoadButton_Click(object sender, EventArgs e) 
        {
            //pickup multithread for loading
            
            string excelFilePath = this.excelFilePathTextbox.Text;

            if (excelFilePath == "") { MessageBox.Show("File path is missing!"); return; }
            //System.Data.DataTable excelSheetDataTable;

            string queryString = queryTextBox.Text;
            if (queryString == "") { queryString = "SELECT * FROM [" + this.excelFileSheetsComboBox.Text.Replace("'", "") + "$]"; /*+ "$B1:G3]"*/ }
            this.queryTextBox.Text = queryString;

            FillPreviewGridView(excelFilePath, ref queryString);

            queryTextBox.Text = queryString;
        }
        private void FillPreviewGridView(string excelFilePath, ref string queryString)
        {
            System.Data.DataTable excelSheetDataTable;
            string excelFileConnectionString = FunRepository.GetConnectionString(excelFilePath);
            FunRepository.SetCustomSqlFunctions(ref queryString, excelFilePath);
            excelSheetDataTable = FunRepository.GetDataTable(excelFileConnectionString, queryString);
            try { this.previewGridView.DataSource = excelSheetDataTable; } catch { }
            UpdateTotalRecords(excelSheetDataTable);
        }
        private void UpdateTotalRecords(DataTable dt = null)
        { if (dt != null) { this.totalRecordsLabel.Text = "Total records: " + dt.Rows.Count; } else { this.totalRecordsLabel.Text = ""; } }
        private void mainForm_Load(object sender, EventArgs e) {  }
        private void quitButton_Click(object sender, EventArgs e) { ClearTempDB(); Environment.Exit(1); }
        private void ClearTempDB()
        {
            var storageDbCatalog = new ADOX.Catalog();
            var tempDBConnection = new ADODB.Connection();
            tempDBConnection.Open(masterConnString);
            storageDbCatalog.ActiveConnection = tempDBConnection;

            List<string> tableNames = FunRepository.GetOleDbSchema(accessStorageDbPath);
            foreach (string currentTableName in tableNames)
            { storageDbCatalog.Tables.Delete(currentTableName); }
            tempDBConnection.Close();
        }
        private void masterButton_Click(object sender, EventArgs e) { DemSwitchez(2,"Master report"); }
        private void DemSwitchez(int switchCode, string flexiLabelText = "")
        {
            /*Upper-right controls should be visible/hidden depending on context*/
            if (flexiLabelText == "") { flexiLabelText = flexiLabel.Text; }
            flexiLabel.Text = flexiLabelText;

            bool masterBool = false;
            bool subBool = false;
            switch (switchCode) 
            { 
                case 0:
                    if (masterQueryTextBox.Visible) { masterBool = true; }
                    break;
                case 1: subBool = true; break; 
                case 2: masterBool = true; break; 
                default: break; 
            }

            masterQueryLoadButton.Visible = masterBool;
            masterQueryTextBox.Visible = masterBool;

            dataSourceGroupBox.Visible = subBool;
            foreach (Control ctrl in dataSourceGroupBox.Controls) { ctrl.Visible = subBool; }
        }
        private void masterQueryLoadButton_Click(object sender, EventArgs e)
        {
            ClearTempDB();
            
            masterConnString = FunRepository.GetConnectionString(accessStorageDbPath);

            var storageDbCatalog = new ADOX.Catalog();
            var con = new OleDbConnection(masterConnString);
            var tempDBConnection = new ADODB.Connection();
            tempDBConnection.Open(masterConnString);
            storageDbCatalog.ActiveConnection = tempDBConnection;

            /* Create tables it temp DB with data */
            
            try
            {
                foreach (var key in dataObjectCollecion.Keys)
                {
                    /* Create empty table */
                    DataObject currentDataObject = dataObjectCollecion[key];
                    DataTable currentDatatable = currentDataObject.DataTable;

                    var newTable = FunRepository.GetNewAdoxTable(currentDatatable, currentDataObject.Name);
                    storageDbCatalog.Tables.Append(newTable);

                    /* Create new connection - required because we have added a table */
                    con = new OleDbConnection(masterConnString);

                    cmdDraft = new OleDbCommand();
                    cmdDraft.Connection = con;
                    cmdDraft.CommandType = CommandType.Text; /* StoredProcedure is an alternative */

                    /* Create insert command with parameters */
                    cmdDraft.CommandText = FunRepository.GetSqlInsertNonQuery(currentDataObject.DataTable, currentDataObject.Name);
                    FunRepository.SetOleDbCommandParameters(ref cmdDraft, currentDatatable);

                    /* Fill the new table with data */
                    OleDbCommand cmd;
                    con.Open();
                    foreach (DataRow row in currentDatatable.Rows)
                    {
                        cmd = cmdDraft;
                        int paramIteration = 1;
                        foreach (DataColumn col in currentDatatable.Columns)
                        { cmd.Parameters[@"@" + paramIteration].Value = row[col.ColumnName]; paramIteration += 1; }
                        cmd.ExecuteNonQuery(); 
                    }
                    con.Close();
                }

                /* Temp DB is ready, so let's run the master query */
                string masterQuery = masterQueryTextBox.Text;
                if (masterQuery != "") 
                { 
                    var masterDataTable = FunRepository.GetDataTable(masterConnString, masterQuery);
                    try { this.previewGridView.DataSource = masterDataTable; } catch { }
                    UpdateTotalRecords(masterDataTable);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, ex.Source); }
            finally { con.Dispose(); tempDBConnection.Close() ; storageDbCatalog = null; } 
        }
        private void dataObjectsListView_SelectedIndexChanged(object sender, EventArgs e) { PreviewMode(); }
        //private void dataObjectsListView_MouseUp(object sender, MouseEventArgs e) /* Not working at the moment */
        //{ if (e.Button == MouseButtons.Right) { } }
        private void PreviewMode()
        {
            //queryLoadButton.Enabled = false;
            //queryTextBox.Enabled = false;
            //saveDataObjectButton.Visible = false;
            foreach (Control ctrl in dataObjectGroupBox.Controls) { ctrl.Enabled = false; }

            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            DataObject currentDataObject = dataObjectCollecion[dataObjectsListView.SelectedItems[0].SubItems[0].Text];
            try { previewGridView.DataSource = currentDataObject.DataTable; } catch { }
            queryTextBox.Text = currentDataObject.SqlQuery;

            DemSwitchez(0);

            excelFilePathTextbox.Text = currentDataObject.ExcelFilePath;
            excelFileSheetsComboBox.Text = currentDataObject.ExcelFileSheet;
            persStorageCheckBox.Checked = currentDataObject.PersStorage;
            descriptionDOTextBox.Text = currentDataObject.Description;

            UpdateTotalRecords(currentDataObject.DataTable);
        }
        private void tableDeleteButton_Click(object sender, EventArgs e)
        {
            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            string currentDataObjectName = dataObjectsListView.SelectedItems[0].SubItems[0].Text;
            if (MessageBox.Show("Are you sure that you want to remove the " + currentDataObjectName + " data object?", "Confirmation required", MessageBoxButtons.YesNo) == DialogResult.No) { return; }
            dataObjectCollecion.Remove(currentDataObjectName);
            dataObjectsListView.SelectedItems[0].Remove();

            excelFilePathTextbox.Text = "";
            queryTextBox.Text = "";
            descriptionDOTextBox.Text = "";
            excelFileSheetsComboBox.DataSource = null; 
            excelFileSheetsComboBox.Items.Clear();
            this.previewGridView.DataSource = null; 
            DemSwitchez(99);
            UpdateTotalRecords();
        }
        private void tableEditButton_Click(object sender, EventArgs e) { EditMode(); }
        private void EditMode()
        {
            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            foreach (Control ctrl in dataObjectGroupBox.Controls) { ctrl.Enabled = true; }
            DemSwitchez(1,"Choose the source file and sheet");
        }
        private void saveDataObjectButton_Click(object sender, EventArgs e)
        {
            DataObject currentDataObject = dataObjectCollecion[dataObjectsListView.SelectedItems[0].SubItems[0].Text];
            currentDataObject.DataTable = (System.Data.DataTable)previewGridView.DataSource;
            currentDataObject.SqlQuery = queryTextBox.Text;
            currentDataObject.ExcelFilePath=excelFilePathTextbox.Text;
            currentDataObject.ExcelFileSheet=excelFileSheetsComboBox.Text;
            currentDataObject.PersStorage = persStorageCheckBox.Checked;
            currentDataObject.Description = descriptionDOTextBox.Text;
            currentDataObject.RunLoad = autoRunCheckBox.Checked;
            MessageBox.Show("Saved successfully.");
        }
        private void exportFromGridViewButton_Click(object sender, EventArgs e)
        {
            if (this.previewGridView.DataSource == null) { return; }
            string savePath = FunRepository.BrowseSavePath("xls");
            if (savePath == "") { return; }
            FunRepository.DataTableToExcelFileWithInterop(this.previewGridView.DataSource as System.Data.DataTable, savePath);
        }
        private void tableAddButton_Click(object sender, EventArgs e)
        {
            string newDataSourceName = FunRepository.SummonInputBox("Provide name of the new data source.", "Enter name");

            if (!(CheckDataSourceName(ref newDataSourceName))) { return; }

            var newDataObject = new DataObject(newDataSourceName);
            dataObjectCollecion.Add(newDataSourceName, newDataObject);

            string[] nameArray = new string[] { newDataSourceName };
            AddDataObjectsNamesToListView(nameArray);

            DemSwitchez(1, "Choose the source file and sheet");
            EditMode();
        }
        private void AddDataObjectsNamesToListView(string[] newDataSourceNames)
        {
            foreach (string currentItem in newDataSourceNames) { dataObjectsListView.Items.Add(currentItem); }

            dataObjectsListView.Items[dataObjectsListView.Items.Count - 1].Selected = true;
            dataObjectsListView.Select();
        }
        private void loadPresetsButton_Click(object sender, EventArgs e)
        {
            string masterQuery = this.masterQueryTextBox.Text;
            FunRepository.ReadSettings(ref masterQuery, ref dataObjectCollecion);
            this.masterQueryTextBox.Text = masterQuery;

            this.dataObjectsListView.Items.Clear();
            string[] keys = dataObjectCollecion.Keys.ToArray();
            AddDataObjectsNamesToListView(keys);

            foreach(var key in dataObjectCollecion.Keys)
            {
                DataObject currentDataObject = dataObjectCollecion[key];
                if(currentDataObject.RunLoad && currentDataObject.ExcelFilePath != "")
                { currentDataObject.DataTable = FunRepository.GetDataTable(FunRepository.GetConnectionString(currentDataObject.ExcelFilePath), currentDataObject.SqlQuery); }
            }
        }
        private void tableRenameButton_Click(object sender, EventArgs e)
        {
            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            string selectedDataSourceName = dataObjectsListView.SelectedItems[0].SubItems[0].Text;
            string selectedDataSourceNewName = FunRepository.SummonInputBox("Provide the new name of the data source.", "Enter name", selectedDataSourceName);//currentDataObject.Name);
            if (!(CheckDataSourceName(ref selectedDataSourceNewName))) { return; }

            var newDataObject = dataObjectCollecion[selectedDataSourceName];
            dataObjectCollecion.Add(selectedDataSourceNewName, newDataObject);
            dataObjectCollecion.Remove(selectedDataSourceName);

            dataObjectCollecion[selectedDataSourceNewName].Name = selectedDataSourceNewName;
            
            dataObjectsListView.SelectedItems[0].SubItems[0].Text = selectedDataSourceNewName;
        }
        private bool CheckDataSourceName(ref string name)
        {
            if (name == "") { return false; }//MessageBox.Show("Blank name!");
            string correctedName = FunRepository.GetCleanAccessObjectName(name, true);
            if (name != correctedName) { MessageBox.Show("The name has incorrect characters and has been corrected to: " + correctedName); name = correctedName; }
            if (dataObjectCollecion.ContainsKey(name)) { MessageBox.Show("There is already a data object with that name!"); return false; }
            return true;
        }
        private void savePresetsButton_Click(object sender, EventArgs e)
        { FunRepository.WriteSettings(this.masterQueryTextBox.Text, dataObjectCollecion); }
        private void previewGridView_CellContentClick(object sender, DataGridViewCellEventArgs e) { }
        private void sourceExcelRadioButton_CheckedChanged_1(object sender, EventArgs e) { }
        private void sourceSharePointRadioButton_CheckedChanged(object sender, EventArgs e) { }
        private void currentStatusLabel_Click(object sender, EventArgs e) { }
        private void persStorageCheckBox_CheckedChanged(object sender, EventArgs e)
        { if (persStorageCheckBox.Checked == true) { autoRunCheckBox.Enabled = true; } else { autoRunCheckBox.Enabled = false; } }
        private void resetButton_Click(object sender, EventArgs e) { ResetSettings(); }
        private void getSheetsButton_Click(object sender, EventArgs e) { RefreshExcelSheets(this.excelFilePathTextbox.Text); }
    }
}