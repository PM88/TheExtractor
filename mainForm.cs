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
        public const string appVersion = "2.0.0.0 BETA";
        //public ADOX.Catalog storageDbCatalog;
        public Dictionary<string, DataObject> dataObjectCollecion;
        public string currentFolder;
        public OleDbCommand cmdDraft;
        public string masterConnString;
        public string accessStorageDbPath;
        //public string masterQuery;
        //public List<string> tempDBList;

        public mainForm() 
        { 
            InitializeComponent();
            //ResetSettings();
            flexiLabel.Text = "Welcome to the Extractor v " + appVersion;
            this.Text = "The Extractor v " + appVersion; /* App label */
            dataObjectCollecion = new Dictionary<string, DataObject>();
            currentFolder = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\";
            accessStorageDbPath = currentFolder + "ExtractorStorage.accdb";
            masterConnString = JazzyFunctionsByPatryk.GetConnectionString(accessStorageDbPath);
            CheckNecessaryFiles();
            ClearTempDB();
        }

        //private void ResetSettings()
        //{
        //    flexiLabel.Text = "Welcome to the Extractor v " + appVersion;
        //    this.Text = "The Extractor v " + appVersion;
        //    dataObjectCollecion = new Dictionary<string, DataObject>();
        //    storageDbCatalog = null;
        //    string accessStorageDbPath = string.Empty;
        //    public OleDbCommand cmdDraft;
        //    public string masterConnString;
        //}
        private void CheckNecessaryFiles()
        {
            string adoDbLibrary = currentFolder + "adodb.dll";
            string vbLibrary = currentFolder + "Microsoft.VisualBasic.dll";
            string[] criticalFiles = { adoDbLibrary, vbLibrary, accessStorageDbPath };

            foreach (string filePath in criticalFiles)
            { if (!(System.IO.File.Exists(filePath))) { MessageBox.Show("Critical file was not found. The application will exit." + Environment.NewLine + filePath); Environment.Exit(1); } }
        }
        private void excelFilePathButton_Click(object sender, EventArgs e)
        {
            ///*Allow the user to browse for eligible Excel file*/
            //string workbookPath = JazzyFunctionsByPatryk.BrowseFilePath("Excel Files|*.xlsx;*.xls;*.xlsm");//;*.csv
            //this.excelFilePathTextbox.Text = workbookPath;
            //if (workbookPath == "") { return; }

            ///*Fill the sheets list*/
            ////string connectionString = JazzyFunctionsByPatryk.GetConnectionStringExcel(workbookPath);
            ////if (connectionString == "") { MessageBox.Show("Could not load the sheets from the chosen Excel file."); return; }

            ////var excelFileSheetsList = JazzyFunctionsByPatryk.ListSheetInExcel(connectionString);
            //RefreshExcelSheets(workbookPath);

            //this.queryLoadButton.Enabled = true;
        }
        private void RefreshExcelSheets(string workbookPath)
        {
            var excelFileSheetsList = JazzyFunctionsByPatryk.GetOleDbSchema(workbookPath);
            if (excelFileSheetsList == null) { MessageBox.Show("Could not load the sheets. Is the Excel file correct?"); return; }
            this.excelFileSheetsComboBox.DataSource = excelFileSheetsList;
        }
        private void excelFileBrowsePathButton_Click(object sender, EventArgs e)
        {
            /*Allow the user to browse for eligible Excel file*/
            string workbookPath = JazzyFunctionsByPatryk.BrowseFilePath("Excel Files|*.xlsx;*.xls;*.xlsm");//;*.csv
            this.excelFilePathTextbox.Text = workbookPath;
            if (workbookPath == "") { return; }

            /*Fill the sheets list*/
            //string connectionString = JazzyFunctionsByPatryk.GetConnectionStringExcel(workbookPath);
            //if (connectionString == "") { MessageBox.Show("Could not load the sheets from the chosen Excel file."); return; }

            //var excelFileSheetsList = JazzyFunctionsByPatryk.ListSheetInExcel(connectionString);
            RefreshExcelSheets(workbookPath);

            this.queryLoadButton.Enabled = true;
        }
        private void excelFileRefreshSheetsButton_Click_1(object sender, EventArgs e) { RefreshExcelSheets(this.excelFilePathTextbox.Text); }
        //private void excelFileSheetsComboBox_SelectedIndexChanged(object sender, EventArgs e) { /*Fill the fields list*/ }
        private void queryLoadButton_Click(object sender, EventArgs e) 
        {
            string excelFilePath = this.excelFilePathTextbox.Text;

            if (excelFilePath == "") { MessageBox.Show("File path is missing!"); return; }
            System.Data.DataTable excelSheetDataTable;
            /*Prepare SQL Query*/
            string queryString = queryTextBox.Text;
            if (queryString == "") { queryString = "SELECT * FROM [" + this.excelFileSheetsComboBox.Text.Replace("'", "") + "$]"; /*+ "$B1:G3]"*/ }
            this.queryTextBox.Text = queryString;

            /*Preparing database connection*/
            string excelFileConnectionString = JazzyFunctionsByPatryk.GetConnectionString(excelFilePath);

            //Create data table
            excelSheetDataTable = JazzyFunctionsByPatryk.GetDataTable(excelFileConnectionString, queryString);

            //foreach (DataColumn col in excelSheetDataTable.Columns) { MessageBox.Show(col.ColumnName + " Type is: " + col.DataType); }

            /*Extract data into data grid on form */
            //if (excelSheetDataTable != null)
            //{ if (excelSheetDataTable.Rows.Count > 0) { this.excelQueryGridView.DataSource = excelSheetDataTable; queryTextBox.Text = queryString; } }
            this.previewGridView.DataSource = excelSheetDataTable; queryTextBox.Text = queryString;
            UpdateTotalRecords(excelSheetDataTable);
        }
        //private void LoadQueryFromTextBoxToGridView(Control currentQueryTextBox, string excelFilePath)
        //{

        //}
        private void UpdateTotalRecords(DataTable dt = null)
        { if (dt != null) { this.totalRecordsLabel.Text = "Total records: " + dt.Rows.Count; } else { this.totalRecordsLabel.Text = ""; } }
        private void mainForm_Load(object sender, EventArgs e) { /*DemSwitchez(0);*/ }
        private void quitButton_Click(object sender, EventArgs e)
        {
            /*Terminate the storage database connection and exit*/
            //TerminateDB(true);
            //var con = new OleDbConnection(masterConnString);
            //tempDBConnection.ConnectionString = masterConnString;

            ClearTempDB();

            Environment.Exit(1);
        }

        private void ClearTempDB()
        {
            var storageDbCatalog = new ADOX.Catalog();
            var tempDBConnection = new ADODB.Connection();
            tempDBConnection.Open(masterConnString);
            storageDbCatalog.ActiveConnection = tempDBConnection;

            List<string> tableNames = JazzyFunctionsByPatryk.GetOleDbSchema(accessStorageDbPath);
            foreach (string currentTableName in tableNames)
            { storageDbCatalog.Tables.Delete(currentTableName); }//string currentTableName in storageDbCatalog.Tables.
            tempDBConnection.Close();
        }
        //private void TerminateDB(bool question = false)
        //{

        //    //if (System.IO.File.Exists(accessStorageDbPath))
        //    //{
        //        //ADODB.Connection con = null;
        //        //try
        //        //{
        //        //    var con = storageDbCatalog.ActiveConnection as OleDbConnection;//ADODB.Connection
        //        //    if (con != null) {con.Dispose();}//con.Close(); con = null;
        //        //}
        //        //catch (Exception e) { }

        //        //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(con);//storageDbCatalog.ActiveConnection
        //        //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(storageDbCatalog);
        //        DialogResult dialogResult = DialogResult.Yes;
        //        if (question) { dialogResult = MessageBox.Show("Would you like to delete the temporary databases?", "", MessageBoxButtons.YesNo); }// + Environment.NewLine + accessStorageDbPath, "", MessageBoxButtons.YesNo); }
                
        //        //string accessStorageLinkPath = accessStorageDbPath.Replace(".accdb", ".laccdb");
        //        //var nl = Environment.NewLine;
        //        //string separator = Environment.NewLine + "---" + Environment.NewLine;
        //        if (dialogResult == DialogResult.Yes) 
        //        {
        //            foreach (string fileDB in tempDBList)
        //            {
        //                try { System.IO.File.Delete(fileDB); }
        //                catch (Exception e) { MessageBox.Show("Could not delete the temporary databases. Please do it manually.");} 
        //                //:" + separator + accessStorageDbPath + separator + accessStorageLinkPath); }
        //            }
        //        }
        //    //}
        //}
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
            //masterQueryExportToExcelButton.Visible = masterBool;
            masterQueryTextBox.Visible = masterBool;

            dataSourceGroupBox.Visible = subBool;
            foreach (Control ctrl in dataSourceGroupBox.Controls) { ctrl.Visible = subBool; }
            //excelFileBrowsePathButton.Visible = subBool;
            //excelFileRefreshSheetsButton.Visible = subBool;
            //excelFilePathTextbox.Visible = subBool;
            //excelFileSheetsComboBox.Visible = subBool;
            //sourceTypesGroupBox.Visible = subBool;
        }
        //private string GetCleanParameterName(string name) { return GetCleanColumnName(name).Replace(" ", ""); }
        //private bool ConnectToTemporaryDatabase()
        //{
        //    try
        //    {
        //        string accessStorageDbPath = GetStringAccessStorageDB();
        //        masterConnString = JazzyFunctionsByPatryk.GetConnectionString(accessStorageDbPath);

                //var storageDbCatalog = new ADOX.Catalog();
                //storageDbCatalog.Create(masterConnString);

                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(storageDbCatalog);
                //var con = storageDbCatalog.ActiveConnection as OleDbConnection;//ADODB.Connection
                //con.Dispose();
                //storageDbCatalog = null;

                //tempDBList.Add(accessStorageDbPath);
                //try
                //{
                //    var con = storageDbCatalog.ActiveConnection as OleDbConnection;//ADODB.Connection
                //    con.Dispose();
                //}
                //catch (Exception e) { MessageBox.Show(e.Message); return false; }

        //        return true;
        //    }
        //    catch (Exception e) { MessageBox.Show(e.Message.ToString()); return false; }
        //}
        //private Catalog OpenDatabase()
        //{
        //    Catalog catalog = new Catalog();
        //    var connection = new ADODB.Connection();

        //    try
        //    {
        //        connection.Open(_ConnectionString);
        //        catalog.ActiveConnection = connection;
        //    }
        //    catch (Exception)
        //    {
        //        catalog.Create(_ConnectionString);
        //    }
        //    return catalog;
        //}
        private void masterQueryLoadButton_Click(object sender, EventArgs e)
        {
            //if(!(ConnectToTemporaryDatabase())) {return;}
            masterConnString = JazzyFunctionsByPatryk.GetConnectionString(accessStorageDbPath);

            //bool askIfDeleteDB = false;
            //string lastTempDB = tempDBList[tempDBList.Count - 1];
            var storageDbCatalog = new ADOX.Catalog();
            var con = new OleDbConnection(masterConnString);
            var tempDBConnection = new ADODB.Connection();
            //tempDBConnection.ConnectionString = masterConnString;
            tempDBConnection.Open(masterConnString);
            storageDbCatalog.ActiveConnection = tempDBConnection;
            //storageDbCatalog.Create(masterConnString);

            /* Create tables it temp DB with data */
            
            try
            {
                foreach (var key in dataObjectCollecion.Keys)
                {
                    /* Create empty table */
                    DataObject currentDataObject = dataObjectCollecion[key];
                    DataTable currentDatatable = currentDataObject.DataTable;

                    var newTable = JazzyFunctionsByPatryk.GetNewAdoxTable(currentDatatable, currentDataObject.Name);
                    storageDbCatalog.Tables.Append(newTable);

                    /* Create new connection - required because we have added a table */
                    con = new OleDbConnection(masterConnString);

                    cmdDraft = new OleDbCommand();
                    cmdDraft.Connection = con;
                    cmdDraft.CommandType = CommandType.Text; /* StoredProcedure is an alternative */

                    /* Create insert command with parameters */
                    cmdDraft.CommandText = JazzyFunctionsByPatryk.GetSqlInsertNonQuery(currentDataObject.DataTable, currentDataObject.Name);
                    JazzyFunctionsByPatryk.SetOleDbCommandParameters(ref cmdDraft, currentDatatable);

                    /* Fill the new table with data */
                    OleDbCommand cmd;
                    con.Open();
                    foreach (DataRow row in currentDatatable.Rows)
                    {
                        cmd = cmdDraft;
                        int paramIteration = 1;
                        foreach (DataColumn col in currentDatatable.Columns)
                        { cmd.Parameters[@"@" + paramIteration].Value = row[col.ColumnName]; paramIteration += 1; }//GetCleanParameterName(col.ColumnName)
                        cmd.ExecuteNonQuery(); 
                    }
                    con.Close();
                }

                /* Temp DB is ready, so let's run the master query */
                //bool dbManualMode = false;
                string masterQuery = masterQueryTextBox.Text;
                if (masterQuery != "") 
                { 
                    var masterDataTable = JazzyFunctionsByPatryk.GetDataTable(masterConnString, masterQuery);
                    this.previewGridView.DataSource = masterDataTable;
                    UpdateTotalRecords(masterDataTable);
                }
                //else { askIfDeleteDB = true; }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, ex.Source); }
            finally { con.Dispose(); tempDBConnection.Close() ; storageDbCatalog = null; }//if (con != null)  con.Close(); con.Close(); con = null; };
            //TerminateDB(askIfDeleteDB);     
        }

        //private string GetStringAccessStorageDB()
        //{
        //    return 
        //    //string accessStorageDbPath = currentFolder + "ExtractorStorage.accdb";
        //    //int i = 1;
        //    //do { accessStorageDbPath = currentFolder + "(" + i + ")" + "ExtractorStorage.accdb"; i += 1; } while (System.IO.File.Exists(accessStorageDbPath));
        //    //return accessStorageDbPath;
        //}
        private void dataObjectsListView_SelectedIndexChanged(object sender, EventArgs e) { PreviewMode(); }
        private void dataObjectsListView_MouseUp(object sender, MouseEventArgs e) /* Not working at the moment */
        { if (e.Button == MouseButtons.Right) { } }
        private void PreviewMode()
        {
            queryLoadButton.Enabled = false;
            //exportToExcelButton.Enabled = false;
            queryTextBox.Enabled = false;
            saveDataObjectButton.Visible = false;

            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            DataObject currentDataObject = dataObjectCollecion[dataObjectsListView.SelectedItems[0].SubItems[0].Text];
            previewGridView.DataSource = currentDataObject.DataTable;
            queryTextBox.Text = currentDataObject.SqlQuery;

            DemSwitchez(0);

            //excelFileBrowsePathButton.Visible = false;
            //excelFilePathTextbox.Text = "";
            //excelFileSheetsComboBox.DataSource = null; excelFileSheetsComboBox.Items.Clear();
            excelFilePathTextbox.Text = currentDataObject.ExcelFilePath;
            excelFileSheetsComboBox.Text = currentDataObject.ExcelFileSheet;
            persStorageCheckBox.Checked = currentDataObject.PersStorage;

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
            excelFileSheetsComboBox.DataSource = null; excelFileSheetsComboBox.Items.Clear();
            DemSwitchez(99);
            UpdateTotalRecords();
        }
        private void tableEditButton_Click(object sender, EventArgs e) { EditMode(); }
        private void EditMode()
        {
            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            queryLoadButton.Enabled = true;
            //exportToExcelButton.Enabled = true;
            queryTextBox.Enabled = true;
            saveDataObjectButton.Visible = true;
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
            MessageBox.Show("Saved successfully.");
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e) { }//excelFileBrowsePathButton.Visible = false; }
        private void sourceExcelRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            //excelFileBrowsePathButton.Visible = true;
            //excelFileSheetsComboBox.Visible = true;
        }
        private void exportFromGridViewButton_Click(object sender, EventArgs e)
        {
            if (this.previewGridView.DataSource == null) { return; }
            string savePath = JazzyFunctionsByPatryk.BrowseSavePath("xls");
            if (savePath == "") { return; }
            JazzyFunctionsByPatryk.DataTableToExcelFileWithInterop(this.previewGridView.DataSource as System.Data.DataTable, savePath);
            //JazzyFunctionsByPatryk.DataTableToExcelFile(this.previewGridView.DataSource as System.Data.DataTable, savePath);  //try { }
            //catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        private void tableAddButton_Click(object sender, EventArgs e)
        {
            string newDataSourceName = JazzyFunctionsByPatryk.SummonInputBox("Provide name of the new data source.", "Enter name");

            if (!(CheckDataSourceName(ref newDataSourceName))) { return; }
            //var newTable = new ADOX.Table();

            var newDataObject = new DataObject(newDataSourceName);
            dataObjectCollecion.Add(newDataSourceName, newDataObject);

            string[] nameArray = new string[] { newDataSourceName };
            AddDataObjectNameToListView(nameArray);

            DemSwitchez(1, "Choose the source file and sheet");
            EditMode();
        }
        private void AddDataObjectNameToListView(string[] newDataSourceNames)
        {
            //var newListViewItem = new ListViewItem(newDataSourceName);
            //ListViewItem newItem;
            foreach (string currentItem in newDataSourceNames) { dataObjectsListView.Items.Add(currentItem); }

            dataObjectsListView.Items[dataObjectsListView.Items.Count - 1].Selected = true;
            dataObjectsListView.Select();
        }
        private void loadPresetsButton_Click(object sender, EventArgs e)
        {
            string masterQuery = this.masterQueryTextBox.Text;
            JazzyFunctionsByPatryk.ReadSettings(ref masterQuery, ref dataObjectCollecion);
            this.masterQueryTextBox.Text = masterQuery;

            this.dataObjectsListView.Items.Clear();
            string[] keys = dataObjectCollecion.Keys.ToArray();
            AddDataObjectNameToListView(keys);
            //foreach (KeyValuePair<string, DataObject> entry in dataObjectCollecion) { AddDataObjectNameToListView(entry.Key); }//entry.Key; entry.Value
        }
        private void tableRenameButton_Click(object sender, EventArgs e)
        {
            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            string selectedName = dataObjectsListView.SelectedItems[0].SubItems[0].Text;
            DataObject currentDataObject = dataObjectCollecion[selectedName];

            string selectedDataSourceNewName = JazzyFunctionsByPatryk.SummonInputBox("Provide the new name of the data source.", "Enter name", currentDataObject.Name);

            if (!(CheckDataSourceName(ref selectedDataSourceNewName))) { return; }

            dataObjectCollecion[selectedName].Name = selectedDataSourceNewName;
            dataObjectsListView.SelectedItems[0].SubItems[0].Text = selectedDataSourceNewName;
        }
        private bool CheckDataSourceName(ref string name)
        {
            if (name == "") { MessageBox.Show("Blank name!"); return false; }
            string correctedName = JazzyFunctionsByPatryk.GetCleanAccessObjectName(name, true);
            if (name != correctedName) { MessageBox.Show("The name has incorrect characters and has been corrected to: " + correctedName); name = correctedName; }
            if (dataObjectCollecion.ContainsKey(name)) { MessageBox.Show("There is already a data object with that name!"); return false; }
            return true;
        }
        //private void excelFileRefreshSheetsButton_Click(object sender, EventArgs e) { RefreshExcelSheets(this.excelFilePathTextbox.Text); }

        //private void button1_Click(object sender, EventArgs e) { JazzyFunctionsByPatryk.WriteSettings(); }
        private void savePresetsButton_Click(object sender, EventArgs e)
        { JazzyFunctionsByPatryk.WriteSettings(this.masterQueryTextBox.Text, dataObjectCollecion); }

        private void previewGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        //private void resetButton_Click(object sender, EventArgs e)
        //{
        //    this.excelFilePathTextbox;
        //    this.excelFileSheetsComboBox;
        //    this.queryTextBox;
        //    this.dataObjectsListView;
        //}
    }
}