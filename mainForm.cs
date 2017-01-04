using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.VisualBasic;
using ADOX;
using System.Data.OleDb;

/*Developer Comments*/
//Rough notes (ignore it)

namespace Report_generator
{
    public partial class mainForm : Form
    {
        public const string appVersion = "2.0.0.0 BETA";
        public ADOX.Catalog storageDbCatalog;
        public Dictionary<string, DataObject> dataObjectCollecion;
        public string currentFolder;
        public string accessStorageDbPath;
        public OleDbCommand cmdDraft;
        public string masterConnString;

        public mainForm() 
        { 
            InitializeComponent();
            flexiLabel.Text = "Welcome to the Extractor v " + appVersion;
            this.Text = "The Extractor v " + appVersion;
            dataObjectCollecion = new Dictionary<string, DataObject>();
            currentFolder = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\";
            CheckNecessaryFiles();
        }
        private void CheckNecessaryFiles()
        {
            string adoDbLibrary = currentFolder + "adodb.dll";
            string vbLibrary = currentFolder + "Microsoft.VisualBasic.dll";
            string[] criticalFiles = { adoDbLibrary, vbLibrary };

            foreach (string filePath in criticalFiles)
            { if (!(System.IO.File.Exists(filePath))) { MessageBox.Show("Critical file was not found. The application will exit." + Environment.NewLine + accessStorageDbPath); Environment.Exit(1); } }
        }
        private void excelFilePathButton_Click(object sender, EventArgs e)
        {
            /*Allow the user to browse for eligible Excel file*/
            string workbookPath = JazzyFunctionsByPatryk.BrowseFilePath("Excel Files|*.xlsx;*.xls;*.xlsm");//;*.csv
            this.excelFilePathTextbox.Text = workbookPath;
            if (workbookPath == "") { return; }

            /*Fill the sheets list*/
            //string connectionString = JazzyFunctionsByPatryk.GetConnectionStringExcel(workbookPath);
            //if (connectionString == "") { MessageBox.Show("Could not load the sheets from the chosen Excel file."); return; }

            //var excelFileSheetsList = JazzyFunctionsByPatryk.ListSheetInExcel(connectionString);
            var excelFileSheetsList = JazzyFunctionsByPatryk.ListSheetInExcel(workbookPath);
            if (excelFileSheetsList == null) { MessageBox.Show("Could not load the sheets. Is the Excel file correct?"); return; }
            this.excelFileSheetsComboBox.DataSource = excelFileSheetsList;

            this.queryLoadButton.Enabled = true;
        }
        private void excelFileSheetsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*Fill the fields list*/
        }
        private void queryLoadButton_Click(object sender, EventArgs e) { LoadQueryFromTextBoxToGridView(queryTextBox, this.excelFilePathTextbox.Text); }
        private void LoadQueryFromTextBoxToGridView(Control currentQueryTextBox, string excelFilePath)
        {
            if (excelFilePath == "") { MessageBox.Show("File path is missing!"); }
            System.Data.DataTable excelSheetDataTable;
            /*Prepare SQL Query*/
            string queryString = currentQueryTextBox.Text;
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
            this.previewGridView.DataSource = excelSheetDataTable; currentQueryTextBox.Text = queryString;
            UpdateTotalRecords(excelSheetDataTable);
        }
        private void UpdateTotalRecords(DataTable dt = null)
        { if (dt != null) { this.totalRecordsLabel.Text = "Total records: " + dt.Rows.Count; } else { this.totalRecordsLabel.Text = ""; } }
        private void mainForm_Load(object sender, EventArgs e) { /*DemSwitchez(0);*/ }
        private void quitButton_Click(object sender, EventArgs e)
        {
            /*Terminate the storage database connection and exit*/
            TerminateDB(true);
            Environment.Exit(1);
        }
        private void TerminateDB(bool question = false)
        {
            if (System.IO.File.Exists(accessStorageDbPath))
            {
                //ADODB.Connection con = null;
                try
                {
                    var con = storageDbCatalog.ActiveConnection as OleDbConnection;//ADODB.Connection
                    if (con != null) con.Close();
                }
                catch (Exception e) { }

                //System.Runtime.InteropServices.Marshal.FinalReleaseComObject(con);//storageDbCatalog.ActiveConnection
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(storageDbCatalog);
                DialogResult dialogResult = DialogResult.Yes;
                if (question) { dialogResult = MessageBox.Show("Would you like to delete the temporary database?" + Environment.NewLine + accessStorageDbPath, "", MessageBoxButtons.YesNo); }
                
                string accessStorageLinkPath = accessStorageDbPath.Replace(".accdb", ".laccdb");
                //var nl = Environment.NewLine;
                string separator = Environment.NewLine + "---" + Environment.NewLine;
                if (dialogResult == DialogResult.Yes) 
                { 
                    try { System.IO.File.Delete(accessStorageDbPath); }
                    catch (Exception e)
                    { MessageBox.Show("Could not delete the temporary database. Please do it manually:" + separator + accessStorageDbPath + separator + accessStorageLinkPath); }
                }
            }
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
            //masterQueryExportToExcelButton.Visible = masterBool;
            masterQueryTextBox.Visible = masterBool;

            excelFileBrowsePathButton.Visible = subBool;
            excelFilePathTextbox.Visible = subBool;
            excelFileSheetsComboBox.Visible = subBool;
            sourceTypesGroupBox.Visible = subBool;
        }
        private void tableAddButton_Click(object sender, EventArgs e)
        {
            string newDataSourceName = Interaction.InputBox("Provide name of the new data source.", "Enter name"); //, "Default Text"
            if(dataObjectCollecion.ContainsKey(newDataSourceName)) { MessageBox.Show("There is already a data object with that name!"); return; }
            if (newDataSourceName == "") { MessageBox.Show("Incorrect name!"); return; }
            //var newTable = new ADOX.Table();
            
            var newDataObject = new DataObject(newDataSourceName);
            dataObjectCollecion.Add(newDataSourceName, newDataObject);

            var newListViewItem = new ListViewItem(newDataSourceName);
            dataObjectsListView.Items.Add(newListViewItem);
            DemSwitchez(1,"Choose the source file and sheet");

            dataObjectsListView.Items[dataObjectsListView.Items.Count - 1].Selected = true;
            dataObjectsListView.Select();
            EditMode();
        }
        private OleDbParameter GetOleDbParam(string name, int code)
        {
            OleDbParameter newParameter = new OleDbParameter();
            newParameter.ParameterName = @"@" + name;// GetCleanParameterName(name);
            switch(code)
            {
                case 0: newParameter.OleDbType = OleDbType.VarWChar;
                break;
                case 1: newParameter.OleDbType = OleDbType.Date;
                break;
                case 2: newParameter.OleDbType = OleDbType.Boolean;
                break;
                case 3: newParameter.OleDbType = OleDbType.Double;
                break;
            }
            newParameter.Value = null;
            return newParameter;
        }
        //private string GetCleanParameterName(string name) { return GetCleanColumnName(name).Replace(" ", ""); }
        private bool CreateTemporaryDatabase()
        {
            try
            {          
                storageDbCatalog = new Catalog();
                SetStringAccessStorageDB();
                masterConnString = JazzyFunctionsByPatryk.GetConnectionString(accessStorageDbPath);
                storageDbCatalog.Create(masterConnString);
                return true;
            }
            catch (Exception e) { MessageBox.Show(e.Message.ToString()); return false; }
        }
        private void masterQueryLoadButton_Click(object sender, EventArgs e)
        {
            if(!(CreateTemporaryDatabase())) {return;}

            /* Create tables it temp DB with data */
            OleDbConnection con = new OleDbConnection();
            try
            {
                foreach (var key in dataObjectCollecion.Keys)
                {
                    /* Create empty table */
                    DataObject currentDataObject = dataObjectCollecion[key];
                    var newTable = new ADOX.Table();
                    newTable.Name = currentDataObject.Name;
                    DataTable currentDatatable = currentDataObject.DataTable;

                    foreach (DataColumn col in currentDatatable.Columns)
                    {
                        ADOX.Column dbField = new Column();
                        dbField.Name = GetCleanColumnName(col.ColumnName);
                        dbField.Attributes = ColumnAttributesEnum.adColNullable;
                        switch (col.DataType.ToString())
                        {
                            case "System.String":
                            case "System.Char":
                            case "System.Guid": dbField.Type = ADOX.DataTypeEnum.adVarWChar; break;//newTable.Columns.Append(cleanedColumnName, ADOX.DataTypeEnum.adVarWChar);
                            case "System.DateTime":
                            case "System.TimeSpan": dbField.Type = ADOX.DataTypeEnum.adDate; break;
                            case "System.Boolean": dbField.Type = ADOX.DataTypeEnum.adBoolean; break;
                            default: dbField.Type = ADOX.DataTypeEnum.adDouble; break;/*"System.Double", "System.Decimal","System.Byte","System.Int16","System.Int32","System.Int64","System.SByte","System.Single","System.UInt16","System.UInt32","System.UInt64" */
                        }
                        newTable.Columns.Append(dbField);
                    }
                    storageDbCatalog.Tables.Append(newTable);
                    /* Create new connection - required because we have added a table */
                    con = new OleDbConnection(masterConnString);

                    cmdDraft = new OleDbCommand();
                    cmdDraft.Connection = con;
                    cmdDraft.CommandType = CommandType.Text; /* StoredProcedure is an alternative */

                    /* Create insert command with parameters */
                    int paramIteration = 1;
                    cmdDraft.CommandText = GetSqlInsertNQuery(currentDataObject.DataTable, currentDataObject.Name);
                    foreach (DataColumn col in currentDatatable.Columns)
                    {
                        switch (col.DataType.ToString())
                        {
                            case "System.String":
                            case "System.Char":
                            case "System.Guid": cmdDraft.Parameters.Add(GetOleDbParam(paramIteration.ToString(), 0)); break;
                            case "System.DateTime":
                            case "System.TimeSpan": cmdDraft.Parameters.Add(GetOleDbParam(paramIteration.ToString(), 1)); break;
                            case "System.Boolean": cmdDraft.Parameters.Add(GetOleDbParam(paramIteration.ToString(), 2)); break;
                            default: cmdDraft.Parameters.Add(GetOleDbParam(paramIteration.ToString(), 3)); break;
                        }
                        paramIteration += 1;
                    }

                    /* Fill the new table with data */
                    OleDbCommand cmd;
                    con.Open();
                    foreach (DataRow row in currentDatatable.Rows)
                    {
                        cmd = cmdDraft;
                        paramIteration = 1;
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
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, ex.Source); }
            finally { TerminateDB(true); }//if (con != null)  con.Close();
                    
        }

        private string GetCleanColumnName(string oldName)
        {
            string[] illegals = { @"!", @"@", @"#", @"$", @"%", @"^", @"&", @"*", @"/", @"\" };//@"(", @")",
            string newName = oldName;
            foreach (string currentIllegal in illegals) { newName = newName.Replace(currentIllegal, " "); }
            return newName;
        }
        private string GetSqlInsertNQuery(DataTable dt, string tableName)
        {
            string columnsNames = string.Empty;
            string columnsValues = string.Empty;
            int paramIteration = 1;
            foreach (DataColumn col in dt.Columns)
            { columnsNames = columnsNames + "[" + GetCleanColumnName(col.ColumnName) + "]" + ", "; columnsValues = columnsValues + @"@" + paramIteration + ", "; paramIteration += 1; }//GetCleanParameterName(col.ColumnName)
            columnsNames = columnsNames.Remove(columnsNames.Length - 2);/*For removing final comma*/
            columnsValues = columnsValues.Remove(columnsValues.Length - 2);

            string fullReturnString = "INSERT INTO [" + tableName + "] (" + columnsNames + ") VALUES(" + columnsValues + ");";
            return fullReturnString;
        }
        private void SetStringAccessStorageDB()
        {
            accessStorageDbPath = currentFolder + "ExtractorStorage.accdb";
            int i = 1;
            do { accessStorageDbPath = currentFolder + "(" + i + ")" + "ExtractorStorage.accdb"; i += 1; } while (System.IO.File.Exists(accessStorageDbPath));
        }
        private void dataObjectsListView_SelectedIndexChanged(object sender, EventArgs e) { PreviewMode(); }
        private void dataObjectsListView_MouseUp(object sender, MouseEventArgs e) /* Not working at the moment */
        {
            if (e.Button == MouseButtons.Right)
            {
                //if (listView1.FocusedItem.Bounds.Contains(e.Location) == true)
                //{ contextMenuStrip1.Show(Cursor.Position); }
                if (dataObjectsListView.SelectedItems.Count == 0) { return; }
                string selectedName = dataObjectsListView.SelectedItems[0].SubItems[0].Text;
                DataObject currentDataObject = dataObjectCollecion[selectedName];

                string selectedDataSourceNewName = Interaction.InputBox("Provide the new name of the data source.", "Enter name", currentDataObject.Name); //, "Default Text"
                if (dataObjectCollecion.ContainsKey(selectedDataSourceNewName)) { MessageBox.Show("There is already a data object with that name!"); return; }
                if (selectedDataSourceNewName == "") { MessageBox.Show("Incorrect name!"); return; }

                dataObjectCollecion[selectedName].Name = selectedDataSourceNewName;
            }
        }
        private void PreviewMode()
        {
            queryLoadButton.Enabled = false;
            //exportToExcelButton.Enabled = false;
            queryTextBox.Enabled = false;
            saveDataObjectButton.Visible = false;

            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            DataObject currentDataObject = dataObjectCollecion[dataObjectsListView.SelectedItems[0].SubItems[0].Text];
            previewGridView.DataSource = currentDataObject.DataTable;
            queryTextBox.Text = currentDataObject.sqlQuery;

            DemSwitchez(0);

            //excelFileBrowsePathButton.Visible = false;
            excelFilePathTextbox.Text = "";
            excelFileSheetsComboBox.DataSource = null; excelFileSheetsComboBox.Items.Clear();
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
            currentDataObject.sqlQuery = queryTextBox.Text;
            MessageBox.Show("Saved successfully.");
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e) { excelFileBrowsePathButton.Visible = false; }
        private void sourceExcelRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            excelFileBrowsePathButton.Visible = true;
            excelFileSheetsComboBox.Visible = true;
        }
        private void exportFromGridViewButton_Click(object sender, EventArgs e)
        {
            if (this.previewGridView.DataSource == null) { return; }
            string savePath = JazzyFunctionsByPatryk.BrowseSavePath("xls");
            try { JazzyFunctionsByPatryk.DataTableToCSVFile(this.previewGridView.DataSource as System.Data.DataTable, savePath); MessageBox.Show("Saved successfully."); }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
    }
}
