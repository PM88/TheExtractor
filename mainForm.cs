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
        public ADOX.Catalog storageDbCatalog;
        public Dictionary<string, DataObject> dataObjectCollecion;
        public string currentFolder;
        public string accessStorageDB;
        //public AdodbCommandDraft cmdDraft;
        //public ADODB.Command cmdDraft;
        public OleDbCommand cmdDraft;

        public mainForm() 
        { 
            InitializeComponent(); dataObjectCollecion = new Dictionary<string, DataObject>();
            currentFolder = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\";
            CheckNecessaryFiles();
        }
        private void CheckNecessaryFiles()
        {
            string adoDbLibrary = currentFolder + "adodb.dll";
            string vbLibrary = currentFolder + "Microsoft.VisualBasic.dll";
            string[] criticalFiles = { adoDbLibrary, vbLibrary };

            foreach (string filePath in criticalFiles)
            { if (!(System.IO.File.Exists(filePath))) { MessageBox.Show("Critical file was not found. The application will exit." + Environment.NewLine + accessStorageDB); Environment.Exit(1); } }
        }
        private void excelFilePathButton_Click(object sender, EventArgs e)
        {
            /*Allow the user to browse for eligible Excel file*/
            string workbookPath = JazzyFunctionsByPatryk.BrowseFilePath("Excel Files|*.xlsx;*.xls;*.xlsm;*.csv");
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
        }
        private void mainForm_Load(object sender, EventArgs e) { /*DemSwitchez(0);*/ }
        private void quitButton_Click(object sender, EventArgs e)
        {
            /*Terminate the storage database connection and exit*/

            TerminateDB(true);
            Environment.Exit(1);
        }
        private void TerminateDB(bool question = false)
        {
            if (System.IO.File.Exists(accessStorageDB))
            {
                //var con = storageDbCatalog.ActiveConnection as ADODB.Connection;
                //if (con != null) con.Close();
                DialogResult dialogResult = DialogResult.Yes;
                if (question) { dialogResult = MessageBox.Show("Would you like to delete the temporary database?" + Environment.NewLine + accessStorageDB, "", MessageBoxButtons.YesNo); }
                if (dialogResult == DialogResult.Yes ) { System.IO.File.Delete(accessStorageDB); }
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
            newParameter.ParameterName = @"@" + name;
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
        private void CreateTemporaryDatabase()
        {
            try
            {          
                //using(var con = new OleDbConnection(JazzyFunctionsByPatryk.GetConnectionString(accessStorageDB))
                //{

                
                //    con.Open();

                storageDbCatalog = new Catalog();
                SetStringAccessStorageDB();
                storageDbCatalog.Create(JazzyFunctionsByPatryk.GetConnectionString(accessStorageDB));
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(storageDbCatalog.ActiveConnection);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(storageDbCatalog);
                    //}
                //storageDbCatalog.ActiveConnection.Close();
                //storageDbCatalog = null;

            }
            catch (Exception e) { MessageBox.Show(e.Message.ToString()); }
        }
        private void masterQueryLoadButton_Click(object sender, EventArgs e)
        {
            CreateTemporaryDatabase();
            var con = new OleDbConnection(JazzyFunctionsByPatryk.GetConnectionString(accessStorageDB));
            con.Open();
            //if (con != null) con.Close();
            //var con = new OleDbConnection(JazzyFunctionsByPatryk.GetConnectionString(accessStorageDB));
            //con.Open();
            //MessageBox.Show(con.ConnectionString);
            //MessageBox.Show(con.DataSource);
            //if (con != null) { MessageBox.Show(con.State.ToString()); }

            foreach (var key in dataObjectCollecion.Keys)
            {
                DataObject currentDataObject = dataObjectCollecion[key];
                var newTable = new ADOX.Table();
                newTable.Name = currentDataObject.Name;
                DataTable currentDatatable = currentDataObject.DataTable;
                //ADODB.Parameter newParameter;
                //OleDbParameter newParameter;
                try
                {
                    cmdDraft = new OleDbCommand();//new AdodbCommandDraft(storageDbCatalog.ActiveConnection);
                    cmdDraft.Connection = con;
                    cmdDraft.CommandType = CommandType.StoredProcedure;
                    foreach (DataColumn col in currentDatatable.Columns)
                    {
                       //cmdDraft.ActiveConnection=storageDbCatalog.ActiveConnection;
                        switch(col.DataType.ToString())
                        {
                            case "System.String": case "System.Char": case "System.Guid":
                                newTable.Columns.Append(col.ColumnName, ADOX.DataTypeEnum.adVarWChar);//ADOX.DataTypeEnum.adVarWChar
                                cmdDraft.Parameters.Add(GetOleDbParam(col.ColumnName,0));
                                //cmdDraft.CreateParameter(@"@" + col.ColumnName, ADODB.DataTypeEnum.adVarWChar, ADODB.ParameterDirectionEnum.adParamInput, -1, null);
                                //newParameter = cmdDraft.CreateParameter(@"@" + col.ColumnName, ADODB.DataTypeEnum.adVarWChar, ADODB.ParameterDirectionEnum.adParamInput, -1, null);
                                //cmdDraft.Parameters.Append(newParameter); 
                                break;
                            case "System.DateTime": case "System.TimeSpan":
                                newTable.Columns.Append(col.ColumnName, ADOX.DataTypeEnum.adDate);
                                cmdDraft.Parameters.Add(GetOleDbParam(col.ColumnName, 1));
                                //cmdDraft.CreateParameter(@"@" + col.ColumnName, ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, -1, null);
                                //newParameter = cmdDraft.CreateParameter(@"@" + col.ColumnName, ADODB.DataTypeEnum.adDate, ADODB.ParameterDirectionEnum.adParamInput, -1, null);
                                //cmdDraft.Parameters.Append(newParameter); 
                                break;
                            case "System.Boolean":
                                newTable.Columns.Append(col.ColumnName, ADOX.DataTypeEnum.adBoolean);
                                cmdDraft.Parameters.Add(GetOleDbParam(col.ColumnName, 2));
                                //newParameter = cmdDraft.CreateParameter(@"@" + col.ColumnName, ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, -1, null);
                                //cmdDraft.Parameters.Append(newParameter); 
                                break;
                            default:/*"System.Double", "System.Decimal","System.Byte","System.Int16","System.Int32","System.Int64","System.SByte","System.Single","System.UInt16","System.UInt32","System.UInt64" */
                                newTable.Columns.Append(col.ColumnName, ADOX.DataTypeEnum.adDouble);
                                cmdDraft.Parameters.Add(GetOleDbParam(col.ColumnName, 3));
                                //cmdDraft.CreateParameter(@"@" + col.ColumnName, ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput, -1, null);
                                //newParameter = cmdDraft.CreateParameter(@"@" + col.ColumnName, ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput, -1, null);
                                //cmdDraft.Parameters.Append(newParameter); 
                                break;
                        }
                    }
                    storageDbCatalog.Tables.Append(newTable);
                    //MessageBox.Show(cmdDraft.Parameters.Count.ToString() + ": " + cmdDraft.Parameters[@"@Code"].Value);
                
                    //object obj = new object();
                    string sqlString = string.Empty;
                    //ADODB.Parameter currentParameter;
                    int currentRowNumber;
                    OleDbCommand cmd;

                        foreach (DataRow row in currentDatatable.Rows)
                        { 
                            currentRowNumber = currentDatatable.Rows.IndexOf(row);
                            //con.Execute(GetSqlInsertNQuery(currentDataObject.DataTable, currentDataObject.Name, currentDatatable.Rows.IndexOf(row)), out obj); 
                            sqlString = GetSqlInsertNQuery(currentDataObject.DataTable, currentDataObject.Name, currentRowNumber);
                            //var cmd = new System.Data.OleDb.OleDbCommand(sqlString, con);
                            //using(con)//System.Data.OleDb.OleDbCommand())
                            //{
                            cmd = cmdDraft;
                            //cmd.Connection.Open();
                            //MessageBox.Show(cmd.Connection.State.ToString());
                            //cmd.Connection = storageDbCatalog.ActiveConnection as OleDbConnection;
                            //cmd.Connection.Open();// = cmdDraft.Connection;
                            // ADODB.CommandTypeEnum.adCmdText for oridinary SQL statements;  
                            // ADODB.CommandTypeEnum.adCmdStoredProc for stored procedures. 
                            //cmd.CommandType = ADODB.CommandTypeEnum.adCmdText; 
                            foreach (DataColumn col in currentDatatable.Columns)
                            {
                                cmd.Parameters[@"@" + col.ColumnName].Value = row[col.ColumnName];
                                //currentParameter = cmd.Parameters[col.ColumnName];
                                //currentParameter.Value = row[col.ColumnName];//currentDatatable.Rows[currentRowNumber]
                                //ADODB.Parameter param = cmd.CreateParameter()
                            }
                            cmd.ExecuteNonQuery();
                            //}
                        }
                        
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, ex.Source); }
                finally { TerminateDB(); }
            }
                       
        }
        private string GetSqlInsertNQuery(DataTable dt, string tableName, int rowNum)
        {
            string columnsNames = string.Empty;
            string columnsValues = string.Empty;
            foreach (DataColumn col in dt.Columns)
            { columnsNames = columnsNames + "[" + col.ColumnName + "]" + ", "; columnsValues = columnsValues + @"@" + col.ColumnName + ", "; }
            columnsNames = columnsNames.Remove(columnsNames.Length - 2);/*For removing comma*/
            columnsValues = columnsValues.Remove(columnsValues.Length - 2);

            //foreach(var item in dt.Rows[rowNum].ItemArray)  
            //{ 
            //    switch(item.GetType().ToString())
            //    {
            //        case "System.String":
            //        case "System.Char":
            //        case "System.Guid":
            //            columnsValues = columnsValues + @""" + item + @""" + ", ";
            //            break;
            //        case "System.DateTime":
            //        case "System.TimeSpan":
            //            columnsValues = columnsValues + Convert.ToDouble(item) + ", ";
            //            break;
            //        default:
            //            columnsValues = columnsValues + item + ", ";
            //            break;
            //    }
            //}

            string fullReturnString = "INSERT INTO " + tableName + "(" + columnsNames + ") VALUES(" + columnsValues + ");";
            return fullReturnString;
                            //conPeople.Execute("INSERT INTO Persons(LastName, Gender, FirstName) " +
                            //  "VALUES('Germain', 'Male', 'Ndongo');", out obj, 0);
        }
        private void SetStringAccessStorageDB()
        {
            accessStorageDB = currentFolder + "ExtractorStorage.accdb";
            int i = 1;
            do { accessStorageDB = currentFolder + "(" + i + ")" + "ExtractorStorage.accdb"; } while (System.IO.File.Exists(accessStorageDB));
        }
        private void dataObjectsListView_SelectedIndexChanged(object sender, EventArgs e) { PreviewMode(); }
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

            //DataObject currentDataObject = dataObjectCollecion[dataObjectsListView.SelectedItems[0].SubItems[0].Text];
            //previewGridView.DataSource = currentDataObject.DataTable;
            //queryTextBox.Text = currentDataObject.sqlQuery;
        }
        private void saveDataObjectButton_Click(object sender, EventArgs e)
        {
            DataObject currentDataObject = dataObjectCollecion[dataObjectsListView.SelectedItems[0].SubItems[0].Text];
            currentDataObject.DataTable = (System.Data.DataTable)previewGridView.DataSource;
            currentDataObject.sqlQuery = queryTextBox.Text;
            MessageBox.Show("Saved successfully.");
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)//sourceSharePointRadioButton; annoying
        {
            excelFileBrowsePathButton.Visible = false;
            //excelFileSheetsComboBox.Visible = false;
        }
        private void sourceExcelRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            excelFileBrowsePathButton.Visible = true;
            excelFileSheetsComboBox.Visible = true;
        }
        private void exportToCsvButton_Click(object sender, EventArgs e)
        {
            if (previewGridView.DataSource == null) { return; }
            string savePath = JazzyFunctionsByPatryk.BrowseSavePath("csv");
            JazzyFunctionsByPatryk.DataTableToCSVFile(this.previewGridView.DataSource as System.Data.DataTable, savePath);
        }
    }
}
