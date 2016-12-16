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
using Microsoft.Office.Interop.Excel;

/*Developer Comment*/
//Rough note (ignore it)

namespace Report_generator
{
    public partial class mainForm : Form
    {
        //public ADOX.Catalog tempCatalog;
        //public JazzyFunctionsByPatryk jf;
        public Dictionary<string, DataObject> dataObjectCollecion;
        Microsoft.Office.Interop.Excel.Application excelApp;
        Microsoft.Office.Interop.Excel.Workbook tempExcelWorkbook;
        public mainForm() { InitializeComponent(); dataObjectCollecion = new Dictionary<string, DataObject>(); }
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
            var excelFileSheetsList = JazzyFunctionsByPatryk.ListSheetInExcelInterop(workbookPath);
            if (excelFileSheetsList == null) { MessageBox.Show("Could not load the sheets. Is the Excel file correct?"); return; }
            this.excelFileSheetsComboBox.DataSource = excelFileSheetsList;

            this.queryLoadButton.Enabled = true;
        }
        private void excelFileSheetsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*Fill the fields list*/
        }
        private void queryLoadButton_Click(object sender, EventArgs e)
        {
            /*Verify user input*/
            /*Load query to data grid or return connection error*/

            LoadQueryFromTextBoxToGridView(queryTextBox, this.excelFilePathTextbox.Text);
            //this.exportToExcelButton.Enabled = true;
        }

        private void LoadQueryFromTextBoxToGridView(Control currentQueryTextBox, string excelFilePath)
        {
            System.Data.DataTable excelSheetDataTable;
            /*Prepare SQL Query*/
            string queryString = currentQueryTextBox.Text;
            if (queryString == "") { queryString = "SELECT * FROM [" + this.excelFileSheetsComboBox.Text.Replace("'", "") + "$]"; /*+ "$B1:G3]"*/ }
            this.queryTextBox.Text = queryString;

            /*Preparing database connection*/
            string excelFileConnectionString = JazzyFunctionsByPatryk.GetConnectionStringExcel(excelFilePath);

            //Create data table
            excelSheetDataTable = JazzyFunctionsByPatryk.GetDataTable(excelFileConnectionString, queryString);

            /*Extract data into data grid on form */
            //if (excelSheetDataTable != null)
            //{ if (excelSheetDataTable.Rows.Count > 0) { this.excelQueryGridView.DataSource = excelSheetDataTable; queryTextBox.Text = queryString; } }
            this.previewGridView.DataSource = excelSheetDataTable; currentQueryTextBox.Text = queryString;
        }
        private void exportToExcelButton_Click(object sender, EventArgs e)
        {
            /*Ask for the target location*/
            /*Save in XLS*/

            if (previewGridView.DataSource == null) { return; }
            string savePath = JazzyFunctionsByPatryk.BrowseSavePath("csv");
            //BindingSource bs = (BindingSource)this.excelQueryGridView.DataSource; // Se convierte el DataSource 
            //DataTable tCxC = (DataTable)bs.DataSource;
            JazzyFunctionsByPatryk.DataTableToCSVFile(this.previewGridView.DataSource as System.Data.DataTable, savePath);
        }
        private void mainForm_Load(object sender, EventArgs e) { /*DemSwitchez(0);*/ }
        private void quitButton_Click(object sender, EventArgs e)
        {
            /*Delete the temp database and exit*/
            //var con = tempCatalog.ActiveConnection as ADODB.Connection;
            //if (con != null) con.Close();

            //JazzyFunctionsByPatryk.ReleaseObject(tempExcelWorkbook);
            try { System.IO.File.Delete(tempExcelWorkbook.FullName); }//if (tempExcelWorkbook.FullName.Length > 0) { 
            catch { }
            finally { Environment.Exit(1); }
            
        }
        private void CreateTemporaryDatabase()
        {
            string extractorFolderPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            string temporaryExcelFilePath = extractorFolderPath + @"\ExtractorTempFile.xls";

            excelApp = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook tempExcelWorkbook = null;
            tempExcelWorkbook = excelApp.Workbooks.Add();
            tempExcelWorkbook.SaveAs(temporaryExcelFilePath);
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
        private void masterQueryLoadButton_Click(object sender, EventArgs e)
        {
            CreateTemporaryDatabase();
            foreach (var key in dataObjectCollecion.Keys)
            {
                DataObject currentDataObject = dataObjectCollecion[key];
                JazzyFunctionsByPatryk.DataTableToXLSFile(currentDataObject.DataTable as System.Data.DataTable, tempExcelWorkbook.FullName, currentDataObject.Name);
            }
            tempExcelWorkbook.Close(true);


            //JazzyFunctionsByPatryk.ReleaseObject(tempExcelWorkbook);
            //JazzyFunctionsByPatryk.ReleaseObject(excelApp);
            JazzyFunctionsByPatryk.KillTask("EXCEL");
            LoadQueryFromTextBoxToGridView(masterQueryTextBox, tempExcelWorkbook.FullName);
            System.IO.File.Delete(tempExcelWorkbook.FullName);

            //MessageBox.Show(tempExcelWorkbook.FullName);
            //cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + "; Jet OLEDB:Engine Type=5");
            //tempCatalog.Tables.Append(newTable);
        }


//         try
//  {
//    m_COMObject.SomeMethod();
//  }

//  Exception(exception exception)
//  {
//    DisposeCOMObject();
//    InitializeCOMOBject();
//    COMObject.Somethod();
//  }


// public void DisposeCOMObject()
//{
//  m_COMObject = null;
//  var process = Process.GetProcessesByNames("COM .exe").FirstDefault();

//   if(process != null)
//    {
//         process.kill();
//       }
//}

// public void InitializeCOMObject()
//{
//  m_COMObject = null;
//  m_COMObject = new COMObject();
//}
        private void dataObjectsListView_SelectedIndexChanged(object sender, EventArgs e) { PreviewMode(); }
        private void PreviewMode()//int listIndex
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
            excelFileSheetsComboBox.DataSource = null;
            excelFileSheetsComboBox.Items.Clear();
        }
        private void tableDeleteButton_Click(object sender, EventArgs e)
        {
            if (dataObjectsListView.SelectedItems.Count == 0) { return; }
            string currentDataObjectName = dataObjectsListView.SelectedItems[0].SubItems[0].Text;
            if (MessageBox.Show("Are you sure that you want to remove the " + currentDataObjectName + " data object?", "Confirmation required", MessageBoxButtons.YesNo) == DialogResult.No) { return; }
            dataObjectCollecion.Remove(currentDataObjectName);
            dataObjectsListView.SelectedItems[0].Remove();
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

        private void exportToExcelButton2_Click(object sender, EventArgs e)
        {
            if (previewGridView.DataSource == null) { return; }
            string savePath = JazzyFunctionsByPatryk.BrowseSavePath("csv");
            JazzyFunctionsByPatryk.DataTableToXLSFile(this.previewGridView.DataSource as System.Data.DataTable, savePath);
        }
    }
}
