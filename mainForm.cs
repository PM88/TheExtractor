using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ADOX; //Requires Microsoft ADO Ext. 2.8 for DDL and Security
using ADODB;

using Microsoft.VisualBasic;

/*Developer Comment*/
//Rough note (ignore it)

namespace Report_generator
{
    public partial class mainForm : Form
    {
        public ADOX.Catalog tempCatalog;
        /**/
        public JazzyFunctionsByPatryk jf;
        public mainForm() { InitializeComponent(); jf = new JazzyFunctionsByPatryk(); }
        private void excelFilePathButton_Click(object sender, EventArgs e)
        {
            /*Allow the user to browse for eligible Excel file*/
            string workbookPath = jf.BrowseFilePath("Excel Files|*.xlsx;*.xls;*.xlsm");
            this.excelFilePathTextbox.Text = workbookPath;

            /*Fill the sheets list*/
            string connectionString = jf.GetConnectionStringExcel(workbookPath);
            if (connectionString == "") { MessageBox.Show("Could not load the sheets from the chosen Excel file."); return; }

            var excelFileSheetsList = jf.ListSheetInExcel(connectionString);
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

            //test below
            System.Data.DataTable excelSheetDataTable;
            //Prepare SQL Query
            string queryString = queryTextBox.Text;
            if (queryString == "") { queryString = "SELECT * FROM [" + this.excelFileSheetsComboBox.Text.Replace("'","") + "$]"; /*+ "$B1:G3]"*/ }
            this.queryTextBox.Text = queryString;

            //Preparing database connection
            var jf = new JazzyFunctionsByPatryk();
            string excelFileConnectionString = jf.GetConnectionStringExcel(this.excelFilePathTextbox.Text);

            //Create data table
            excelSheetDataTable = jf.GetDataTable(excelFileConnectionString, queryString);

            //Extract data into data grid on form
            //if (excelSheetDataTable != null)
            //{ if (excelSheetDataTable.Rows.Count > 0) { this.excelQueryGridView.DataSource = excelSheetDataTable; queryTextBox.Text = queryString; } }
            this.previewGridView.DataSource = excelSheetDataTable; queryTextBox.Text = queryString;
            this.exportToExcelButton.Enabled = true;
        }
        private void exportToExcelButton_Click(object sender, EventArgs e)
        {
            /*Ask for the target location*/
            /*Save in XLS*/

            //test
            string savePath = jf.BrowseSavePath("csv");
            //BindingSource bs = (BindingSource)this.excelQueryGridView.DataSource; // Se convierte el DataSource 
            //DataTable tCxC = (DataTable)bs.DataSource;
            jf.DataTableToCSVFile(this.previewGridView.DataSource as DataTable, savePath);
        }
        private void mainForm_Load(object sender, EventArgs e) 
        {
            DemSwitchez(flexiLabel.Text, 0); 
            createTempDatabase(); 
        }
        private void quitButton_Click(object sender, EventArgs e)
        {
            /*Delete the temp database and exit*/
            var con = tempCatalog.ActiveConnection as ADODB.Connection;
            if (con != null) con.Close();

            Environment.Exit(1);
        }
        private void createTempDatabase()
        {
            tempCatalog = new ADOX.Catalog();

            //ADOX.Catalog cat = new ADOX.Catalog();
            //ADOX.Table table = new ADOX.Table();

            ////Create the table and it's fields. 
            //table.Name = "Table1";
            //table.Columns.Append("Field1");
            //table.Columns.Append("Field2");

            //try
            //{
            //    cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + "; Jet OLEDB:Engine Type=5");
            //    cat.Tables.Append(table);

            //    //Now Close the database
            //    ADODB.Connection con = cat.ActiveConnection as ADODB.Connection;
            //    if (con != null)
            //        con.Close();

            //    result = true;
            //}
            //catch (Exception ex)
            //{
            //    result = false;
            //}
        }

        private void masterButton_Click(object sender, EventArgs e) { DemSwitchez("Master report", 2); }

        private void DemSwitchez(string flexiLabelText, int switchCode)
        {
            /*Upper-right controls should be visible/hidden depending on context*/
            flexiLabel.Text = flexiLabelText;

            bool masterBool = false;
            bool subBool = false;
            switch (switchCode) { case 1: subBool = true; break; case 2: masterBool = true; break; default: break; }

            masterQueryLoadButton.Visible = masterBool;
            masterQueryExportToExcelButton.Visible = masterBool;
            masterQueryTextBox.Visible = masterBool;

            excelFileBrowsePathButton.Visible = subBool;
            excelFilePathTextbox.Visible = subBool;
            excelFileSheetsComboBox.Visible = subBool;
        }

        private void tableAddButton_Click(object sender, EventArgs e)
        {
            DemSwitchez("Choose the source file and sheet", 1);
            string newTableName = Interaction.InputBox("Provide name of the new data source?", "Enter name"); //, "Default Text"
            var newTable = new ADOX.Table();
            //cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + fileName + "; Jet OLEDB:Engine Type=5");
            tempCatalog.Tables.Append(newTable);
            var newListViewItem = new ListViewItem(newTableName);
            tableListView.Items.Add(newListViewItem);


        }
    }
}
