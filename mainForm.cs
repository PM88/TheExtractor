using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

/*Developer Comment*/
//Rough note (ignore it)

namespace Report_generator
{
    public partial class mainForm : Form
    {
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
            this.excelQueryGridView.DataSource = excelSheetDataTable; queryTextBox.Text = queryString;
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
            jf.DataTableToCSVFile(this.excelQueryGridView.DataSource as DataTable, savePath);
        }
    }
}
