using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

namespace Report_generator
{
    public partial class mainForm : Form
    {
        public string queryString;
        public System.Data.DataTable excelSheetDataTable;
        public mainForm()
        {
            InitializeComponent();
        }

        private void queryLoadButton_Click(object sender, EventArgs e)
        {
            //Prepare SQL Query
            if (queryTextBox.Text == "")
                queryTextBox.Text = "SELECT * FROM [" + excelFileSheetsComboBox.Text + "$]"; //+ "$B1:G3]"

            queryString = this.queryTextBox.Text;

            //Preparing database connection
            string excelFilePath = System.IO.Path.GetExtension(this.excelFilePathTextbox.Text);
            string connectionPropertiesExcelVersion = string.Empty;

            switch (excelFilePath)
            {
                case ".xls":
                    connectionPropertiesExcelVersion = "\"Excel 8.0";
                    break;
                case ".xlsx":
                    connectionPropertiesExcelVersion = "\"Excel 12.0 Xml";
                    break;
                case ".xlsm":
                    connectionPropertiesExcelVersion = "\"Excel 12.0 Macro";
                    break;
                default:
                    MessageBox.Show("Invalid data type. The only acceptable extensions are: .xls .xlsx .xlsm");
                    return;
            }

            string connectionProperties = connectionPropertiesExcelVersion + "; HDR=YES\";"; //HDR means that the first row is header
            string excelFileConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + this.excelFilePathTextbox.Text
                + "; Extended Properties=" + connectionProperties;

            //Establish connection
            var excelFileConnection = new System.Data.OleDb.OleDbConnection(excelFileConnectionString);
            var excelSheetDataAdapter = new System.Data.OleDb.OleDbDataAdapter(queryString, excelFileConnection);
            excelSheetDataTable = new System.Data.DataTable();

            //Extract data into data grid on form
            try
            {
                excelSheetDataAdapter.Fill(excelSheetDataTable);
                this.excelQueryGridView.DataSource = excelSheetDataTable;
            }
            catch
            {
                MessageBox.Show("Connection failed. Check the SQL query.");
            }
            finally
            {
                excelFileConnection.Close();
            }
        }

        private void excelFilePathButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog excelFilePathOpenFileDialog = new OpenFileDialog();
            excelFilePathOpenFileDialog.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
            if (excelFilePathOpenFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.excelFilePathTextbox.Text = excelFilePathOpenFileDialog.FileName;

                var excelApp = new Microsoft.Office.Interop.Excel.Application();

                string workbookPath = this.excelFilePathTextbox.Text;
                Workbook excelWorkbook = null;
                try
                {
                    excelWorkbook = excelApp.Workbooks.Open(Filename: workbookPath, ReadOnly: true);

                    var excelFileSheetsList = new List<string>();

                    foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in excelWorkbook.Worksheets)
                        excelFileSheetsList.Add(worksheet.Name);

                    this.excelFileSheetsComboBox.DataSource = excelFileSheetsList;
                }
                catch
                {
                    MessageBox.Show("Could not load the sheets from the chosen Excel file.");
                }
                finally
                {
                    excelWorkbook.Close();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //XLWorkbook newWorkBook = new XLWorkbook();
            //newWorkBook.Worksheets.Add(excelSheetDataTable, "Report");
            ////ClosedXML.Excel.IXLWorksheet ws = newWorkBook.Worksheets.;

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string targetPath = System.IO.Path.Combine(desktopPath, "temp.xlsx");
            //newWorkBook.SaveAs(targetPath);
            var fun = new JazzyFunctionsByPatryk_ver030216_1();
            if(fun.DataTableToExcelFile(excelSheetDataTable,targetPath)==true)
            {
                MessageBox.Show("Your file has been created/n" + targetPath);
            }
            else
            {
                MessageBox.Show("Something went wrong");
            }
            
            
            //Environment.Exit(0);
        }
    }
}
