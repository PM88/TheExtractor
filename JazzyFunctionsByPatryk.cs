using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//using Microsoft.Office.Interop.Excel;

namespace Report_generator
{
    public static class JazzyFunctionsByPatryk //ver271216_1
    {
        public static string queryString;
        public static string connectionStringExcel;

        //public void SetQueryString(string qs) { queryString = qs; }
        //public void SetConnectionStringExcel(string excelFilePath) { connectionStringExcel = GetConnectionStringExcel(excelFilePath); }
        public static void DataTableToCSVFile(System.Data.DataTable dt, string targetPath)
        {
            StringBuilder sb = new StringBuilder();

            string[] columnNames = dt.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName).
                                              ToArray();
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt.Rows)
            {
                string[] fields = row.ItemArray.Select(field => field.ToString()).
                                                ToArray();
                sb.AppendLine(string.Join(",", fields));
            }

            File.WriteAllText(targetPath, sb.ToString());
        }
        //public static void DataTableToXLSFile(System.Data.DataTable dt, string targetPath, string sheetName = "")
        //{
        //    var excelApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
        //    if (! File.Exists(targetPath))
        //    {
        //        excelWorkBook = excelApp.Workbooks.Add();
        //        excelWorkBook.SaveAs(targetPath);
        //    }
        //    else { excelWorkBook = excelApp.Workbooks.Open(targetPath); }
            
        //    Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Worksheet)excelWorkBook.Sheets.Add();
            
        //    if(sheetName != "") {excelWorkSheet.Name = sheetName;}

        //    for (int i = 1; i < dt.Columns.Count + 1; i++)
        //    { excelWorkSheet.Cells[1, i] = dt.Columns[i - 1].ColumnName; }

        //    for (int j = 0; j < dt.Rows.Count; j++)
        //    {
        //        for (int k = 0; k < dt.Columns.Count; k++)
        //        { excelWorkSheet.Cells[j + 2, k + 1] = dt.Rows[j].ItemArray[k].ToString(); }
        //    }
        //    excelWorkBook.Save();
        //    KillTask("EXCEL");
        //}
        public static System.Data.DataTable GetDataTable(string connectionString, string queryString)
        {
            var fileConnection = new System.Data.OleDb.OleDbConnection(connectionString);
            var dataTable = new System.Data.DataTable();
            try
            {
            var dataAdapter = new System.Data.OleDb.OleDbDataAdapter(queryString, fileConnection);
            dataAdapter.Fill(dataTable); }
            catch (Exception e) { MessageBox.Show(e.Message.ToString()); } //"Connection failed. Check the SQL query."
            finally { fileConnection.Close(); }

            return dataTable;
        }
        public static string GetConnectionString(string filePath)
        {
            var sbConnection = new OleDbConnectionStringBuilder();
            string strExtendedProperties = string.Empty;
            sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
            //string strDataSource = string.Empty;
            string sourceFileExtension = System.IO.Path.GetExtension(filePath);
            //int sourceType;
            //string connectionPropertiesExcelVersion = string.Empty;

            switch (sourceFileExtension)
            {
                case ".xls": strExtendedProperties = "Excel 8.0; HDR = Yes; IMEX = 1;";  break; 
                case ".xlsx": strExtendedProperties = "Excel 12.0 Xml; HDR = Yes; IMEX = 1;";  break;
                case ".xlsm": strExtendedProperties = "Excel 12.0 Macro; HDR = Yes; IMEX = 1;Integrated Security=True;READONLY=1;"; break; //test
                //case ".accdb": break;//strDataSource = "|DataDirectory|"; //strExtendedProperties = "Persist Security Info = False;" //sbConnection.PersistSecurityInfo = false;
                default: break; //sbConnection.Provider = "Microsoft.Jet.Oledb.4.0";
            }
            //strDataSource +=filePath;
            sbConnection.DataSource = filePath;
            //if (sourceFileExtension == ".accdb") { sbConnection.DataSource = "|DataDirectory|" + filePath; } else { sbConnection.DataSource = filePath; }
            
            if (!(strExtendedProperties == string.Empty)) { sbConnection.Add("Extended Properties", strExtendedProperties); }
            return sbConnection.ToString(); //excelFileConnectionString;
            //string connectionProperties = connectionPropertiesExcelVersion + "; HDR=YES\";"; //HDR means that the first row is header
            //string excelFileConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + excelFilePath
            //    + "; Extended Properties=" + connectionProperties;
        }
        public static List<string> ListSheetInExcel(string excelFilePath)
        {
            var listSheet = new List<string>();
            string connectionString = GetConnectionString(excelFilePath);
            using (var conn = new OleDbConnection(connectionString)) //sbConnection.ToString()))
            {
                conn.Open();
                System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))//checks whether row contains '_xlnm#_FilterDatabase' or sheet name(i.e. sheet name always ends with $ sign)
                    { listSheet.Add(drSheet["TABLE_NAME"].ToString().Replace("$", "")); }
                }
                conn.Close();
            }
            return listSheet;
        }
        //public static List<string> ListSheetInExcelInterop(string excelFilePath)
        //{
        //    var listSheet = new List<string>();
        //    var excelApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
        //    try 
        //    { 
        //        excelWorkbook = excelApp.Workbooks.Open(excelFilePath, 0, true);
        //        var excelWorksheets = excelWorkbook.Worksheets;
        //        foreach (Worksheet worksheet in excelWorksheets) { listSheet.Add(worksheet.Name); }
        //    }
        //    catch { listSheet = null; }
        //    KillTask("EXCEL");
        //    return listSheet;
        //    //ReleaseObject(excelApp); ReleaseObject(excelWorkbook); ReleaseObject(excelWorksheets); //ReleaseObject(worksheet);
        //}
        public static string BrowseSavePath(string extension = "") //BrowseSavePath and BrowseFilePath will be refactored into one
        {
            string filter;
            if (extension == "") { filter = "All files (*.*)|*.*"; } else { filter = "(*." + extension + ")|*." + extension; }

            var sfd = new SaveFileDialog();
            sfd.Filter = filter;

            if (sfd.ShowDialog() == DialogResult.OK) { return sfd.FileName; } else { return ""; }
        }
        public static string BrowseFilePath(string browseFilter = "") //string extension = "", 
        {
            string filter;
            if (browseFilter == "") { filter = "All files (*.*)|*.*"; } else { filter = browseFilter; }// "(*." + extension + ")|*." + extension; }
            //if (extension == "") { filter = "All files (*.*)|*.*"; } else { filter = "(*." + extension + ")|*." + extension; }

            var ofd = new OpenFileDialog();
            ofd.Filter = filter;

            if (ofd.ShowDialog() == DialogResult.OK) { return ofd.FileName; } else { return ""; }
        }
        //public static void ReleaseObject(object obj)
        //{
        //    try
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //        obj = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj = null;
        //        MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
        //    }
        //    finally
        //    { GC.Collect(); }
        //}
        public static void KillTask(string ProcessesName)
        {
            /*Kill the all process obj from the Task Manager(Process)*/
            System.Diagnostics.Process[] objProcesses = System.Diagnostics.Process.GetProcessesByName(ProcessesName);

            if (objProcesses.Length > 0)
            {
                var objHashtable = new System.Collections.Hashtable();

                /* check to kill the right process*/
                foreach (System.Diagnostics.Process process in objProcesses)
                { if (objHashtable.ContainsKey(process.Id) == false) { process.Kill(); } }
                objProcesses = null;
            }

            //Quick Answer:
            //foreach (var process in Process.GetProcessesByName("whatever"))
            //{ process.Kill(); }

            /*In case of you want to quit what you have created the Excel object from your application //just use below condition in above,
                if (processInExcel.MainWindowTitle.ToString() == "") ;*/
        }

        //public string GetHTMLStringFromDataTable(System.Data.DataTable dt, bool enableOuterMarkupTags = true)
        //{
        //    //Convert date columns to string w/o time if 
        //    foreach (System.Data.DataColumn myColumn in dt.Columns)
        //    {
        //        //todo ; convert dates to string and cut " 00:00:00" that way will keep dates with time other than 0

        //    }

        //    StringBuilder strHTMLBuilder = new StringBuilder();

        //    //Open structure tags
        //    if (enableOuterMarkupTags)
        //    {
        //        strHTMLBuilder.Append("<html >");
        //        strHTMLBuilder.Append("<head>");
        //        strHTMLBuilder.Append("</head>");
        //        strHTMLBuilder.Append("<body>");
        //    }

        //    //Table tags
        //    //Table properties
        //    strHTMLBuilder.Append("<table >"/* + 
        //        "border='1px' " +
        //        "cellpadding='5px' " +
        //        "cellspacing='0px' " +
        //        "bgcolor='lightyellow' " +
        //        "style='font-family:Garamond; font-size:smaller'>"*/
        //        );

        //    //Header
        //    strHTMLBuilder.Append("<tr >");
        //    foreach (System.Data.DataColumn myColumn in dt.Columns)
        //    {
        //        strHTMLBuilder.Append("<td >");
        //        strHTMLBuilder.Append(myColumn.ColumnName);
        //        strHTMLBuilder.Append("</td>");

        //    }
        //    strHTMLBuilder.Append("</tr>");

        //    //Rows
        //    foreach (System.Data.DataRow myRow in dt.Rows)
        //    {

        //        strHTMLBuilder.Append("<tr >");
        //        //Columns
        //        foreach (System.Data.DataColumn myColumn in dt.Columns)
        //        {
        //            strHTMLBuilder.Append("<td >");
        //            strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
        //            strHTMLBuilder.Append("</td>");

        //        }
        //        strHTMLBuilder.Append("</tr>");
        //    }
        //    strHTMLBuilder.Append("</table>");

        //    //Close tags
        //    if (enableOuterMarkupTags)
        //    {
        //        strHTMLBuilder.Append("</body>");
        //        strHTMLBuilder.Append("</html>");
        //    }

        //    //Output
        //    string returnString = strHTMLBuilder.ToString();
        //    return returnString;
        //}

    }
}
