﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

namespace Report_generator
{
    public static class JazzyFunctionsByPatryk //ver111216_1
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
        public static System.Data.DataTable GetDataTable(string connectionString, string queryString)
        {
            //Establish connection
            var fileConnection = new System.Data.OleDb.OleDbConnection(connectionString);
            var dataAdapter = new System.Data.OleDb.OleDbDataAdapter(queryString, fileConnection);
            var dataTable = new System.Data.DataTable();

            //Run the command to fill the data table
            try { dataAdapter.Fill(dataTable); }
            catch (Exception e) { MessageBox.Show(e.Message.ToString()); } //"Connection failed. Check the SQL query."
            finally { fileConnection.Close(); }

            return dataTable;
        }
        public static string GetConnectionStringExcel(string excelFilePath)
        {
            var sbConnection = new OleDbConnectionStringBuilder();
            string strExtendedProperties = string.Empty;

            string excelFileExtension = System.IO.Path.GetExtension(excelFilePath);
            //string connectionPropertiesExcelVersion = string.Empty;

            switch (excelFileExtension)
            {
                case ".xls":
                    strExtendedProperties = "Excel 8.0;HDR=Yes;IMEX=1";
                    //connectionPropertiesExcelVersion = "\"Excel 8.0";
                    break;
                case ".xlsx":
                    strExtendedProperties = "Excel 12.0 Xml;HDR=Yes;IMEX=1";
                    //connectionPropertiesExcelVersion = "\"Excel 12.0 Xml";
                    break;
                case ".xlsm":
                    strExtendedProperties = "Excel 12.0 Macro;HDR=Yes;IMEX=1";
                    //connectionPropertiesExcelVersion = "\"Excel 12.0 Macro";
                    break;
                default:
                    // MessageBox.Show("Invalid data type. The only acceptable extensions are: .xls .xlsx .xlsm");
                    return "";
            }

            sbConnection.DataSource = excelFilePath;
            sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
            sbConnection.Add("Extended Properties", strExtendedProperties);
            return sbConnection.ToString(); //excelFileConnectionString;
            //string connectionProperties = connectionPropertiesExcelVersion + "; HDR=YES\";"; //HDR means that the first row is header
            //string excelFileConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + excelFilePath
            //    + "; Extended Properties=" + connectionProperties;
        }
        public static List<string> ListSheetInExcel(string connectionString)
        {
            var listSheet = new List<string>();
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
        public static List<string> ListSheetInExcelInterop(string excelFilePath)
        {
            var listSheet = new List<string>();
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWorkbook = null;
            try { excelWorkbook = excelApp.Workbooks.Open(excelFilePath, 0, true); }
            catch { listSheet = null; ReleaseObject(excelApp); ReleaseObject(excelWorkbook); return listSheet; }

            var excelWorksheets = excelWorkbook.Worksheets;
            foreach (Worksheet worksheet in excelWorksheets) { listSheet.Add(worksheet.Name); }

            ReleaseObject(excelApp); ReleaseObject(excelWorkbook); ReleaseObject(excelWorksheets); //ReleaseObject(worksheet);
            return listSheet;
        }
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
        public static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            { GC.Collect(); }
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
