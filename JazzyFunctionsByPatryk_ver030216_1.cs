using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.OleDb;

namespace Report_generator
{
    class JazzyFunctionsByPatryk_ver030216_1
    {
        public bool DataTableToExcelFile(System.Data.DataTable dt, string targetPath)
        {
            //const bool dontSave = false;
            bool success = true;

            //Exit if there is no rows to export
            if (dt.Rows.Count == 0) return false;

            //object misValue = System.Reflection.Missing.Value;
            List<int> dateColIndex = new List<int>();
            var excelApp = new Microsoft.Office.Interop.Excel.Application();//.Excel.Application
            var excelWorkBook = excelApp.Workbooks.Add(Type.Missing);//.Excel.Workbook
            var excelWorkSheet = excelWorkBook.Sheets["sheet1"];//.Excel.Worksheet
            
            try
            {
                for (int i = -1; i <= dt.Rows.Count - 1; i++) 
                {
                    for (int j = 0; j <= dt.Columns.Count - 1; j++) 
                    {
                        if (i < 0) 
                        {
                            //Take special care with Date columns
                            if (dt.Columns[j].DataType is typeof(DateTime)) 
                            {
                                excelWorkSheet.Cells(1, j + 1).EntireColumn.NumberFormat = "dd/mm/yyyy;@";
                                dateColIndex.Add(j);
                            } 
                            //else if ... Feel free to add more Formats
                            else 
                            {
                                //Otherwise Format the column as text
                                excelWorkSheet.Cells(1, j + 1).EntireColumn.NumberFormat = "@";
                            }
                            excelWorkSheet.Cells[1, j + 1] = dt.Columns[j].Caption;
                        } 
                        else if (dateColIndex.IndexOf(j) > -1) {
                            excelWorkSheet.Cells[i + 2, j + 1] = Convert.ToDateTime(dt.Rows[i].ItemArray[j]).ToString("dd/mm/yyyy");
                        } 
                        else {
                            excelWorkSheet.Cells[i + 2, j + 1] = dt.Rows[i].ItemArray[j].ToString();
                        }
                    }
                }

                //Add Autofilters to the Excel work sheet  
                /*excelWorkSheet.Cells.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);*/
                //Autofit columns for neatness

                //Populate the Excel work sheet with DataTable converted to DataSet
                /*var ds = new System.Data.DataSet();
                ds.Tables.Add(dt);
                excelWorkSheet.Range["A1"].CopyFromRecordset(ds);*/
                
                excelWorkSheet.Columns.AutoFit();
                if (File.Exists(targetPath)) File.Delete(targetPath);
                excelWorkSheet.SaveAs(targetPath);
            } 
            catch 
            {
                success = false;
            } 
            finally 
            {
                //Do this irrespective of whether there was an exception or not. 
                excelWorkBook.Close(SaveChanges:false);
                excelApp.Quit();
                /*releaseObject(excelWorkSheet);
                releaseObject(excelWorkBook);
                releaseObject(excelApp);*/
            }
            return success;
        }
        public string HTMLStringFromDataTable(System.Data.DataTable dt, bool enableOuterMarkupTags = true)
        {
            //Convert date columns to string w/o time if 
            foreach (System.Data.DataColumn myColumn in dt.Columns)
            {
                //pickup ; convert dates to string and cut " 00:00:00" that way will keep dates with time other than 0

            }

            StringBuilder strHTMLBuilder = new StringBuilder();

            //Open structure tags
            if (enableOuterMarkupTags)
            {
                strHTMLBuilder.Append("<html >");
                strHTMLBuilder.Append("<head>");
                strHTMLBuilder.Append("</head>");
                strHTMLBuilder.Append("<body>");
            }

            //Table tags
            //Table properties
            strHTMLBuilder.Append("<table " +
                "border='10px' " +
                "cellpadding='10px' " +
                "cellspacing='10px' >" /*+
                "bgcolor='lightyellow' " +
                "style='font-family:Garamond; font-size:smaller'>"*/
                );

            //Header
            strHTMLBuilder.Append("<tr >");
            foreach (System.Data.DataColumn myColumn in dt.Columns)
            {
                strHTMLBuilder.Append("<td >");
                strHTMLBuilder.Append(myColumn.ColumnName);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");

            //Rows
            foreach (System.Data.DataRow myRow in dt.Rows)
            {

                strHTMLBuilder.Append("<tr >");
                //Columns
                foreach (System.Data.DataColumn myColumn in dt.Columns)
                {
                    strHTMLBuilder.Append("<td >");
                    strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    strHTMLBuilder.Append("</td>");

                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</table>");

            //Close tags
            if (enableOuterMarkupTags)
            {
                strHTMLBuilder.Append("</body>");
                strHTMLBuilder.Append("</html>");
            }

            //Output
            string returnString = strHTMLBuilder.ToString();
            return returnString;
        }
    }
}
