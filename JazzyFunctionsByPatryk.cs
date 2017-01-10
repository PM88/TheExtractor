using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ADOX;
using Microsoft.VisualBasic;

using Microsoft.Office.Interop.Excel;

/* Developer comment */
// Rough notes

namespace Report_generator
{
    public static class JazzyFunctionsByPatryk //ver090116_1
    {
        public static string queryString;
        public static string connectionStringExcel;

        public static void WriteSettings(string masterQuery, Dictionary<string, DataObject> dataObjectCollecion) 
        {
            string separatorDO = @"!";
            string fileName = JazzyFunctionsByPatryk.BrowseSavePath("txt");
            StreamWriter sw = new StreamWriter(fileName, false); /* True stands for appending (not overwrt) */
            //sw.WriteLine("*******");
            if (masterQuery != "") { sw.WriteLine("Master_Query = " + masterQuery); }
            foreach (DataObject currentDO in dataObjectCollecion.Values)
            {
                sw.WriteLine("Data_Object" + separatorDO + currentDO.Name + separatorDO + "SQL_Query = " + currentDO.SqlQuery);
                if(currentDO.PersStorage)
                {
                    sw.WriteLine("Data_Object" + separatorDO + currentDO.Name + separatorDO + "Source_path = " + currentDO.ExcelFilePath);
                    sw.WriteLine("Data_Object" + separatorDO + currentDO.Name + separatorDO + "Source_table = " + currentDO.ExcelFileSheet);
                }
            }
            //sw.WriteLine(presetsName);
            //sw.WriteLine(operationType);
            //sw.WriteLine(tag);
            //sw.WriteLine(sqlQuery);
            sw.Close();
        }
        public static void ReadSettings(ref string masterQuery, ref Dictionary<string, DataObject> dataObjectCollecion)
        {
            string fileName = JazzyFunctionsByPatryk.BrowseFilePath("Preset files in txt (*.txt)|*.txt");
            dataObjectCollecion.Clear();
            StreamReader sr = new StreamReader(fileName);

            string separator = " = ";
            string separatorDO = @"!";
            string line;
            string[] presetTypeArray;
            int prefixLength;
            string presetType;
            string presetSubType;
            string presetValue;
            string dataObjectName;
            DataObject currentDataObject;
            

            while ((line = sr.ReadLine()) != null)
            {
                presetType = line.Substring(0, line.IndexOf(separator, 0));
                prefixLength = separator.Length + line.IndexOf(separator, 0);
                presetValue = line.Substring(prefixLength, line.Length - prefixLength);
                switch (presetType)
                {
                    case "Master_Query": masterQuery = presetValue; break;
                    default: if(presetType.Contains(separatorDO))
                        {
                            presetTypeArray = presetType.Split('!');
                            dataObjectName = presetTypeArray[1];
                            presetSubType = presetTypeArray[2];
                            //dataObjectName = presetType.Substring(presetType.IndexOf(separatorDO, 0), );
                            if (!(dataObjectCollecion.ContainsKey(dataObjectName)))
                            { dataObjectCollecion.Add(dataObjectName, new DataObject(dataObjectName)); }

                            currentDataObject = dataObjectCollecion[dataObjectName];
                            switch(presetSubType)
                            {
                                case "SQL_Query": currentDataObject.SqlQuery = presetValue; break;
                                case "Source_path": 
                                    currentDataObject.ExcelFilePath = presetValue;
                                    currentDataObject.PersStorage = true;
                                    break;
                                case "Source_table": 
                                    currentDataObject.ExcelFileSheet = presetValue;
                                    currentDataObject.PersStorage = true;
                                    break;
                            }
                        } break;
                }
            }
            sr.Close();
        }
        public static int IndexOfNth(this string input, string value, int startIndex, int nth)
        {
            if (nth < 1)
                throw new NotSupportedException("Param 'nth' must be greater than 0!");
            if (nth == 1)
                return input.IndexOf(value, startIndex);
            var idx = input.IndexOf(value, startIndex);
            if (idx == -1)
                return -1;
            return input.IndexOfNth(value, idx + 1, --nth);
        }
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
        public static string GetCleanAccessObjectName(string oldName, bool isTable = false)
        {
            string newName = oldName;
            string[] illegals = { @"`", @"!", @".", @"[", @"]" };
            foreach (string currentIllegal in illegals) { newName = newName.Replace(currentIllegal, " "); }
            if (isTable) { newName = newName.Replace(@"""", " "); } /* Double quotes are illegal for tables, views and stored procedures */
            return newName;
        }
        public static ADOX.Table GetNewAdoxTable(System.Data.DataTable dt, string name)
        {
            var newTable = new ADOX.Table();
            newTable.Name = name;

            foreach (DataColumn col in dt.Columns)
            {
                ADOX.Column dbField = new Column();
                dbField.Name = JazzyFunctionsByPatryk.GetCleanAccessObjectName(col.ColumnName);
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
            return newTable;
        }
        public static string GetSqlInsertNonQuery(System.Data.DataTable dt, string tableName)
        {
            string columnsNames = string.Empty;
            string columnsValues = string.Empty;
            int paramIteration = 1;
            foreach (DataColumn col in dt.Columns)
            { columnsNames += "[" + JazzyFunctionsByPatryk.GetCleanAccessObjectName(col.ColumnName) + "]" + ", "; columnsValues += @"@" + paramIteration + ", "; paramIteration += 1; }//GetCleanParameterName(col.ColumnName)
            columnsNames = columnsNames.Remove(columnsNames.Length - 2);/*For removing final comma*/
            columnsValues = columnsValues.Remove(columnsValues.Length - 2);

            string fullReturnString = "INSERT INTO [" + tableName + "] (" + columnsNames + ") VALUES(" + columnsValues + ");";
            return fullReturnString;
        }
        public static void SetOleDbCommandParameters(ref OleDbCommand cmd, System.Data.DataTable dt)
        {
            int paramIteration = 1;
            foreach (DataColumn col in dt.Columns)
            {
                switch (col.DataType.ToString())
                {
                    case "System.String":
                    case "System.Char":
                    case "System.Guid": cmd.Parameters.Add(GetOleDbParam(paramIteration.ToString(), 0)); break;
                    case "System.DateTime":
                    case "System.TimeSpan": cmd.Parameters.Add(GetOleDbParam(paramIteration.ToString(), 1)); break;
                    case "System.Boolean": cmd.Parameters.Add(GetOleDbParam(paramIteration.ToString(), 2)); break;
                    default: cmd.Parameters.Add(GetOleDbParam(paramIteration.ToString(), 3)); break;
                }
                paramIteration += 1;
            }
        }
        private static OleDbParameter GetOleDbParam(string name, int code)
        {
            OleDbParameter newParameter = new OleDbParameter();
            newParameter.ParameterName = @"@" + name;// GetCleanParameterName(name);
            switch (code)
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
        public static string SummonInputBox(string prompt, string title = "", string defaultResponse="") { return Interaction.InputBox(prompt, title, defaultResponse); }
        public static void DataTableToExcelFile(System.Data.DataTable dt, string targetPath) /* Experimental, breaks at connection open due to unrecognized format */
        {
            File.Create(targetPath).Close();
            var cmd = new OleDbCommand();
            //var con = new OleDbConnection(GetConnectionString(targetPath));
            using (OleDbConnection con = new OleDbConnection(GetConnectionString(targetPath)))
            {
                con.Open();
                cmd.Connection = con;
                //Create table (sheet)
                string tableName = "Sheet1";
                cmd.CommandText = GetSqlCreateNonQuery(dt, tableName);

                try
                { 
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = GetSqlInsertNonQuery(dt, tableName);
                    JazzyFunctionsByPatryk.SetOleDbCommandParameters(ref cmd, dt);

                    foreach (DataRow row in dt.Rows)
                    {
                        int paramIteration = 1;
                        foreach (DataColumn col in dt.Columns)
                        { cmd.Parameters[@"@" + paramIteration].Value = row[col.ColumnName]; paramIteration += 1; }//GetCleanParameterName(col.ColumnName)
                        cmd.ExecuteNonQuery();
                    }
                    MessageBox.Show("Saved successfully.");
                }
                catch(Exception e){MessageBox.Show(e.Message);}
                finally { con.Close(); }
            }
        }
        private static string GetSqlCreateNonQuery(System.Data.DataTable dt, string tableName)
        {
            //cmd.CommandText = "CREATE TABLE [table1] (id INT, name VARCHAR, datecol DATE );";
            string columnsNamesWithType = string.Empty;
            foreach (DataColumn col in dt.Columns)
            { columnsNamesWithType += JazzyFunctionsByPatryk.GetColumnNameAndTypeForSqlCreate(col.ColumnName, col) + ", "; }
            columnsNamesWithType = columnsNamesWithType.Remove(columnsNamesWithType.Length - 2);/*For removing final comma*/

            string fullReturnString = "CREATE TABLE [" + tableName + "] (" + columnsNamesWithType + ");";
            return fullReturnString;
        }
        private static string GetColumnNameAndTypeForSqlCreate(string columnName, DataColumn col)
        {
            string columnNameAndType = "[" + columnName + "] ";
            switch (col.DataType.ToString())
            {
                case "System.String":
                case "System.Char":
                case "System.Guid": columnNameAndType += "TEXT(255)"; break;
                case "System.DateTime":
                case "System.TimeSpan": columnNameAndType += "DATETIME"; break;
                case "System.Boolean": columnNameAndType += "YESNO"; break;
                default: columnNameAndType += "DECIMAL"; break;
            }
            return columnNameAndType;
        }
        public static void DataTableToExcelFileWithInterop(System.Data.DataTable dt, string targetPath, string sheetName = "Sheet1")
        {
            var dialogResult = MessageBox.Show("Excel must be closed lest you lose your work progress. Continue?" , "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.No) { return; }

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
            if (!File.Exists(targetPath))
            {
                excelWorkBook = excelApp.Workbooks.Add();
                excelWorkBook.SaveAs(targetPath);
            }
            else { excelWorkBook = excelApp.Workbooks.Open(targetPath); }

            Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = (Worksheet)excelWorkBook.Sheets.Add();

            excelWorkSheet.Name = sheetName; //if (sheetName != "") { }

            for (int i = 1; i < dt.Columns.Count + 1; i++)
            { excelWorkSheet.Cells[1, i] = dt.Columns[i - 1].ColumnName; }

            for (int j = 0; j < dt.Rows.Count; j++)
            {
                for (int k = 0; k < dt.Columns.Count; k++)
                { excelWorkSheet.Cells[j + 2, k + 1] = dt.Rows[j].ItemArray[k].ToString(); }
            }
            excelWorkBook.Save();
            KillTask("EXCEL");
            MessageBox.Show("Saved successfully to: " + Environment.NewLine + targetPath);
        }
        //### FUNCTIONS MORGUE ####
        #region Local vars
        //public void SetQueryString(string qs) { queryString = qs; }
        //public void SetConnectionStringExcel(string excelFilePath) { connectionStringExcel = GetConnectionStringExcel(excelFilePath); }
        #endregion
        //####################################################################################################################################
        #region public static List<string> ListSheetInExcelInterop(string excelFilePath)
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
        #endregion
        //####################################################################################################################################
        #region public static void ReleaseObject(object obj)
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
        #endregion
        //####################################################################################################################################
        #region  public string GetHTMLStringFromDataTable(System.Data.DataTable dt, bool enableOuterMarkupTags = true)
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
        #endregion 
        //####################################################################################################################################
        #region public static void DataTableToCSVFile(System.Data.DataTable dt, string targetPath)
        //public static void DataTableToCSVFile(System.Data.DataTable dt, string targetPath)
        //{
        //    StringBuilder sb = new StringBuilder();

        //    string[] columnNames = dt.Columns.Cast<DataColumn>().
        //                                      Select(column => column.ColumnName).
        //                                      ToArray();
        //    sb.AppendLine(string.Join("\t", columnNames));

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        string[] fields = row.ItemArray.Select(field => field.ToString()).
        //                                        ToArray();
        //        sb.AppendLine(string.Join("\t", fields));
        //    }

        //    File.WriteAllText(targetPath, sb.ToString());
        //}
        #endregion 
        //####################################################################################################################################



    }
}
