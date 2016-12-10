using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report_generator
{
    class DataObject
    {
        public string Name;
        public string sqlQuery;
        public System.Data.DataTable DataTable;

        /* Constructor that takes one argument.*/
        public DataObject(string newName) { Name = newName; }
        // Method
        public void SetQueryAndDataTable(string sql, string excelFilePath)
        {
            sqlQuery = sql;
            DataTable = JazzyFunctionsByPatryk.GetDataTable(JazzyFunctionsByPatryk.GetConnectionStringExcel(excelFilePath), sql);
        }
    }
}
