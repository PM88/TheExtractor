using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report_generator
{
    public class DataObject
    {
        public string Name; //Datatable has tablename property but Data Object not always has datatable
        public string SqlQuery;
        public System.Data.DataTable DataTable;
        public string ExcelFilePath;
        public string ExcelFileSheet;
        public bool PersStorage;
        public string Description;
        //public string sourceAddress;
        //public string sourceExcelSheet;

        /* Constructor that takes one argument.*/
        public DataObject(string newName) { Name = newName; }
        // Method
        public void SetQueryAndDataTable(string sql, string excelFilePath)
        {
            SqlQuery = sql;
            DataTable = FunRepository.GetDataTable(FunRepository.GetConnectionString(excelFilePath), sql);
        }
    }
}
