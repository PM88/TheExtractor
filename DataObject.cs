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
        public bool RunLoad;
        public string Password;
        //public string sourceAddress;
        //public string sourceExcelSheet;

        /* Constructor that takes one argument.*/
        public DataObject(string newName) 
        { 
            Name = newName;
            SqlQuery = string.Empty; /* Should not be null */
            ExcelFilePath = string.Empty;
            ExcelFileSheet = string.Empty;
            Description = string.Empty;
        }
        // Method
        public void SetQueryAndDataTable(string sql, string excelFilePath)
        {
            SqlQuery = sql;
            DataTable = FunRepository.GetDataTable(FunRepository.GetConnectionString(excelFilePath), sql);
        }
        public DataObject CloneMe(string cloneName)
        {
            var newClone = new DataObject(cloneName);
            newClone.Name = Name; 
            newClone.SqlQuery = SqlQuery;
            newClone.DataTable = DataTable;
            newClone.ExcelFilePath = ExcelFilePath;
            newClone.ExcelFileSheet = ExcelFileSheet;
            newClone.PersStorage = PersStorage;
            newClone.Description = Description;
            newClone.RunLoad = RunLoad;
            newClone.Password = Password;

            return newClone;
        }
    }
}
