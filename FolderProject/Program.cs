using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FolderProject
{
    /// <summary>
    /// Excel 4.5 library url:https://github.com/ExcelDataReader/ExcelDataReader
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            FileStream stream = File.Open(ConfigurationManager.AppSettings["sourceFile"], FileMode.Open, FileAccess.Read);
            string basePath = ConfigurationManager.AppSettings["basePath"];

            // Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            // DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();

            // Getting distinct values
            var distinctValues = result.Tables[0].AsEnumerable()
                        .Select(row => new
                        {
                            Mem = row.Field<double>("Mem#"),
                            DocumentTypeId = row.Field<string>("DocumentTypeId")
                        })
                        .Distinct();

            // Create reead data tables and create folders
            foreach (var d in distinctValues)
            {
                // Generate folder path
                string folderPath = Convert.ToInt32(d.Mem).ToString()+@"\"+d.DocumentTypeId;

                // Append to base folder
                string pathString = System.IO.Path.Combine(basePath, folderPath);

                // Create folder
                System.IO.Directory.CreateDirectory(pathString);
            }

            // Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
        }
    }
}
