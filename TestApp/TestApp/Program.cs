using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string saveLocation = Path.GetFullPath(Path.Combine(path, @"..\..\..\Test\Test.xlsx"));
            Console.WriteLine(saveLocation);
            Console.WriteLine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location));
            var dtContent = GetDataTableFromExcel(saveLocation);
            foreach (DataRow dr in dtContent.Rows)
            {
                Console.WriteLine(dr["Key"].ToString()+" "+ dr["Company"].ToString());
            }
            Console.ReadLine();
        }
        private static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                int count = pck.Workbook.Worksheets.Count;
                Console.WriteLine(count);
                var ws = pck.Workbook.Worksheets["US"];
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }
    }
}
