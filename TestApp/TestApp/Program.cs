using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using TestApp.Test;
namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            AddressData addressData = new AddressData();
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string saveLocation = Path.GetFullPath(Path.Combine(path, @"..\..\..\Test\AddressData.xlsx"));
            Console.WriteLine(saveLocation);
            Console.WriteLine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location));
            string[] worksheets = {"US", "Canada", "Germany"};
            foreach(string worksheet in worksheets)
            {
                var dtContent = GetDataTableFromExcel(saveLocation, worksheet);

                foreach (DataRow dr in dtContent.Rows)
                {
                    addressData.Name = dr["Name"].ToString();
                    addressData.Street = dr["Street"].ToString();
                    addressData.Street1 = dr["Street1"].ToString();
                    addressData.Street2 = dr["Street2"].ToString();
                    addressData.City = dr["City"].ToString();
                    addressData.Country = dr["Country"].ToString();
                    addressData.State = dr["State"].ToString();
                    addressData.Zipcode = dr["Zipcode"].ToString();
                    addressData.Telephone = dr["Telephone"].ToString();
                    addressData.VAT = dr["VAT"].ToString();
                    Console.WriteLine(dr["Name"].ToString() + " " + dr["State"].ToString());

                }
            }
           
            Console.ReadLine();
        }
        private static DataTable GetDataTableFromExcel(string path,string sheetName, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                int count = pck.Workbook.Worksheets.Count;
                Console.WriteLine(count);
                var ws = pck.Workbook.Worksheets[sheetName];
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
