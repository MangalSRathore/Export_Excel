using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string newExcelFile = @"M:\Person.xlsx";
            new Program().Export(newExcelFile);
        }
        private void Export(string file)
        {
            var list = new List<Person>
            {
                new Person{Name="Mangal",Last_Name="Rathore",Street="03",State="M.P.",Zip="474006"}

            };
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using(ExcelPackage pck = new ExcelPackage())
            {
                pck.Workbook.Worksheets.Add("Person").Cells[1, 1].LoadFromCollection(list, true);
                pck.SaveAs(new System.IO.FileInfo(file));
            }
        }
    }
    public class Person  
    {
       public string Name { get; set; }
        public string Last_Name { get; set; }
        public string Street  { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
    }
}
