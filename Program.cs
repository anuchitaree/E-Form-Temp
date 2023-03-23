using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Eform
{
    internal class Program
    {
        static void Main(string[] args)
        {

            string path = @"D:\Project\Eform\Form\exsample.xlsm";  // PpasMaster


            //string path = @"C:\Users\Administrator\Desktop\PpasMaster.xlsx";

            try
            {


                FileInfo excelFile = new FileInfo(path);
                ExcelPackage excel = new ExcelPackage(excelFile);

                //var worksheet1 = excel.Workbook.Worksheets;

                var worksheet = excel.Workbook.Worksheets["Sheet1"];
                //ExcelWorksheet anotherWorksheet = excel.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                {
                    return;
                }
                //using (var package = new ExcelPackage(excelFile.))
                //{
                //    var worksheet = package.Workbook.Worksheets["Sheet1"];
                worksheet.Cells["H1"].Value = "completed";
                excel.Save();
                System.Diagnostics.Process.Start(path);
                //}
            }
            catch (Exception ex)
            {
                string message = ex.Message;
                //throw;
            }
        }
    }
}
