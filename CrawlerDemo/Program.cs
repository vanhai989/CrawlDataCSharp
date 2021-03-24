using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using HtmlAgilityPack;
using System.Text;
using System.Data.Odbc;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace CrawlerDemo
{
    class Program
    {
       static SqlConnection con = new SqlConnection();
        static void Main(string[] args)
        {
            startCrawlerasync();
            Console.ReadLine();
            
        }

        private static async Task startCrawlerasync()
        {
           
            try
            {

             //   var url = "https://www.datingcelebs.com/who-is-shane-filan-dating/";
                //the url of the page we want to test
                  var url = "https://tinbanxe.vn/gia-xe-oto";

                var httpClient = new HttpClient();
                var html = await httpClient.GetStringAsync(url);
                var htmlDocument = new HtmlDocument();
                htmlDocument.LoadHtml(html);

                // a list to add all the list of cars and the various prices 
                var cars = new List<Car>();
                var divs =
                htmlDocument.DocumentNode.Descendants("div")
                    .Where(node => node.GetAttributeValue("class", "").Equals("td_module_flex td_module_flex_5 td_module_wrap td-animation-stack")).ToList();

                foreach (var div in divs)
                {

                    var car = new Car
                    {

                        Name = div.Descendants("h6").FirstOrDefault().InnerText,
                        ImageUrl = div.Descendants("img").FirstOrDefault().ChildAttributes("data-src").FirstOrDefault().Value
                    };

                    cars.Add(car);
                }

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                   Console.WriteLine("Excel is not properly installed!!");
                    return;
                }


                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                foreach(var item in cars)
                {

                }
                xlWorkSheet.Cells[1, 1] = "ID";
                xlWorkSheet.Cells[1, 2] = "Name";
                xlWorkSheet.Cells[2, 1] = "1";
                xlWorkSheet.Cells[2, 2] = "One";
                xlWorkSheet.Cells[3, 1] = "2";
                xlWorkSheet.Cells[3, 2] = "Two";



                xlWorkBook.SaveAs("d:\\csharp-Excel.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

               Console.WriteLine("Excel file created , you can find the file d:\\csharp-Excel.xls");
            }
            catch (Exception ex )
            {
                Console.WriteLine(ex);
            }

            // Connection string 
       //     string MyConnection = "Server =desktop-15ipd1k; Database =CrawlData; Trusted_Connection = True; MultipleActiveResultSets = true";
  

        

        }
       
    }
}
