using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
namespace ConsoleApplication9
{
    class Program
    {
        static void Main(string[] args)
        {
       
            string in_name = args[0];
            string out_name = args[1];
           in_name= System.IO.Path.GetFullPath(in_name);
           out_name = System.IO.Path.GetFullPath(out_name);
          ApplicationClass excelapp = new ApplicationClass();
          Workbook excelworkbook = excelapp.Workbooks.Open(in_name,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          ((Microsoft.Office.Interop.Excel._Worksheet)excelworkbook.Worksheets.get_Item(1)).Activate();
          excelworkbook.SaveAs(out_name, XlFileFormat.xlWorkbookDefault, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          excelworkbook.Close(true);
          if (excelworkbook != null)
          {
              System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkbook);
          }
          excelapp.Quit();
          if (excelapp != null)
          {
              System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp);
          }
            
          excelworkbook = null;
          excelapp = null;
          GC.Collect();
          GC.WaitForPendingFinalizers();
         // System.Console.WriteLine("ok");

        }
    }
}
