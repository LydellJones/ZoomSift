using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Consoletestwork
{
    public class ExcelPrint : ExcelParse
    {
        private string[,] Payload;
        private string load;
        public ExcelPrint(string[,] payload)
        {
            Payload = payload;
        }

        public void Print()
        {
            var shopapp = new Excel.Application();
            shopapp.Visible = true;
            shopapp.Workbooks.Add();
            Excel._Worksheet sheet = (Excel.Worksheet)shopapp.ActiveSheet;

            for (int o = 0; o < fullset.GetLength(0); o++)
            {
                for (int i = 0; i < fullset.GetLength(1); i++)
                {
                    
                    load = (string)Payload[o, i];
                    if (load != null) {
                        sheet.Cells[o + 1, i+1] = load;
                    }
                }
            }
            Console.WriteLine("Done!");
        }                    

    }
}
