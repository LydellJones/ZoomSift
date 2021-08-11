using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Consoletestwork
{
    public class ExcelPrint : ExcelParse
    {
        private string[] colindexary = {"A", "B", "C", "D", "E", "F", "G", "H", "I",
            "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W",
            "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI",
            "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW",
            "AX", "AY", "AZ"};
        private int rowindex = 1;
        private int colindex = 0;
        private int searchrindex = 1;
        private int searchcindex = 0;
        private bool samefound;
        private string authorcheck;
        private int rowindexcheck = 1;
        
        Excel.Application shopapp = new Excel.Application();
        
        public ExcelPrint()
        {
            Excel._Worksheet sheet = (Excel.Worksheet)shopapp.ActiveSheet;
            shopapp.Visible = true;
        }
        public void Print()
        {
            searchcindex = 0;
            rowindexcheck = 1;

            Excel._Worksheet sheet = (Excel.Worksheet)shopapp.ActiveSheet;
            while (sheet.Cells[1, colindexary[searchcindex]] != null)
            {
                authorcheck = (string)sheet.Cells[1, colindexary[searchcindex]];
                if (authorcheck == author)
                {
                    samefound = true;
                    break;
                }
                searchcindex++;
            }
            if (samefound == true)
            {
                while (sheet.Cells[rowindexcheck, colindexary[searchcindex]] != null)
                {
                    rowindexcheck++;
                }
                for (int i = 0; i < printableresults.Length; i++)
                {
                    if (printableresults != null)
                    {
                        sheet.Cells[rowindexcheck, colindexary[searchcindex]] = printableresults[i];
                        rowindexcheck++;
                    }
                    else
                    {
                        rowindexcheck = 0;
                        break;
                    }
                    rowindexcheck = 0;
                }

            }
            else
            {
                while (sheet.Cells[1, colindexary[colindex]] != null)
                {
                    colindex++;
                }
                sheet.Cells[1, colindex] = author;
                for (int i = 0; i < printableresults.Length; i++)
                {
                    if (printableresults != null)
                    {
                        sheet.Cells[rowindexcheck, colindexary[searchcindex]] = printableresults[i];
                        rowindexcheck++;
                    }
                    else
                    {
                        rowindexcheck = 0;
                        break;
                    }
                    rowindexcheck = 0;
                }
            }
            rowindexcheck = 1;
            searchcindex = 0;

        }
    }
}
