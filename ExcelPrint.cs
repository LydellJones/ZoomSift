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
            "AX", "AY", "AZ"};//list of column identifiers
        private int rowindex = 1;//regular indexs
        private int colindex = 0;
        private int searchrindex = 1;//index for searching out cells
        private int searchcindex = 0;
        private bool samefound;//bool for if there is a match
        private string authorcheck;//checks against the actual author of the message
        private int rowindexcheck = 1;//indexing for checking rows
        
        Excel.Application shopapp = new Excel.Application();
        
        public ExcelPrint()
        {
            Excel._Worksheet sheet = (Excel.Worksheet)shopapp.ActiveSheet;
            shopapp.Visible = true;
        }
        public void Print()//checcks sfor an existing author first
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
            if (samefound == true)//if existing author exists then it should add values under the author's name or under the preexisting values
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
            else //if not then it should create the author on the first row, then under the name, provide values
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
