using System;

namespace ZoomSift
{
    public class Program
    {
        static public void Main(string[] args)
        {
            Console.WriteLine("Input File Path...:");
            string path = Console.ReadLine();
            ExcelParse workload = new ExcelParse();
            workload.SiftStart(path);
        }
    }
}
