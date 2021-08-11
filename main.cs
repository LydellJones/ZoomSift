namespace Consoletestwork
{
    public class Program
    {
        static public void Main(string[] args)
        { 
            ExcelParse workload = new ExcelParse();//path for file
            workload.SiftStart("example.txt");//enter a sample file
        }
    }
}
