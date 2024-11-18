using ExTest;

internal class Program
{
    private static void Main(string[] args)
    {
        string path = "";
        using ExcelHelper ehelper = new ExcelHelper();
        ehelper.Open(path);
        ehelper.Save();
    }
}