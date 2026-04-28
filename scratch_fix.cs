using System;
using System.IO;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        string path = @"c:\repo\ExcelCSIToolBoxAddIn\UI\Views\EtabsToolboxWindow.xaml";
        string content = File.ReadAllText(path);
        string pattern = @"Content=""(Run|By UniqueName|By Coor|By Point|Get|Set)""\s+Foreground=""Blue""";
        string replacement = "Content=\"$1\"\r\n                                    HorizontalAlignment=\"Stretch\"\r\n                                    Foreground=\"Blue\"";
        content = Regex.Replace(content, pattern, replacement);
        File.WriteAllText(path, content);
    }
}

