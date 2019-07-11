using System;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace StoryConv2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            using (SpreadsheetDocument document = SpreadsheetDocument.Open("D:/dev/wod/svn/オリジナルデータ/p1_UI元素材/Unity/GUI/StoryDemo/StoryDemo_1.xlsm", false))
            {
                foreach(var sheetPart in document.WorkbookPart.WorksheetParts)
                {
                    var a = sheetPart.Worksheet;
                    var rows = a.GetFirstChild<SheetData>().Elements<Row>();
                    foreach(var r in rows)
                    {
                        foreach(var c in r.Elements<Cell>())
                        {
                            Console.WriteLine("{0}", c.InnerText);
                        }
                    }
                }
                // foreach(var sheet in document.WorkbookPart.Workbook.Sheets)
                // {
                //     var rows = sheet.Elements<Row>();
                //     foreach(var row in rows)
                //     {
                //         Console.WriteLine("a");
                //     }
                //     // foreach(var attr in sheet.GetAttributes())
                //     // {
                //     //     Console.WriteLine("{0},{1}", attr.LocalName, attr.Value);
                //     // }
                // }
            }
        }
    }
}
