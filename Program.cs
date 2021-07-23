using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEdit
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excel = null;
            Excel.Workbooks books = null;
            Excel.Workbook book = null;
            Excel.Workbook book1 = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            Excel.Range cells = null;
            Excel.Range range = null;
            // 命名文件路径
            string file = @"C:\Users\xyy\Desktop\book.xlsx";
            try
            {
                excel = new Excel.Application();
                books = excel.Workbooks;
                try
                {
                    book = books.Open(file);
                    sheets = book.Worksheets;
                    sheet = sheets[1];
                    cells = sheet.Cells;
                    // 创建以前三行内容为文件名的文件
                    for (int i = 1; i < 4; i++)
                    {
                        range = cells[i, 1];
                        string cellContent = range.Value;
                        book1 = books.Add();
                        // 设置新创建的文件路径和文件名(文件名是命名文件的每行内容)
                        string b = @"C:\Users\xyy\Desktop\" + $"{cellContent}" + ".xlsx";
                        book1.SaveAs2(b);
                        Marshal.FinalReleaseComObject(range);
                    }
                    book.Close(true);
                    excel.Quit();
                }
                finally
                {
                    Marshal.FinalReleaseComObject(cells);
                    Marshal.FinalReleaseComObject(sheet);
                    Marshal.FinalReleaseComObject(sheets);
                    Marshal.FinalReleaseComObject(book);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(books);
                Marshal.FinalReleaseComObject(excel);
            }
        }
    }
}
