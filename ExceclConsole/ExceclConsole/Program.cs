using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExceclConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Excel excel = new Excel(@"File Path", "Sheet Name");
            var result = excel.sheet.Cells[3, 3].Value;
            Console.WriteLine(result);

        }
    }
    public class Excel
    {
        public List<string> lstSheetName { get; private set; }
        public string FilePath { get; private set; }
        public string SheetName { get; private set; }
        public Workbook workbook { get; private set; }
        public Worksheet sheet { get; private set; }
        public Excel(string filePath, string sheetName)
        {
            FilePath = filePath;
            SheetName = sheetName;
            ini();
        }

        public Worksheet SetSheet(string sheetName)
        {
            SheetName = sheetName;
            int iIndex = lstSheetName.IndexOf(sheetName);
            sheet = workbook.Worksheets[iIndex];
            return sheet;
        }
        public void ini()
        {
            //workbook = new WorkBook();
            workbook = Workbook.Load(FilePath);
            lstSheetName = workbook.Worksheets.Select(c => c.Name).ToList();
            int iIndex = lstSheetName.IndexOf(this.SheetName);
            sheet = workbook.Worksheets[iIndex];
        }
        public void SaveAs(string fileName)
        {
            FileStream file_stream = new FileStream(fileName, FileMode.Create);
            workbook.SaveToStream(file_stream);
            file_stream.Close();
        }
    }
}
