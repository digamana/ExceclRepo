using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace ReadExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //初始化時，設定要讀取的Excel路徑跟工作表名稱
            Excel excel = new Excel(@"File Path", "SheetName");
            var result = excel.sheet["A3"].StringValue;
            Console.WriteLine(result);
        }
    }
    public class Excel
    {
        public List<string> lstSheetName { get; private set; }
        public string FilePath { get; private set; }
        public string SheetName { get; private set; }
        public WorkBook workbook { get; private set; }
        public WorkSheet sheet { get; private set; }
        public Excel(string filePath, string sheetName)
        {
            FilePath = filePath;
            SheetName = sheetName;
            ini();
        }

        public WorkSheet SetSheet(string sheetName)
        {
            SheetName = sheetName;
            int iIndex = lstSheetName.IndexOf(sheetName);
            sheet = workbook.WorkSheets[iIndex];
            return sheet;
        }
        public void ini()
        {
            workbook = WorkBook.Load(FilePath);
            lstSheetName = workbook.WorkSheets.Select(c => c.Name).ToList();
            int iIndex = lstSheetName.IndexOf(this.SheetName);
            sheet = workbook.WorkSheets[iIndex];
        }
    }
}