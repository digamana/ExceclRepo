using Spire.Xls;
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

            //設定要讀取的儲存格
            var Cell_Value = excel.sheet[3, 3].Value;

            //使用SetSheet("SheetName")可以變更讀取的工作表
            //SetSheet("SheetName")
            Console.WriteLine(Cell_Value);
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
            int iIndex = lstSheetName.IndexOf(sheetName);
            sheet = workbook.Worksheets[iIndex];
            return sheet;
        }
        public void ini()
        {
            workbook = new Workbook();
            workbook.LoadFromFile(FilePath);
            lstSheetName = workbook.Worksheets.Select(c => c.Name).ToList();
            int iIndex = lstSheetName.IndexOf(this.SheetName);
            sheet = workbook.Worksheets[iIndex];
        }
    }
}