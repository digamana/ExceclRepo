using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
namespace ReadExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {


            Excel excel = new Excel(@"File Path", "Sheet Name");
            int rowCount = excel.xlRange.Rows.Count;
            int colCount = excel.xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (excel.xlRange.Cells[i, j] != null && excel.xlRange.Cells[i, j].Value2 != null)
                    {
                        var test = excel.xlRange.Cells[i, j].Value2.ToString();
                        Console.Write(excel.xlRange.Cells[i, j].Value2.ToString() + "\t");
                    }

                }
            }

        }
    }
    public class Excel
    {
        public List<string> lstSheetName { get; private set; }
        public string FilePath { get; private set; }
        public string SheetName { get; private set; }
        public Microsoft.Office.Interop.Excel.Application xlApp { private set; get; }
        public Microsoft.Office.Interop.Excel.Workbook workbook { get; private set; }
        public Microsoft.Office.Interop.Excel.Sheets sheet { get; private set; }
        public Microsoft.Office.Interop.Excel.Range xlRange { private set; get; }
        public Microsoft.Office.Interop.Excel.Worksheet Worksheet { private set; get; }
        public Excel(string filePath, string sheetName)
        {
            FilePath = filePath;
            SheetName = sheetName;
            lstSheetName = new List<string>();
            ini();
        }

        public Microsoft.Office.Interop.Excel.Sheets SetSheet(string sheetName)
        {
            SheetName = sheetName;
            int iIndex = lstSheetName.IndexOf(sheetName);
            sheet = workbook.Sheets[iIndex];
            return sheet;
        }
        public void ini()
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            //xlApp.Visible = false;
            workbook = xlApp.Workbooks.Open(FilePath);
            //xlApp.Visible = false;

            foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in workbook.Worksheets)
            {
                lstSheetName.Add(wSheet.Name);
            }
            int iIndex = lstSheetName.IndexOf(this.SheetName);
            Worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[iIndex];
            xlRange = Worksheet.UsedRange;
        }
        public void Dispose()
        {
            if (workbook != null) workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            if (xlApp != null) xlApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            if (xlRange != null) releaseObject(xlRange);
            if (sheet != null) releaseObject(sheet);
            if (workbook != null) releaseObject(workbook);
            if (xlApp != null) releaseObject(xlApp);
            //xlApp.Quit();
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
               Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}