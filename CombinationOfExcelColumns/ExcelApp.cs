using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace CombinationOfExcelColumns
{
    public class ExcelApp
    {
        public ExcelApp(string excelFileName, int excelSheetNumber)
        {
            CreateExcelApp(excelFileName, excelSheetNumber);
        }
        public int ColumnCount { get; private set; }
        public int RowCount { get; private set; }
        public Excel.Sheets Sheets { get; private set; }
        public Excel.Worksheet WorkSheet { get; private set; }
        public Excel.Application Application { get; private set; }
        public Excel.Workbook Workbook { get; private set; }
        public Excel.Range Range { get; private set; }

        public Excel.Worksheet CreateNewSheet(string sheetName)
        {
            try
            {
                var newSheet = (Excel.Worksheet)Sheets.Add(
                        System.Reflection.Missing.Value,
                        Sheets[Sheets.Count],
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value
                        );
                newSheet.Name = sheetName;
                return newSheet;
            }
            catch
            {
                throw new Exception("Excel çalışma kitabınızdaki çalışama sayfalarınız farklı isimde olmalıdır!\r\nLütfen kombinasyon için daha önce oluşturulan sayfaları siliniz...");
            }
        }

        public void CloseExcel()
        {
            try
            {
                Workbook.Save();
                Workbook.Close();
                Application.Quit();
                Marshal.ReleaseComObject(WorkSheet);
                Marshal.ReleaseComObject(Sheets);
                Marshal.ReleaseComObject(Workbook);
                Marshal.ReleaseComObject(Application);
            }
            catch
            {

                throw new Exception("Excel dosyası kapatılırken bilinmeyen bir hata oluştu.\r\nLütfen işlemi tekrarlayınız.");
            }
        }

        public void ShowSheet(Excel.Application application, int sheetNum)
        {
            application.Visible = true;//show excel file
            var currentWorkSheet = (Excel.Worksheet)Sheets.get_Item(sheetNum);
            currentWorkSheet.Select();//select sheet
        }

        private void CreateExcelApp(string fileName, int sheetNum)
        {
            try
            {
                var result = new Microsoft.Office.Interop.Excel.Application();
                if (result != null)
                {
                    Application = result;
                    Application.DisplayAlerts = false;
                    //current workbook
                    Workbook = Application.Workbooks.Open(fileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", false, false, 0, true, false, false);
                    //
                    Sheets = Workbook.Worksheets;//all worksheets
                    //
                    WorkSheet = Sheets.get_Item(sheetNum);//current worksheet
                    //
                    Range = WorkSheet.UsedRange;
                    ColumnCount = Range.Columns.Count;
                    RowCount = Range.Rows.Count;
                    //
                }
            }
            catch
            {
                throw new Exception("Excel açılırken bir hata oluştu!\r\nÇalışma sayfa numarasını doğru belirttiğinizden emin olunuz.");
            }
        }
    }
}
