using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding.ApplicatinsExcel
{
    interface App
    {
        Excel.Application GetExcelApp();
        void CloseExcelApp();
        Excel.Workbook GetBookFromApp(string bookName);
        void CloseBookFromApp(bool save);
    }
}
