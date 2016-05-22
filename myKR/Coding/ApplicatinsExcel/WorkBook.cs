using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding.ApplicatinsExcel
{
    class WorkBook
    {
        public Excel.Workbook Book;
        public Dictionary<string, Excel.Worksheet> Sheets;
    }
}
