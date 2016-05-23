using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using myKR.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding.ExcelApplication
{
    public class ExcelApplication
    {
        private static readonly log4net.ILog Log =
            log4net.LogManager.GetLogger("Form1.cs");

        private static ExcelApplication _excelApp;
        public Object LastUsedObject;
       
        public static ExcelApplication CreateExcelApplication()
        {
            return _excelApp ?? (_excelApp = new ExcelApplication());
        }

        private static Excel.Application _app;
        
        private ExcelApplication()
        {
            try
            {
                _app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
            }
            catch (COMException)
            {
                MessageBox.Show(Resources.notFoundExcelApp);
                Environment.Exit(-1);
            }

            if (_app != null) return;

            MessageBox.Show(Resources.incorrectConnectToExcel);
            Environment.Exit(-1);
        }

        public void CloseApp(bool save)
        {
            _app.Quit();
            Kill();
            _app = null;
        }

        public void SetVisibilityForApp(bool visibil)
        {
            _app.Visible = visibil;
        }

        public void CloseBook(Excel.Workbook book, bool save)
        {
            if (book == null)
                return;
            try
            {
                book.Close(save);
            }
            catch (COMException)
            {
                // logger
            }
            catch (Exception)
            {
                // logger
            }
        }

        public Excel.Workbook OpenBook(string pathToBook)
        {
            LastUsedObject = null;

            if (string.IsNullOrEmpty(pathToBook))
                return null;

            if (!File.Exists(pathToBook))
            {
                // logger
                return null;
            }
            Excel.Workbook workbook = null;
            try
            {
                return workbook = _app.Workbooks.Open(pathToBook);
            }
            catch (COMException)
            {
                // logger
                if (workbook == null)
                {
                    string bookName = pathToBook.Substring(pathToBook.LastIndexOf("\\", StringComparison.Ordinal) + 1,
                        pathToBook.LastIndexOf(".", StringComparison.Ordinal) - 3);

                    LastUsedObject = IfBookExist(bookName);

                    return (Excel.Workbook) LastUsedObject;
                }

                CloseBook(workbook, false);
                return null;
            }
            catch (Exception)
            {
                // logger
                return null;
            }
        }

        public Excel.Worksheet OpenWorksheet(Excel.Workbook book, string sheetName)
        {
            LastUsedObject = null;

            if (book == null || string.IsNullOrEmpty(sheetName))
            {
                // logger
                return null;
            }

            LastUsedObject = IfSheetExist(book, sheetName);

            return (Excel.Worksheet) LastUsedObject;
        }

        public Excel.Worksheet OpenWorksheet(Excel.Workbook book, int index)
        {
            LastUsedObject = null;

            if (book == null)
            {
                // logger
                return null;
            }
            if (book.Worksheets.Count > 0 && book.Worksheets.Count <= index)
            {
                LastUsedObject = book.Worksheets[index];

                return (Excel.Worksheet) LastUsedObject;
            }
            // logger
            return null;
        }

        public Excel.Worksheet CreateNewSheet(Excel.Workbook book, string sheetName)
        {
            LastUsedObject = null;

            if (book == null || string.IsNullOrEmpty(sheetName))
                // logger
                return null;

            if (IfSheetExist(book, sheetName) == null)
                // logger
                return null;

            Excel.Worksheet sheet = book.Worksheets.Add(Type.Missing);
            sheet.Name = sheetName;

            LastUsedObject = sheet;

            return sheet;
        }

        private Excel.Worksheet IfSheetExist(Excel.Workbook book, string sheetName)
        {
            return book.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name.Equals(sheetName));
        }

        private Excel.Workbook IfBookExist(string bookName)
        {
            return _app.Workbooks.Cast<Excel.Workbook>().FirstOrDefault(book =>
                bookName.Equals(book.Name.Substring(0, book.Name.LastIndexOf(".", StringComparison.Ordinal))));
        }

        private void Kill()
        {
            if (_app == null)
                return;
            try
            {
                int excelProcessId;
                GetWindowThreadProcessId(_app.Hwnd, out excelProcessId);
                Process p = Process.GetProcessById(excelProcessId);
                p.Kill();
                _app = null;
            }
            catch
            {
                // ignored
            }
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
    }
}
