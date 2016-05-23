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
            Log.Info("Create Excel Application");
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
            catch (COMException e)
            {
                Log.Error(Resources.notFoundExcelApp, e);
                MessageBox.Show(Resources.notFoundExcelApp);
                Environment.Exit(-1);
            }

            if (_app != null) return;

            Log.Error(Resources.incorrectConnectToExcel);
            MessageBox.Show(Resources.incorrectConnectToExcel);
            Environment.Exit(-1);
        }

        public void CloseApp(bool save)
        {
            Log.Info(LoggetConstats.ENTER);
            _app.Quit();
            Kill();
            _app = null;
            Log.Info(LoggetConstats.EXIT);
        }

        public void SetVisibilityForApp(bool visibil)
        {
            Log.Info(LoggetConstats.ENTER_EXIT);
            _app.Visible = visibil;
        }

        public void CloseBook(Excel.Workbook book, bool save)
        {
            Log.Info(LoggetConstats.ENTER);
            if (book == null)
            {
                Log.Warn(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return;
            }
            try
            {
                Log.Info("Book name is `" + book.Name + "`");
                book.Close(save);
            }
            catch (COMException e)
            {
                Log.Warn("COMException with closing `" + book.Name + "`", e);
            }
            catch (Exception e)
            {
                Log.Warn("Exception with closing `" + book.Name + "`", e);
            }
            Log.Info(LoggetConstats.EXIT);
        }

        public Excel.Workbook OpenBook(string pathToBook)
        {
            Log.Info(LoggetConstats.ENTER);
            LastUsedObject = null;

            if (string.IsNullOrEmpty(pathToBook))
            {
                Log.Info(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }

            if (!File.Exists(pathToBook))
            {
                Log.Warn(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }
            Excel.Workbook workbook = null;
            try
            {
                Log.Info(LoggetConstats.EXIT);
                return workbook = _app.Workbooks.Open(pathToBook);
            }
            catch (COMException e)
            {
                Log.Warn("COMException while opening BOOK", e);
                if (workbook == null)
                {
                    Log.Info("Try find opened book with same name");

                    string bookName = pathToBook.Substring(pathToBook.LastIndexOf("\\", StringComparison.Ordinal) + 1,
                        pathToBook.LastIndexOf(".", StringComparison.Ordinal) - 3);

                    LastUsedObject = IfBookExist(bookName);

                    Log.Info(LoggetConstats.EXIT);
                    return (Excel.Workbook) LastUsedObject;
                }
                CloseBook(workbook, false);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }
            catch (Exception e)
            {
                Log.Warn("Exception while opening BOOK", e);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }
        }

        public Excel.Worksheet OpenWorksheet(Excel.Workbook book, string sheetName)
        {
            Log.Info(LoggetConstats.ENTER);
            LastUsedObject = null;

            if (book == null || string.IsNullOrEmpty(sheetName))
            {
                Log.Warn(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }

            LastUsedObject = IfSheetExist(book, sheetName);

            Log.Info(LoggetConstats.EXIT);
            return (Excel.Worksheet) LastUsedObject;
        }

        public Excel.Worksheet OpenWorksheet(Excel.Workbook book, int index)
        {
            Log.Info(LoggetConstats.ENTER);
            LastUsedObject = null;

            if (book == null)
            {
                Log.Warn(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }
            if (book.Worksheets.Count <= 0 || book.Worksheets.Count > index)
            {
                Log.Warn(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }

            LastUsedObject = book.Worksheets[index];

            Log.Info(LoggetConstats.EXIT);
            return (Excel.Worksheet) LastUsedObject;
        }

        public Excel.Worksheet CreateNewSheet(Excel.Workbook book, string sheetName)
        {
            Log.Info(LoggetConstats.ENTER);
            LastUsedObject = null;

            if (book == null || string.IsNullOrEmpty(sheetName))
            {
                Log.Warn(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }

            if (IfSheetExist(book, sheetName) == null)
            {
                Log.Warn(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return null;
            }

            Excel.Worksheet sheet = book.Worksheets.Add(Type.Missing);
            sheet.Name = sheetName;

            LastUsedObject = sheet;

            Log.Info(LoggetConstats.EXIT);
            return sheet;
        }

        private Excel.Worksheet IfSheetExist(Excel.Workbook book, string sheetName)
        {
            Log.Info(LoggetConstats.ENTER_EXIT);
            return book.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name.Equals(sheetName));
        }

        private Excel.Workbook IfBookExist(string bookName)
        {
            Log.Info(LoggetConstats.ENTER_EXIT);
            return _app.Workbooks.Cast<Excel.Workbook>().FirstOrDefault(book =>
                bookName.Equals(book.Name.Substring(0, book.Name.LastIndexOf(".", StringComparison.Ordinal))));
        }

        private void Kill()
        {
            Log.Info(LoggetConstats.ENTER);
            if (_app == null)
            {
                Log.Warn(LoggetConstats.BAD_VALIDATION);
                Log.Info(LoggetConstats.EXIT);
                return;
            }
            try
            {
                int excelProcessId;
                GetWindowThreadProcessId(_app.Hwnd, out excelProcessId);
                Process p = Process.GetProcessById(excelProcessId);
                p.Kill();
                _app = null;
            }
            catch(Exception e)
            {
                Log.Warn("Exception while killed process", e);
                Log.Info(LoggetConstats.EXIT);
            }
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
    }
}
