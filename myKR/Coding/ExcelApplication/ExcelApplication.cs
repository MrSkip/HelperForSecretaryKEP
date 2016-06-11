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

        private static ExcelApplication ExcelApp;
        public Object LastUsedObject;
       
        public static ExcelApplication CreateExcelApplication()
        {
            Log.Info("Create Excel Application");
            return ExcelApp ?? (ExcelApp = new ExcelApplication());
        }

        public static Excel.Application App;
        
        private ExcelApplication()
        {
            try
            {
                App = new Excel.Application
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

            if (App != null) return;

            Log.Error(Resources.incorrectConnectToExcel);
            MessageBox.Show(Resources.incorrectConnectToExcel);
            Environment.Exit(-1);
        }

        public void CloseApp()
        {
            Log.Info(LoggerConstants.ENTER);
            App.Quit();
            Kill(App);
            App = null;
            Log.Info(LoggerConstants.EXIT);
        }

        public void SetVisibilityForApp(bool visibil)
        {
            Log.Info(LoggerConstants.ENTER_EXIT);
            App.Visible = visibil;
        }

        public void CloseBook(Excel.Workbook book, bool save)
        {
            Log.Info(LoggerConstants.ENTER);
            if (book == null)
            {
                Log.Warn(LoggerConstants.BAD_VALIDATION);
                Log.Info(LoggerConstants.EXIT);
                return;
            }
            try
            {
                Log.Info("Book name is `" + book.Name + "`");
                if (save)
                    book.Save();
                book.Close();
            }
            catch (COMException e)
            {
                Log.Warn("COMException with closing `" + book.Name + "`", e);
            }
            catch (Exception e)
            {
                Log.Warn("Exception with closing `" + book.Name + "`", e);
            }
            Log.Info(LoggerConstants.EXIT);
        }

        public Excel.Workbook OpenBook(string pathToBook)
        {
            Log.Info(LoggerConstants.ENTER);
            LastUsedObject = null;

            if (string.IsNullOrEmpty(pathToBook) || !File.Exists(pathToBook))
            {
                Log.Info(LoggerConstants.BAD_VALIDATION + ": " + pathToBook);
                Log.Info(LoggerConstants.EXIT);
                return null;
            }

            Excel.Workbook workbook = null;
            try
            {
                Log.Info(LoggerConstants.EXIT);
                return workbook = App.Workbooks.Open(pathToBook);
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

                    Log.Info(LoggerConstants.EXIT);
                    return (Excel.Workbook) LastUsedObject;
                }
                CloseBook(workbook, false);
                Log.Info(LoggerConstants.EXIT);
                return null;
            }
            catch (Exception e)
            {
                Log.Warn("Exception while opening BOOK", e);
                Log.Info(LoggerConstants.EXIT);
                return null;
            }
        }

        public Excel.Worksheet OpenWorksheet(Excel.Workbook book, string sheetName)
        {
            Log.Info(LoggerConstants.ENTER);
            LastUsedObject = null;

            if (book == null || string.IsNullOrEmpty(sheetName))
            {
                Log.Warn(LoggerConstants.BAD_VALIDATION);
                Log.Info(LoggerConstants.EXIT);
                return null;
            }

            LastUsedObject = IfSheetExist(book, sheetName);

            Log.Info(LoggerConstants.EXIT);
            return (Excel.Worksheet) LastUsedObject;
        }

        public Excel.Worksheet OpenWorksheet(Excel.Workbook book, int index)
        {
            Log.Info(LoggerConstants.ENTER);
            LastUsedObject = null;

            if (book == null)
            {
                Log.Warn(LoggerConstants.BAD_VALIDATION);
                Log.Info(LoggerConstants.EXIT);
                return null;
            }
            if (book.Worksheets.Count <= 0 || book.Worksheets.Count > index)
            {
                Log.Warn(LoggerConstants.BAD_VALIDATION);
                Log.Info(LoggerConstants.EXIT);
                return null;
            }

            LastUsedObject = book.Worksheets[index];

            Log.Info(LoggerConstants.EXIT);
            return (Excel.Worksheet) LastUsedObject;
        }

        public Excel.Worksheet CreateNewSheet(Excel.Workbook book, string sheetName)
        {
            Log.Info(LoggerConstants.ENTER);
            LastUsedObject = null;

            if (book == null || string.IsNullOrEmpty(sheetName))
            {
                Log.Warn(LoggerConstants.BAD_VALIDATION);
                Log.Info(LoggerConstants.EXIT);
                return null;
            }

            if (IfSheetExist(book, sheetName) != null)
            {
                Log.Warn(LoggerConstants.BAD_VALIDATION + "; bookName = " + book.Name + "; SheetName = " + sheetName);
                Log.Info(LoggerConstants.EXIT);
                return null;
            }

            Excel.Worksheet sheet = book.Worksheets.Add(Type.Missing);
            sheet.Name = sheetName;

            LastUsedObject = sheet;

            Log.Info(LoggerConstants.EXIT);
            return sheet;
        }

        private Excel.Worksheet IfSheetExist(Excel.Workbook book, string sheetName)
        {
            Log.Info(LoggerConstants.ENTER_EXIT);
            return book.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(sheet => sheet.Name.Equals(sheetName));
        }

        private Excel.Workbook IfBookExist(string bookName)
        {
            Log.Info(LoggerConstants.ENTER_EXIT);
            return App.Workbooks.Cast<Excel.Workbook>().FirstOrDefault(book =>
                bookName.Equals(book.Name.Substring(0, book.Name.LastIndexOf(".", StringComparison.Ordinal))));
        }

        public static void Kill(Excel.Application app)
        {
            Log.Info(LoggerConstants.ENTER);
            if (app == null)
            {
                Log.Warn(LoggerConstants.BAD_VALIDATION);
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            try
            {
                int excelProcessId;
                GetWindowThreadProcessId(app.Hwnd, out excelProcessId);
                Process p = Process.GetProcessById(excelProcessId);
                p.Kill();
            }
            catch(Exception e)
            {
                Log.Warn("Exception while killed process", e);
                Log.Info(LoggerConstants.EXIT);
            }
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
    }
}
