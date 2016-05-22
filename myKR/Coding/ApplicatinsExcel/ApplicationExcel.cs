using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using myKR.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding.ApplicatinsExcel
{
    public class ApplicationExcel
    {
        private static ApplicationExcel _excelApp;
       
        private static ApplicationExcel CreateAppImpl()
        {
            return _excelApp ?? new ApplicationExcel();
        }

        private static Excel.Application _app;

        private ApplicationExcel()
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
       
        public class BookExcel
        {
            public void Close(Excel.Workbook book, bool save)
            {
                try
                {
                    book.Close(save);
                }
                catch (COMException)
                {
                    // ignore
                }
                catch (Exception)
                {
                    // ignore
                }
            }

            public Excel.Workbook Open(string pathToBook)
            {
                
            }
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
