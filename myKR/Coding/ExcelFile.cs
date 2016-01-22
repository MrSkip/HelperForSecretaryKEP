using System;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding
{
    public static class ExcelFile
    {
        public static void ReadRobPlan(string pathToRobPlan)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Open(pathToRobPlan);

            foreach (Excel.Worksheet sheet in book.Worksheets)
            {
                if (sheet.Name.Trim().Length == 8 && sheet.Name.Trim().IndexOf('-') == 2 &&
                    sheet.Name.Trim().LastIndexOf('-') == 5)
                {
                    Manager.Groups.Add(ReadSheetFromRobPlan(sheet));
                }
            }

            app.Quit();
        }

        private static Group ReadSheetFromRobPlan(Excel.Worksheet sheet)
        {
            Group group = new Group();

            group.Name = sheet.Name;

            //Read "Напряму підготовки"
            string s = sheet.Cells[6, "R"].Value;

            if (string.IsNullOrEmpty(s) || s.Count(c => c.Equals('"')) != 2)
            {
                s = "ВВЕДІТЬ НАПРЯМ ПІДГОТОВКИ";
                MassageError(sheet.Name, "R6", "Напряму підготовки 6.050103   \"Програмна інженерія\"");
            }
            else
            {
                try
                {
                    s = s.Substring(s.IndexOf("\"", StringComparison.Ordinal) + 1,
                        s.LastIndexOf("\"", StringComparison.Ordinal) - s.IndexOf("\"", StringComparison.Ordinal) - 1);
                }
                catch (Exception)
                {
                    MassageError(sheet.Name, "R6", "Напряму підготовки 6.050103   \"Програмна інженерія\""); 
                }
            }

            group.TrainingDirection = s;

            //Read "Спеціальність"
            s = sheet.Cells[7, "R"].Value;

            if (string.IsNullOrEmpty(s) || s.Count(c => c.Equals('"')) != 2)
            {
                s = "ВВЕДІТЬ НАЗВУ СПЕЦІАЛЬНОСТІ";
                MassageError(sheet.Name, "R7", "Спеціальності  5.05010301 \"Розробка програмного забезпечення\"");
            }
            else
                try
                {
                    group.Speciality = s.Substring(s.IndexOf("\"", StringComparison.Ordinal) + 1,
                        s.LastIndexOf("\"", StringComparison.Ordinal) - s.IndexOf("\"", StringComparison.Ordinal) - 1);
                }
                catch (Exception)
                {
                    MassageError(sheet.Name, "R7", "Спеціальності  5.05010301 \"Розробка програмного забезпечення\"");
                }

            //Код спеціальності
            if (s.Equals("ВВЕДІТЬ НАЗВУ СПЕЦІАЛЬНОСТІ"))
            {
                MassageError(sheet.Name, "R7", "Спеціальності  5.05010301 \"Розробка програмного забезпечення\"");
                s = "КОД";
            }
            else
            {
                try
                {
                    s = s.Trim().Substring(0, s.Trim().IndexOf(" \"", StringComparison.Ordinal));
                    s = s.Substring(s.IndexOf(" ", StringComparison.Ordinal)).Trim();
                }
                catch (Exception)
                {
                    MassageError(sheet.Name, "R7", "Спеціальності  5.05010301 \"Розробка програмного забезпечення\"");
                }
            }

            group.CodeOfSpeciality = s;

            //read "Курс"
            s = sheet.Cells[9, "R"].Value;

            if (string.IsNullOrEmpty(s) || !s.Trim().StartsWith("Курс"))
            {
                s = "ВВЕДІТЬ КУРС";
                MassageError(sheet.Name, "R9", "Курс __II____          Група __ПІ-_14-01___");
            }
            else
            {
                try
                {
                    s = s.Trim().Substring(GetPositionForCellCource(s.Trim()));
                    s = s.Remove(s.IndexOf("_", StringComparison.Ordinal));
                }
                catch (Exception)
                {
                    MassageError(sheet.Name, "R9", "Курс __II____          Група __ПІ-_14-01___");
                }
            }

            group.Course = s;

            //Read "Рік"
            s = sheet.Cells[6, "B"].Value;

            if (string.IsNullOrEmpty(s))
            {
                s = "ВВЕДІТЬ РIK";
                MassageError(sheet.Name, "B6", "\"  28  \"       серпня            2015 року");
            }
            else
            {
                try
                {
                    s = s.Substring(s.Length - 9, 4);
                }
                catch (Exception)
                {
                    MassageError(sheet.Name, "B6", "\"  28  \"       серпня            2015 року");
                    throw;
                }
            }

            group.Year = s;

            ReadSubject(group.Subject, sheet);

            return group;
        }

        private static void ReadSubject(Subject subject, Excel.Worksheet sheet)
        {
            //[0] - Hours; [1] - Cursova; [2] - Ispyt (Examen) [3] - DyfZalikOrZalic; [4] - DyfZalik (if exist)
            string[]
                firstSemestr = {"Y", "AK", "AO", "AQ", "AR"},
                secondSemestr = {"AS", "BE", "BI", "BK", "BL" };
            int n = 14;
            while (true)
            {
                n++;
                string value = sheet.Cells[n, "C"].Value;
                if (string.IsNullOrEmpty(value))
                    continue;
                if (value.Trim().ToLower().Equals("разом"))
                    break;
                value = value.Trim();

            }
            
        }

        private static void MassageError(string sheetName, string cell, string format)
        {
            MessageBox.Show("Помилка у робочому листі [" + sheetName + "]!\nСлід дотримуватися цього формату для клітини [" + cell + "]:\n" + format);
        }

        private static int GetPositionForCellCource(string str)
        {
            int x = -1;
            foreach (var c in str)
            {
                x++;
                if (c.Equals('I') || c.Equals('V') || c.Equals('І'))
                    break;
            }
            return x;
        }
    }
}
