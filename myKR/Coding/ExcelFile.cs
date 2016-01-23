using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding
{
    public static class ExcelFile
    {
        public static Excel.Application App;
        public static void ReadRobPlan(string pathToRobPlan)
        {
            App = new Excel.Application();
            Excel.Workbook book = App.Workbooks.Open(pathToRobPlan);

            foreach (Excel.Worksheet sheet in book.Worksheets)
            {
                if (sheet.Name.Trim().Length == 8 && sheet.Name.Trim().IndexOf('-') == 2 &&
                    sheet.Name.Trim().LastIndexOf('-') == 5)
                {
                    Manager.Groups.Add(ReadSheetFromRobPlan(sheet));
                }
            }
            book.Close();
            App.Quit();
        }

        public static void ReadStudentsAndOlicAndCurators(string pathToDb)
        {
            App = new Excel.Application();
            Excel.Workbook book = App.Workbooks.Open(pathToDb);

            // Read [База студентів]
//            try
//            {
//                List<Student> students = ReadStudents((Excel.Worksheet)book.Worksheets.Item["База студентів"]);
//                foreach (Group group in Manager.Groups)
//                {
//                    group.Students = students.FindAll(student => student.Group.Equals(group.Name));
//                }
//            }
//            catch (Exception e)
//            {
//                MassageError("База студентів", "", "Щось не гаразд із зчитуванням студентів\nМожливо, лист [База студентів] відсунтій - створіть його\n" + e);
//            }

            // Read [Реєстраційна відомість (журнал)]
            try
            {
                List<NumberOfOblic> oblics = ReadNumbersOfOblic((Excel.Worksheet)book.Worksheets.Item["Реєстраційна відомість (журнал)"]);
                foreach (Group group in Manager.Groups)
                {
                    foreach (NumberOfOblic numberOfOblic in oblics.FindAll(oblic => CustomEquals(oblic.Group, @group.Name)))
                    {
                        Subject find = @group.Subjects.Find(subject => CustomEquals(subject.Name, numberOfOblic.Subject));
                        if (find != null)
                            find.NumberOfOlic = numberOfOblic.Number;

                        Practice practice =
                            group.Practice.Find(practice1 => CustomEquals(practice1.Name, numberOfOblic.Subject));
                        if (practice != null)
                            practice.NumberOfOlic = numberOfOblic.Number;
                    }
                }
            }
            catch (Exception e)
            {
                MassageError("Реєстраційна відомість (журнал)", "", "Щось не гаразд із зчитуванням реєстраційної відомості\nМожливо, лист [Реєстраційна відомість (журнал)] відсунтій - створіть його\n" + e);
            }

            // Read [Куратори]
            try
            {
                List<string[]> list =
                    ReadCurator((Excel.Worksheet) book.Worksheets.Item["Куратори"]);
                foreach (Group group in Manager.Groups)
                {
                    string[] s = list.Find(strings => CustomEquals(strings[1], group.Name));
                    if (s != null)
                        group.Curator = s[0];
                }
            }
            catch (Exception e)
            {
                MassageError("[Куратори]", "", "Щось не гараз із зчитування кураторів груп\n" + e);
            }

            book.Close();
            App.Quit();
        }

        public static List<string[]> ReadCurator(Excel.Worksheet sheet)
        {
            List<string[]> curators = new List<string[]>();
            try
            {
                int n = 1;
                while (true)
                {
                    n++;
                    var value = sheet.Cells[n, "A"].Value;
                    if (value == null || string.IsNullOrEmpty(value.ToString()))
                        break;
                    string curatorName = value.ToString();

                    value = sheet.Cells[n, "B"].Value;
                    if (value == null || string.IsNullOrEmpty(value.ToString()))
                        continue;
                    string groupName = value.ToString();

                    curators.Add(new[] {curatorName.Trim(), groupName.Trim()});
                }
            }
            catch (Exception e)
            {
                MassageError(sheet.Name, "", "Щось не гараз із зчитування кураторів груп\n" + e);
            }
            return curators;
        }

        private static bool CustomEquals(string first, string second)
        {
            first = first.ToLower().Trim().Replace("*", "");
            second = second.ToLower().Trim().Replace("*", "");
            return first.Equals(second);
        }

        public static List<NumberOfOblic> ReadNumbersOfOblic(Excel.Worksheet sheet)
        {
            List<NumberOfOblic> oblics = new List<NumberOfOblic>();

            int n = 2;
            try
            {
                while (true)
                {
                    n++;
                    var value = sheet.Cells[n, "A"].Value;
                    if (value == null || string.IsNullOrEmpty(value.ToString()))
                        break;
                    string number = value.ToString();

                    value = sheet.Cells[n, "B"].Value;
                    if (value == null || string.IsNullOrEmpty(value))
                        continue;
                    string sujectName = value.ToString();

                    value = sheet.Cells[n, "D"].Value;
                    if (value == null || string.IsNullOrEmpty(value))
                        continue;
                    string groupName = value.ToString();

                    oblics.Add(new NumberOfOblic
                    {
                        Number = number,
                        Subject = sujectName,
                        Group = groupName
                    });
                }
            }
            catch (Exception e)
            {
                MassageError(sheet.Name, "", "Щось не гараз із зчитуванням реєстрації відомостей (журнал)\n" + e);
            }
            return oblics;
        }

        public static List<Student> ReadStudents(Excel.Worksheet sheet)
        {
            List<Student> students = new List<Student>();

            int n = 1;

            while (true)
            {
                try
                {
                    n++;
                    var value = sheet.Cells[n, "C"].Value;
                    if (value == null || string.IsNullOrEmpty(value.ToString()))
                        if (n >= 5) break;
                        else continue;

                    Student student = new Student {Pib = value.ToString().Trim()};

                    value = sheet.Cells[n, "D"].Value;
                    if (value != null)
                        student.NumberOfBook = value.ToString();

                    value = sheet.Cells[n, "E"].Value;
                    if (value != null)
                        student.Group = value.ToString();

                    value = sheet.Cells[n, "G"].Value;
                    if (value != null)
                        student.FormaTeaching = value.ToString();

                    value = sheet.Cells[n, "L"].Value;
                    if (value != null)
                        student.Benefits = value.ToString();

                    students.Add(student);
                }
                catch (Exception e)
                {
                    MassageError(sheet.Name, "", "Щось не гаразд із зчитуванням студентів\n" + e);
                }
            }
            return students;
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

            group.Subjects = ReadSubject(sheet);
            group.Practice = ReadPractice(sheet);
            group.StateExamination = ReadStateExamination(sheet);

            return group;
        }

        private static List<Subject> ReadSubject(Excel.Worksheet sheet)
        {
            List<Subject> subjects = new List<Subject>();
            //[0] - Hours; [1] - Cursova; [2] - Ispyt (Examen) [3] - DyfZalikOrZalic; [4] - DyfZalik (if exist)
            string[]
                firstSemestr = { "Y", "AK", "AO", "AQ", "AR" },
                secondSemestr = { "AS", "BE", "BI", "BK", "BL" };
            bool existingDyfZalik = false;
            bool ifDufZalic = true;

            // check the cells
            var s = sheet.Cells[15, "C"].Value;
            if (s == null || string.IsNullOrWhiteSpace(s.ToString()) || !s.Trim().ToLower().Equals("назви навчальних  дисциплін"))
            {
                MassageError(sheet.Name, "C15", "Назви навчальних  дисциплін");
                return new List<Subject>();
            }

            // check the cells
            s = sheet.Cells[18, "AR"].Value;
            if (!string.IsNullOrEmpty(s) && s.Trim().ToLower().Equals("диф  залік"))
                ifDufZalic = false;

            // check the cells
            s = sheet.Cells[18, "AQ"].Value;
            if (!string.IsNullOrEmpty(s) && s.Trim().ToLower().Equals("диф  залік"))
                existingDyfZalik = true;

            int n = 14;

            while (true)
            {
                n++;
                
                var subjectName = sheet.Cells[n, "C"].Value;
                if (subjectName == null || string.IsNullOrEmpty(subjectName.ToString()) || n == 15)
                    continue;
                if (subjectName.ToString().Trim().ToLower().Equals("разом"))
                    break;

                string teacher = sheet.Cells[n, "BN"].Value;
                if (string.IsNullOrEmpty(teacher))
                    teacher = "";

                
                try
                {
                    Subject subject = null;
                    for (int i = 0; i < 2; i++)
                    {
                        string[] list = i == 0 ? firstSemestr :  secondSemestr;
                        var ss = sheet.Cells[n, list[0]].Value;

                        // if cursova robota have same of the pas the not continue
                        bool bl = true;
                        var kp = sheet.Cells[n, list[1]].Value;
                        if (kp != null && !string.IsNullOrEmpty(kp.ToString()))
                            bl = false;

                        if ((ss == null || string.IsNullOrEmpty(ss.ToString())) && bl)
                            continue;
                        subject = new Subject
                        {
                            Name = subjectName.ToString().Trim(),
                            Teacher = teacher
                        };

                        Semestr semestr = new Semestr();

                        if (ss == null || string.IsNullOrEmpty(ss.ToString()))
                            semestr.CountOfHours = 0;
                        else semestr.CountOfHours = ss;

                        ss = sheet.Cells[n, list[1]].Value;
                        if (ss != null && (!string.IsNullOrEmpty(ss.ToString()) || !ss.ToString().Equals("0")))
                            semestr.CursovaRobota = ss;

                        ss = sheet.Cells[n, list[2]].Value;
                        if (ss != null && (!string.IsNullOrEmpty(ss.ToString()) || !ss.ToString().Equals("0")))
                            semestr.Isput = ss;

                        ss = sheet.Cells[n, list[3]].Value;
                        if (ss != null && (!string.IsNullOrEmpty(ss.ToString()) || !ss.ToString().Equals("0")))
                        {
                            if (ifDufZalic) semestr.DyfZalik = ss;
                            else semestr.Zalic = ss;
                        }

                        if (!existingDyfZalik)
                        {
                            ss = sheet.Cells[n, list[4]].Value;
                            if (ss != null && (!string.IsNullOrEmpty(ss.ToString()) || !ss.ToString().Equals("0")))
                                semestr.DyfZalik = ss;
                        }

                        if (i == 0) subject.FirstSemestr = semestr;
                        else subject.SecondSemestr = semestr;
                    }

                    if (subject != null) subjects.Add(subject);
                }
                catch (Exception exception)
                {
                    MassageError(sheet.Name, "", "Щось не гараз із записами про предмет\n" + exception);
                }
            }
            return subjects;
        }

        private static List<Practice> ReadPractice(Excel.Worksheet sheet)
        {
            List<Practice> practices = new List<Practice>();
            string[][] position =
            {
                // Rows - PositionNameOfPractice - Semest - CountOfHours - FormaControlling - Teacher1 - Teacher2
                new[] {"40", "C", "A", "AA", "AE", "AJ", "AS"},
                new[] {"41", "C", "A", "AA", "AE", "AJ", "AS"},
                new[] {"43", "C", "A", "AA", "AE", "AJ", "AS"},
                new[] {"47", "C", "A", "AA", "AE", "AJ", "AS"},
                new[] {"49", "C", "A", "AA", "AE", "AJ", "AS"}
            };
            foreach (string[] strings in position)
            {
                var value = sheet.Cells[int.Parse(strings[0]), strings[1]].Value;
                if (value != null && value.ToString().Trim().ToLower().Equals("назва практики"))
                {
                    int n = int.Parse(strings[0]);
                    while (true)
                    {
                        n++;
                        value = sheet.Cells[n, strings[1]].Value;

                        if (value == null || string.IsNullOrEmpty(value.ToString()))
                            break;

                        if (!value.ToString().Trim().ToLower().Equals("навчальна") &&
                            !value.ToString().Trim().ToLower().Equals("виробнича"))
                        {
                            try
                            {
                                Practice practice = new Practice();
                                List<string> list = new List<string>();

                                practice.Name = value.ToString().Trim();

                                value = sheet.Cells[n, strings[2]].Value;
                                practice.Semestr = value == null ? "" : value.ToString();

                                value = sheet.Cells[n, strings[3]].Value;
                                if (value == null || string.IsNullOrEmpty(value.ToString()))
                                    practice.CountOfHours = 0;
                                else practice.CountOfHours = double.Parse(value.ToString());

                                value = sheet.Cells[n, strings[4]].Value;
                                practice.FormOfControll = value == null ? "" : value.ToString();

                                value = sheet.Cells[n, strings[5]].Value;
                                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                    list.Add(value.ToString());

                                value = sheet.Cells[n, strings[6]].Value;
                                if (value != null && !string.IsNullOrEmpty(value.ToString()))
                                    list.Add(value.ToString());

                                practice.Teacher = list;
                                practices.Add(practice);
                            }
                            catch (Exception e)
                            {
                                MassageError(sheet.Name, "", "Щось не гаразд із зчитуванням практики\n" + e);
                            }
                        }
                    }
                    break;
                }
            }
            return practices;
        }

        private static List<StateExamination> ReadStateExamination(Excel.Worksheet sheet)
        {
            List<StateExamination> examinations = new List<StateExamination>();
            string[][] position =
            {
                new []{"49", "BE", "BO"},
                new []{"40", "BE", "BO"},
                new []{"40", "AX", "BO"},
                new []{"47", "BE", "BO"},
                new []{"41", "BE", "BO"},
                new []{"43", "BE", "BO"},
                new []{"38", "AX", "BO"}
            };

            foreach (string[] strings in position)
            {
                try
                {
                    var value = sheet.Cells[int.Parse(strings[0]), strings[1]].Value;
                    if (value != null && value.ToString().Trim().ToLower().Equals("назва"))
                    {
                        int n = Int32.Parse(strings[0]);
                        while (true)
                        {
                            n++;
                            value = sheet.Cells[n, strings[2]].Value;
                            if (value == null || string.IsNullOrEmpty(value.ToString()))
                                break;
                            StateExamination examination = new StateExamination();
                            var nameOfExamen = sheet.Cells[n, strings[1]].Value;
                            if (nameOfExamen != null && !string.IsNullOrEmpty(nameOfExamen.ToString()))
                            {
                                examination.Name = nameOfExamen.ToString();
                                examination.Semestr = value.ToString();
                                examinations.Add(examination);
                            }
                        }
                        break;
                    }
                }
                catch (Exception e)
                {
                    MassageError(sheet.Name, "", "Щось не гаразд із зчитуванням державних екзаменів\n" + e);
                }
            }
            return examinations;
        }

        private static void MassageError(string sheetName, string cell, string format)
        {
            MessageBox.Show("Помилка у робочому листі [" + sheetName + "]!\nСлід дотримуватися цього формату для клітини [" + cell + "]:\n[" + format + "]");
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