using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding
{
    public static class ExcelFile
    {
        public static Excel.Application App = new Excel.Application();
        public static string CurrentFolder = Environment.CurrentDirectory + "\\";
        private static int CountofUse = 0;

        public static void ReadRobPlan(string pathToRobPlan)
        {
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
        }

        public static void ReadStudentsAndOlicAndCurators(string pathToDb)
        {
            Excel.Workbook book = App.Workbooks.Open(pathToDb);

            // Read [База студентів]
            try
            {
                List<Student> students = ReadStudents((Excel.Worksheet)book.Worksheets.Item["База студентів"]);
                foreach (Group group in Manager.Groups)
                {
                    group.Students = students.FindAll(student => student.Group.Equals(group.Name));
                }
            }
            catch (Exception e)
            {
                MassageError("База студентів", "", "Щось не гаразд із зчитуванням студентів\nМожливо, лист [База студентів] відсунтій - створіть його\n" + e);
            }

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
                catch (Exception e)
                {
                    MassageError(sheet.Name, "R6", "Напряму підготовки 6.050103   \"Програмна інженерія\"\n" + e); 
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
                catch (Exception e)
                {
                    MassageError(sheet.Name, "R7", "Спеціальності  5.05010301 \"Розробка програмного забезпечення\"\n" + e);
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
                catch (Exception e)
                {
                    MassageError(sheet.Name, "R7", "Спеціальності  5.05010301 \"Розробка програмного забезпечення\"\n" + e);
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
                catch (Exception e)
                {
                    MassageError(sheet.Name, "R9", "Курс __II____          Група __ПІ-_14-01___\n" + e);
                }
            }

            group.Course = s;

            //Read "Рік"
            s = sheet.Cells[6, "B"].Value;

            if (string.IsNullOrEmpty(s))
            {
                s = "Введіть рік";
                MassageError(sheet.Name, "B6", "\"  28  \"       серпня            2015 року");
            }
            else
            {
                try
                {
                    s = s.Substring(s.Length - 9, 4);
                }
                catch (Exception e)
                {
                    MassageError(sheet.Name, "B6", "\"  28  \"       серпня            2015 року\n" + e);
                    throw;
                }
            }

            group.Year = s;

            //Read "Семестр для першого півріччя"
            s = sheet.Cells[15, "Y"].Value;

            if (string.IsNullOrEmpty(s))
            {
                s = "ВВЕДІТЬ СЕМЕСТР";
                MassageError(sheet.Name, "Y15", "VII семестр        12  навчальних тижнів");
            }
            else
            {
                try
                {
                    s = s.Trim().Substring(0, s.Trim().IndexOf(' '));
                }
                catch (Exception e)
                {
                    MassageError(sheet.Name, "Y15", "VII семестр        12  навчальних тижнів\n" + e);
                    throw;
                }
            }

            group.FirstRomeSemestr = s;

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

            bool dyfZalikOrNot = false;
            bool ifDufZalicIsExist = true;

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
                ifDufZalicIsExist = false;

            // check the cells
            s = sheet.Cells[18, "AQ"].Value;
            if (!string.IsNullOrEmpty(s) && s.Trim().ToLower().Equals("диф  залік"))
                dyfZalikOrNot = true;

            int n = 14;

            while (true)
            {
                try
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

                    Subject subject = new Subject
                    {
                        Name = subjectName.ToString().Trim(),
                        Teacher = teacher
                    };
                    bool addToList = false;

                    for (int i = 0; i < 2; i++)
                    {
                        string[] list = i == 0 ? firstSemestr : secondSemestr;
                        var ss = sheet.Cells[n, list[0]].Value;

                        // if cursova robota have same of the pas the not continue
                        bool bl = true;
                        var kp = sheet.Cells[n, list[1]].Value;
                        if (kp != null && !string.IsNullOrEmpty(kp.ToString()))
                            bl = false;

                        if ((ss == null || string.IsNullOrEmpty(ss.ToString())) && bl)
                            continue;

                        addToList = true;
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
                            if (ifDufZalicIsExist) semestr.DyfZalik = ss;
                            else semestr.Zalic = ss;
                        }

                        if (!dyfZalikOrNot)
                        {
                            ss = sheet.Cells[n, list[4]].Value;
                            if (ss != null && (!string.IsNullOrEmpty(ss.ToString()) || !ss.ToString().Equals("0")))
                                semestr.DyfZalik = ss;
                        }

                        if (i == 0) subject.FirstSemestr = semestr;
                        else subject.SecondSemestr = semestr;
                    }
                    if (addToList) subjects.Add(subject);
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
                                practice.FormOfControl = value == null ? "" : value.ToString();

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

        /*
         *      Create Oblic Uspishosti
         *      if `groupName` is null or empty and `subjectName` is null or empty than create for all groups
         *      else if `groupName` is not null and not empty and `subjectName` is null or empty than create for one group
         *      else if `groupName` is not null and not empty and `subjectName` is not null and not empty than create for one subject
        */
        public static void CreateOblicUspishnosti(string groupName, string subjectName, int pivricha)
        {
            if (!File.Exists(CurrentFolder + "Data\\DataToProgram.xls"))
            {
                MessageBox.Show("В папці [" + CurrentFolder + "Data] повинен знайходитися файл [DataToProgram.xls]");
                return;
            }

            if (CountofUse == 0)
            {
                MovePracticeAndStateExam();
                CountofUse++;
            }

            Excel.Workbook bookCore = App.Workbooks.Open(CurrentFolder + "Data\\DataToProgram.xls");
            if (string.IsNullOrEmpty(groupName) && string.IsNullOrEmpty(subjectName))
                foreach (Group group in Manager.Groups)
                {
                    foreach (Subject subject in @group.Subjects)
                    {
                        Semestr semestr = pivricha == 1 ? subject.FirstSemestr : subject.SecondSemestr;
                        if (semestr != null)
                            CreateOblicForOneSubject(bookCore, group, @group.Name, pivricha);
                    }
                }
            else if (string.IsNullOrEmpty(subjectName) && !string.IsNullOrEmpty(groupName))
            {
                Group gropu = Manager.Groups.Find(group => group.Name.Equals(groupName));
                if (gropu != null)
                    foreach (Subject subject in gropu.Subjects)
                    {
                        Semestr semestr = pivricha == 1 ? subject.FirstSemestr : subject.SecondSemestr;
                        if (semestr != null)
                            CreateOblicForOneSubject(bookCore, gropu, subject.Name, pivricha);
                    }
            }
            else if (!string.IsNullOrEmpty(subjectName) && !string.IsNullOrEmpty(groupName))
            {
                Group gropu = Manager.Groups.Find(group => group.Name.Equals(groupName));
                if (gropu != null)
                    CreateOblicForOneSubject(bookCore, gropu, subjectName, pivricha);
            }

            Control.IfShow = false;
            bookCore.Close();
        }

        private static void MovePracticeAndStateExam()
        {
            foreach (Group group in Manager.Groups)
            {
                group.FirstRomeSemestr = ArabNormalize(group.FirstRomeSemestr);
                if (group.Practice != null)
                    foreach (Practice practice in @group.Practice)
                    {
                        practice.Semestr = ArabNormalize(practice.Semestr);

                        Subject subject = new Subject()
                        {
                            Name = practice.Name,
                            NumberOfOlic = practice.NumberOfOlic,
                            Teacher = practice.Teacher.Aggregate("", (current, s) => current + s)
                        };
                        Semestr semestr = new Semestr
                        {
                            CountOfHours = practice.CountOfHours,
                            PracticeFormOfControl = practice.FormOfControl
                        };

                        if (practice.Semestr.Equals(group.FirstRomeSemestr))
                            subject.FirstSemestr = semestr;
                        else subject.SecondSemestr = semestr;

                        group.Subjects.Add(subject);
                    }

                if (group.StateExamination != null)
                    foreach (StateExamination stateExamination in @group.StateExamination)
                    {
                        stateExamination.Semestr = ArabNormalize(stateExamination.Semestr);
                        foreach (Subject subject in @group.Subjects)
                        {
                            if (CustomEquals(subject.Name, stateExamination.Name))
                            {
                                if (group.FirstRomeSemestr.Equals(stateExamination.Semestr) &&
                                    subject.FirstSemestr != null)
                                {
                                    subject.FirstSemestr.StateExamination = subject.FirstSemestr.Isput;
                                    subject.FirstSemestr.Isput = 0;
                                }
                                else if (subject.SecondSemestr != null)
                                {
                                    subject.SecondSemestr.StateExamination = subject.SecondSemestr.Isput;
                                    subject.SecondSemestr.Isput = 0;
                                }
                            }
                        }
                    }
            }
        }

        private static void CreateOblicForOneSubject(Excel.Workbook book, Group group, string subjectName, int pivricha)
        {
            try
            {
                Excel.Workbook bookOfOblic = null;
                Excel.Worksheet sheetOfOblic;
                string nameOfOblic = CreateSheetName(subjectName);
                bool exist = false;

                if (File.Exists(CurrentFolder + "User Data\\Облік успішності\\" + group.Name + ".xls"))
                {
                    bookOfOblic =
                        App.Workbooks.Open(CurrentFolder + "User Data\\Облік успішності\\" + group.Name + ".xls");
                    exist = true;
                }

                if (!exist)
                {
                    bookOfOblic = App.Workbooks.Add(Type.Missing);
                    sheetOfOblic = bookOfOblic.Worksheets[1];
                    sheetOfOblic.Name = nameOfOblic;
                    bookOfOblic.SaveAs(CurrentFolder + "User Data\\Облік успішності\\" + group.Name,
                        Excel.XlFileFormat.xlAddIn8);
                }
                else
                {
                    exist =
                        bookOfOblic.Worksheets.Cast<object>()
                            .Any(sheet => ((Excel.Worksheet) sheet).Name.Equals(nameOfOblic));
                    if (exist)
                    {
                        if (!Control.IfShow)
                        {
                            Control control =
                                new Control("Група [" + group.Name + "]. Уже існує облік успішності для предмету:\n" +
                                            subjectName);
                            control.ShowDialog();
                            if (Control.ButtonClick == 1)
                            {
                                Excel.Application newApp = new Excel.Application() {Visible = true};
                                ((Excel.Worksheet)
                                    newApp.Workbooks.Open(CurrentFolder + "User Data\\Облік успішності\\" + group.Name +
                                                          ".xls").Worksheets[subjectName]).Select();

                                Control.ButtonClick = 0;
                                control.SetButtonReseachEnabled(false);
                                control.ShowDialog();
                            }
                            if (Control.ButtonClick == 2)
                            {
                                Control.ButtonClick = 0;
                                return;
                            }

                            sheetOfOblic = bookOfOblic.Worksheets[nameOfOblic];
                            sheetOfOblic.Cells.Delete();
                            Control.ButtonClick = 0;
                        }
                        else
                        {
                            if (Control.ButtonClick == 2)
                                return;
                            sheetOfOblic = bookOfOblic.Worksheets[nameOfOblic];
                            sheetOfOblic.Cells.Delete();
                        }
                    }
                    else
                    {
                        sheetOfOblic = bookOfOblic.Worksheets.Add(Type.Missing);
                        sheetOfOblic.Name = nameOfOblic;
                    }
                }

               foreach (Subject subject in @group.Subjects)
               {
                   Semestr semestr = pivricha == 1 ? subject.FirstSemestr : subject.SecondSemestr;
                   if (semestr != null)
                   {
                       if (semestr.DyfZalik > 0 || semestr.Zalic > 0 || semestr.Isput > 0)
                       {
                           CreateZalicExamenAndDufZalic();
                           bookOfOblic.Close();
                       }
                       else if (semestr.StateExamination > 0)
                       {
                           CreateStateExamen();
                           bookOfOblic.Close();
                       }
                       else if (semestr.CursovaRobota > 0 || !string.IsNullOrEmpty(semestr.PracticeFormOfControl))
                       {
                           CreateKpOrPractice();
                           bookOfOblic.Close();
                       }
                   }
               }
            }
            catch (Exception e)
            {
                MassageError("", "", "Щось не гаразд із підключенням до обліків успішності:\n" + e);
            }
        }

        private static void CreateKpOrPractice()
        {
//            MessageBox.Show("Practica or KP");
        }

        private static void CreateStateExamen()
        {
//            MessageBox.Show("Statement examen");
        }

        private static void CreateZalicExamenAndDufZalic()
        {
//            MessageBox.Show("Zalic examen duf");
        }

        private static string CreateSheetName(string s)
        {
            return s.Length <= 32 ? s.Replace("*", "&") : s.Substring(0, 31).Replace("*", "&");
        }

        private static int FromRomeToArab(string rome)
        {
            int arab = 0;
            rome = ArabNormalize(rome);

            if (rome.Equals("I")) arab = 1;
            else if (rome.Equals("II")) arab = 2;
            else if (rome.Equals("III")) arab = 3;
            else if (rome.Equals("IV")) arab = 4;
            else if (rome.Equals("V")) arab = 5;
            else if (rome.Equals("VI")) arab = 6;
            else if (rome.Equals("VII")) arab = 7;
            else if (rome.Equals("VIII")) arab = 8;

            return arab;
        }

        public static string ArabToRome(int arab)
        {
            var rome = "";
            switch (arab)
            {
                case 1:
                    rome = "I";
                    break;
                case 2:
                    rome = "II";
                    break;
                case 3:
                    rome = "III";
                    break;
                case 4:
                    rome = "IV";
                    break;
                case 5:
                    rome = "V";
                    break;
                case 6:
                    rome = "VI";
                    break;
                case 7:
                    rome = "VII";
                    break;
                case 8:
                    rome = "VIII";
                    break;
            }
            return rome;
        }

        private static string ArabNormalize(string str)
        {
            char[] ch = str.ToCharArray();
            for (int i = 0; i < str.Length; i++)
            {
                var arg = (int)ch[i];
                if (arg == 1030) arg = 73;
                ch[i] = (char)arg;
            }
            return new string(ch);
        }

    }
}