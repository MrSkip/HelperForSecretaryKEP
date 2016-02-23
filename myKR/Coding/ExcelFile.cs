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
        public static Excel.Application App = new Excel.Application
        {
            Visible = false,
            DisplayAlerts = false 
        };

        public static string CurrentFolder = Environment.CurrentDirectory + "\\";

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
            MovePracticeAndStateExam();
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
                        Name = RemoveSymbolFromSubjectName(subjectName.ToString().Trim()),
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

        private static string RemoveSymbolFromSubjectName(string subjectName)
        {
            if (string.IsNullOrEmpty(subjectName)) return "";
            string withoutSymbol = "";
            foreach (char c in subjectName)
            {
                if (c != '*')
                    withoutSymbol += c;
            }
            return withoutSymbol;
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

            Excel.Workbook bookCore = App.Workbooks.Open(CurrentFolder + "Data\\DataToProgram.xls");
            if (string.IsNullOrEmpty(groupName) && string.IsNullOrEmpty(subjectName))
                foreach (Group group in Manager.Groups)
                {
                    foreach (Subject subject in @group.Subjects)
                    {
                        Semestr semestr = pivricha == 1 ? subject.FirstSemestr : subject.SecondSemestr;
                        if (semestr != null)
                            CreateOblicForOneSubject(bookCore, group, subject.Name, pivricha);
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

                        Subject subject = new Subject
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
                                if (subject.FirstSemestr != null && group.FirstRomeSemestr.Equals(stateExamination.Semestr))
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
                
                Subject subjectFind = group.Subjects.Find(subject => subject.Name.Equals(subjectName));
                string nameOfOblic = "";
                if (subjectFind != null)
                {
                    Semestr semestrFindSemestr = pivricha == 1 ? subjectFind.FirstSemestr : subjectFind.SecondSemestr;
                    if (semestrFindSemestr != null)
                    {
                        nameOfOblic = semestrFindSemestr.CursovaRobota > 0
                            ? CreateSheetName("КП" + subjectName)
                            : CreateSheetName(subjectName);
                    }
                    else return;
                }
                else
                {
                    MessageBox.Show("У групі [" + group.Name + "] не знайдено предмет - [" + subjectName + "]");
                    return;
                }
                bool exist = false;

                if (File.Exists(CurrentFolder + "User Data\\Облік успішності\\" + group.Name + ".xls"))
                {
                    bookOfOblic =
                        App.Workbooks.Open(CurrentFolder + "User Data\\Облік успішності\\" + group.Name + ".xls");
                    exist = true;
                }

                if (!exist)
                {
                    if (!File.Exists(CurrentFolder + "Data\\WithMacros.xls"))
                    {
                        MessageBox.Show("Файл не існує - [" + CurrentFolder + "Data\\WithMacros.xls" + "]");
                        return;
                    }

                    File.Copy(CurrentFolder + "Data\\WithMacros.xls", CurrentFolder +  "User Data\\Облік успішності\\" + group.Name + ".xls");

                    bookOfOblic = App.Workbooks.Open(CurrentFolder + "User Data\\Облік успішності\\" + group.Name + ".xls");
                    sheetOfOblic = (Excel.Worksheet)bookOfOblic.Worksheets[1];
                    sheetOfOblic.Name = nameOfOblic;
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
                                                          ".xls").Worksheets[nameOfOblic]).Select();

                                Control.ButtonClick = 0;
                                control.SetButtonReseachEnabled(false);
                                control.ShowDialog();

                                newApp.Quit();
                            }
                            if (Control.ButtonClick == 2)
                                return;

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
                   if (semestr != null && subject.Name.Equals(subjectName))
                   {
                       if (semestr.DyfZalik > 0 || semestr.Zalic > 0 || semestr.Isput > 0)
                       {
                           CreateZalicExamenAndDufZalic(book.Worksheets["Залік - ДифЗалік - Екзамен"], sheetOfOblic, group, subject, semestr, pivricha);
                       }
                       else if (semestr.StateExamination > 0)
                       {
                           CreateStateExamen(book.Worksheets["Державний екзамен"], sheetOfOblic, group, subject, semestr, pivricha);
                       }
                       else if (semestr.CursovaRobota > 0 || !string.IsNullOrEmpty(semestr.PracticeFormOfControl))
                       {
                           CreateKpOrPractice(book.Worksheets["КП - Технологічна практика"], sheetOfOblic, group, subject, semestr, pivricha);
                       }
                       break;
                   }
               }
                bookOfOblic.Save();
            }
            catch (Exception e)
            {
                MassageError("", "", "Щось не гаразд із підключенням до обліків успішності:\n" + e);
            }
        }

        private static void CreateKpOrPractice(Excel.Worksheet sheetTamplate, Excel.Worksheet sheet, Group group, Subject subject, Semestr semestr, int pivricha)
        {
            sheet.Cells.PasteSpecial(sheetTamplate.Cells.Copy());

            sheet.Cells[13, "E"].Value = group.TrainingDirection.Equals("Програмна інженерія") ? "Програмної інженерії"
                : "Метрології та інформаційно-вимірювальної технології";
            sheet.Cells[15, "F"].Value = group.Speciality;
            sheet.Cells[17, "D"].Value = group.Course;
            sheet.Cells[17, "G"].Value = group.Name;
            sheet.Cells[19, "I"].Value = group.Year + "-" + (int.Parse(group.Year.Trim()) + 1);
            sheet.Cells[26, "F"].Value = subject.Name;
            sheet.Cells[28, "D"].Value = pivricha == 1 ? group.FirstRomeSemestr : ArabToRome(FromRomeToArab(group.FirstRomeSemestr) + 1);
            sheet.Cells[22, "M"].Value = CreateNumberOfOblic(subject.NumberOfOlic, pivricha == 1 ? group.Year : (int.Parse(group.Year.Trim()) + 1) + "");
            sheet.Cells[30, "Q"].Value = semestr.CountOfHours;
            sheet.Cells[30, "F"].Value = FormaZdachi(semestr);
//            sheet.Cells[32, "K"].Value = subject.Teacher + "_____";
//            sheet.Cells[100, "N"].Value = subject.Teacher;

            int n = 45;
            foreach (Student student in @group.Students)
            {
                sheet.Cells[n, "C"].Value = student.Pib;
                sheet.Cells[n, "H"].Value = student.NumberOfBook;
                n++;
            }
            if (n != 75)
                sheet.Range["B" + n, "Q" + 74].Delete();
        }

        private static void CreateStateExamen(Excel.Worksheet sheetTamplate, Excel.Worksheet sheet, Group group, Subject subject, Semestr semestr, int pivricha)
        {
            sheet.Cells.PasteSpecial(sheetTamplate.Cells.Copy());
            sheet.Cells[4, "H"].Value = subject.Name;
            sheet.Cells[9, "C"].Value = group.Name;
            sheet.Cells[20, "G"].Value = subject.Teacher + "_________________________________";
            sheet.Cells[84, "H"].Value = subject.Teacher + "__";

            int n = 46;
            foreach (Student student in @group.Students)
            {
                sheet.Cells[n, "C"].Value = student.Pib;
                n++;
            }

            if (n != 76)
                sheet.Range["B" + n, "Q" + 75].Delete();

            // Count of students in group
            sheet.Cells[12, "G"] = "__" + (n - 46) + "__";
        }

        private static void CreateZalicExamenAndDufZalic(Excel.Worksheet sheetTamplate, Excel.Worksheet sheet, 
            Group group, Subject subject, Semestr semestr, int pivricha)
        {
            sheet.Cells.PasteSpecial(sheetTamplate.Cells.Copy());

            sheet.Cells[13, "E"].Value = group.TrainingDirection.Equals("Програмна інженерія") ? "Програмної інженерії" 
                : "Метрології та інформаційно-вимірювальної технології";
            sheet.Cells[15, "F"].Value = group.Speciality;
            sheet.Cells[17, "D"].Value = group.Course;
            sheet.Cells[17, "G"].Value = group.Name;
            sheet.Cells[19, "I"].Value = group.Year + "-" + (int.Parse(group.Year.Trim()) + 1);
            sheet.Cells[26, "F"].Value = subject.Name;
            sheet.Cells[28, "D"].Value = pivricha == 1 ? group.FirstRomeSemestr : ArabToRome(FromRomeToArab(group.FirstRomeSemestr) + 1);
            sheet.Cells[22, "M"].Value = CreateNumberOfOblic(subject.NumberOfOlic, pivricha == 1 ? group.Year : (int.Parse(group.Year.Trim()) + 1) + "");
            sheet.Cells[30, "Q"].Value = semestr.CountOfHours;
            sheet.Cells[30, "F"].Value = FormaZdachi(semestr);
            sheet.Cells[32, "E"].Value = subject.Teacher;
            sheet.Cells[94, "N"].Value = subject.Teacher;

            int n = 39;
            foreach (Student student in @group.Students)
            {
                sheet.Cells[n, "C"].Value = student.Pib;
                sheet.Cells[n, "H"].Value = student.NumberOfBook;
                n++;
            }
            if (n != 69)
                sheet.Range["B" + n, "Q" + 68].Delete();
        }

        private static string CreateNumberOfOblic(string number, string currentYear)
        {
            int x = 0;
            if (!string.IsNullOrEmpty(number))
            x = int.Parse(number.Trim());

            if (x < 10) number = "00" + number;
            else if (x < 100) number = "0" + number;

            int n = int.Parse(currentYear.Trim()) - 2000;

            number = "" + n + "." + number;
            return number;
        }

        private static string FormaZdachi(Semestr semestr)
        {
            string s = "";

            if (semestr.CursovaRobota > 0) s = "курсовий проект";
            else if (semestr.DyfZalik > 0) s = "диф залік";
            else if (semestr.Isput > 0) s = "екзамен";
            else if (!string.IsNullOrEmpty(semestr.PracticeFormOfControl)) s = semestr.PracticeFormOfControl;
            else if (semestr.Zalic > 0) s = "залік";
            else if (semestr.StateExamination > 0) s = "протокол";

            return s;
        }

        private static string CreateSheetName(string s)
        {
            string s2 = "";
            foreach (char c in s)
            {
                if (c.Equals('[') || c.Equals(']') || c.Equals('[') || c.Equals('/') || c.Equals('\\') || c.Equals('?') ||
                    c.Equals('*'))
                    continue;
                s2 += c;
            }

            return s2.Length <= 31 ? s2 : s2.Substring(0, 31);
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

        private static string ArabToRome(int arab)
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


        // Creating ZvedVidomist

        // Read Oblics Uspisnosti
        public static void CreateVidomist(Group group, int pivricha)
        {
            if (!File.Exists(CurrentFolder + "User Data\\Облік успішності\\" + group.Name + ".xls"))
            {
                MessageBox.Show("У вас немає обліків успішності для групи - " + group.Name + " за " + pivricha +
                                " півріччя");
            }
            else
            {
                try
                {
                    Excel.Workbook book = App.Workbooks.Open(CurrentFolder + "User Data\\Облік успішності\\" + group.Name + ".xls");
                    foreach (object sheetO in book.Worksheets)
                    {
                        Excel.Worksheet sheet = (Excel.Worksheet) sheetO;

                        var protocol = sheet.Cells[3, "H"].Value;
                        if (protocol == null || string.IsNullOrEmpty(protocol))
                        {
                            var formaZdachi = sheet.Cells[30, "F"].Value;
                            if (formaZdachi == null || string.IsNullOrEmpty(formaZdachi.ToString())) continue;

                            if (formaZdachi.ToString().Equals("диф залік") || formaZdachi.ToString().Equals("екзамен") ||
                                formaZdachi.ToString().Equals("залік"))
                                ReadOcinkaFromOblics(group, sheet, pivricha, 1);
                            else
                            {
                                ReadOcinkaFromOblics(group, sheet, pivricha, 2);
                            }
                        }
                        else
                        {
                            ReadOcinkaFromOblics(group, sheet, pivricha, 3);
                        }
                    }
                    book.Save();
                    book.Close();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Щось не так із методом CreateVidomist()\n" + e);
                }
            }

            CreateZvedeniaVidomist(group, pivricha);
        }

        // if type == 1 than DufZalicZalic else if == 2 than PracticeOrKP else StateExamen
        private static void ReadOcinkaFromOblics(Group @group, Excel.Worksheet sheet, int pivricha, int type)
        {
            try
            {
                List<Ocinka> list = new List<Ocinka>();
                string subjectName;

                string currentSemestr = pivricha == 1
                    ? group.FirstRomeSemestr
                    : ArabToRome(FromRomeToArab(ArabNormalize(group.FirstRomeSemestr.Trim())) + 1);

                if (type <= 2)
                {
                    var subjectNameV = sheet.Cells[26, "F"].Value;
                    if (subjectNameV == null || string.IsNullOrEmpty(subjectNameV)) return;
                    subjectName = subjectNameV.ToString();

                    var semestr = sheet.Cells[28, "D"].Value;
                    if (semestr == null || string.IsNullOrEmpty(semestr.ToString())) return;

                    if (!currentSemestr.Equals(semestr.ToString()))
                    {
                        return;
                    }
                }
                else
                {
                    var protocol = sheet.Cells[3, "H"].Value;
                    if (protocol == null || string.IsNullOrEmpty(protocol)) return;
                    var subjectNameV = sheet.Cells[4, "H"].Value;

                    if (subjectNameV == null || string.IsNullOrEmpty(subjectNameV)) return;
                    subjectName = subjectNameV.ToString();

                    Subject subjectT = group.Subjects.Find(subject1 => subject1.Name.Equals(subjectName));

                    if (subjectT == null) return;
                    Semestr semestr = pivricha == 1
                        ? subjectT.FirstSemestr
                        : subjectT.SecondSemestr;

                    if (semestr == null) return;
                    if (semestr.StateExamination <= 0) return;
                }


                int n;
                switch (type)
                {
                    case 1:
                        n = 38;
                        break;
                    case 2:
                        n = 44;
                        break;
                    default:
                        n = 45;
                        break;
                }

                string ocinkaPositio = type <= 2 ? "L" : "J";
                while (true)
                {
                    n++;
                    var studentName = sheet.Cells[n, "C"].Value;
                    if (studentName == null || string.IsNullOrEmpty(studentName.ToString())) break;
                    var pas = sheet.Cells[n, ocinkaPositio].Value ?? "";
                    list.Add(new Ocinka
                    {
                        Name = studentName,
                        Number = pas + ""
                    });
                }
                Subject subject = group.Subjects.Find(subject1 => subject1.Name.Trim().Equals(subjectName.Trim()));
                if (subject != null)
                {
                    subject.Ocinka = list;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Щось не гаразд із зчитуванням оцінок із Обліків Успішності\n" + e);
            }
        }

        private static void CreateZvedeniaVidomist(Group group, int pivricha)
        {
            Excel.Workbook bookTamplate = null;
            Excel.Workbook book = null;

            try
            {
                string pathToVidomist = CurrentFolder +
                                        "User Data\\Зведена відомість успішності\\Зведена відомість успішності за " +
                                        (pivricha == 1 ? "1-ше півріччя.xls" : "2-ге півріччя.xls");

                if (!File.Exists(CurrentFolder + "Data\\DataToProgram.xls"))
                {
                    MessageBox.Show("Відсутній файл:\n" + CurrentFolder + "Data\\DataToProgram.xls");
                    return;
                }

                bookTamplate = App.Workbooks.Open(CurrentFolder + "Data\\DataToProgram.xls");
                Excel.Worksheet sheetTamplate = (Excel.Worksheet) bookTamplate.Worksheets["Зведена відомість"];

                Excel.Worksheet sheet;

                if (!File.Exists(pathToVidomist))
                {
                    if (!File.Exists(CurrentFolder + "Data\\WithMacros.xls"))
                    {
                        MessageBox.Show("Файл не існує - [" + CurrentFolder + "Data\\WithMacros.xls" + "]");
                        return;
                    }
                    File.Copy(CurrentFolder + "Data\\WithMacros.xls", pathToVidomist);
                    book = App.Workbooks.Open(pathToVidomist);
                    sheet = (Excel.Worksheet) book.Worksheets[1];
                    sheet.Name = group.Name;
                }
                else
                {
                    book = App.Workbooks.Open(pathToVidomist);
                    bool exist =
                        book.Worksheets.Cast<object>()
                            .Any(sheet2 => ((Excel.Worksheet) sheet2).Name.Equals(group.Name));
                    if (exist)
                    {
                        if (!Control.IfShow)
                        {
                            Control control =
                                new Control("Уже існує зведена відомість для групи:\n" + group.Name);
                            control.ShowDialog();
                            if (Control.ButtonClick == 1)
                            {

                                Excel.Application newApp = new Excel.Application {Visible = true};
                                ((Excel.Worksheet)
                                    newApp.Workbooks.Open(pathToVidomist).Worksheets[group.Name]).Select();

                                Control.ButtonClick = 0;
                                control.SetButtonReseachEnabled(false);
                                control.ShowDialog();

                                newApp.Quit();
                            }
                            if (Control.ButtonClick == 2)
                                return;

                            sheet = book.Worksheets[group.Name];
                            sheet.Cells.Delete();
                        }
                        else
                        {
                            if (Control.ButtonClick == 2)
                                return;
                            sheet = book.Worksheets[group.Name];
                            sheet.Cells.Delete();
                        }
                    }
                    else
                    {
                        sheet = book.Worksheets.Add(Type.Missing);
                        sheet.Name = group.Name;
                    }
                }

                string semestrCurrent = pivricha == 1
                    ? group.FirstRomeSemestr
                    : ArabToRome(FromRomeToArab(group.FirstRomeSemestr) + 1);

                int yearCurrent = string.IsNullOrEmpty(group.Year)
                    ? 0
                    : int.Parse(group.Year.Trim()) + 1;

                sheet.Cells.PasteSpecial(sheetTamplate.Cells.Copy());

                sheet.Cells[4, "C"].Value = "спеціальності \"" + group.Speciality + "\"";

                sheet.Cells[5, "D"].Value = "групи " + group.Name + " за " + semestrCurrent + " семестр " + group.Year +
                                            "-" +
                                            yearCurrent + " навчального року";

                sheet.Cells[45, "K"].Value = "/ " + group.Curator + " /";

                List<Subject> subjects = pivricha == 1
                    ? @group.Subjects.FindAll(subject => subject.FirstSemestr != null)
                    : @group.Subjects.FindAll(subject => subject.SecondSemestr != null);

                int count = -1;
                char[] c = {'F', 'F'};
                bool practiceExist = false;

                for (int i = 1; i < 6; i++)
                {
                    List<Subject> list = null;
                    if (i == 1)
                    {
                        list =
                            subjects.FindAll(
                                subject =>
                                    pivricha == 1
                                        ? subject.FirstSemestr.Isput > 0 || subject.FirstSemestr.StateExamination > 0
                                        : subject.SecondSemestr.Isput > 0 || subject.SecondSemestr.StateExamination > 0);
                    }
                    else if (i == 2)
                    {
                        list =
                            subjects.FindAll(
                                subject =>
                                    pivricha == 1
                                        ? subject.FirstSemestr.DyfZalik > 0
                                        : subject.SecondSemestr.DyfZalik > 0);
                    }
                    else if (i == 3)
                    {
                        list =
                            subjects.FindAll(
                                subject =>
                                    pivricha == 1
                                        ? subject.FirstSemestr.CursovaRobota > 0
                                        : subject.SecondSemestr.CursovaRobota > 0);
                    }
                    else if (i == 4)
                    {
                        list =
                            subjects.FindAll(
                                subject =>
                                    pivricha == 1 ? subject.FirstSemestr.Zalic > 0 : subject.SecondSemestr.Zalic > 0);
                    }
                    else if (i == 5)
                    {
                        practiceExist = true;
                        list =
                            subjects.FindAll(
                                subject =>
                                    pivricha == 1
                                        ? !string.IsNullOrEmpty(subject.FirstSemestr.PracticeFormOfControl)
                                        : !string.IsNullOrEmpty(subject.SecondSemestr.PracticeFormOfControl));
                    }
                    if (list != null && list.Count > 0)
                    {
                        foreach (Subject subject in list)
                        {
                            count++;
                            sheet.Cells[9, c[1].ToString()] = subject.Name;
                            sheet.Cells[9, c[1].ToString()].ColumnWidth = ColumnWidth(subject.Name);
                            sheet.Cells[43, c[1].ToString()] = subject.Teacher;
                            if (group.Students != null)
                            {
                                sheet.Cells[41, c[1].ToString()].Value = "=Uspishnist(" + count + "," +
                                                                         group.Students.Count + ")";
                                sheet.Cells[42, c[1].ToString()].Value = "=Quality(" + count + "," +
                                                                         group.Students.Count + ")";
                            }

                            int n = 10;
                            if (subject.Ocinka != null)
                                foreach (Ocinka ocinka in subject.Ocinka)
                                {
                                    n++;
                                    sheet.Cells[n, c[1].ToString()].Value = ocinka.Number;

                                    group.Students[n - 11].Ocinkas.Add(new Ocinka
                                    {
                                        Name = subject.Name,
                                        Number = ocinka.Number
                                    });
                                }
                            c[1]++;
                        }
                    }
                    else continue;

                    switch (i)
                    {
                        case 1:
                        {
                            sheet.Cells[8, c[0].ToString()].Value = "Іспит";
                            char ch = c[1];
                            ch--;
                            sheet.Range[c[0].ToString() + 8, ch.ToString() + 8].Merge();
                            sheet.Range[c[0].ToString() + 8, ch.ToString() + 8].HorizontalAlignment =
                                Excel.XlHAlign.xlHAlignCenter;
                        }
                            break;
                        case 2:
                        {
                            sheet.Cells[8, c[0].ToString()].Value = "Д/З";
                            char ch = c[1];
                            ch--;
                            sheet.Range[c[0].ToString() + 8, ch.ToString() + 8].Merge();
                            sheet.Range[c[0].ToString() + 8, ch.ToString() + 8].HorizontalAlignment =
                                Excel.XlHAlign.xlHAlignCenter;
                        }
                            break;
                        case 4:
                        {
                            sheet.Cells[8, c[0].ToString()].Value = "Залік";
                            char ch = c[1];
                            ch--;
                            sheet.Range[c[0].ToString() + 8, ch.ToString() + 8].Merge();
                            sheet.Range[c[0].ToString() + 8, ch.ToString() + 8].HorizontalAlignment =
                                Excel.XlHAlign.xlHAlignCenter;
                        }
                            break;
                    }

                    if (!practiceExist)
                    {
                        char ch = c[1];
                        ch--;
                        sheet.Cells[7, "F"].Value = "Предмети";
                        sheet.Range["F" + 7, ch.ToString() + 7].Merge();
                        sheet.Range["F" + 7, ch.ToString() + 7].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    }

                    c[0] = c[1];
                }

                int row = 10;
                sheet.Cells[9, c[1].ToString()].Value = "Середній бал";
                char cBenefics = c[1];
                c[0] = c[1];
                c[0]--;

                cBenefics++;
                cBenefics++;

                foreach (Student student in @group.Students)
                {
                    row++;
                    sheet.Cells[row, "D"].Value = student.Pib;
                    sheet.Cells[row, "E"].Value = student.FormaTeaching;
                    sheet.Cells[row, cBenefics.ToString()].Value = student.Benefits;
                    sheet.Cells[row, c[1].ToString()].Formula = "=AVERAGE(" + "F" + row + ":" + c[0] + row + ") - 0.5";
                    sheet.Cells[row, c[1].ToString()].NumberFormatLocal = "##";

                    bool hight = true;
                    int sum = 0;
                    int countOf = 0;

                    if (student.Ocinkas.Count >= 1)
                    {
                        foreach (Ocinka ocinka in student.Ocinkas)
                        {
                            if (string.IsNullOrEmpty(ocinka.Number)) countOf++;
                            int number;
                            if (!int.TryParse(ocinka.Number, out number)) continue;
                            if (number < 10) hight = false;
                            sum += number;
                        }
                        if (student.FormaTeaching.Equals("п")) hight = false;

                        char stupendiaColumnPosution = c[1];
                        stupendiaColumnPosution++;

                        if (group.Students[0].Ocinkas.Count - countOf == 0)
                            hight = false;
                        else if (sum / (group.Students[0].Ocinkas.Count - countOf) >= 7  && !student.FormaTeaching.Equals("п"))
                        sheet.Cells[row, stupendiaColumnPosution.ToString()].Value = 1;

                        if (hight)
                            sheet.Cells[row, stupendiaColumnPosution.ToString()].Interior.Color =
                                System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    }

                    sheet.Cells[row, cBenefics].Value = student.Benefits;
                }

                sheet.Range["C7", c[1].ToString() + 44].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                if (group.Students.Count < 30)
                    sheet.Range["A" + (group.Students.Count + 11), "IV" + 40].Delete(
                        Excel.XlDeleteShiftDirection.xlShiftUp);

                // Add vidomist to arhive
                ArhiveZvedVidomist(sheet, semestrCurrent);
            }
            catch (Exception e)
            {
                MessageBox.Show(e + "" + "\n" + group.Name);
            }
            finally
            {
                if (book != null)
                {
                    book.Save();
                    book.Close();
                }
                if (bookTamplate != null)
                    bookTamplate.Close();
            }
            Control.IfShow = false;
        }

        private static double ColumnWidth(string s)
        {
            if (s.Length <= 21) return 5.57;
            if (s.Length <= 40) return 9.70;
            if (s.Length <= 55) return 11;
            return 13.43;
        }

        private static void ArhiveZvedVidomist(Excel.Worksheet sheet, string semesterRome)
        {
            Excel.Workbook book = null;
            try
            {
                if (string.IsNullOrEmpty(semesterRome)) return;

                string pathToArhiveFile = CurrentFolder + "User Data\\Зведена відомість успішності\\Архів\\" +
                                          sheet.Name +
                                          ".xls";

                if (!File.Exists(pathToArhiveFile))
                    File.Copy(CurrentFolder + "Data\\WithMacros.xls", pathToArhiveFile);

                book = App.Workbooks.Open(pathToArhiveFile);
                Excel.Worksheet sheetArhive;

                string sheetNameOfArhive = semesterRome + " семестр";

                bool bl = true;
                foreach (Excel.Worksheet worksheet in book.Worksheets)
                {
                    if (worksheet.Name.Equals(sheetNameOfArhive))
                    {
                        bl = false;
                        break;
                    }
                }

                if (bl)
                {
                    sheetArhive = book.Sheets.Count == 1
                        ? book.Worksheets.Item[1]
                        : book.Worksheets.Add(Type.Missing);
                    sheetArhive.Name = sheetNameOfArhive;
                }
                else
                {
                    sheetArhive = book.Sheets[sheetNameOfArhive];
                    sheetArhive.Cells.Delete();
                }

                sheetArhive.Cells.PasteSpecial(sheet.Cells.Copy());
            }
            catch (Exception e)
            {
                MessageBox.Show("Помилка!\n Архівування зведеної відомості: " + sheet.Name + "\nКод помилки:\n" + e);
            }
            finally
            {
                if (book != null)
                {
                    book.Save();
                    book.Close();
                }
            }
        }


    }
}