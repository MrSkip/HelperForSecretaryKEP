﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace myKR.Coding
{
    public static class ExcelFile
    {
        private static readonly PathsFile PathsFile = PathsFile.GetPathsFile();

        private static readonly log4net.ILog Log =
            log4net.LogManager.GetLogger("ExcelFile.cs");

        public static ExcelApplication.ExcelApplication App = ExcelApplication.ExcelApplication.CreateExcelApplication();

        public static void ReadRobPlan(string pathToRobPlan)
        {
            Log.Info(LoggerConstants.ENTER);

            Excel.Workbook book = App.OpenBook(pathToRobPlan);
            if (book == null)
            {
                Log.Error("Can`t opet book from path: " + pathToRobPlan);
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            foreach (Excel.Worksheet sheet in book.Worksheets)
            {
                if (sheet.Name.Trim().Length == 8 && sheet.Name.Trim().IndexOf('-') == 2 &&
                    sheet.Name.Trim().LastIndexOf('-') == 5)
                {
                    Log.Info("Add group to program with name `" + sheet.Name);
                    Manager.Groups.Add(ReadSheetFromRobPlan(sheet));
                }
            }
            MovePracticeAndStateExam();
            book.Close();
            Log.Info(LoggerConstants.EXIT);
        }

        public static void ReadStudentsAndOlicAndCurators(string pathToDb)
        {
            Log.Info(LoggerConstants.ENTER);

            Excel.Workbook book = App.OpenBook(pathToDb);

            if (book == null)
            {
                Log.Error("Path to DB with students, obliks and curators not correct: " + pathToDb);
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            // Read [База студентів]
            App.OpenWorksheet(book, "База студентів");
            if (App.LastUsedObject != null)
            {
                List<Student> students = ReadStudents((Excel.Worksheet) App.LastUsedObject);

                foreach (Group group in Manager.Groups)
                {
                    group.Students = students.FindAll(student => student.Group.Equals(group.Name));
                }
            }
            else
                Log.Error("Can`t open sheet `База студентів`");

            // Read [Реєстраційна відомість (журнал)]
            App.OpenWorksheet(book, "Реєстраційна відомість (журнал)");
            if (App.LastUsedObject != null)
            {

                List<NumberOfOblic> oblics
                    = ReadNumbersOfOblic((Excel.Worksheet) App.LastUsedObject);

                foreach (Group group in Manager.Groups)
                {
                    foreach (
                        NumberOfOblic numberOfOblic in oblics.FindAll(oblic => CustomEquals(oblic.Group, @group.Name)))
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
            else
                Log.Error("Can`t open sheet `Реєстраційна відомість (журнал)`");

            // Read [Куратори]
            App.OpenWorksheet(book, "Куратори");
            if (App.LastUsedObject == null)
            {
                Log.Info("Can`t open sheet `Куратори`");
                App.CloseBook(book, false);
                Log.Info(LoggerConstants.EXIT);
                return;
            }
            List<string[]> list =
                ReadCurator((Excel.Worksheet) App.LastUsedObject);
            foreach (Group group in Manager.Groups)
            {
                string[] s = list.Find(strings => CustomEquals(strings[1], @group.Name));
                if (s != null)
                    @group.Curator = s[0];
            }

            App.CloseBook(book, false);
            Log.Info(LoggerConstants.EXIT);
        }

        public static List<string[]> ReadCurator(Excel.Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
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
                Log.Warn("Something wrong with reading curators", e);
            }
            Log.Info(LoggerConstants.EXIT);
            return curators;
        }

        private static bool CustomEquals(string first, string second)
        {
            Log.Info(LoggerConstants.ENTER);
            first = first.ToLower().Trim().Replace("*", "");
            second = second.ToLower().Trim().Replace("*", "");
            Log.Info(LoggerConstants.EXIT);
            return first.Equals(second);
        }

        public static List<NumberOfOblic> ReadNumbersOfOblic(Excel.Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
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
                Log.Warn("Something wrang", e);
            }
            Log.Info(LoggerConstants.EXIT);
            return oblics;
        }

        public static List<Student> ReadStudents(Excel.Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
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
                    Log.Warn("Something wrong", e);
                    MassageError(sheet.Name, "", "Щось не гаразд із зчитуванням студентів\n" + e);
                }
            }

            Log.Info(LoggerConstants.EXIT);
            return students;
        }

        private static Group ReadSheetFromRobPlan(Excel.Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);

            Group group = new Group();
            group.Name = sheet.Name;

            //Read "Напряму підготовки"
            string s = sheet.Cells[6, "R"].Value;

            bool exist = true;

            if (string.IsNullOrEmpty(s) || s.Count(c => c.Equals('"')) != 2)
                exist = false;
            else
            {
                int beginSlash = s.IndexOf("\"", StringComparison.Ordinal);
                int lastSlash = s.LastIndexOf("\"", StringComparison.Ordinal);

                s = s.Substring(beginSlash + 1, beginSlash - lastSlash - 1);

                if (string.IsNullOrEmpty(s))
                    exist = false;
            }

            if (!exist)
            {
                sheet.Cells[6, "R"].Interior.Color =
                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                s = "ВВЕДІТЬ НАПРЯМ ПІДГОТОВКИ";
                Log.Error("Expected direction of training in sheet '" + sheet.Name + "'");
            }

            exist = true;

            group.TrainingDirection = s;

            //Read "Спеціальність"
            s = sheet.Cells[7, "R"].Value;

            if (string.IsNullOrEmpty(s) || s.Count(c => c.Equals('"')) != 2)
                exist = false;
            else
            {
                int beginSlash = s.IndexOf("\"", StringComparison.Ordinal);
                int lastSlash = s.LastIndexOf("\"", StringComparison.Ordinal);

                s = s.Substring(beginSlash + 1, beginSlash - lastSlash - 1);

                if (string.IsNullOrEmpty(s))
                    exist = false;
            }

            if (!exist)
            {
                sheet.Cells[7, "R"].Interior.Color =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                s = "ВВЕДІТЬ НАЗВУ СПЕЦІАЛЬНОСТІ";
                Log.Error("Expected spesiality in sheet '" + sheet.Name + "'");
            }

            group.Speciality = s;

            s = sheet.Cells[7, "R"].Value;

            //Код спеціальності
            if (exist)
            {
                s = s.Trim().Substring(0, s.Trim().IndexOf(" \"", StringComparison.Ordinal));

                if (!s.Contains(" "))
                {
                    sheet.Cells[7, "R"].Interior.Color =
                        System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    s = "КОД";
                    Log.Error("Code of spesiality is incorrect in sheet '" + sheet.Name + "'");
                }
                else
                    s = s.Substring(s.IndexOf(" ", StringComparison.Ordinal)).Trim();
            }

            exist = true;
            group.CodeOfSpeciality = s;

            //read "Курс"
            s = sheet.Cells[9, "R"].Value;

            if (string.IsNullOrEmpty(s) || !s.Trim().StartsWith("Курс"))
                exist = false;
            else
            {
                int coursePosition = GetPositionForCellCource(s.Trim());

                if (coursePosition == -1)
                    exist = false;
                else
                {
                    s = s.Trim().Substring(coursePosition);
                    if (!s.Contains("_"))
                        exist = false;
                    else
                    {
                        s = s.Remove(s.IndexOf("_", StringComparison.Ordinal));
                    }
                }
            }

            if (!exist)
            {
                s = "ВВЕДІТЬ КУРС";
                sheet.Cells[9, "R"].Interior.Color 
                    = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                Log.Error("Course is incorrect in sheet '" + sheet.Name + "'");
            }

            exist = true;
            group.Course = s;

            //Read "Рік"
            s = sheet.Cells[6, "B"].Value;

            if (string.IsNullOrEmpty(s))
                exist = false;
            else
            {
                try
                {
                    s = s.Substring(s.Length - 9, 4);
                }
                catch (IndexOutOfRangeException e)
                {
                    exist = false;
                    Log.Warn("IndexOutOfRangeException when try to find year", e);
                }
            }

            if (!exist)
            {
                s = "ВВЕДІТЬ РІК";
                sheet.Cells[6, "B"].Interior.Color
                    = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                Log.Error("Year is incorrect in sheet '" + sheet.Name + "'");
            }

            exist = true;
            group.Year = s;

            //Read "Семестр для першого півріччя"
            s = sheet.Cells[15, "Y"].Value;

            if (string.IsNullOrEmpty(s))
                exist = false;
            else
            {
                try
                {
                    if (!s.Contains(' '))
                        exist = false;
                    else
                        s = s.Trim().Substring(0, s.Trim().IndexOf(' '));
                }
                catch (IndexOutOfRangeException e)
                {
                    exist = false;
                    Log.Warn("IndexOutOfRangeException when try to get semestr", e);
                }
            }
            if (!exist)
            {
                s = "ВВЕДІТЬ СЕМЕСТР";
                sheet.Cells[15, "Y"].Interior.Color
                    = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                Log.Error("Semestr is incorrect in sheet '" + sheet.Name + "'");
            }

            group.FirstRomeSemestr = s;

            group.Subjects = ReadSubject(sheet);
            group.Practice = ReadPractice(sheet);
            group.StateExamination = ReadStateExamination(sheet);

            Log.Info(LoggerConstants.EXIT);
            return group;
        }

        private static List<Subject> ReadSubject(Excel.Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
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
                sheet.Cells[15, "C"].Interior.Color =
                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                Log.Error("Expekted at the `C15` value 'Назви навчальних  дисциплін'");
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
                    if (n == 100)
                    {
                        Log.Warn("Exit with bad parameter");
                        break;
                    }

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
                    Log.Warn("Something wrong", exception);
                }
            }
            Log.Info(LoggerConstants.EXIT);
            return subjects;
        }

        private static string RemoveSymbolFromSubjectName(string subjectName)
        {
            Log.Info(LoggerConstants.ENTER);
            if (string.IsNullOrEmpty(subjectName)) return "";
            string withoutSymbol = "";
            foreach (char c in subjectName)
            {
                if (c != '*')
                    withoutSymbol += c;
            }
            Log.Info(LoggerConstants.EXIT);
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
                                Log.Warn("Something wrong while was reading practice", e);
                            }
                        }
                    }
                    break;
                }
            }
            Log.Info(LoggerConstants.EXIT);
            return practices;
        }

        private static List<StateExamination> ReadStateExamination(Excel.Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
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
                    Log.Warn("Something wrong while was reading state examinations", e);
                }
            }
            Log.Info(LoggerConstants.EXIT);
            return examinations;
        }

        private static void MassageError(string sheetName, string cell, string format)
        {
            MessageBox.Show("Помилка у робочому листі [" + sheetName + "]!\nСлід дотримуватися цього формату для клітини [" + cell + "]:\n[" + format + "]");
        }

        private static int GetPositionForCellCource(string str)
        {
            Log.Info(LoggerConstants.ENTER);
            int x = -1;
            foreach (var c in str)
            {
                x++;
                if (c.Equals('I') || c.Equals('V') || c.Equals('І'))
                    break;
            }
            Log.Info(LoggerConstants.EXIT);
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
            Log.Info(LoggerConstants.ENTER);

            Excel.Workbook bookCore = App.OpenBook(PathsFile.PathsDto.PathToExcelDataForProgram);
            if (bookCore == null)
            {
                Log.Error("Path to Excel file with all templates are not exist");
                Log.Info(LoggerConstants.EXIT);
                return;
            }

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
            App.CloseBook(bookCore, false);

            Log.Info(LoggerConstants.EXIT);
        }

        private static void MovePracticeAndStateExam()
        {
            Log.Info(LoggerConstants.ENTER);
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
            Log.Info(LoggerConstants.EXIT);
        }

        private static void CreateOblicForOneSubject(Excel.Workbook book, Group group, string subjectName, int pivricha)
        {
            Log.Info(LoggerConstants.ENTER);
            Excel.Workbook bookOfOblic = null;
            try
            {
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
                    else
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }
                }
                else
                {
                    Log.Warn("For group `" + group.Name + "` don`t find any subjects");
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                bookOfOblic = App.OpenBook(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti 
                    + group.Name + PathsFile.PathsDto.ExcelExtensial);

                Excel.Worksheet sheetOfOblic;

                if (bookOfOblic == null)
                {
                    if (!File.Exists(PathsFile.PathsDto.PathToFileWithMacros))
                    {
                        Log.Error(LoggerConstants.FILE_NOT_EXIST + ": " + PathsFile.PathsDto.PathToFileWithMacros);
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }

                    File.Copy(PathsFile.PathsDto.PathToFileWithMacros,
                        PathsFile.PathsDto.PathToFolderWithOblicUspishnosti + group.Name + ".xls");

                    bookOfOblic = App.OpenBook(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti + group.Name + ".xls");

                    sheetOfOblic = App.OpenWorksheet(bookOfOblic, 1);
                    sheetOfOblic.Name = nameOfOblic;
                }
                else
                {
                    var exist = bookOfOblic.Worksheets.Cast<object>()
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

                                Excel.Application newApp = new Excel.Application
                                {
                                    Visible = true
                                };
                                ((Excel.Worksheet)
                                    newApp.Workbooks.Open(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti + group.Name 
                                    +".xls").Worksheets[nameOfOblic]).Select();

                                Control.ButtonClick = 0;
                                control.SetButtonReseachEnabled(false);
                                control.ShowDialog();

                                newApp.Quit();
                                App.Kill(newApp);
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
                            {
                                Log.Info(LoggerConstants.EXIT);
                                return;
                            }
                            sheetOfOblic = bookOfOblic.Worksheets[nameOfOblic];
                            sheetOfOblic.Cells.Delete();
                        }
                    }
                    else
                    {
                        sheetOfOblic = App.CreateNewSheet(bookOfOblic, nameOfOblic);
                    }
                }


                foreach (Subject subject in @group.Subjects)
                {
                    Semestr semestr = pivricha == 1 ? subject.FirstSemestr : subject.SecondSemestr;
                    if (semestr != null && subject.Name.Equals(subjectName))
                    {
                        if (semestr.DyfZalik > 0 || semestr.Zalic > 0 || semestr.Isput > 0)
                        {
                            CreateZalicExamenAndDufZalic(book.Worksheets["Залік - ДифЗалік - Екзамен"], sheetOfOblic,
                                group, subject, semestr, pivricha);
                        }
                        else if (semestr.StateExamination > 0)
                        {
                            CreateStateExamen(book.Worksheets["Державний екзамен"], sheetOfOblic, group, subject,
                                semestr, pivricha);
                        }
                        else if (semestr.CursovaRobota > 0 || !string.IsNullOrEmpty(semestr.PracticeFormOfControl))
                        {
                            CreateKpOrPractice(book.Worksheets["КП - Технологічна практика"], sheetOfOblic, group,
                                subject, semestr, pivricha);
                        }
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Log.Warn("Something wrong while reading obliks uspishnosti", e);
            }
            finally
            {
                App.CloseBook(bookOfOblic, true);
                Log.Info(LoggerConstants.EXIT);
            }
        }

        private static void CreateKpOrPractice(Excel.Worksheet sheetTamplate, Excel.Worksheet sheet, Group group, Subject subject, Semestr semestr, int pivricha)
        {
            Log.Info(LoggerConstants.ENTER);
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
            Log.Info(LoggerConstants.EXIT);
        }

        private static void CreateStateExamen(Excel.Worksheet sheetTamplate, Excel.Worksheet sheet, Group group, Subject subject, Semestr semestr, int pivricha)
        {
            Log.Info(LoggerConstants.ENTER);
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
            Log.Info(LoggerConstants.EXIT);
        }

        private static void CreateZalicExamenAndDufZalic(Excel.Worksheet sheetTamplate, Excel.Worksheet sheet, 
            Group group, Subject subject, Semestr semestr, int pivricha)
        {
            Log.Info(LoggerConstants.ENTER);
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
            Log.Info(LoggerConstants.EXIT);
        }

        private static string CreateNumberOfOblic(string number, string currentYear)
        {
            Log.Info(LoggerConstants.ENTER);
            int x = 0;
            if (!string.IsNullOrEmpty(number))
            x = int.Parse(number.Trim());

            if (x < 10) number = "00" + number;
            else if (x < 100) number = "0" + number;

            int n = int.Parse(currentYear.Trim()) - 2000;

            number = "" + n + "." + number;
            Log.Info(LoggerConstants.EXIT);
            return number;
        }

        private static string FormaZdachi(Semestr semestr)
        {
            Log.Info(LoggerConstants.ENTER);
            string s = "";

            if (semestr.CursovaRobota > 0) s = "курсовий проект";
            else if (semestr.DyfZalik > 0) s = "диф залік";
            else if (semestr.Isput > 0) s = "екзамен";
            else if (!string.IsNullOrEmpty(semestr.PracticeFormOfControl)) s = semestr.PracticeFormOfControl;
            else if (semestr.Zalic > 0) s = "залік";
            else if (semestr.StateExamination > 0) s = "протокол";

            Log.Info(LoggerConstants.EXIT);
            return s;
        }

        private static string CreateSheetName(string s)
        {
            Log.Info(LoggerConstants.ENTER);
            string s2 = "";
            foreach (char c in s)
            {
                if (c.Equals('[') || c.Equals(']') || c.Equals('[') || c.Equals('/') || c.Equals('\\') || c.Equals('?') ||
                    c.Equals('*'))
                    continue;
                s2 += c;
            }
            Log.Info(LoggerConstants.EXIT);
            return s2.Length <= 31 ? s2 : s2.Substring(0, 31);
        }

        private static int FromRomeToArab(string rome)
        {
            Log.Info(LoggerConstants.ENTER);
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

            Log.Info(LoggerConstants.EXIT);
            return arab;
        }

        private static string ArabToRome(int arab)
        {
            Log.Info(LoggerConstants.ENTER);
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
            Log.Info(LoggerConstants.EXIT);
            return rome;
        }

        private static string ArabNormalize(string str)
        {
            Log.Info(LoggerConstants.ENTER);
            char[] ch = str.ToCharArray();
            for (int i = 0; i < str.Length; i++)
            {
                var arg = (int)ch[i];
                if (arg == 1030) arg = 73;
                ch[i] = (char)arg;
            }
            Log.Info(LoggerConstants.EXIT);
            return new string(ch);
        }


        // Creating ZvedVidomist

        // Read Oblics Uspisnosti
        public static void CreateVidomist(Group group, int pivricha, string month)
        {
            Log.Info(LoggerConstants.ENTER);
            if (!File.Exists(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti + group.Name + ".xls"))
            {
                Log.Warn("Don`t find any obliks yspishnosti");
            }
            else
            {
                try
                {
                    Excel.Workbook book = App.OpenBook(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti
                        + group.Name + ".xls");
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
                            else if (string.IsNullOrEmpty(month))
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
                    Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
                }
            }
            CreateZvedeniaVidomist(group, pivricha, month);
            Log.Info(LoggerConstants.EXIT);
        }

        // if type == 1 than DufZalicZalic else if == 2 than PracticeOrKP else StateExamen
        private static void ReadOcinkaFromOblics(Group @group, Excel.Worksheet sheet, int pivricha, int type)
        {
            Log.Info(LoggerConstants.ENTER);
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
                Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
            }
            Log.Info(LoggerConstants.EXIT);
        }

        private static void CreateZvedeniaVidomist(Group @group, int pivricha, string mount)
        {
            Excel.Workbook bookTamplate = null;
            Excel.Workbook book = null;

            try
            {
                string stringPivricha = pivricha == 1 ? "1-ше півріччя.xls" : "2-ге півріччя.xls";
                string pathToVidomist = PathsFile.PathsDto.PathToFolderWithZvedeniaVidomistUspishnosti
                    + "Зведена відомість успішності за " + (string.IsNullOrEmpty(mount) ? stringPivricha : mount + ".xls");

                if (!File.Exists(PathsFile.PathsDto.PathToExcelDataForProgram))
                {
                    Log.Error(LoggerConstants.FILE_NOT_EXIST + ": DataToProgram");
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                bookTamplate = App.OpenBook(PathsFile.PathsDto.PathToExcelDataForProgram);
                Excel.Worksheet sheetTamplate = App.OpenWorksheet(bookTamplate, "Зведена відомість");

                if (sheetTamplate == null)
                {
                    Log.Error("DataToProgram must contains sheet with name `Зведена відомість`");
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                Excel.Worksheet sheet;

                if (!File.Exists(pathToVidomist))
                {
                    if (!File.Exists(PathsFile.PathsDto.PathToFileWithMacros))
                    {
                        Log.Warn("Empty Excel file with macros not find");
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }
                    File.Copy(PathsFile.PathsDto.PathToFileWithMacros, pathToVidomist);

                    book = App.OpenBook(pathToVidomist);
                    sheet = App.OpenWorksheet(book, group.Name);
                }
                else
                {
                    book = App.OpenBook(pathToVidomist);
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
                                App.Kill(newApp);
                            }
                            if (Control.ButtonClick == 2)
                            {
                                Log.Info(LoggerConstants.EXIT);
                                return;
                            }

                            sheet = App.OpenWorksheet(book, group.Name);
                            if (sheet != null)
                                sheet.Cells.Delete();
                            else
                            {
                                Log.Warn("Some sheet == null");
                            }
                        }
                        else
                        {
                            if (Control.ButtonClick == 2)
                            {
                                Log.Info(LoggerConstants.EXIT);
                                return;
                            }
                            sheet = App.OpenWorksheet(book, group.Name);
                            if (sheet != null)
                                sheet.Cells.Delete();
                            else
                            {
                                Log.Warn("Some sheet == null");
                            }
                        }
                    }
                    else
                    {
                        sheet = App.CreateNewSheet(book, group.Name);
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

                sheet.Cells[46, "K"].Value = "/ " + group.Curator + " /";

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
                    else if (i == 3 && string.IsNullOrEmpty(mount))
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
                    else if (i == 5 && string.IsNullOrEmpty(mount))
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
                            sheet.Cells[44, c[1].ToString()] = pivricha == 1
                                ? subject.FirstSemestr.CountOfHours : subject.SecondSemestr.CountOfHours;
                            
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

                    if (student.Ocinkas.Count >= 1 && string.IsNullOrEmpty(mount))
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

                    if (string.IsNullOrEmpty(mount))
                        sheet.Cells[row, cBenefics].Value = student.Benefits;
                }

                sheet.Range["C7", c[1].ToString() + 45].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                if (group.Students.Count < 30)
                    sheet.Range["A" + (group.Students.Count + 11), "IV" + 40].Delete(
                        Excel.XlDeleteShiftDirection.xlShiftUp);

                // Add vidomist to arhive
                if (string.IsNullOrEmpty(mount))
                    ArhiveZvedVidomist(sheet, semestrCurrent);
                else
                    sheet.Cells[6, "D"].Value = "Зведена відомість успішності за " + mount;
            }
            catch (Exception e)
            {
                Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
            }
            finally
            {
                App.CloseBook(book, true);
                App.CloseBook(bookTamplate, true);
            }
            Control.IfShow = false;
        }

        private static double ColumnWidth(string s)
        {
            Log.Info(LoggerConstants.ENTER_EXIT);
            if (s.Length <= 21) return 5.57;
            if (s.Length <= 40) return 9.70;
            if (s.Length <= 55) return 11;
            return 13.43;
        }

        private static void ArhiveZvedVidomist(Excel.Worksheet sheet, string semesterRome)
        {
            Log.Info(LoggerConstants.ENTER);
            Excel.Workbook book = null;
            try
            {
                if (string.IsNullOrEmpty(semesterRome))
                {
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                string pathToArhiveFile = PathsFile.PathsDto.PathToArhive + sheet.Name + ".xls";

                bool existSheet = false;

                if (!File.Exists(pathToArhiveFile))
                {
                    File.Copy(PathsFile.PathsDto.PathToFileWithMacros, pathToArhiveFile);
                    existSheet = true;
                }

                book = App.OpenBook(pathToArhiveFile);
                Excel.Worksheet sheetArhive;

                string sheetNameOfArhive = semesterRome + " семестр";

                bool bl = false;
                foreach (Excel.Worksheet worksheet in book.Worksheets)
                {
                    if (worksheet.Name.Equals(sheetNameOfArhive))
                    {
                        bl = true;
                        break;
                    }
                }

                if (bl)
                {
                    sheetArhive = book.Sheets[sheetNameOfArhive];
                    sheetArhive.Cells.Delete();
                }
                else
                {
                    if (existSheet)
                        sheetArhive = book.Sheets[1];
                    else
                    {
                        sheetArhive = App.CreateNewSheet(book, sheetNameOfArhive);
                    }
                }

                sheetArhive.Cells.PasteSpecial(sheet.Cells.Copy());
            }
            catch (Exception e)
            {
                Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
            }
            finally
            {
                App.CloseBook(book, true);
                Log.Info(LoggerConstants.EXIT);
            }
        }

//      Creating atestat

        private static List<NewSubject> GetSubjectsForAtestat()
        {
            Log.Info(LoggerConstants.ENTER);
            Excel.Workbook bookTemplate = null;
            List<NewSubject> subjects = new List<NewSubject>();
            try
            {
                if (!File.Exists(PathsFile.PathsDto.PathToExcelDataForProgram))
                {
                    Log.Error("Excel file DataToProgram not find");
                    Log.Info(LoggerConstants.EXIT);
                    return new List<NewSubject>();
                }

                bookTemplate = App.OpenBook(PathsFile.PathsDto.PathToExcelDataForProgram);
                Excel.Worksheet sheet = App.OpenWorksheet(bookTemplate, "Формування атестату - предмети");

                int startRow = 2;

                while (true)
                {
                    string cellValue = sheet.Cells[startRow, "B"].Value + "";
                    if (string.IsNullOrEmpty(cellValue))
                        break;
                    subjects.Add(new NewSubject
                    {
                        Name = cellValue.Trim()
                    });
                    startRow++;
                }
                if (subjects.Count == 0)
                {
                    Log.Error("In book DataToProgram not wrote any subjects for atestat");
                    Log.Info(LoggerConstants.EXIT);
                    return subjects;
                }

                char startColumn = 'D';
                startRow = 2;

                while (true)
                {
                    string cellsGroupInizial = sheet.Cells[startRow, startColumn.ToString()].Value + "";
                    if (string.IsNullOrEmpty(cellsGroupInizial))
                        break;

                    int startRowForState = 3;

                    while (true)
                    {
                        string cellStateSubject = sheet.Cells[startRowForState, startColumn.ToString()].Value;
                        if (string.IsNullOrEmpty(cellStateSubject))
                            break;
                        foreach (NewSubject newSubject in subjects)
                        {
                            if (newSubject.Name.Equals(cellStateSubject.Trim()))
                            {
                                newSubject.GroupPrefixStatemets.Add(cellsGroupInizial.Trim());
                            }
                        }
                        startRowForState++;
                    }
                    startColumn++;
                }
            }
            catch (Exception e)
            {
                Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
                Log.Info(LoggerConstants.EXIT);
                return subjects;
            }
            finally
            {
                App.CloseBook(bookTemplate, false);
            }
            Log.Info(LoggerConstants.EXIT);
            return subjects;
        }

        public static List<NewSubject> ReadAllNeedSheetsFromArhiveZVtoAtestat(string groupName)
        {
            Log.Info(LoggerConstants.ENTER);
            Excel.Workbook book = null;
            List<NewSubject> subjects = GetSubjectsForAtestat();

            try
            {
                string pathToArhive = PathsFile.PathsDto.PathToArhive + groupName + ".xls";
                if (!File.Exists(pathToArhive))
                {
                    Log.Warn("Dont have vidomostey yspishnosti for group: " + groupName);
                }
                book = App.OpenBook(pathToArhive);

                foreach (Excel.Worksheet sheet in book.Worksheets)
                {
                    int semestr = sheet.Name.Trim().IndexOf(" ") > 0
                        ? FromRomeToArab(sheet.Name.Trim().Substring(0, sheet.Name.Trim().IndexOf(" ")) + "")
                        : 0;
                    if (semestr != 0 && semestr <= 4)
                        ReadOneSheetFromArhiveZVtoAtestat(subjects, sheet, semestr, groupName);
                }
            }
            catch (Exception e)
            {
                Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
                Log.Info(LoggerConstants.EXIT);
                return subjects;
            }
            finally
            {
                App.CloseBook(book, false);
            }

            Log.Info(LoggerConstants.EXIT);
            return subjects;
        }

        private static void ReadOneSheetFromArhiveZVtoAtestat(List<NewSubject> subjects, Excel.Worksheet sheet, int semestr, string groupName)
        {
            Log.Info(LoggerConstants.ENTER);
            try
            {
                char startColumn = 'F';
                byte startRowForStudent = 11;
                byte bt = 0;
                byte countOfStudent = 0;

                while (true)
                {
                    string cellStudentNumber = sheet.Cells[startRowForStudent + countOfStudent, "C"].Value + "";
                    if (string.IsNullOrEmpty(cellStudentNumber))
                        break;
                    countOfStudent++;
                }
                while (true)
                {
                    Excel.Range mergeCells = sheet.Cells[8, startColumn.ToString()];
                    string pas = mergeCells.Value + "";

                    if (string.IsNullOrEmpty(sheet.Cells[9, startColumn.ToString()].Value))
                    {
                        break;
                    }
                    for (int i = 1; i <= mergeCells.MergeArea.Columns.Count; i++)
                    {
                        string subjectName = sheet.Cells[9, startColumn.ToString()].Value + "";
                        if (string.IsNullOrEmpty(subjectName))
                        {
                            startColumn++;
                            break;
                        }

                        string countOfHour = sheet.Cells[countOfStudent + 14, startColumn.ToString()].Value + "";

                        NewSubject newSubjectRef = null;

                        foreach (NewSubject newSubject in subjects)
                        {
                            if (newSubject.Name.Equals(subjectName))
                            {
                                double hour;
                                newSubject.Semestrs.Add(new NewSemestr
                                {
                                    NumberOfSemestr = semestr,
                                    StateExamenExist = pas.Equals("Іспит") && newSubject.GroupExist(groupName),
                                    CountOfHours = double.TryParse(countOfHour, out hour) ? hour : 0
                                });
                                newSubjectRef = newSubject;
                                break;
                            }
                        }

                        if (newSubjectRef == null)
                        {
                            startColumn++;
                            continue;
                        }

                        for (int j = 11; j < 11 + countOfStudent; j++)
                        {
                            newSubjectRef.Semestrs[newSubjectRef.Semestrs.Count - 1].Ocinkas.Add(sheet.Cells[j, startColumn.ToString()].Value + "");
                        }
                        newSubjectRef.Teacher = sheet.Cells[13 + countOfStudent, startColumn.ToString()].Value + "";
                        startColumn++;
                    }
                }
            }
            catch (Exception e)
            {
                Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
            }
            Log.Info(LoggerConstants.EXIT);
        }

        public static void CreateAtestatForOneGroup(string groupName)
        {
            Log.Info(LoggerConstants.ENTER);
            Excel.Workbook
                book = null,
                bookTemplate = null;
            List<NewSubject> subjects = ReadAllNeedSheetsFromArhiveZVtoAtestat(groupName);

            if (subjects.Count == 0) return;

            try
            {
                string pathToAtestat = PathsFile.PathsDto.PathToAtestatFolder + groupName + ".xls";
                string pathToTemplateWithMacros = PathsFile.PathsDto.PathToFileWithMacros;
                string pathToTemplateWithSheet = PathsFile.PathsDto.PathToExcelDataForProgram;

                if (!File.Exists(pathToTemplateWithSheet) || !File.Exists(pathToTemplateWithMacros))
                {
                    MessageBox.Show("Немає потрібних книг\n" + pathToTemplateWithSheet + "\n" + pathToTemplateWithMacros);
                    Log.Warn("Don`t have needed books");
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }
                bool exist = true;
                if (!File.Exists(pathToAtestat))
                {
                    File.Copy(pathToTemplateWithMacros, pathToAtestat);
                    exist = false;
                }

                book = App.OpenBook(pathToAtestat);
                bookTemplate = App.OpenBook(pathToTemplateWithSheet);
                Excel.Worksheet
                    sheet = !exist ? book.Sheets[1] : null,
                    sheetTempPZVY = bookTemplate.Sheets["Підсумкова ЗВУ"],
                    sheetTempPVY = bookTemplate.Sheets["Підсумкова ВУ"];

                foreach (NewSubject newSubject in subjects)
                {
                    string sheetName = CreateSheetName(newSubject.Name);
                    bool sheetEquals = false;

                    foreach (Excel.Worksheet sh in book.Worksheets)
                    {
                        if (CustomEquals(sh.Name, sheetName))
                        {
                            sheetEquals = true;
                            sheet = sh;
                            break;
                        }
                    }
                    if (sheetEquals)
                        sheet.Cells.Delete();
                    else
                    {
                        if (!exist)
                        {
                            sheet.Name = sheetName;
                            exist = true;
                        }
                        else
                        {
                            sheet = book.Worksheets.Add(Type.Missing);
                            sheet.Name = sheetName;
                        }
                    }

                    sheet.Cells.PasteSpecial(sheetTempPVY.Cells.Copy());
                    InsertValuesIntoPVY(sheet, groupName, newSubject);
                }

                // create Pidsumkova Zvedena Vidomist Uspishnosti
                sheet = null;
                foreach (Excel.Worksheet worksheet in book.Worksheets)
                {
                    if (worksheet.Name.Equals("Загальна"))
                    {
                        sheet = worksheet;
                        sheet.Cells.Delete();
                        break;
                    }
                }

                if (sheet == null)
                {
                    sheet = book.Worksheets.Add(Type.Missing);
                    sheet.Name = "Загальна";
                }
                sheet.Cells.PasteSpecial(sheetTempPZVY.Cells.Copy());

                InsertValuesIntoPZVY(sheet, groupName, subjects);
            }
            catch (Exception e)
            {
                Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
            }
            finally
            {
                App.CloseBook(bookTemplate, false);
                App.CloseBook(book, true);
            }
            Log.Info(LoggerConstants.EXIT);
        }

        private static Group GetGroupByName(string groupName)
        {
            foreach (Group @group in Manager.Groups)
            {
                if (group.Name.Equals(groupName))
                    return group;
            }
            return null;
        }

        private static void InsertValuesIntoPVY(Excel.Worksheet sheet, string groupName, NewSubject newSubject)
        {
            Log.Info(LoggerConstants.ENTER);
            Group group = GetGroupByName(groupName);
            if (group == null)
            {
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            byte countOfStudent = 0;
            
            sheet.Cells[6, "B"].Value = "з дисципліни " + newSubject.Name;
            sheet.Cells[7, "B"].Value = "Група " + groupName + " (" + groupName
                .Substring(groupName.IndexOf("-", StringComparison.Ordinal) + 1, groupName.IndexOf("-", StringComparison.Ordinal)) + ")";
            sheet.Cells[8, "B"].Value = "Спеціальність: \"" + group.Speciality + "\"";
            sheet.Cells[9, "B"].Value = "Викладач " + newSubject.Teacher;
            
            foreach (Student student in @group.Students)
            {
                sheet.Cells[12 + countOfStudent, "C"].Value = student.Pib;
                countOfStudent++;
            }
            
            newSubject.Semestrs = newSubject.Semestrs.OrderBy(semestr => semestr.NumberOfSemestr).ToList();

            char pasPosition = 'D';
            byte countOfSemestrWithoutStateExame = 0;
            int columnOfStateExame = -1;
            for (byte i = 0; i < newSubject.Semestrs.Count; i++)
            {
                if (newSubject.Semestrs[i].StateExamenExist)
                {
                    columnOfStateExame = i;
                    continue;
                }
                countOfSemestrWithoutStateExame++;

                sheet.Cells[10, pasPosition.ToString()].Value = newSubject.Semestrs[i].CountOfHours;
                sheet.Cells[11, pasPosition.ToString()].Value = ArabToRome(newSubject.Semestrs[i].NumberOfSemestr) + " семестр Оцінка в балах";

                byte rowForOcinkaStart = 12;

                foreach (string ocinka in newSubject.Semestrs[i].Ocinkas)
                {
                    sheet.Cells[rowForOcinkaStart, pasPosition.ToString()].Value = ocinka;
                    rowForOcinkaStart++;
                }
                pasPosition++;
            }

            if (countOfSemestrWithoutStateExame == 0) return;
            char average = pasPosition;
            average--;

            sheet.Cells[10, pasPosition.ToString()].Formula = "=SUM(D10:" + average + 10 + ")";
            sheet.Cells[11, pasPosition.ToString()].Value = "Підсумкова оцінка";
            for (int i = 0; i < countOfStudent; i++)
            {
                string formula = "0";
                if (countOfSemestrWithoutStateExame == 1)
                    formula = "(D10*D" + (i + 12) + ")/" + pasPosition + "10";
                else if (countOfSemestrWithoutStateExame == 2)
                    formula = "(D10*D" + (i + 12) + "+E10*E" + (i + 12) + ")/" + pasPosition + "10";
                else if (countOfSemestrWithoutStateExame == 3)
                    formula = "(D10*D" + (i + 12) + "+E10*E" + (i + 12)  + "+F10*F" + (i + 12) + ")/" + pasPosition + "10";
                else if (countOfSemestrWithoutStateExame == 4)
                    formula = "(D10*D" + (i + 12) + "+E10*E" + (i + 12) + "+F10*F" + (i + 12) + "+G10*G" + (i + 12) + ")/" + pasPosition + "10";

                sheet.Cells[12 + i, pasPosition.ToString()].Formula = "=ROUND(" + formula + ", 0)";
            }

            // insert stateExamen
            if (columnOfStateExame >= 0)
            {
                pasPosition++;
                sheet.Cells[11, pasPosition.ToString()].Value = "Державна підсумкова атестація";

                for (byte i = 0; i < newSubject.Semestrs[columnOfStateExame].Ocinkas.Count; i++)
                {
                    sheet.Cells[12 + i, pasPosition.ToString()].Value =
                        newSubject.Semestrs[columnOfStateExame].Ocinkas[i];
                }
            }

            // delete range of cells
            if (countOfStudent < 30)
                sheet.Range["A" + (12 + countOfStudent), "IV41"].Delete();

            pasPosition++;
            if (pasPosition < 'J')
                sheet.Range[pasPosition.ToString() + 1, "I65536"].Delete();
            Log.Info(LoggerConstants.EXIT);
        }

        private static void InsertValuesIntoPZVY(Excel.Worksheet sheet, string groupName, List<NewSubject> newSubject)
        {
            Log.Info(LoggerConstants.ENTER);
            Group group = GetGroupByName(groupName);
            if (group == null)
            {
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            sheet.Cells[7, "B"].Value = "Група " + groupName + " (" + groupName
                .Substring(groupName.IndexOf("-", StringComparison.Ordinal) + 1, groupName.IndexOf("-", StringComparison.Ordinal)) + ")";
            sheet.Cells[8, "B"].Value = "Спеціальність: \"" + group.Speciality + "\"";

            byte startRow = 12;
            byte startColumn = 4;

            foreach (Student student in @group.Students)
            {
                sheet.Cells[startRow, "C"].Value = student.Pib;
                startRow++;
            }

            if (startRow < 41)
                sheet.Range[sheet.Cells[startRow, "A"], "IV41"].Delete();

            foreach (NewSubject subject in newSubject)
            {
                if (subject.GroupExist(groupName))
                    sheet.Cells[11, startColumn++].Value = subject.Name + "\n" + "ДА";

                foreach (NewSemestr newSemestr in subject.Semestrs)
                {
                    if (newSemestr.StateExamenExist)
                    {
                        byte startRowForOcinka = 12;
                        foreach (string ocinka in newSemestr.Ocinkas)
                        {
                            sheet.Cells[startRowForOcinka, startColumn - 1].Value = ocinka;
                            startRowForOcinka++;
                        }
                        break;
                    }
                }
            }

            foreach (NewSubject subject in newSubject)
            {
                sheet.Cells[11, startColumn].Value = subject.Name;

                byte startRowForOcinka = 12;
                foreach (string s in subject.GetPidsumkovaOcinka())
                {
                    sheet.Cells[startRowForOcinka, startColumn].Value = s;
                    startRowForOcinka++;
                }
                startColumn++;
            }
            Log.Info(LoggerConstants.EXIT);
        }
    }
}