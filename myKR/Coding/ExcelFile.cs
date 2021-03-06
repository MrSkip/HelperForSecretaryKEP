﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using log4net;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace myKR.Coding
{
    public static class ExcelFile
    {
        private static readonly PathsFile PathsFile = PathsFile.GetPathsFile();

        private static readonly ILog Log =
            LogManager.GetLogger("ExcelFile.cs");

        public static ExcelApplication.ExcelApplication App = ExcelApplication.ExcelApplication.CreateExcelApplication();

        public static void ReadRobPlan(string pathToRobPlan)
        {
            Log.Info(LoggerConstants.ENTER);

            var book = App.OpenBook(pathToRobPlan);
            if (book == null)
            {
                Log.Error("Can`t opet book from path: " + pathToRobPlan);
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            foreach (Worksheet sheet in book.Worksheets)
            {
                if (sheet.Name.Trim().Length == 8 && sheet.Name.Trim().IndexOf('-') == 2 &&
                    sheet.Name.Trim().LastIndexOf('-') == 5)
                {
                    Log.Info("Add group to program with name `" + sheet.Name);
                    Manager.Groups.Add(ReadSheetFromRobPlan(sheet));
                }
            }
            MovePracticeAndStateExam();
            App.CloseBook(book, true);
            Log.Info(LoggerConstants.EXIT);
        }

        public static void ReadStudentsAndOlicAndCurators(string pathToDb)
        {
            Log.Info(LoggerConstants.ENTER);

            var book = App.OpenBook(pathToDb);

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
                var students = ReadStudents((Worksheet) App.LastUsedObject);

                foreach (var group in Manager.Groups)
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
                var oblics
                    = ReadNumbersOfOblic((Worksheet) App.LastUsedObject);

                foreach (var group in Manager.Groups)
                {
                    foreach (
                        var numberOfOblic in oblics.FindAll(oblic => CustomEquals(oblic.Group, @group.Name)))
                    {
                        var find = @group.Subjects.Find(subject => CustomEquals(subject.Name, numberOfOblic.Subject));
                        if (find != null)
                            find.NumberOfOlic = numberOfOblic.Number;

                        var practice =
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
            var list =
                ReadCurator((Worksheet) App.LastUsedObject);
            foreach (var group in Manager.Groups)
            {
                var s = list.Find(strings => CustomEquals(strings[1], @group.Name));
                if (s != null)
                    @group.Curator = s[0];
            }

            App.CloseBook(book, false);
            Log.Info(LoggerConstants.EXIT);
        }

        public static List<string[]> ReadCurator(Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
            var curators = new List<string[]>();
            try
            {
                var n = 1;
                while (true)
                {
                    n++;
                    var value = sheet.Cells[n, "A"].Value;
                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                        break;

                    string curatorName = value.ToString();

                    value = sheet.Cells[n, "B"].Value;

                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
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

        public static List<NumberOfOblic> ReadNumbersOfOblic(Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
            var oblics = new List<NumberOfOblic>();

            var n = 2;
            try
            {
                while (true)
                {
                    n++;
                    var value = sheet.Cells[n, "A"].Value;
                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                        break;
                    string number = value.ToString();

                    value = sheet.Cells[n, "B"].Value;
                    if (value == null || string.IsNullOrWhiteSpace(value))
                        continue;
                    string sujectName = value.ToString();

                    value = sheet.Cells[n, "D"].Value;
                    if (value == null || string.IsNullOrWhiteSpace(value))
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

        public static List<Student> ReadStudents(Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
            var students = new List<Student>();

            var n = 1;
            // if amountOfAvoidStudent appiarance 10 count then stop
            byte amountOfAvoidStudent = 0;

            while (true)
            {
                try
                {
                    n++;
                    var value = sheet.Cells[n, "C"].Value;

                    if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                    {
                        if (amountOfAvoidStudent > 20)
                        {
                            break;
                        }
                        amountOfAvoidStudent++;
                        continue;
                    }

                    amountOfAvoidStudent = 0;

                    var student = new Student
                    {
                        Pib = value.ToString().Trim()
                    };

                    // group name
                    value = sheet.Cells[n, "E"].Value;
                    if (value != null)
                        student.Group = value.ToString();
                    else
                    {
                        continue;
                    }

                    value = sheet.Cells[n, "D"].Value;
                    if (value != null)
                        student.NumberOfBook = value.ToString();

                    value = sheet.Cells[n, "G"].Value;
                    if (value != null)
                        student.FormaTeaching = value.ToString();

                    value = sheet.Cells[n, "L"].Value;
                    if (value != null)
                        student.Benefits = value.ToString();

                    value = sheet.Cells[n, "M"].Value;
                    if (value != null)
                        student.PibChanged = value.ToString();

                    students.Add(student);
                }
                catch (Exception e)
                {
                    Log.Warn("Something wrong", e);
                }
            }

            Log.Info(LoggerConstants.EXIT);
            return students;
        }

        private static Group ReadSheetFromRobPlan(Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);

            var group = new Group();
            group.Name = sheet.Name;

            //Read "Напряму підготовки"
            string s = sheet.Cells[6, "R"].Value;

            var exist = true;

            if (string.IsNullOrWhiteSpace(s) || s.Count(c => c.Equals('"')) < 2)
                exist = false;
            else
            {
                var beginSlash = s.IndexOf("\"", StringComparison.Ordinal);
                var lastSlash = s.LastIndexOf("\"", StringComparison.Ordinal);

                s = s.Substring(beginSlash + 1, lastSlash - beginSlash - 1);

                if (string.IsNullOrWhiteSpace(s))
                    exist = false;
            }

            if (!exist)
            {
                sheet.Cells[6, "R"].Interior.Color =
                    ColorTranslator.ToOle(Color.Red);
                s = "ВВЕДІТЬ НАПРЯМ ПІДГОТОВКИ";
                Log.Error("Expected direction of training in sheet '" + sheet.Name + "'");
            }

            exist = true;

            group.TrainingDirection = s;

            //Read "Спеціальність"
            s = sheet.Cells[7, "R"].Value;

            if (string.IsNullOrWhiteSpace(s) || s.Count(c => c.Equals('"')) != 2)
                exist = false;
            else
            {
                var beginSlash = s.IndexOf("\"", StringComparison.Ordinal);
                var lastSlash = s.LastIndexOf("\"", StringComparison.Ordinal);

                s = s.Substring(beginSlash + 1, lastSlash - beginSlash - 1);

                if (string.IsNullOrWhiteSpace(s))
                    exist = false;
            }

            if (!exist)
            {
                sheet.Cells[7, "R"].Interior.Color =
                    ColorTranslator.ToOle(Color.Red);
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
                        ColorTranslator.ToOle(Color.Red);
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

            if (string.IsNullOrWhiteSpace(s) || !s.Trim().StartsWith("Курс"))
                exist = false;
            else
            {
                var coursePosition = GetPositionForCellCource(s.Trim());

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
                    = ColorTranslator.ToOle(Color.Red);
                Log.Error("Course is incorrect in sheet '" + sheet.Name + "'");
            }

            exist = true;
            group.Course = s;

            //Read "Рік"
            s = sheet.Cells[6, "B"].Value;

            if (string.IsNullOrWhiteSpace(s))
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
                    = ColorTranslator.ToOle(Color.Red);
                Log.Error("Year is incorrect in sheet '" + sheet.Name + "'");
            }

            exist = true;
            group.Year = s;

            //Read "Семестр для першого півріччя"
            s = sheet.Cells[15, "Y"].Value;

            if (string.IsNullOrWhiteSpace(s))
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
                    = ColorTranslator.ToOle(Color.Red);
                Log.Error("Semestr is incorrect in sheet '" + sheet.Name + "'");
            }

            group.FirstRomeSemestr = s;

            group.Subjects = ReadSubject(sheet);
            group.Practice = ReadPractice(sheet);
            group.StateExamination = ReadStateExamination(sheet);

            Log.Info(LoggerConstants.EXIT);
            return group;
        }

        private static List<Subject> ReadSubject(Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
            var subjects = new List<Subject>();
            //[0] - Hours; [1] - Cursova; [2] - Ispyt (Examen) [3] - DyfZalikOrZalic; [4] - DyfZalik (if exist)
            string[]
                firstSemestr = {"Y", "AK", "AO", "AQ", "AR"},
                secondSemestr = {"AS", "BE", "BI", "BK", "BL"};

            var dyfZalikOrNot = false;
            var ifDufZalicIsExist = true;

            // check the cells
            var s = sheet.Cells[15, "C"].Value;
            if (s == null || string.IsNullOrWhiteSpace(s.ToString()) ||
                !s.Trim().ToLower().Equals("назви навчальних  дисциплін"))
            {
                sheet.Cells[15, "C"].Interior.Color =
                    ColorTranslator.ToOle(Color.Red);
                Log.Error("Expekted at the `C15` value 'Назви навчальних  дисциплін'");
                return new List<Subject>();
            }

            // check the cells
            s = sheet.Cells[18, "AR"].Value;
            if (!string.IsNullOrWhiteSpace(s) && s.Trim().ToLower().Equals("диф  залік"))
                ifDufZalicIsExist = false;

            // check the cells
            s = sheet.Cells[18, "AQ"].Value;
            if (!string.IsNullOrWhiteSpace(s) && s.Trim().ToLower().Equals("диф  залік"))
                dyfZalikOrNot = true;

            var n = 14;

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
                    if (subjectName == null || string.IsNullOrWhiteSpace(subjectName.ToString()) || n == 15)
                        continue;
                    if (subjectName.ToString().Trim().ToLower().Equals("разом"))
                        break;

                    string teacher = sheet.Cells[n, "BN"].Value;
                    if (string.IsNullOrWhiteSpace(teacher))
                        teacher = "";

                    var subject = new Subject
                    {
                        Name = RemoveSymbolFromSubjectName(subjectName.ToString().Trim()),
                        Teacher = teacher
                    };
                    var addToList = false;

                    for (var i = 0; i < 2; i++)
                    {
                        var list = i == 0 ? firstSemestr : secondSemestr;
                        var ss = sheet.Cells[n, list[0]].Value;

                        // if cursova robota have same of the pas the not continue
                        var bl = true;
                        var kp = sheet.Cells[n, list[1]].Value;
                        if (kp != null && !string.IsNullOrWhiteSpace(kp.ToString()))
                            bl = false;

                        if ((ss == null || string.IsNullOrWhiteSpace(ss.ToString())) && bl)
                            continue;

                        addToList = true;
                        var semestr = new Semestr();

                        if (ss == null || string.IsNullOrWhiteSpace(ss.ToString()))
                            semestr.CountOfHours = 0;
                        else semestr.CountOfHours = ss;

                        ss = sheet.Cells[n, list[1]].Value;
                        if (ss != null && (!string.IsNullOrWhiteSpace(ss.ToString()) || !ss.ToString().Equals("0")))
                            semestr.CursovaRobota = ss;

                        ss = sheet.Cells[n, list[2]].Value;
                        if (ss != null && (!string.IsNullOrWhiteSpace(ss.ToString()) || !ss.ToString().Equals("0")))
                            semestr.Isput = ss;

                        ss = sheet.Cells[n, list[3]].Value;
                        if (ss != null && (!string.IsNullOrWhiteSpace(ss.ToString()) || !ss.ToString().Equals("0")))
                        {
                            if (ifDufZalicIsExist) semestr.DyfZalik = ss;
                            else semestr.Zalic = ss;
                        }

                        if (!dyfZalikOrNot)
                        {
                            ss = sheet.Cells[n, list[4]].Value;
                            if (ss != null && (!string.IsNullOrWhiteSpace(ss.ToString()) || !ss.ToString().Equals("0")))
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
            if (string.IsNullOrWhiteSpace(subjectName)) return "";
            var withoutSymbol = "";
            foreach (var c in subjectName)
            {
                if (c != '*')
                    withoutSymbol += c;
            }
            Log.Info(LoggerConstants.EXIT);
            return withoutSymbol;
        }

        private static List<Practice> ReadPractice(Worksheet sheet)
        {
            var practices = new List<Practice>();
            string[][] position =
            {
                // Rows - PositionNameOfPractice - Semest - CountOfHours - FormaControlling - Teacher1 - Teacher2
                new[] {"40", "C", "A", "AA", "AE", "AJ", "AS"},
                new[] {"41", "C", "A", "AA", "AE", "AJ", "AS"},
                new[] {"43", "C", "A", "AA", "AE", "AJ", "AS"},
                new[] {"47", "C", "A", "AA", "AE", "AJ", "AS"},
                new[] {"49", "C", "A", "AA", "AE", "AJ", "AS"}
            };
            foreach (var strings in position)
            {
                var value = sheet.Cells[int.Parse(strings[0]), strings[1]].Value;
                if (value != null && value.ToString().Trim().ToLower().Equals("назва практики"))
                {
                    var n = int.Parse(strings[0]);
                    while (true)
                    {
                        n++;
                        value = sheet.Cells[n, strings[1]].Value;

                        if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                            break;

                        if (!value.ToString().Trim().ToLower().Equals("навчальна") &&
                            !value.ToString().Trim().ToLower().Equals("виробнича"))
                        {
                            try
                            {
                                var practice = new Practice();
                                var list = new List<string>();

                                practice.Name = value.ToString().Trim();

                                value = sheet.Cells[n, strings[2]].Value;
                                practice.Semestr = value == null ? "" : value.ToString();

                                value = sheet.Cells[n, strings[3]].Value;
                                if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                                    practice.CountOfHours = 0;
                                else practice.CountOfHours = double.Parse(value.ToString());

                                value = sheet.Cells[n, strings[4]].Value;
                                practice.FormOfControl = value == null ? "" : value.ToString();

                                value = sheet.Cells[n, strings[5]].Value;
                                if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                                    list.Add(value.ToString());

                                value = sheet.Cells[n, strings[6]].Value;
                                if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
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

        private static List<StateExamination> ReadStateExamination(Worksheet sheet)
        {
            Log.Info(LoggerConstants.ENTER);
            var examinations = new List<StateExamination>();
            string[][] position =
            {
                new[] {"49", "BE", "BO"},
                new[] {"40", "BE", "BO"},
                new[] {"40", "AX", "BO"},
                new[] {"47", "BE", "BO"},
                new[] {"41", "BE", "BO"},
                new[] {"43", "BE", "BO"},
                new[] {"38", "AX", "BO"}
            };

            foreach (var strings in position)
            {
                try
                {
                    var value = sheet.Cells[int.Parse(strings[0]), strings[1]].Value;
                    if (value != null && value.ToString().Trim().ToLower().Equals("назва"))
                    {
                        var n = int.Parse(strings[0]);
                        while (true)
                        {
                            n++;
                            value = sheet.Cells[n, strings[2]].Value;
                            if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                                break;
                            var examination = new StateExamination();
                            var nameOfExamen = sheet.Cells[n, strings[1]].Value;
                            if (nameOfExamen != null && !string.IsNullOrWhiteSpace(nameOfExamen.ToString()))
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

        private static int GetPositionForCellCource(string str)
        {
            Log.Info(LoggerConstants.ENTER);
            var x = -1;
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

            var bookCore = App.OpenBook(PathsFile.PathsDto.PathToExcelDataForProgram);
            if (bookCore == null)
            {
                Log.Error("Path to Excel file with all templates are not exist");
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            if (string.IsNullOrWhiteSpace(groupName) && string.IsNullOrWhiteSpace(subjectName))
                foreach (var group in Manager.Groups)
                {
                    foreach (var subject in @group.Subjects)
                    {
                        var semestr = pivricha == 1 ? subject.FirstSemestr : subject.SecondSemestr;
                        if (semestr != null)
                            CreateOblicForOneSubject(bookCore, group, subject.Name, pivricha);
                    }
                }
            else if (string.IsNullOrWhiteSpace(subjectName) && !string.IsNullOrWhiteSpace(groupName))
            {
                var gropu = Manager.Groups.Find(group => group.Name.Equals(groupName));
                if (gropu != null)
                    foreach (var subject in gropu.Subjects)
                    {
                        var semestr = pivricha == 1 ? subject.FirstSemestr : subject.SecondSemestr;
                        if (semestr != null)
                            CreateOblicForOneSubject(bookCore, gropu, subject.Name, pivricha);
                    }
            }
            else if (!string.IsNullOrWhiteSpace(subjectName) && !string.IsNullOrWhiteSpace(groupName))
            {
                var gropu = Manager.Groups.Find(group => group.Name.Equals(groupName));
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
            foreach (var group in Manager.Groups)
            {
                group.FirstRomeSemestr = ArabNormalize(group.FirstRomeSemestr);
                if (group.Practice != null)
                    foreach (var practice in @group.Practice)
                    {
                        practice.Semestr = ArabNormalize(practice.Semestr);

                        var subject = new Subject
                        {
                            Name = practice.Name,
                            NumberOfOlic = practice.NumberOfOlic,
                            Teacher = practice.Teacher.Aggregate("", (current, s) => current + s)
                        };
                        var semestr = new Semestr
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
                    foreach (var stateExamination in @group.StateExamination)
                    {
                        stateExamination.Semestr = ArabNormalize(stateExamination.Semestr);
                        foreach (var subject in @group.Subjects)
                        {
                            if (CustomEquals(subject.Name, stateExamination.Name))
                            {
                                if (subject.FirstSemestr != null &&
                                    group.FirstRomeSemestr.Equals(stateExamination.Semestr))
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

        private static void CreateOblicForOneSubject(Workbook book, Group group, string subjectName, int pivricha)
        {
            Log.Info(LoggerConstants.ENTER);
            Workbook bookOfOblic = null;
            try
            {
                var subjectFind = group.Subjects.Find(subject => subject.Name.Equals(subjectName));
                var nameOfOblic = "";
                if (subjectFind != null)
                {
                    var semestrFindSemestr = pivricha == 1 ? subjectFind.FirstSemestr : subjectFind.SecondSemestr;
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

                Worksheet sheetOfOblic;

                if (bookOfOblic == null)
                {
                    if (!File.Exists(PathsFile.PathsDto.PathToFileWithMacros))
                    {
                        Log.Error(LoggerConstants.FILE_NOT_EXIST + ": " + PathsFile.PathsDto.PathToFileWithMacros);
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }

                    File.Copy(PathsFile.PathsDto.PathToFileWithMacros,
                        PathsFile.PathsDto.PathToFolderWithOblicUspishnosti + group.Name +
                        PathsFile.PathsDto.ExcelExtensial);

                    bookOfOblic =
                        App.OpenBook(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti + group.Name +
                                     PathsFile.PathsDto.ExcelExtensial);

                    sheetOfOblic = App.OpenWorksheet(bookOfOblic, 1);
                    sheetOfOblic.Name = nameOfOblic;
                }
                else
                {
                    var exist = bookOfOblic.Worksheets.Cast<object>()
                        .Any(sheet => ((Worksheet) sheet).Name.Equals(nameOfOblic));
                    if (exist)
                    {
                        if (!Control.IfShow)
                        {
                            var control =
                                new Control("Група [" + group.Name + "]. Уже існує облік успішності для предмету:\n" +
                                            subjectName);
                            control.ShowDialog();
                            if (Control.ButtonClick == 1)
                            {
                                var newApp = new Application
                                {
                                    Visible = true
                                };
                                ((Worksheet)
                                    newApp.Workbooks.Open(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti +
                                                          group.Name
                                                          + PathsFile.PathsDto.ExcelExtensial).Worksheets[nameOfOblic])
                                    .Select();

                                Control.ButtonClick = 0;
                                control.SetButtonReseachEnabled(false);
                                control.ShowDialog();

                                newApp.Quit();
                                ExcelApplication.ExcelApplication.Kill(newApp);
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


                foreach (var subject in @group.Subjects)
                {
                    var semestr = pivricha == 1 ? subject.FirstSemestr : subject.SecondSemestr;
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
                        else if (semestr.CursovaRobota > 0 || !string.IsNullOrWhiteSpace(semestr.PracticeFormOfControl))
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

        private static void CreateKpOrPractice(Worksheet sheetTamplate, Worksheet sheet, Group group, Subject subject,
            Semestr semestr, int pivricha)
        {
            Log.Info(LoggerConstants.ENTER);
            sheet.Cells.PasteSpecial(sheetTamplate.Cells.Copy());

            sheet.Cells[13, "E"].Value = group.TrainingDirection.Equals("Програмна інженерія")
                ? "Програмної інженерії"
                : "Метрології та інформаційно-вимірювальної технології";
            sheet.Cells[15, "F"].Value = group.Speciality;
            sheet.Cells[17, "D"].Value = group.Course;
            sheet.Cells[17, "G"].Value = group.Name;
            sheet.Cells[19, "I"].Value = group.Year + "-" + (int.Parse(group.Year.Trim()) + 1);
            sheet.Cells[26, "F"].Value = subject.Name;
            sheet.Cells[28, "D"].Value = pivricha == 1
                ? group.FirstRomeSemestr
                : ArabToRome(FromRomeToArab(group.FirstRomeSemestr) + 1);
            sheet.Cells[22, "M"].Value = CreateNumberOfOblic(subject.NumberOfOlic,
                pivricha == 1 ? group.Year : int.Parse(@group.Year.Trim()) + 1 + "");
            sheet.Cells[30, "Q"].Value = semestr.CountOfHours;
            sheet.Cells[30, "F"].Value = FormaZdachi(semestr);
//            sheet.Cells[32, "K"].Value = subject.Teacher + "_____";
//            sheet.Cells[100, "N"].Value = subject.Teacher;

            var n = 45;
            foreach (var student in @group.Students)
            {
                sheet.Cells[n, "C"].Value = string.IsNullOrWhiteSpace(student.PibChanged)
                    ? student.Pib
                    : student.PibChanged;

                sheet.Cells[n, "H"].Value = student.NumberOfBook;
                n++;
            }
            if (n != 75)
                sheet.Range["B" + n, "Q" + 74].Delete();
            Log.Info(LoggerConstants.EXIT);
        }

        private static void CreateStateExamen(Worksheet sheetTamplate, Worksheet sheet, Group group, Subject subject,
            Semestr semestr, int pivricha)
        {
            Log.Info(LoggerConstants.ENTER);
            sheet.Cells.PasteSpecial(sheetTamplate.Cells.Copy());
            sheet.Cells[4, "H"].Value = subject.Name;
            sheet.Cells[9, "C"].Value = group.Name;
            sheet.Cells[20, "G"].Value = subject.Teacher + "_________________________________";
            sheet.Cells[84, "H"].Value = subject.Teacher + "__";

            var n = 46;
            foreach (var student in @group.Students)
            {
                sheet.Cells[n, "C"].Value = string.IsNullOrWhiteSpace(student.PibChanged)
                    ? student.Pib
                    : student.PibChanged;

                n++;
            }

            if (n != 76)
                sheet.Range["B" + n, "Q" + 75].Delete();

            // Count of students in group
            sheet.Cells[12, "G"] = "__" + (n - 46) + "__";
            Log.Info(LoggerConstants.EXIT);
        }

        private static void CreateZalicExamenAndDufZalic(Worksheet sheetTamplate, Worksheet sheet,
            Group group, Subject subject, Semestr semestr, int pivricha)
        {
            Log.Info(LoggerConstants.ENTER);
            sheet.Cells.PasteSpecial(sheetTamplate.Cells.Copy());

            sheet.Cells[13, "E"].Value = group.TrainingDirection.Equals("Програмна інженерія")
                ? "Програмної інженерії"
                : "Метрології та інформаційно-вимірювальної технології";
            sheet.Cells[15, "F"].Value = group.Speciality;
            sheet.Cells[17, "D"].Value = group.Course;
            sheet.Cells[17, "G"].Value = group.Name;
            sheet.Cells[19, "I"].Value = group.Year + "-" + (int.Parse(group.Year.Trim()) + 1);
            sheet.Cells[26, "F"].Value = subject.Name;
            sheet.Cells[28, "D"].Value = pivricha == 1
                ? group.FirstRomeSemestr
                : ArabToRome(FromRomeToArab(group.FirstRomeSemestr) + 1);
            sheet.Cells[22, "M"].Value = CreateNumberOfOblic(subject.NumberOfOlic,
                pivricha == 1 ? group.Year : int.Parse(@group.Year.Trim()) + 1 + "");
            sheet.Cells[30, "Q"].Value = semestr.CountOfHours;
            sheet.Cells[30, "F"].Value = FormaZdachi(semestr);
            sheet.Cells[32, "E"].Value = subject.Teacher;
            sheet.Cells[94, "N"].Value = subject.Teacher;

            var n = 39;
            foreach (var student in @group.Students)
            {
                sheet.Cells[n, "C"].Value = string.IsNullOrWhiteSpace(student.PibChanged)
                    ? student.Pib
                    : student.PibChanged;

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
            var x = 0;
            if (!string.IsNullOrWhiteSpace(number))
                x = int.Parse(number.Trim());

            if (x < 10) number = "00" + number;
            else if (x < 100) number = "0" + number;

            var n = int.Parse(currentYear.Trim()) - 2000;

            number = "" + n + "." + number;
            Log.Info(LoggerConstants.EXIT);
            return number;
        }

        private static string FormaZdachi(Semestr semestr)
        {
            Log.Info(LoggerConstants.ENTER);
            var s = "";

            if (semestr.CursovaRobota > 0) s = ConstantExcel.KursovyiProekt;
            else if (semestr.DyfZalik > 0) s = ConstantExcel.DyfZalik;
            else if (semestr.Isput > 0) s = ConstantExcel.Examen;
            else if (!string.IsNullOrWhiteSpace(semestr.PracticeFormOfControl)) s = semestr.PracticeFormOfControl;
            else if (semestr.Zalic > 0) s = ConstantExcel.Zalik;
            else if (semestr.StateExamination > 0) s = ConstantExcel.Protokol;

            Log.Info(LoggerConstants.EXIT);
            return s;
        }

        private static string CreateSheetName(string s)
        {
            Log.Info(LoggerConstants.ENTER);
            var s2 = "";
            foreach (var c in s)
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
            var arab = 0;
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
            var ch = str.ToCharArray();
            for (var i = 0; i < str.Length; i++)
            {
                var arg = (int) ch[i];
                if (arg == 1030) arg = 73;
                ch[i] = (char) arg;
            }
            Log.Info(LoggerConstants.EXIT);
            return new string(ch);
        }


        // Creating ZvedVidomist

        // Read Oblics Uspisnosti
        public static void CreateVidomist(Group group, int pivricha, string month)
        {
            Log.Info(LoggerConstants.ENTER);
            if (
                !File.Exists(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti + group.Name +
                             PathsFile.PathsDto.ExcelExtensial))
            {
                Log.Warn("Don`t find any obliks uspishnosti");
                //TODOO enter path to folder with ObLisk Uspishnosti
            }
            else
            {
                foreach (Student student in @group.Students)
                {
                    student.Ocinkas.Clear();
                }

                try
                {
                    var book = App.OpenBook(PathsFile.PathsDto.PathToFolderWithOblicUspishnosti
                                            + group.Name + PathsFile.PathsDto.ExcelExtensial);
                    foreach (var sheetO in book.Worksheets)
                    {
                        var sheet = (Worksheet) sheetO;

                        var protocol = sheet.Cells[3, "H"].Value;
                        if (protocol == null || string.IsNullOrWhiteSpace(protocol))
                        {
                            var formaZdachi = sheet.Cells[30, "F"].Value;
                            if (formaZdachi == null || string.IsNullOrWhiteSpace(formaZdachi.ToString())) continue;

                            if (formaZdachi.ToString().Equals(ConstantExcel.DyfZalik) ||
                                formaZdachi.ToString().Equals(ConstantExcel.Examen) ||
                                formaZdachi.ToString().Equals(ConstantExcel.Zalik))
                                ReadOcinkaFromOblics(group, sheet, pivricha, 1);
                            else if (string.IsNullOrWhiteSpace(month))
                            {
                                ReadOcinkaFromOblics(group, sheet, pivricha, 2);
                            }
                        }
                        else
                        {
                            ReadOcinkaFromOblics(group, sheet, pivricha, 3);
                        }
                    }
                    App.CloseBook(book, false);
                }
                catch (Exception e)
                {
                    Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
                }
            }
            CreateZvedeniaVidomist(group, pivricha, month);
            Control.IfShow = false;
            Log.Info(LoggerConstants.EXIT);
        }

        // if type == 1 than DufZalicZalic else if == 2 than PracticeOrKP else StateExamen
        private static void ReadOcinkaFromOblics(Group @group, Worksheet sheet, int pivricha, int type)
        {
            Log.Info(LoggerConstants.ENTER);
            try
            {
                string subjectName;

                var currentSemestr = pivricha == 1
                    ? group.FirstRomeSemestr
                    : ArabToRome(FromRomeToArab(ArabNormalize(group.FirstRomeSemestr.Trim())) + 1);

                if (type <= 2)
                {
                    var subjectNameV = sheet.Cells[26, "F"].Value;

                    if (subjectNameV == null || string.IsNullOrWhiteSpace(subjectNameV.ToString()))
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }

                    subjectName = subjectNameV.ToString();
                    var semestr = sheet.Cells[28, "D"].Value;

                    if (semestr == null || string.IsNullOrWhiteSpace(semestr.ToString()))
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }

                    if (!currentSemestr.Equals(semestr.ToString()))
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }
                }
                else
                {
                    var protocol = sheet.Cells[3, "H"].Value;

                    if (protocol == null || string.IsNullOrWhiteSpace(protocol.ToString()))
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }

                    var subjectNameV = sheet.Cells[4, "H"].Value;

                    if (subjectNameV == null || string.IsNullOrWhiteSpace(subjectNameV.ToString()))
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }

                    subjectName = subjectNameV.ToString();
                    var subjectT = group.Subjects.Find(subject1 => subject1.Name.Equals(subjectName));

                    if (subjectT == null)
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }

                    var semestr = pivricha == 1
                        ? subjectT.FirstSemestr
                        : subjectT.SecondSemestr;

                    if (semestr == null)
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }

                    if (semestr.StateExamination <= 0)
                    {
                        Log.Info(LoggerConstants.EXIT);
                        return;
                    }
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

                var ocinkaPositio = type <= 2 ? "L" : "J";
                byte counter = 0;

                while (true)
                {
                    n++;
                    var studentName = sheet.Cells[n, "C"].Value;

                    if (studentName == null || string.IsNullOrWhiteSpace(studentName.ToString()))
                    {
                        break;
                    }

                    var pas = sheet.Cells[n, ocinkaPositio].Value ?? "";

                    Student student = group.Students.Find(student1 => student1.Pib.Equals(studentName.ToString().Trim()));

                    if (student == null)
                    {
                        foreach (Student student1 in @group.Students.Where(student1 => !string.IsNullOrWhiteSpace(student1.PibChanged)
                                                                                       && student1.PibChanged.Equals(studentName.ToString().Trim())))
                        {
                            student = student1;
                            break;
                        }

                        if (counter >= 10)
                        {
                            break;
                        }

                        if (student == null)
                        {
                            counter++;
                            continue;
                        }
                    }

                    counter = 0;

                    student.Ocinkas.Add(new Ocinka
                    {
                        SubjectName = subjectName,
                        StudentName = student.Pib,
                        Mark = pas + ""
                    });
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
            Log.Info(LoggerConstants.ENTER);

            Workbook bookTamplate = null;
            Workbook book = null;

            try
            {
                var stringPivricha = pivricha == 1 ? "1-ше півріччя.xls" : "2-ге півріччя.xls";
                var pathToVidomist = PathsFile.PathsDto.PathToFolderWithZvedeniaVidomistUspishnosti
                                     + "Зведена відомість успішності за " + (string.IsNullOrWhiteSpace(mount)
                                         ? stringPivricha
                                         : mount + PathsFile.PathsDto.ExcelExtensial);

                if (!File.Exists(PathsFile.PathsDto.PathToExcelDataForProgram))
                {
                    Log.Error(LoggerConstants.FILE_NOT_EXIST + ": DataToProgram");
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                bookTamplate = App.OpenBook(PathsFile.PathsDto.PathToExcelDataForProgram);
                var sheetTamplate = App.OpenWorksheet(bookTamplate, "Зведена відомість");

                if (sheetTamplate == null)
                {
                    Log.Error("DataToProgram must contains sheet with name `Зведена відомість`");
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                Worksheet sheet;

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
                    sheet = App.OpenWorksheet(book, 1);

                    sheet.Cells.Delete();
                    sheet.Name = group.Name;
                }
                else
                {
                    book = App.OpenBook(pathToVidomist);
                    var exist =
                        book.Worksheets.Cast<object>()
                            .Any(sheet2 => ((Worksheet) sheet2).Name.Equals(group.Name));
                    if (exist)
                    {
                        if (!Control.IfShow)
                        {
                            var control =
                                new Control("Уже існує зведена відомість для групи:\n" + group.Name);
                            control.ShowDialog();
                            if (Control.ButtonClick == 1)
                            {
                                var newApp = new Application {Visible = true};
                                ((Worksheet)
                                    newApp.Workbooks.Open(pathToVidomist).Worksheets[group.Name]).Select();

                                Control.ButtonClick = 0;
                                control.SetButtonReseachEnabled(false);
                                control.ShowDialog();

                                newApp.Quit();
                                ExcelApplication.ExcelApplication.Kill(newApp);
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

                var semestrCurrent = pivricha == 1
                    ? group.FirstRomeSemestr
                    : ArabToRome(FromRomeToArab(group.FirstRomeSemestr) + 1);

                var yearCurrent = string.IsNullOrWhiteSpace(group.Year)
                    ? 0
                    : int.Parse(group.Year.Trim()) + 1;

                sheet.Cells.PasteSpecial(sheetTamplate.Cells.Copy());

                sheet.Cells[4, "C"].Value = "спеціальності \"" + group.Speciality + "\"";

                sheet.Cells[5, "D"].Value = "групи " + group.Name + " за " + semestrCurrent + " семестр " + group.Year +
                                            "-" +
                                            yearCurrent + " навчального року";

                sheet.Cells[46, "K"].Value = "/ " + group.Curator + " /";

                var subjects = pivricha == 1
                    ? @group.Subjects.FindAll(subject => subject.FirstSemestr != null)
                    : @group.Subjects.FindAll(subject => subject.SecondSemestr != null);

                if (subjects.Count == 0)
                {
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                var subjectList = new Dictionary<string, List<Subject>>();

                foreach (Subject subject in @group.Subjects)
                {
                    subject.Ocinka.Clear();
                }

                List<Ocinka> ocinkas = new List<Ocinka>();
                foreach (Student student in @group.Students)
                {
                    ocinkas.AddRange(student.Ocinkas);
                }

                SortSubject(pivricha, mount, subjects, subjectList, ocinkas, group.Students);

                var count = -1;
                char[] c = { 'F', 'F' };

                if (subjectList.Count(pair => pair.Value.Count > 0) > 0)
                {
                    foreach (var keyValuePair in subjectList.Where(pair => pair.Value.Count != 0))
                    {
                        foreach (var subject in keyValuePair.Value)
                        {
                            count++;
                            sheet.Cells[9, c[1].ToString()] = subject.Name;
                            sheet.Cells[9, c[1].ToString()].ColumnWidth = ColumnWidth(subject.Name);
                            sheet.Cells[43, c[1].ToString()] = subject.Teacher;
                            sheet.Cells[44, c[1].ToString()] = pivricha == 1
                                ? subject.FirstSemestr.CountOfHours
                                : subject.SecondSemestr.CountOfHours;

                            var n = 10;

                            if (subject.Ocinka != null)
                            {
                                foreach (var ocinka in subject.Ocinka)
                                {
                                    n++;
                                    sheet.Cells[n, c[1].ToString()].Value = ocinka.Mark;
                                }
                            }

                            if (group.Students != null && group.Students.Count != 0)
                            {
                                sheet.Cells[41, c[1].ToString()].Value = "=Uspishnist(" + count + "," +
                                                                         group.Students.Count + ")";
                                sheet.Cells[42, c[1].ToString()].Value = "=Quality(" + count + "," +
                                                                         group.Students.Count + ")";
                            }

                            c[1]++;
                        }

                        CaseForMergingSubject(keyValuePair, sheet, c);

                        if (!keyValuePair.Key.Equals(ConstantExcel.Practice))
                        {
                            sheet.Cells[7, "F"].Value = "Предмети";
                            sheet.Range["F" + 7, ((char)(c[1] - (char)1)).ToString() + 7].Merge();
                            sheet.Range["F" + 7, ((char)(c[1] - (char)1)).ToString() + 7].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        }

                        c[0] = c[1];
                    }
                }

                var row = 10;
                sheet.Cells[9, c[1].ToString()].Value = "Середній бал";
                var cBenefics = c[1];
                c[0] = c[1];
                c[0]--;

                cBenefics++;
                cBenefics++;

                foreach (var student in @group.Students)
                {
                    row++;

                    sheet.Cells[row, "D"].Value = string.IsNullOrWhiteSpace(student.PibChanged)
                        ? student.Pib
                        : student.PibChanged;

                    sheet.Cells[row, "E"].Value = student.FormaTeaching;
                    sheet.Cells[row, cBenefics.ToString()].Value = student.Benefits;
                    sheet.Cells[row, c[1].ToString()].Formula = "=AVERAGE(" + "F" + row + ":" + c[0] + row + ") - 0.5";
                    sheet.Cells[row, c[1].ToString()].NumberFormatLocal = "##";

                    var hight = true;
                    var sum = 0;
                    var countOf = 0;

                    if (student.Ocinkas.Count >= 1 && string.IsNullOrWhiteSpace(mount))
                    {
                        foreach (var ocinka in student.Ocinkas)
                        {
                            if (string.IsNullOrWhiteSpace(ocinka.Mark))
                            {
                                countOf++;
                            }

                            int number;

                            if (!int.TryParse(ocinka.Mark, out number))
                            {
                                continue;
                            }

                            if (number < 10)
                            {
                                hight = false;
                            }

                            sum += number;
                        }

                        if (student.FormaTeaching.Equals("п"))
                        {
                            hight = false;
                        }

                        var stupendiaColumnPosution = c[1];
                        stupendiaColumnPosution++;

                        if (group.Students[0].Ocinkas.Count - countOf == 0)
                        {
                            hight = false;
                        }
                        else if (sum/(group.Students[0].Ocinkas.Count - countOf) >= 7 &&
                                 !student.FormaTeaching.Equals("п"))
                        {
                            sheet.Cells[row, stupendiaColumnPosution.ToString()].Value = 1;
                        }

                        if (hight)
                        {
                            sheet.Cells[row, stupendiaColumnPosution.ToString()].Interior.Color =
                                ColorTranslator.ToOle(Color.Yellow);
                        }
                    }

                    if (string.IsNullOrWhiteSpace(mount))
                    {
                        sheet.Cells[row, cBenefics].Value = student.Benefits;
                    }
                }

                sheet.Range["C7", c[1].ToString() + 45].Borders.LineStyle = XlLineStyle.xlContinuous;

                if (group.Students.Count < 30)
                    sheet.Range["A" + (group.Students.Count + 11), "IV" + 40].Delete(
                        XlDeleteShiftDirection.xlShiftUp);

                // Add vidomist to arhive
                if (string.IsNullOrWhiteSpace(mount))
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
        }

        private static void CaseForMergingSubject(KeyValuePair<string, List<Subject>> keyValuePair, Worksheet sheet,
            char[] c)
        {
            switch (keyValuePair.Key)
            {
                case ConstantExcel.Ispyt:
                {
                    sheet.Cells[8, c[0].ToString()].Value = ConstantExcel.Ispyt;
                    sheet.Range[c[0].ToString() + 8, ((char)(c[1] - (char)1)).ToString() + 8].Merge();
                    sheet.Range[c[0].ToString() + 8, ((char)(c[1] - (char)1)).ToString() + 8].HorizontalAlignment =
                        XlHAlign.xlHAlignCenter;
                    break;
                }
                case ConstantExcel.Dz:
                {
                    sheet.Cells[8, c[0].ToString()].Value = ConstantExcel.Dz;
                    sheet.Range[c[0].ToString() + 8, ((char)(c[1] - (char)1)).ToString() + 8].Merge();
                    sheet.Range[c[0].ToString() + 8, ((char)(c[1] - (char)1)).ToString() + 8].HorizontalAlignment =
                        XlHAlign.xlHAlignCenter;
                    break;
                }
                case ConstantExcel.ZalicFirstSymbolIsUpperCase:
                {
                    sheet.Cells[8, c[0].ToString()].Value = ConstantExcel.ZalicFirstSymbolIsUpperCase;
                    sheet.Range[c[0].ToString() + 8, ((char)(c[1] - (char)1)).ToString() + 8].Merge();
                    sheet.Range[c[0].ToString() + 8, ((char)(c[1] - (char)1)).ToString() + 8].HorizontalAlignment =
                        XlHAlign.xlHAlignCenter;
                    break;
                }
            }
        }

        private static void SortSubject(int pivricha, string mount, List<Subject> subjects,
            Dictionary<string, List<Subject>> subjectList, List<Ocinka> ocinkas, List<Student> students)
        {
            try
            {
                List<Subject> tempSubjectList = subjects.FindAll(
                    subject =>
                        pivricha == 1
                            ? subject.FirstSemestr.Isput > 0 || subject.FirstSemestr.StateExamination > 0
                            : subject.SecondSemestr.Isput > 0 || subject.SecondSemestr.StateExamination > 0);

                SetSubjectsNameToSubjectStructure(tempSubjectList, ocinkas, students);
            subjectList.Add(ConstantExcel.Ispyt, tempSubjectList);

            tempSubjectList = subjects.FindAll(
                subject =>
                    pivricha == 1
                        ? subject.FirstSemestr.DyfZalik > 0
                        : subject.SecondSemestr.DyfZalik > 0);
            SetSubjectsNameToSubjectStructure(tempSubjectList, ocinkas, students);
            subjectList.Add(ConstantExcel.Dz, tempSubjectList);

            if (string.IsNullOrWhiteSpace(mount))
            {
                tempSubjectList = subjects.FindAll(
                    subject =>
                        pivricha == 1
                            ? subject.FirstSemestr.CursovaRobota > 0
                            : subject.SecondSemestr.CursovaRobota > 0);
                SetSubjectsNameToSubjectStructure(tempSubjectList, ocinkas, students);
                subjectList.Add(ConstantExcel.KursovyiProekt, tempSubjectList);
            }

            tempSubjectList = subjects.FindAll(
                subject =>
                    pivricha == 1
                        ? subject.FirstSemestr.Zalic > 0
                        : subject.SecondSemestr.Zalic > 0);
            SetSubjectsNameToSubjectStructure(tempSubjectList, ocinkas, students);
            subjectList.Add(ConstantExcel.ZalicFirstSymbolIsUpperCase, tempSubjectList);

            if (string.IsNullOrWhiteSpace(mount))
            {
                tempSubjectList = subjects.FindAll(
                    subject =>
                        pivricha == 1
                            ? !string.IsNullOrWhiteSpace(subject.FirstSemestr.PracticeFormOfControl)
                            : !string.IsNullOrWhiteSpace(subject.SecondSemestr.PracticeFormOfControl));
                SetSubjectsNameToSubjectStructure(tempSubjectList, ocinkas, students);
                subjectList.Add(ConstantExcel.Practice, tempSubjectList);
            }
            }
            catch (Exception e)
            {
                Log.Error(LoggerConstants.SOMETHING_WRONG, e);
            }
        }

        private static void SetSubjectsNameToSubjectStructure(List<Subject> tempSubjectList, List<Ocinka> ocinkas, List<Student> students)
        {
            if (tempSubjectList.Count <= 0 || ocinkas.Count == 0) return;

            foreach (Subject subject in tempSubjectList)
            {
                var subject1 = subject;

                List<Ocinka> ocinkasForStud = ocinkas.FindAll(ocinka1 => ocinka1.SubjectName.Equals(subject1.Name));

                if (ocinkasForStud.Count == 0) continue;

                if (ocinkasForStud.Count != students.Count)
                {

                    for (int index = 0; index < students.Count; index++)
                    {
                        Student student = students[index];

                        if (!student.Pib.Equals(ocinkasForStud[index].StudentName))
                            ocinkasForStud.Insert(index, new Ocinka());
                    }
                }

                subject.Ocinka.AddRange(ocinkasForStud);
            }
        }

        private static double ColumnWidth(string s)
        {
            Log.Info(LoggerConstants.ENTER_EXIT);
            if (s.Length <= 21) return 5.57;
            if (s.Length <= 40) return 9.70;
            if (s.Length <= 55) return 11;
            return 13.43;
        }

        private static void ArhiveZvedVidomist(Worksheet sheet, string semesterRome)
        {
            Log.Info(LoggerConstants.ENTER);
            Workbook book = null;
            try
            {
                if (string.IsNullOrWhiteSpace(semesterRome))
                {
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                var pathToArhiveFile = PathsFile.PathsDto.PathToArhive + sheet.Name + PathsFile.PathsDto.ExcelExtensial;

                var existSheet = false;

                if (!File.Exists(pathToArhiveFile))
                {
                    File.Copy(PathsFile.PathsDto.PathToFileWithMacros, pathToArhiveFile);
                    existSheet = true;
                }

                book = App.OpenBook(pathToArhiveFile);
                Worksheet sheetArhive;

                var sheetNameOfArhive = semesterRome + " семестр";

                var sheetExist = book.Worksheets.Cast<Worksheet>().Any(worksheet => worksheet.Name.Equals(sheetNameOfArhive));

                if (sheetExist)
                {
                    sheetArhive = book.Sheets[sheetNameOfArhive];
                    sheetArhive.Cells.Delete();
                }
                else
                {
                    sheetArhive = existSheet ? book.Sheets[1] : App.CreateNewSheet(book, sheetNameOfArhive);
                    sheetArhive.Name = sheetNameOfArhive;
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

        private static List<SubjectForAtestat> ReadAllSubjectsForAtestat()
        {
            Log.Info(LoggerConstants.ENTER);
            Workbook bookTemplate = null;
            var subjects = new List<SubjectForAtestat>();
            try
            {
                if (!File.Exists(PathsFile.PathsDto.PathToExcelDataForProgram))
                {
                    Log.Error("Excel file DataToProgram does`t find");
                    Log.Info(LoggerConstants.EXIT);
                    return new List<SubjectForAtestat>();
                }

                bookTemplate = App.OpenBook(PathsFile.PathsDto.PathToExcelDataForProgram);
                var sheet = App.OpenWorksheet(bookTemplate, "Формування атестату - предмети");

                var startRow = 2;

                while (true)
                {
                    string cellValue = sheet.Cells[startRow, "B"].Value + "";

                    if (string.IsNullOrWhiteSpace(cellValue))
                        break;

                    subjects.Add(new SubjectForAtestat
                    {
                        SubjectName = cellValue.Trim()
                    });
                    startRow++;
                }

                if (subjects.Count == 0)
                {
                    Log.Error("In book DataToProgram (Формування атестату - предмети) without subjects for atestat");
                    Log.Info(LoggerConstants.EXIT);
                    return subjects;
                }

                var startColumn = 'D';
                startRow = 2;
                byte counter = 0;

                while (true)
                {
                    string cellsGroupInizial = sheet.Cells[startRow, startColumn.ToString()].Value + "";

                    if (string.IsNullOrWhiteSpace(cellsGroupInizial))
                    {
                        counter++;

                        if (counter > 5)
                            break;

                        continue;
                    }

                    counter = 0;

                    var subjectBeginRow = 3;

                    while (true)
                    {
                        string cellStateSubject = sheet.Cells[subjectBeginRow, startColumn.ToString()].Value;

                        if (string.IsNullOrWhiteSpace(cellStateSubject))
                            break;

                        SubjectForAtestat foundSubjectForAtestat = subjects.Find(subject => subject.SubjectName.ToLower(CultureInfo.CurrentCulture).Trim()
                            .Equals(cellStateSubject.ToLower(CultureInfo.CurrentCulture).Trim()));

                        if (foundSubjectForAtestat != null)
                            foundSubjectForAtestat.GroupPrefixForExam.Add(cellsGroupInizial.Trim());
                        else
                        {
                            sheet.Cells[subjectBeginRow, startColumn.ToString()].Interior.Color =
                                ColorTranslator.ToOle(Color.Red);
                            Log.Error(LoggerConstants.SOMETHING_WRONG + ". Path [" + PathsFile.PathsDto.PathToExcelDataForProgram + "]"
                                + "[Формування атестату - предмети]" + "[column - " + startColumn + ", row - " + subjectBeginRow + "]");
                        }

                        subjectBeginRow++;
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

        public static List<SubjectForAtestat> AnalyseAllSheetsInArhiveZvForOneGroup(string groupName)
        {
            Log.Info(LoggerConstants.ENTER);
            Workbook book = null;
            var subjects = ReadAllSubjectsForAtestat();

            if (subjects == null || subjects.Count == 0)
            {
                Log.Info(LoggerConstants.EXIT);
                return null;
            }

            try
            {
                var pathToArhive = PathsFile.PathsDto.PathToArhive + groupName + PathsFile.PathsDto.ExcelExtensial;

                if (!File.Exists(pathToArhive))
                {
                    Log.Warn("Dont have vidomostey yspishnosti for group: " + groupName);
                    return null;
                }

                book = App.OpenBook(pathToArhive);

                foreach (Worksheet sheet in book.Worksheets)
                {
                    var semestr = sheet.Name.Trim().Contains(" ")
                        ? FromRomeToArab(sheet.Name.Trim().Split(' ')[0])
                        : 0;

                    if (semestr != 0 && semestr <= 4)
                        AnalyseOneSheetInArhiveZv(subjects, sheet, semestr, groupName);
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

        private static void AnalyseOneSheetInArhiveZv(List<SubjectForAtestat> subjects, Worksheet sheet, int semestr,
            string groupName)
        {
            Log.Info(LoggerConstants.ENTER);
            try
            {
                var startColumn = 'F';
                const string studentNamePosition = "D";
                byte studentRow = 11;
                byte countOfStudent = 0;

                while (true)
                {
                    string cellStudentNumber = sheet.Cells[studentRow + countOfStudent, "C"].Value + "";

                    if (string.IsNullOrWhiteSpace(cellStudentNumber))
                        break;

                    countOfStudent++;
                }

                while (true)
                {
                    Range mergeCells = sheet.Cells[8, startColumn.ToString()];
                    string subject = sheet.Cells[11, startColumn.ToString()].Value + "";

                    if (string.IsNullOrWhiteSpace(subject)
                        || subject.Trim().Equals("Середній бал"))
                        break;

                    string pas = mergeCells.Value + "";

                    for (var i = 0; i <= mergeCells.MergeArea.Columns.Count; i++)
                    {
                        string subjectName = sheet.Cells[9, startColumn.ToString()].Value + "";

                        if (string.IsNullOrWhiteSpace(subjectName))
                        {
                            startColumn++;
                            break;
                        }

                        string countOfHour = sheet.Cells[countOfStudent + 14, startColumn.ToString()].Value + "";

                        SubjectForAtestat subjectForAtestatRef = subjects.Find(newSubject =>
                            newSubject.SubjectName.ToLower(CultureInfo.CurrentCulture).Trim()
                                .Equals(subjectName.ToLower(CultureInfo.CurrentCulture).Trim()));

                        if (subjectForAtestatRef == null)
                        {
                            startColumn++;
                            continue;
                        }

                        double hour;

                        subjectForAtestatRef.Semestrs.Add(new SemestrForAtestat
                        {
                            Semestr = semestr,
                            StateExamenExist = pas.Equals(ConstantExcel.Ispyt) && subjectForAtestatRef.GroupExist(groupName.Trim()),
                            CountOfHours = double.TryParse(countOfHour, out hour) ? hour : 0
                        });
                        subjectForAtestatRef.Teacher = sheet.Cells[13 + countOfStudent, startColumn.ToString()].Value + "";


                        for (var j = 11; j < 11 + countOfStudent; j++)
                        {
                            string mark = sheet.Cells[j, startColumn.ToString()].Value + "";
                            string studentName = sheet.Cells[j, studentNamePosition].Value + "";

                            subjectForAtestatRef.Semestrs[subjectForAtestatRef.Semestrs.Count - 1].Marks
                                .Add(new RecordStudmark
                                {
                                    Mark = mark.Trim(),
                                    StudentName = studentName.Trim()
                                });
                        }

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
            Workbook
                book = null,
                bookTemplate = null;
            var subjects = AnalyseAllSheetsInArhiveZvForOneGroup(groupName);

            if (subjects == null || subjects.Count == 0)
            {
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            try
            {
                var pathToAtestat = PathsFile.PathsDto.PathToAtestatFolder + groupName +
                                    PathsFile.PathsDto.ExcelExtensial;
                var pathToTemplateWithMacros = PathsFile.PathsDto.PathToFileWithMacros;
                var pathToTemplateWithSheet = PathsFile.PathsDto.PathToExcelDataForProgram;

                if (!File.Exists(pathToTemplateWithSheet) || !File.Exists(pathToTemplateWithMacros))
                {
                    MessageBox.Show("Немає потрібних книг\n" + pathToTemplateWithSheet + "\n" + pathToTemplateWithMacros);
                    Log.Warn("Don`t have needed books");
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                var exist = false;

                if (!File.Exists(pathToAtestat))
                {
                    File.Copy(pathToTemplateWithMacros, pathToAtestat);
                    exist = true;
                }

                book = App.OpenBook(pathToAtestat);
                bookTemplate = App.OpenBook(pathToTemplateWithSheet);

                Worksheet
                    sheet = exist ? book.Sheets[1] : null,
                    sheetTempPzvy = App.OpenWorksheet(bookTemplate, "Підсумкова ЗВУ"),
                    sheetTempPvy = App.OpenWorksheet(bookTemplate, "Підсумкова ВУ");

                if (sheetTempPvy == null || sheetTempPzvy == null)
                {
                    Log.Error(LoggerConstants.FILE_NOT_EXIST + ": `Підсумкова ЗВУ` or `Підсумкова ВУ`");
                    Log.Info(LoggerConstants.EXIT);
                    return;
                }

                foreach (var newSubject in subjects)
                {
                    var sheetName = CreateSheetName(newSubject.SubjectName);
                    var sheetEquals = false;

                    foreach (Worksheet sh in book.Worksheets.Cast<Worksheet>().Where(sh => CustomEquals(sh.Name, sheetName)))
                    {
                        sheetEquals = true;
                        sheet = sh;
                        break;
                    }

                    if (sheetEquals)
                        sheet.Cells.Delete();
                    else
                    {
                        if (exist)
                        {
                            sheet.Name = sheetName;
                            exist = false;
                        }
                        else
                        {
                            sheet = book.Worksheets.Add(Type.Missing);
                            sheet.Name = sheetName;
                        }
                    }

                    sheet.Cells.PasteSpecial(sheetTempPvy.Cells.Copy());
                    CreatePvyForOneSubject(sheet, groupName, newSubject);
                }

                // create Pidsumkova Zvedena Vidomist Uspishnosti
                sheet = null;
                foreach (Worksheet worksheet in book.Worksheets.Cast<Worksheet>().Where(worksheet => worksheet.Name.Equals("Загальна")))
                {
                    sheet = worksheet;
                    sheet.Cells.Delete();
                    break;
                }

                if (sheet == null)
                {
                    sheet = book.Worksheets.Add(Type.Missing);
                    sheet.Name = "Загальна";
                }

                sheet.Cells.PasteSpecial(sheetTempPzvy.Cells.Copy());
                InsertValuesIntoPzvy(sheet, groupName, subjects, book);
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
            return Manager.Groups.FirstOrDefault(@group => @group.Name.Equals(groupName));
        }

        private static bool ComparePibs(string pib1, string pib2)
        {
            if (string.IsNullOrWhiteSpace(pib1) || string.IsNullOrWhiteSpace(pib2))
                return false;

            pib1 = pib1.Trim().Replace(".", " ").Replace("  ", " ").Trim();
            pib2 = pib2.Trim().Replace(".", " ").Replace("  ", " ").Trim();

            if (pib2.Split(' ').Length != 3 || pib1.Split(' ').Length != 3)
                return false;

            string[] s1 = pib1.Split(' ');
            string[] s2 = pib2.Split(' ');

            string name1 = s1[0] + s1[1][0] + s1[2][0];
            string name2 = s2[0] + s2[1][0] + s2[2][0];

            return name2.Equals(name1);
        }

        private static void CreatePvyForOneSubject(Worksheet sheet, string groupName, SubjectForAtestat subjectForAtestat)
        {
            Log.Info(LoggerConstants.ENTER);
            var group = GetGroupByName(groupName);

            if (group == null)
            {
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            byte countOfStudent = 0;

            sheet.Cells[6, "B"].Value = "з дисципліни " + subjectForAtestat.SubjectName;
            var groupSpexific = "Група " + groupName;

            if (groupName.Split('-').Length == 3)
                groupSpexific += "(" + groupName.Split('-')[1] + ")";

            sheet.Cells[7, "B"].Value = groupSpexific;
            sheet.Cells[8, "B"].Value = "Спеціальність: \"" + group.Speciality + "\"";
            sheet.Cells[9, "B"].Value = "Викладач " + subjectForAtestat.Teacher;

            foreach (var student in @group.Students)
            {
                sheet.Cells[12 + countOfStudent, "C"].Value = student.GetPib();
                countOfStudent++;
            }

            subjectForAtestat.Semestrs = subjectForAtestat.Semestrs.OrderBy(semestr => semestr.Semestr).ToList();

            var markPositionColumn = 'D';
            byte countOfSemestrWithoutStateExame = 0;
            var columnOfStateExame = -1;

            for (byte i = 0; i < subjectForAtestat.Semestrs.Count; i++)
            {
                if (subjectForAtestat.Semestrs[i].StateExamenExist)
                {
                    columnOfStateExame = i;
                    continue;
                }

                countOfSemestrWithoutStateExame++;

                sheet.Cells[10, markPositionColumn.ToString()].Value = subjectForAtestat.Semestrs[i].CountOfHours;
                sheet.Cells[11, markPositionColumn.ToString()].Value = ArabToRome(subjectForAtestat.Semestrs[i].Semestr) +
                                                                " семестр Оцінка в балах";

                byte markPositionRow = 12;

                for (byte i2 = 1; i2 <= countOfStudent; i2++)
                {
                    var studentName = sheet.Cells[markPositionRow + i2 - 1, "C"].Value + "";

                    if (string.IsNullOrWhiteSpace(studentName))
                        continue;

                    RecordStudmark markWithStudName;

                    if (GetMarkByStudentName(subjectForAtestat, i, studentName, @group, out markWithStudName))
                        continue;

                    sheet.Cells[markPositionRow + i2 - 1, markPositionColumn.ToString()].Value = markWithStudName.Mark;
                }

                markPositionColumn++;
            }

            if (countOfSemestrWithoutStateExame == 0)
            {
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            var average = markPositionColumn;
            average--;

            sheet.Cells[10, markPositionColumn.ToString()].Formula = "=SUM(D10:" + average + 10 + ")";
            sheet.Cells[11, markPositionColumn.ToString()].Value = "Підсумкова оцінка";

            for (var i = 0; i < countOfStudent; i++)
            {
                var formula = "0";
                switch (countOfSemestrWithoutStateExame)
                {
                    case 1:
                        formula = "(D10*D" + (i + 12) + ")/" + markPositionColumn + "10";
                        break;
                    case 2:
                        formula = "(D10*D" + (i + 12) + "+E10*E" + (i + 12) + ")/" + markPositionColumn + "10";
                        break;
                    case 3:
                        formula = "(D10*D" + (i + 12) + "+E10*E" + (i + 12) + "+F10*F" + (i + 12) + ")/" + markPositionColumn + "10";
                        break;
                    case 4:
                        formula = "(D10*D" + (i + 12) + "+E10*E" + (i + 12) + "+F10*F" + (i + 12) + "+G10*G" + (i + 12) +
                                  ")/" + markPositionColumn + "10";
                        break;
                }

                string cellValue = sheet.Cells[12 + i, (char)(markPositionColumn - (char)1) + ""].Value + "";
                double tryPasre;

                sheet.Cells[12 + i, markPositionColumn.ToString()].Value =
                    double.TryParse(cellValue.Trim(), out tryPasre)
                        ? "=ROUND(" + formula + ", 0)"
                        : cellValue;
            }

            // insert stateExamen
            if (columnOfStateExame >= 0)
            {
                markPositionColumn++;
                sheet.Cells[11, markPositionColumn.ToString()].Value = "Державна підсумкова атестація";

                for (byte i = 0; i < subjectForAtestat.Semestrs[columnOfStateExame].Marks.Count; i++)
                {
                    var studentName = sheet.Cells[12 + i, "C"].Value + "";

                    if (string.IsNullOrWhiteSpace(studentName))
                        break;


                    RecordStudmark markWithStudName;

                    if (GetMarkByStudentName(subjectForAtestat, columnOfStateExame, studentName, @group, out markWithStudName))
                        continue;

                    sheet.Cells[12 + i, markPositionColumn.ToString()].Value = markWithStudName.Mark;
                }
            }

            // delete range of cells
            if (countOfStudent < 30)
                sheet.Range["A" + (12 + countOfStudent), "IV41"].Delete();

            markPositionColumn++;

            if (markPositionColumn < 'J')
                sheet.Range[markPositionColumn.ToString() + 1, "I65536"].Delete();

            Log.Info(LoggerConstants.EXIT);
        }

        private static bool GetMarkByStudentName(SubjectForAtestat subjectForAtestat, int semestr, string studentName, Group @group,
            out RecordStudmark markWithStudName)
        {
            markWithStudName = subjectForAtestat.Semestrs[semestr].Marks.Find(studmark =>
                ComparePibs(studentName, studmark.StudentName));

            if (markWithStudName != null) return false;

            Student student = @group.Students.Find(student1 =>
                    ComparePibs(student1.Pib, studentName) || ComparePibs(student1.PibChanged, studentName));

            if (student == null)
                return true;

            markWithStudName = subjectForAtestat.Semestrs[semestr].Marks.Find(studmark =>
                ComparePibs(studmark.StudentName, student.Pib)
                || ComparePibs(studmark.StudentName, student.PibChanged));

            return markWithStudName == null;
        }

        public static Dictionary<string, List<string>> GetMarksForSubjectList(List<Student> students, 
            Workbook book, List<SubjectForAtestat> subjects)
        {
            Dictionary<string, List<string>> marks = new Dictionary<string, List<string>>();

            if (subjects.Count == 0)
                return marks;

            const int markRowPosition = 12;
            const int nameRowPosition = 11;

            foreach (SubjectForAtestat subject in subjects)
            {
                string sheetName = CreateSheetName(subject.SubjectName.Trim());
                Worksheet sheet = App.OpenWorksheet(book, sheetName);
                marks.Add(sheetName, new List<string>());

                if (sheet == null)
                    continue;

                int nameColumnPosition = 3;

                for (int i = 0; i < 6; i++)
                {
                    string cellValue = sheet.Cells[nameRowPosition, ++nameColumnPosition].Value + "";

                    if (string.IsNullOrWhiteSpace(cellValue))
                        continue;

                    if (cellValue.Trim().Equals(ConstantExcel.SumMarkForAtestat))
                        break;
                }

                if (nameColumnPosition == 3) continue;

                for (var i = 0; i < students.Count; i++)
                {
                    string cellValue = sheet.Cells[markRowPosition + i, nameColumnPosition].Value + "";
                    marks[sheetName].Add(cellValue);
                }
            }

            return marks;
        }

        private static void InsertValuesIntoPzvy(Worksheet sheet, string groupName, List<SubjectForAtestat> subjects, Workbook book)
        {
            Log.Info(LoggerConstants.ENTER);
            var group = GetGroupByName(groupName);

            if (group == null)
            {
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            var groupSpexific = "Група " + groupName;

            if (groupName.Split('-').Length == 2)
                groupSpexific += " (" + groupName.Split('-')[1] + ")";

            sheet.Cells[7, "B"].Value = groupSpexific;
            sheet.Cells[8, "B"].Value = "Спеціальність: \"" + group.Speciality + "\"";

            byte startRow = 12;
            byte startColumn = 4;

            foreach (var student in @group.Students)
            {
                sheet.Cells[startRow, "C"].Value = student.GetPib();
                startRow++;
            }

            if (startRow < 41)
                sheet.Range[sheet.Cells[startRow, "A"], "IV41"].Delete();

            foreach (var subject in subjects)
            {
                if (subject.GroupExist(groupName))
                    sheet.Cells[11, startColumn++].Value = subject.SubjectName + "\n" + "ДА";

                SemestrForAtestat subjectForAtestat = subject.Semestrs.Find(atestat1 => atestat1.StateExamenExist);

                if (subjectForAtestat == null)
                    continue;

                for (int i = 0; i < group.Students.Count; i++)
                {
                    var studentName = sheet.Cells[12 + i, "C"].Value + "";

                    if (string.IsNullOrWhiteSpace(studentName))
                        break;

                    RecordStudmark markWithStudName;

                    int nSemestr = subject.Semestrs.IndexOf(subjectForAtestat);

                    if (GetMarkByStudentName(subject, nSemestr, studentName, @group, out markWithStudName))
                        continue;

                    sheet.Cells[12 + i, startColumn - 1].Value = markWithStudName.Mark;
                }
            }

            Dictionary<string, List<string>> marks = GetMarksForSubjectList(group.Students, book, subjects);

            if (marks.Count == 0)
            {
                Log.Info(LoggerConstants.EXIT);
                return;
            }

            foreach (var subject in subjects)
            {
                sheet.Cells[11, startColumn].Value = subject.SubjectName;

                List<string> curMarks = marks[subject.SubjectName];

                if (curMarks == null || curMarks.Count == 0)
                {
                    startColumn++;
                    continue;
                }

                byte startRowForOcinka = 12;

                foreach (string mark in curMarks)
                {
                    sheet.Cells[startRowForOcinka++, startColumn].Value = mark;
                }

                startColumn++;
            }

            if (startColumn < 31)
                sheet.Range[sheet.Cells[1, startColumn], "AD" + 74].Delete();

            Log.Info(LoggerConstants.EXIT);
        }

        public static List<Record> GetPidsumkovaOcinka(List<SemestrForAtestat> semestrs, Group group)
        {
            CalculationsDto sum = new CalculationsDto();

            List<SemestrForAtestat> list =
                semestrs.Where(newSemestr => newSemestr.Marks.Count > 0 && !newSemestr.StateExamenExist).ToList();

            foreach (SemestrForAtestat atestat in list)
                sum.CountOfHour += atestat.CountOfHours;

            if (sum.CountOfHour <= 0)
                return new List<Record>();

            foreach (Student student in @group.Students)
            {
                sum.Studmarks.Add(new Record
                {
                    StudentName = student.Pib,
                    StudentNameChanged = student.PibChanged
                });
            }

            foreach (SemestrForAtestat semestr in list)
            {
                foreach (RecordStudmark studMark in semestr.Marks)
                {
                    Record studentExist = sum.Studmarks.Find(record => ComparePibs(record.StudentName, studMark.StudentName)) ??
                                          sum.Studmarks.Find(record => ComparePibs(record.StudentNameChanged, studMark.StudentName));

                    if (studentExist == null)
                        continue;

                    double mark;
                    mark = double.TryParse(studMark.Mark.Trim(), out mark) ? mark : 0;

                    if (mark <= 0)
                    {
                        studentExist.IfMarkCanParse = false;
                        studentExist.Mark = studMark.Mark;
                        continue;
                    }

                    if (studentExist.IfMarkCanParse)
                        continue;

                    studentExist.IfMarkCanParse = true;

                    if (string.IsNullOrWhiteSpace(studentExist.Mark))
                        studentExist.Mark = semestr.CountOfHours*mark + "";
                    else
                    {
                        double mark2;
                        mark2 = double.TryParse(studMark.Mark.Trim(), out mark2) ? mark2 : 0;

                        studentExist.Mark = mark2 <= 0 ? mark2 + "" : mark2 + semestr.CountOfHours*mark + "";
                    }
                }
            }

            foreach (Record studmark in sum.Studmarks)
            {
                if (studmark.IfMarkCanParse)
                {
                    studmark.Mark = Math.Round(double.Parse(studmark.Mark) / sum.CountOfHour, 0) + "";
                }
            }

            return sum.Studmarks;
        }

        public static void KillMainEecelApp()
        {
            App.CloseApp();
        }
    }
}