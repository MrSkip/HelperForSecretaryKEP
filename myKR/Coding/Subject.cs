﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace myKR.Coding
{
    public class Group
    {
        public string Name;
        public string TrainingDirection;
        public string Speciality;
        public string CodeOfSpeciality;
        public string Curator;
        public string Course;
        public string Year;
        public string FirstRomeSemestr;
        public List<Subject> Subjects;
        public List<Practice> Practice;
        public List<StateExamination> StateExamination;
        public List<Student> Students;
    }

    public class Subject
    {
        public string Name;
        public string Teacher;
        public string NumberOfOlic;
        public Semestr
            FirstSemestr,
            SecondSemestr;

        public List<Ocinka> Ocinka = new List<Ocinka>();
    }

    public class SubjectForAtestat
    {
        public string SubjectName;
        public string Teacher;

        public List<SemestrForAtestat> Semestrs = new List<SemestrForAtestat>();
        public List<string> GroupPrefixForExam = new List<string>();
        private List<string> _pidsumkovaOcinka = new List<string>();

        public List<string> GetPidsumkovaOcinka()
        {
            double countOfHour = 0;
            int indexOfLastExistSubject = -1;

            foreach (SemestrForAtestat newSemestr in Semestrs)
            {
                if (newSemestr.Marks.Count > 0 && !newSemestr.StateExamenExist)
                {
                    indexOfLastExistSubject = Semestrs.IndexOf(newSemestr);
                    countOfHour += newSemestr.CountOfHours;
                }
            }

            if (!(countOfHour > 0)) return new List<string>();
            if (indexOfLastExistSubject < 0) return new List<string>();

            for (int i = 0; i < Semestrs[indexOfLastExistSubject].Marks.Count; i++)
            {
                double lastExpression = 0;
                string someString = "";

                foreach (SemestrForAtestat newSemestr in Semestrs)
                {
                    if (newSemestr.Marks.Count > 0 && !newSemestr.StateExamenExist)
                    {
                        double ocinka;
                        lastExpression += newSemestr.CountOfHours
                            * (double.TryParse(newSemestr.Marks[i], out ocinka) ? ocinka : 0);
                        someString += double.TryParse(newSemestr.Marks[i], out ocinka) ? "" : newSemestr.Marks[i];
                    }
                }

                _pidsumkovaOcinka.Add(string.IsNullOrEmpty(someString.Trim()) ? Math.Round(lastExpression/countOfHour, 0) + "": someString);
            }  

            return _pidsumkovaOcinka;
        }

        public bool GroupExist(string groupName)
        {
            return GroupPrefixForExam.Any(groupPrefixStatemet => groupPrefixStatemet.Equals(groupName.Split('-')[0]));
        }
    }

    public class SemestrForAtestat
    {
        public int Semestr;
        public double CountOfHours = 0;
        public bool StateExamenExist = false;
        public List<RecordStudmark> Marks = new List<RecordStudmark>();
    }

    public class Semestr
    {
        public double CountOfHours = 0;

        public double CursovaRobota = 0;
        public double Isput = 0;
        public double Zalic = 0;
        public double DyfZalik = 0;
        public double StateExamination = 0;

        public string PracticeFormOfControl = "";

    }

    public class Practice
    {
        public string Name;
        public string Semestr;
        public string FormOfControl;
        public double CountOfHours;
        public string NumberOfOlic;
        public List<string> Teacher;
    }

    public class StateExamination
    {
        public string Name;
        public string Semestr;
    }

    public class Student
    {
        public string Pib;
        public string NumberOfBook;
        public string Group;
        public string FormaTeaching;
        public string Benefits = "";
        public string PibChanged;
        public List<Ocinka> Ocinkas = new List<Ocinka>();

        public string GetPib()
        {
            return string.IsNullOrWhiteSpace(PibChanged) ? Pib : PibChanged;
        }
    }
    
    public class NumberOfOblic
    {
        public string Number;
        public string Subject;
        public string Group;
    }

    public class Ocinka
    {
        public string Mark;
        public string SubjectName;
        public string StudentName;
    }

    public class RecordStudmark
    {
        public string Mark;
        public string StudentName;
    }
}
