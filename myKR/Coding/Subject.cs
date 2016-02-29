using System;
using System.Collections.Generic;

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

        public List<Ocinka> Ocinka;
    }

    public class NewSubject
    {
        public string Name;
        public string Teacher;
        public List<NewSemestr> Semestrs = new List<NewSemestr>();
        public List<string> GroupPrefixStatemets = new List<string>();
        private List<string> _pidsumkovaOcinka;

        public List<string> GetPidsumkovaOcinka()
        {
            double countOfHour = 0;
            int indexOfLastExistSubject = 0;

            foreach (NewSemestr newSemestr in Semestrs)
            {
                countOfHour += newSemestr.CountOfHours;
                if (newSemestr.Ocinkas.Count > 0)
                    indexOfLastExistSubject = Semestrs.IndexOf(newSemestr);
            }

            if (!(countOfHour > 0)) return new List<string>();
            if (indexOfLastExistSubject <= 0) return new List<string>();

            _pidsumkovaOcinka = new List<string>();

            for (int i = 0; i < Semestrs[indexOfLastExistSubject].Ocinkas.Count; i++)
            {
                double lastExpression = 0;
                string someString = "";

                foreach (NewSemestr newSemestr in Semestrs)
                {
                    if (newSemestr.Ocinkas.Count > 0)
                    {
                        double ocinka;
                        lastExpression += newSemestr.CountOfHours 
                            * (double.TryParse(newSemestr.Ocinkas[i], out ocinka) ? ocinka : 0);
                        someString += double.TryParse(newSemestr.Ocinkas[i], out ocinka) ? "" : newSemestr.Ocinkas[i];
                    }
                }
                _pidsumkovaOcinka.Add(string.IsNullOrEmpty(someString.Trim()) ? Math.Round(lastExpression, 0) + "": someString);
            }

            return _pidsumkovaOcinka;
        }
        public bool GroupExist(string groupName)
        {
            foreach (string groupPrefixStatemet in GroupPrefixStatemets)
            {
                if (groupPrefixStatemet.Equals(groupName.Substring(0, groupName.IndexOf("-", StringComparison.Ordinal) + 1)))
                    return true;
            }
            return false;
        }
    }

    public class NewSemestr
    {
        public int NumberOfSemestr;
        public double CountOfHours = 0;
        public bool StateExamenExist = false;
        public List<string> Ocinkas = new List<string>();
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
        public List<Ocinka> Ocinkas = new List<Ocinka>();
    }

    public class NumberOfOblic
    {
        public string Number;
        public string Subject;
        public string Group;
    }

    public class Ocinka
    {
        public string Name;
        public string Number;
    }
}
