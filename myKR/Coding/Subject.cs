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
        public List<string> GroupPrefixStatemets;
        private List<string> _pidsumkovaOcinka;

        public List<string> GetPidsumkovaOcinka()
        {
            double countOfHour = 0;

            foreach (NewSemestr newSemestr in Semestrs)
            {
                countOfHour += newSemestr.CountOfHours;
            }
            if (!(countOfHour > 0) || Semestrs[0].Ocinkas.Count == 0) return new List<string>();

            _pidsumkovaOcinka = new List<string>();

//            foreach (string ocinka in Semestrs[0].Ocinkas)
//            {
//                _pidsumkovaOcinka.Add();
//            }

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
