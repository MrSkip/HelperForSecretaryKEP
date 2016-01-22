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
        public List<Subject> Subjects;
        public List<Practice> Practice;
        public List<StateExamination> StateExamination;
        public List<Students> Students;

        public int GetYearInIneget()
        {
            return int.Parse(Year);
        }
    }
    public class Subject
    {
        public string Name;
        public string Teacher;
        public Semestr
            FirstSemestr,
            SecondSemestr;
    }

    public class Semestr
    {
        public double CountOfHours = 0;
        public double CursovaRobota = 0;
        public double Isput = 0;
        public double Zalic = 0;
        public double DyfZalik = 0;
    }

    public class Practice
    {
        public string Name;
        public string Semestr;
        public string FormOfControll;
        public double CountOfHours;
        public string Teacher;
    }

    public class StateExamination
    {
        public string Name;
        public int Semestr;
    }

    public class Students
    {
        public string Pib;
        public string NumberOfBook;
        public string FormaTeaching;
        public string Benefits;
    }
}
