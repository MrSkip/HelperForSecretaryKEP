using System;

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
        public Subject Subject;
        public Practice Practice;
        public StateExamination StateExamination;
        public Students Students;

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
        public int CountOfHours;
        public int CursovaRobota;
        public int Isput;
        public int DyfZalikOrZalic;
        public int DyfZalik;
    }

    public class Practice
    {
        public string Name;
        public string Semestr;
        public string FormOfControll;
        public int CountOfHours;
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
