namespace myKR.Coding
{
    public class Group
    {
        public string Name;
        public string DirectOfDirection;
        public string Speciality;
        public string CodeOfSpeciality;
        public string Curator;
        public int Course;
        public int Year;
        public Subject Subject;
        public Practice Practice;
        public StateExamination StateExamination;
    }
    public class Subject
    {
        public string Name;
        public int CountOfHours;
        public string Teacher;
        public Semestr
            FirstSemestr,
            SecondSemestr;
    }

    public class Semestr
    {
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
}
