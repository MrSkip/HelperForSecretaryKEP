using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace myKR.Coding
{
    public class Manager
    {
        public static List<Group> Groups = new List<Group>();

        public static void ReadData(string pathToRobPlan, string pathToStudent)
        {
            if (Groups.Count >= 0) Groups.Clear();
            ExcelFile.ReadRobPlan(pathToRobPlan);
            ExcelFile.SetStudentsIntoGroup(pathToStudent);
        }

        public static List<string> GetlistOfGroupsName()
        {
            return Groups.Select(@group => @group.Name).ToList();
        }
    }

}
