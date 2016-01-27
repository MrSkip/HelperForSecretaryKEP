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
            ExcelFile.ReadStudentsAndOlicAndCurators(pathToStudent);
        }

        public static List<string> GetlistOfGroupsName()
        {
            return Groups.Select(@group => @group.Name).ToList();
        }

        /*
         *      Create Oblic Uspishosti
         *      if `groupName` is null or empty and `subjectName` is null or empty than create for all groups
         *      else if `groupName` is not null and not empty and `subjectName` is null or empty than create for one group
         *      else if `groupName` is not null and not empty and `subjectName` is not null and not empty than create for one subject
        */
        public static void CreateOblicUspishnosti(string groupName, string subjectName, int pivricha)
        {
            ExcelFile.CreateOblicUspishnosti(groupName, subjectName, pivricha);
        }

        public static void CreateVidomistUspishnosti(string groupName, int pivricha)
        {
            foreach (Group @group in Groups)
            {
                if (group.Name.Equals(groupName))
                {
                    ExcelFile.CreateVidomist(group, pivricha);
                    break;
                }
            }
        }
    }

}
