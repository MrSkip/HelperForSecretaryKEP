using System;
using System.Collections.Generic;
using System.Linq;

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

        public static Group GetGroupByName(String name)
        {
            foreach (Group @group in Groups)
            {
                if (group.Name.Equals(name))
                    return group;
            }
            return new Group();
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

        public static void CreateVidomistUspishnosti(string groupName, int pivricha, string month)
        {
            if (!groupName.Equals("Усі групи"))
                ExcelFile.CreateVidomist(GetGroupByName(groupName), pivricha, month);
            else
            {
                foreach (Group @group in Groups)
                {
                    ExcelFile.CreateVidomist(group, pivricha, month);
                }
                Control.IfShow = false;
            }
        }

        public static void CreateAtestat(List<string> groupList)
        {
//            foreach (string s in groupList)
//            {
//                ExcelFile.ReadDataFromArhiveZVtoAtestat(s);
//            }
//            foreach (NewSubject newSubject in ExcelFile.GetSubjectsForAtestat())
//            {
//                Console.WriteLine("nameOfSubject - " + newSubject.Name);
//                Console.WriteLine("groupPrefix - " + newSubject.GroupPrefixStatemets.Count + "\n");
//            }
            ExcelFile.ReadAllNeedSheetsFromArhiveZVtoAtestat(groupList[0]);
        }
    }

}
