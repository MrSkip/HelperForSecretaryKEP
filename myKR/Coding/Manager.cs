﻿using System.Collections.Generic;
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
        }

        public static List<string> GetlistOfGroupsName()
        {
            return Groups.Select(@group => @group.Name).ToList();
        }
    }

}