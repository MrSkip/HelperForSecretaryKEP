using System;
using System.IO;
using System.Windows.Forms;
using myKR.Properties;
using Newtonsoft.Json;

namespace myKR.Coding
{
    public class PathsFile
    {
        private static PathsFile _pathsFile;
        
        public PathsFile GetPathsFile()
        {
            return _pathsFile ?? (_pathsFile = new PathsFile());
        }

        public PathsDTO PathsDto;

        private PathsFile()
        {
            ParseJson(Environment.CurrentDirectory + "\\paths.json");
        }

        private void ParseJson(string currentFolder)
        {
            if (!File.Exists(currentFolder))
            {
                MessageBox.Show(Resources.notFindJsonFileWithPaths);
                Environment.Exit(-1);
            }

            using (StreamReader r = new StreamReader(currentFolder))
            {
                string json = r.ReadToEnd();
                try
                {
                    PathsDto = JsonConvert.DeserializeObject<PathsDTO>(json);
                    CheckObject();
                }
                catch (JsonException)
                {
                    // ignore
                    CheckObject();
                }
            }
        }

        private void CheckObject()
        {
            if (PathsDto == null || IfAllIsEmpty())
            {
                MessageBox.Show(Resources.badParseJsonFile);
                Environment.Exit(-1);
            }
        }

        private bool IfAllIsEmpty()
        {
            return string.IsNullOrEmpty(PathsDto.PathToArhive)
                   && string.IsNullOrEmpty(PathsDto.PathToStudentDb)
                   && string.IsNullOrEmpty(PathsDto.PathToWorkPlan)
                   && string.IsNullOrEmpty(PathsDto.PathToExcelDataForProgram)
                   && string.IsNullOrEmpty(PathsDto.PathToFolderWithOblicUspishnosti)
                   && string.IsNullOrEmpty(PathsDto.PathToFolderWithZvedeniaVidomistUspishnosti)
                   && string.IsNullOrEmpty(PathsDto.PathToArhive);
        }

        public class PathsDTO
        {
            public string PathToStudentDb;
            public string PathToWorkPlan;

            public string PathToExcelDataForProgram;

            public string PathToFolderWithOblicUspishnosti;
            public string PathToFolderWithZvedeniaVidomistUspishnosti;
            public string PathToArhive;

        }
    }
}
