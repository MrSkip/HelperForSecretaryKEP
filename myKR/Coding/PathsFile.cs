using System;
using System.IO;
using System.Windows.Forms;
using myKR.Properties;
using Newtonsoft.Json;

namespace myKR.Coding
{
    public class PathsFile
    {
        private static readonly log4net.ILog Log =
            log4net.LogManager.GetLogger("PathsFile.cs");

        private static PathsFile _pathsFile;
        
        public static PathsFile GetPathsFile()
        {
            Log.Info(LoggerConstants.ENTER_EXIT);
            return _pathsFile ?? (_pathsFile = new PathsFile());
        }

        public static PathsDTO PathsDto;

        private PathsFile()
        {
            ParseJson(Environment.CurrentDirectory + "\\paths.json");
            if (string.IsNullOrEmpty(PathsDto.Work))
                FirstStart();
        }

        private void FirstStart()
        {
            Log.Info(LoggerConstants.ENTER_EXIT);
            string parentFolder = Environment.CurrentDirectory;

            PathsDto.Work = "this is not first uses";
            PathsDto.PathToExcelDataForProgram = parentFolder + "\\Data\\DataToProgram.xls";
            PathsDto.PathToFileWithMacros = parentFolder + "\\Data\\WithMacros.xls";
            PathsDto.PathToStudentDb = parentFolder + "\\LastReadFromHere\\База.xls";
            PathsDto.PathToWorkPlan = parentFolder + "\\LastReadFromHere\\Rp15.xls";
            PathsDto.PathToFolderWithOblicUspishnosti = parentFolder + "\\User Data\\Облік успішності\\";
            PathsDto.PathToFolderWithZvedeniaVidomistUspishnosti = parentFolder +
                                                                   "\\User Data\\Зведена відомість успішності\\";
            PathsDto.PathToArhive = parentFolder + "\\User Data\\Зведена відомість успішності\\Архів\\";
            PathsDto.PathToAtestatFolder = parentFolder + "\\User Data\\Атестат\\";
            PathsDto.ExcelExtensial = ".xls";

            WriteFromObjectToJson();
        }

        public static void WriteFromObjectToJson()
        {
            Log.Info(LoggerConstants.ENTER);
            string currentFolder = Environment.CurrentDirectory + "\\paths.json";
            if (!File.Exists(currentFolder))
            {
                MessageBox.Show(Resources.notFindJsonFileWithPaths);
                Log.Info(LoggerConstants.EXIT);
                Environment.Exit(-1);
            }
            try
            {
                string json = JsonConvert.SerializeObject(PathsDto);

                StreamWriter r = new StreamWriter(currentFolder);
                r.Write(json);
                r.Close();
            }
            catch (Exception e)
            {
                Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
                Log.Info(LoggerConstants.EXIT);
            }
            Log.Info(LoggerConstants.EXIT);
        }

        private void ParseJson(string currentFolder)
        {
            Log.Info(LoggerConstants.ENTER);
            if (!File.Exists(currentFolder))
            {
                MessageBox.Show(Resources.notFindJsonFileWithPaths);
                Log.Info(LoggerConstants.EXIT);
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
                catch (JsonException e)
                {
                    Log.Warn(LoggerConstants.SOMETHING_WRONG, e);
                    CheckObject();
                }
            }
            Log.Info(LoggerConstants.EXIT);
        }

        private void CheckObject()
        {
            Log.Info(LoggerConstants.ENTER);
            if (PathsDto == null || IfAllIsEmpty())
            {
                MessageBox.Show(Resources.badParseJsonFile);
                Log.Info(LoggerConstants.EXIT);
                Environment.Exit(-1);
            }
            Log.Info(LoggerConstants.EXIT);
        }

        private bool IfAllIsEmpty()
        {
            Log.Info(LoggerConstants.ENTER_EXIT);
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
            public string Work;
            public string PathToStudentDb;
            public string PathToWorkPlan;

            public string PathToExcelDataForProgram;

            public string PathToFolderWithOblicUspishnosti;
            public string PathToFolderWithZvedeniaVidomistUspishnosti;
            public string PathToArhive;
            public string PathToAtestatFolder;

            public string PathToFileWithMacros;
            
            public string ExcelExtensial;
        }
    }
}
