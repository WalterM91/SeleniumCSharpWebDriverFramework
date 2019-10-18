using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;

namespace AllokateCJV.Tests
{
    public static class Paths
    {
        private static List<string> goldenCopiesPaths = new List<string>();

        private static string GetPathToFile(string relativePath)
        {
            string outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            return Path.GetFullPath(Path.Combine(outPutDirectory, relativePath));
        }

        private static List<string> GetPathToFiles(string relativePath)
        {
            List<string> directories = new List<string>();
            string outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] middleFolderPaths = Directory.GetDirectories(Path.GetFullPath(Path.Combine(outPutDirectory, relativePath)));
            foreach (string subDir in middleFolderPaths)
            {
                string[] testNamesPath = Directory.GetDirectories(subDir);
                foreach (string finalPath in testNamesPath)
                {
                    directories.Add(finalPath + "\\");
                    goldenCopiesPaths.Add(finalPath + "\\Golden Copies\\");
                }
            }
            return directories;
        }

        public static string GetGoldenCopyPath(string projectName, string filename)
        {
            string chosenGoldenPath = goldenCopiesPaths.Where(path => path.Contains("\\" + projectName + "\\")).FirstOrDefault();
            return GetPathToFile(chosenGoldenPath + filename);

        }

        public static List<string> TestCasesNew(string filename)
        {
            return GetPathToFiles(@"..\..\..\" + filename);
        }

    }
}
