using EYWebDriverFramework.Config;
using EYWebDriverFramework.Utils;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace AllokateCJV.Tests
{
    public class EydsCompareTestCase
    {
        public string ProjectLink { get; internal set; }
        public string ProjectName { get; internal set; }
        public string Set { get; internal set; }
        public string ReportGroup { get; internal set; }
        public string ReportName { get; internal set; }
        public string FullYear { get; internal set; }
        public int YpeNUM { get; internal set; }
        public string BooksetType { get; internal set; }
        public string LayerType { get; internal set; }
        public string GoldenCopyPath { get; internal set; }

        internal static Func<DataRow, string, string, string, EydsCompareTestCase> GetFromExcel = (row, projectLink, projectName, setName) =>
        {
            return new EydsCompareTestCase
            {
                ProjectLink = projectLink,
                ProjectName = projectName,
                Set = setName,
                ReportGroup = row[0].ToString().Split('_')[1] + ":",
                ReportName = row[1].ToString(),
                FullYear = row[2].ToString(),
                YpeNUM = !String.IsNullOrEmpty(row[3].ToString()) ? Int32.Parse(row[3].ToString()) - 1 : -1,
                BooksetType = row[4].ToString(),
                LayerType = row[5].ToString(),
                GoldenCopyPath = row[6].ToString()
            };
        };

        override public string ToString()
        {
            return GoldenCopyPath;
        }

    }

    public class EydsReviewTestCase : TestCaseData
    {
        public string ProjectLink { get; internal set; }
        public string ProjectName { get; internal set; }
        public string Set { get; internal set; }
        public string StartPeriod { get; internal set; }
        public string StartAtAsset { get; internal set; }
        public string StopPeriod { get; internal set; }
        public string StopAtAsset { get; internal set; }

        internal static Func<DataRow, EydsReviewTestCase> GetFromExcel = (row) =>
        {
            return new EydsReviewTestCase
            {
                ProjectLink = row[0].ToString(),
                ProjectName = row[1].ToString(),
                Set = row[2].ToString(),
                StartPeriod = row[3].ToString(),
                StartAtAsset = row[4].ToString(),
                StopPeriod = row[5].ToString(),
                StopAtAsset = row[6].ToString()
            };

        };

        override public string ToString()
        {
            return ProjectName;
        }

    }


    public class CompareTestCase
    {
        public string ProjectLink { get; internal set; }
        public string ProjectName { get; internal set; }
        public string Set { get; internal set; }
        public string ReportGroup { get; internal set; }
        public string ReportName { get; internal set; }
        public string FullYear { get; internal set; }
        public int YpeNUM { get; internal set; }
        public string BooksetType { get; internal set; }
        public string LayerType { get; internal set; }
        public string GoldenCopyPath { get; internal set; }

        internal static Func<DataRow, string, string, string, CompareTestCase> GetFromExcel = (row, projectLink, projectName, setName) =>
        {
            return new CompareTestCase
            {
                ProjectLink = projectLink,
                ProjectName = projectName,
                Set = setName,
                ReportGroup = row[0].ToString().Split('_')[1] + ":",
                ReportName = row[1].ToString(),
                FullYear = row[2].ToString(),
                YpeNUM = !String.IsNullOrEmpty(row[3].ToString()) ? Int32.Parse(row[3].ToString()) - 1 : -1,
                BooksetType = row[4].ToString(),
                LayerType = row[5].ToString(),
                GoldenCopyPath = row[6].ToString()
            };
        };

        override public string ToString()
        {
            return GoldenCopyPath;
        }

    }

    public class ReviewTestCase
    {
        public string ProjectLink { get; internal set; }
        public string ProjectName { get; internal set; }
        public string Set { get; internal set; }
        public string StartPeriod { get; internal set; }
        public string StartAtAsset { get; internal set; }
        public string StopPeriod { get; internal set; }
        public string StopAtAsset { get; internal set; }

        internal static Func<DataRow, ReviewTestCase> GetFromExcel = (row) =>
        {
            return new ReviewTestCase
            {
                ProjectLink = row[0].ToString(),
                ProjectName = row[1].ToString(),
                Set = row[2].ToString(),
                StartPeriod = row[3].ToString(),
                StartAtAsset = row[4].ToString(),
                StopPeriod = row[5].ToString(),
                StopAtAsset = row[6].ToString()
            };
            
        };

        override public string ToString()
        {
            return ProjectName;
        }

    }

    public class TestGenerator {
        private readonly IEnumerable<string> fullNames;
        public TestGenerator(string path)
        {
            fullNames = GetTestFiles(path);
        }

        public TestGenerator(List<string> paths)
        {
            List<string> fileList = new List<string>();
            foreach (string path in paths)
            {
                IEnumerable<string> fileNames = GetTestFiles(path);
                foreach (var fileNameItem in fileNames)
                {
                    fileList.Add(fileNameItem);
                }
            }
            fullNames = fileList.AsEnumerable();
        }

        private IEnumerable<string> GetTestFiles(string path)
        {
            int totalFiles = Directory.EnumerateFiles(@path, "*.xlsx", SearchOption.TopDirectoryOnly).Count();
            DirectoryInfo dirInfo = new DirectoryInfo(@path);
            FileInfo[] files = dirInfo.GetFiles()
                .Where(file => 
                (file.Attributes & FileAttributes.Hidden) == 0 &&
                (file.Extension.Contains("xlsx") || file.Extension.Contains("xlsm"))
                ).ToArray();
            return files.Select(file => file.FullName);
        }

        public List<TestCaseData> RetrieveReportTests()
        {
            List<TestCaseData> testList = new List<TestCaseData>();

            foreach (string filename in fullNames)
            {
                ExcelInputReader<CompareTestCase>.PopulateInReportCollection(filename, CompareTestCase.GetFromExcel);
            }
            var testItems = ExcelInputReader<CompareTestCase>.GetData();

            foreach (var testItem in testItems)
            {
                testList.Add(new TestCaseData(testItem)
                    .SetProperty(nameof(testItem.ProjectName), testItem.ProjectName)
                    .SetProperty(nameof(testItem.Set), testItem.Set));
            }

            return testList;
        }

        public List<TestCaseData> RetrieveReviewCalcsTests()
        {
            List<TestCaseData> testList = new List<TestCaseData>();

            foreach (string filename in fullNames)
            {
                ExcelInputReader<ReviewTestCase>.PopulateInReviewCalcCollection(filename, ReviewTestCase.GetFromExcel);
            }
            var testItems = ExcelInputReader<ReviewTestCase>.GetData();

            foreach (var testItem in testItems)
            {
                testList.Add(new TestCaseData(testItem)
                    .SetProperty(nameof(testItem.ProjectName), testItem.ProjectName)
                    .SetProperty(nameof(testItem.Set), testItem.Set));
            }

            return testList;
        }

        public List<EydsCompareTestCase> RetrieveEYDSReportTests(string projectName)
        {
            foreach (string filename in fullNames)
            {
                ExcelInputReader<EydsCompareTestCase>.PopulateInReportCollection(filename, EydsCompareTestCase.GetFromExcel);
            }
            return ExcelInputReader<EydsCompareTestCase>.GetData()
                .Where(x => x.GoldenCopyPath.Contains("_" + projectName + (ReportComparatorTests.DayA ? Settings.Eyds.DayASuffix : Settings.Eyds.DayBSuffix))).ToList();
        }

        public List<TestCaseData> RetrieveEYDSReviewCalcsTests()
        {
            List<TestCaseData> testList = new List<TestCaseData>();

            ConfigReader.SetEydsSettings();

            foreach (string filename in fullNames)
            {
                ExcelInputReader<EydsReviewTestCase>.PopulateInReviewCalcCollection(filename, EydsReviewTestCase.GetFromExcel);
            }

            //Esta parte final se repite en los 3 métodos (salvo de el eyds) pero no supe como hacerlo genérico :P
            var testItems = ExcelInputReader<EydsReviewTestCase>.GetData();

            foreach (var testItem in testItems)
            {
                testList.Add(new TestCaseData(testItem)
                    .SetProperty(nameof(testItem.ProjectName), testItem.ProjectName)
                    .SetProperty(nameof(testItem.Set), testItem.Set));
            }
            return testList;
        }
    }
}
