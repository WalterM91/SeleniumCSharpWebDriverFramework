using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.XPath;

namespace EYWebDriverFramework.Config
{
    public class ConfigReader
    {
        private static readonly string FrameworkConfigFile = "FrameworkSettings.xml";
        public static string TestResultsPath;
        public static string DownloadsPath;
        public static int NumberOfFiles;
        public static string reportGenerated = null;

        public static void SetEydsSettings()
        {
            //EYDS
            XPathItem dayAValue;
            XPathItem dayBValue;
            XPathItem dayASuffix;
            XPathItem dayBSuffix;
            XPathItem processTime;
            XPathItem reviewMaxProcessTime;
            XPathItem calculationMaxProcessTime;

            string stringFileName = GetDriversPath(FrameworkConfigFile);

            FileStream stream = new FileStream(stringFileName, FileMode.Open);
            XPathDocument document = new XPathDocument(stream);
            XPathNavigator navigator = document.CreateNavigator();

            // Get XMl Details and pass it in XPathItem type variables
            string rootPath = "FrameworkConfiguration/";

            //EYDS
            dayAValue = navigator.SelectSingleNode(rootPath + "EYDS/DayAValue");
            dayBValue = navigator.SelectSingleNode(rootPath + "EYDS/DayBValue");
            dayASuffix = navigator.SelectSingleNode(rootPath + "EYDS/DayASuffix");
            dayBSuffix = navigator.SelectSingleNode(rootPath + "EYDS/DayBSuffix");
            processTime = navigator.SelectSingleNode(rootPath + "EYDS/ProcessTime");
            reviewMaxProcessTime = navigator.SelectSingleNode(rootPath + "EYDS/ReviewMaxProcessTime");
            calculationMaxProcessTime = navigator.SelectSingleNode(rootPath + "EYDS/CalculationMaxProcessTime");

            EYDS eyds = new EYDS
            {
                DayAValue = dayAValue.ValueAsInt,
                DayBValue = dayBValue.ValueAsInt,
                DayASuffix = dayASuffix.ToString(),
                DayBSuffix = dayBSuffix.ToString(),
                ProcessTime = processTime.ValueAsInt,
                ReviewMaxProcessTime = reviewMaxProcessTime.ValueAsInt,
                CalculationMaxProcessTime = calculationMaxProcessTime.ValueAsInt
            };

            Settings.Eyds = eyds;
        }

        public static void SetFrameworkSettings(string testName)
        {
            //General settings
            XPathItem logPath;
            XPathItem downloadPath;
            XPathItem chromeProfile;
            XPathItem testResultsPath;

            //Timeouts
            XPathItem timeoutImplicit;
            XPathItem timeoutExplicit;
            XPathItem timeoutPageLoad;
            XPathItem timeoutProcess;

            string stringFileName = GetDriversPath(FrameworkConfigFile);

            //string stringFileName = Environment.CurrentDirectory.ToString() + "\\Config\\FrameworkSettings.xml";
            FileStream stream = new FileStream(stringFileName, FileMode.Open);
            XPathDocument document = new XPathDocument(stream);
            XPathNavigator navigator = document.CreateNavigator();

            // Get XMl Details and pass it in XPathItem type variables
            string rootPath = "FrameworkConfiguration/";
            logPath = navigator.SelectSingleNode(rootPath + "RunSettings/LogPath");
            downloadPath = navigator.SelectSingleNode(rootPath + "RunSettings/DownloadPath");
            chromeProfile = navigator.SelectSingleNode(rootPath + "RunSettings/ChromeProfilePath");
            testResultsPath = navigator.SelectSingleNode(rootPath + "RunSettings/TestResultsPath");

            timeoutImplicit = navigator.SelectSingleNode(rootPath + "Timeouts/Implicit");
            timeoutExplicit = navigator.SelectSingleNode(rootPath + "Timeouts/Explicit");
            timeoutPageLoad = navigator.SelectSingleNode(rootPath + "Timeouts/PageLoad");
            timeoutProcess = navigator.SelectSingleNode(rootPath + "Timeouts/Process");

            Timeouts timeouts = new Timeouts
            {
                Implicit = timeoutImplicit.ValueAsInt,
                Explicit = timeoutExplicit.ValueAsInt,
                PageLoad = timeoutPageLoad.ValueAsInt,
                Process = timeoutProcess.ValueAsInt
            };

            XPathItem currencyThreshold = navigator.SelectSingleNode(rootPath + "Threshold/Currency");
            XPathItem percentageThreshold = navigator.SelectSingleNode(rootPath + "Threshold/Percentage");

            Threshold threshold = new Threshold
            {
                Currency = currencyThreshold.ValueAsDouble,
                Percentage = percentageThreshold.ValueAsDouble
            };
            // Set XML Details in the property to be used across the Framework
            Settings.LogPath = GetPath(logPath.Value.ToString());
            Settings.DownloadPath = GetTempDownloadPath(downloadPath.ToString(), testName);
            Settings.TestResultsPath = GetPath(testResultsPath.ToString());
            Settings.ChromeProfile = GetPath(chromeProfile.ToString());
            Settings.Timeouts = timeouts;
            Settings.Threshold = threshold;
            TestResultsPath = Settings.TestResultsPath;
            DownloadsPath = Settings.DownloadPath;
            NumberOfFiles = GetNumberOfFiles(Settings.DownloadPath, "*.xlsm");
        }

        private static string GetTempDownloadPath(string downloadPath, string testName)
        {
            return GetPath(downloadPath + @"\" + testName + "_" + DateTime.Now.Ticks.ToString());
        }

        private static string GetDriversPath(string paths)
        {
            string outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string relativePath = @"..\..\..\EYWebDriverFramework\Config";
            return Path.GetFullPath(Path.Combine(outPutDirectory, relativePath, paths));
        }
        public static string GetPath(string paths = "")
        {
            string outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string relativePath = @paths;
            string fullPath = Path.GetFullPath(Path.Combine(outPutDirectory, relativePath));
            bool exists = System.IO.Directory.Exists(fullPath);
            if (!exists)
                System.IO.Directory.CreateDirectory(fullPath);
            return fullPath;
        }

        public static int GetNumberOfFiles(string paths, string extension)
        {
            int fileCount = Directory.EnumerateFiles(@paths, extension, SearchOption.TopDirectoryOnly).Count();
            return fileCount;
        }

    }


}
