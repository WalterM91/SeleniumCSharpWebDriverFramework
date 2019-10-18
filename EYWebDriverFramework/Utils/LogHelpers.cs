using EYWebDriverFramework.Config;
using NUnit.Framework;
using System;
using System.IO;

namespace EYWebDriverFramework.Utils
{
    public class LogHelpers
    {
        // Global declaration
        private static StreamWriter _streamw = null;
        private static string LogPath;

        public static string NowFileName()
        {
            return string.Format("{0:yyyyMMddHHmmssffff}", DateTime.Now);
        }


        // Create a file which can store log information
        public static void CreateLogFile(string testName)
        {
            string dir = Settings.LogPath;

            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            LogPath = dir + testName + NowFileName() + ".log";
            _streamw = File.AppendText(LogPath);
            
        }

        // Create a method which can write the text in the log file
        public static void Write(string logmessage)
        {
            string time = String.Format("{0} {1}", DateTime.Now.ToShortDateString(), DateTime.Now.ToEasternTime().ToLongTimeString());
            _streamw.WriteLine("{0}\t\t {1}", time, logmessage);
            _streamw.Flush();
        }

        public static void CloseLogFile()
        {
            TestContext.AddTestAttachment(LogPath, "Log file of test");
            _streamw.Close();
        }
    }
}
