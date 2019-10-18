using EYWebDriverFramework.Config;
using System;
using System.IO;

namespace EYWebDriverFramework.Utils
{//this could be the class that prepares the content for the email
    public class TxtCreator
    {
        // Global declaration
        private static StreamWriter _streamw = null;
        public static string txtName = "\\#" + NowFileName() + ".log";
        public static string txtPath = Settings.DownloadPath;
        public static string txtFullPath = txtPath + txtName;

        public static string NowFileName()
        {
            return string.Format("{0:MM-dd-yyyy-HH-mm}", DateTime.Now);
        }

        public static void CreateLogFile()
        {
            _streamw = File.AppendText(txtFullPath);
        }

        // Create a method which can write the text in the log file
        public static void Write(string logmessage)
        {
            _streamw.WriteLine(logmessage);
            _streamw.Flush();
        }

        public static void EndEydsRun(string projectLink, string reviewTime, string calculationTime)
        {
            txtName = txtName.Replace("#", "EYDS_DailyRun-");
            txtFullPath = txtPath + txtName;
            CreateLogFile();
            Write(string.Format("Hi All,{0}{0}" +
                "EYDS test run on STG environment completed successfully on {1} at {2}.{0}{0}" +
                "Time taken for processes:{0}" +
                " - Review process time: {3}.{0}" +
                " - Calculation process time: {4}.{0}{0}" +
                "Please find the results attached. {0}{0}" +
                "{5}", Environment.NewLine, DateTime.Now.ToShortDateString(), 
                DateTime.Now.ToEasternTime().ToString("HH:mm:ss EST"), 
                reviewTime, calculationTime, projectLink));
            CloseLogFile();
            txtName = txtName.Replace("EYDS_DailyRun-", "#");
        }

        public static void SlowProcessMsg(string projectLink, string reviewTime, string calculationTime)
        {
            txtName = txtName.Replace("#", "EYDS_DailyRun-SlowProcessWarning-");
            txtFullPath = txtPath + txtName;
            CreateLogFile();
            Write(string.Format("Hi All,{0}{0}" +
                "Today long times were perceived for some tasks in AlloKate.{0}{0}" +
                "Time taken for processes:{0}" +
                " - Review process time: {1}.{0}" +
                " - Calculation process time: {2}.{0}{0}" +
                "The link of the project is: {0}" + "{3}",
                Environment.NewLine, reviewTime, 
                calculationTime, projectLink));
            CloseLogFile();
            txtName = txtName.Replace("EYDS_DailyRun-SlowProcessWarning-", "#");
        }

        public static void CloseLogFile()
        {
            _streamw.Close();
        }
    }
}
