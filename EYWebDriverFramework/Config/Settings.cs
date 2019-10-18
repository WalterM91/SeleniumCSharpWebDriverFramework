namespace EYWebDriverFramework.Config
{
    public class Settings
    {
        public static string LogPath { get; set; }
        public static string DownloadPath { get; set; }
        public static string TestResultsPath { get; set; }
        public static string ChromeProfile { get; set; }
        public static Timeouts Timeouts { get; set; }
        public static Threshold Threshold { get; set; }
        public static EYDS Eyds { get; set; }

    }
    public class Timeouts
    {
        public int Implicit { get; set; }
        public int Explicit { get; set; }
        public int PageLoad { get; set; }
        public int Process { get; set; }
        public int SleepIntervalInMillis { get; set; }

    }

    public class Threshold
    {
        public double Currency { get; internal set; }
        public double Percentage { get; internal set; }
    }

    public class EYDS
    {
        public int DayAValue { get; internal set; }
        public int DayBValue { get; internal set; }
        public string DayASuffix { get; internal set; }
        public string DayBSuffix { get; internal set; }
        public int ProcessTime { get; internal set; }
        public int ReviewMaxProcessTime { get; internal set; }
        public int CalculationMaxProcessTime { get; internal set; }

    }

}
