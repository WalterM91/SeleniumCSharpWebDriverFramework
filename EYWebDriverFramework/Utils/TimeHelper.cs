using System;

namespace EYWebDriverFramework.Utils
{
    public static class TimeHelper
    {
        public static bool IsProcessSlow(DateTime time, int processMaxTime)
        {
            return (((time.Minute * 60) + time.Second) > processMaxTime);
        }

        public static string FormatTimeForEmail(DateTime time)
        {
            return time.Minute + " min " + time.Second + " sec";
        }


        public static bool IsANumber(this char chr)
        {
            return (chr > 47 && chr < 58);
        }

        public static DateTime ToEasternTime(this DateTime dateTime)
        {
            //TODO: Should we throw an exception whenever the time zone is not found? 
            //      If no, what should we return?
            //      Is it possible for Eastern Standard Time to be different or not present?
            DateTime timeUtc = dateTime.ToUniversalTime();
            DateTime estTime;
            try
            {
                TimeZoneInfo estZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                estTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, estZone);
            }
            catch (TimeZoneNotFoundException)
            {
                throw new TimeZoneNotFoundException("The registry does not define the Eastern Time zone.");
            }
            catch (InvalidTimeZoneException)
            {
                throw new InvalidTimeZoneException("Registry data on the Eastern Time zone has been corrupted.");
            }
            return estTime;
        }
    }
}
