using System;
using System.Threading.Tasks;

namespace AutoStats
{
    public static class UnixTimestamp
    {
        public static int ToEpochTime(DateTime dateTime)
        {
            var t = dateTime - new DateTime(1970, 1, 1);
            var secondsSinceEpoch = (int)t.TotalSeconds;

            return secondsSinceEpoch;
        }

        public static DateTime ToDateTime(int epochTime)
        {
            var t = new TimeSpan(0, 0, epochTime);
            return new DateTime(1970, 1, 1) + t;
        }
    }
}