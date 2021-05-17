#undef TEST

using System;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace AutoStats
{
    class Program
    {
#if TEST
        static void Main(string[] args)
        {
            double a = 122.44567;
            Console.WriteLine(Math.Round(a, 1));
        }
#endif

#if !TEST
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var newActivitiesTask = Strava.ActivitiesAfter(ExcelCarnet.LastActivity);
            newActivitiesTask.Wait();
            var newActivities = newActivitiesTask.Result;

            var nbrActivities = 0;

            try
            {
                while (true)
                {
                    var activity = new Activity()
                    {
                        AvrSpeed = newActivities[nbrActivities]["average_speed"] * 3.6,
                        Distance = newActivities[nbrActivities]["distance"] / 1000,
                        Name = newActivities[nbrActivities]["name"],
                        Time = newActivities[nbrActivities]["moving_time"]/60,
                        HeartMax = newActivities[nbrActivities]["max_heartrate"],
                        HeartRate = newActivities[nbrActivities]["average_heartrate"],
                    };

                    var localDateTime = ((string) newActivities[nbrActivities]["start_date_local"]).Split();
                    var date = Array.ConvertAll(localDateTime[0].Split('/'), Convert.ToInt32);
                    var time = Array.ConvertAll(localDateTime[1].Split(':'), int.Parse);

                    var activityDateTime = new DateTime(date[2], date[0], date[1], time[0], time[1], time[2]);

                    ExcelCarnet.SetNewActivity(activityDateTime, activity);
                    nbrActivities++;
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                Console.WriteLine($"Done, Writing {nbrActivities} Activities");
                ExcelCarnet.SaveAndCloseExcel();
            }

            
        }
#endif
    }
}