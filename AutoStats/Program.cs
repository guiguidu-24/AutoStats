using System;
using System.IO;
using OfficeOpenXml;

namespace AutoStats
{
    class Program
    {
        static void Main(string[] args)
        {
            //only necessary in debug mode
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var appSettings = new JsonReader("appsettings.json").Content;
            
            var yearFromJson = (string) appSettings["Config"]["Year"];
            var beginningYear = int.Parse(yearFromJson.Split('-')[0]);
            var carnetFileInfo = new FileInfo((string)appSettings["Config"]["TrainingSheetPath"]);
            var excelCarnet = new ExcelCarnet(carnetFileInfo, beginningYear);
            
            var clientId = (string) appSettings["Config"]["StravaApi"]["ClientId"];
            var clientSecret = (string) appSettings["Config"]["StravaApi"]["ClientSecret"];
            var refreshToken = (string) appSettings["Config"]["StravaApi"]["RefreshToken"];
            
            var newActivityDate = excelCarnet.LastActivity.AddDays(1);
            var newActivities = new Strava(clientId, clientSecret, refreshToken).ActivitiesAfter(UnixTimestamp.ToEpochTime(newActivityDate));
            
            var nbrActivities = 0;
            while (true)
            {
                string[] localDateTime;
                try                     //see if there is more activities
                {
                    localDateTime = ((string) newActivities[nbrActivities]["start_date_local"]).Split();
                }
                catch (ArgumentOutOfRangeException)
                {
                    Console.WriteLine($"Done, Writing {nbrActivities} Activities");
                    if (nbrActivities == 200) Console.WriteLine("There is certainly some other activities. Advice : relaunch the app");
                    excelCarnet.SaveAndCloseExcel();
                    return;
                }
                
                //Write an activity
                var date = Array.ConvertAll(localDateTime[0].Split('/'), Convert.ToInt32);
                var time = Array.ConvertAll(localDateTime[1].Split(':'), int.Parse);
                var activityDateTime = new DateTime(date[2], date[0], date[1], time[0], time[1], time[2]);
                
                var activity = new Activity()
                {
                    AvrSpeed = newActivities[nbrActivities]["average_speed"] * 3.6,
                    Distance = newActivities[nbrActivities]["distance"] / 1000,
                    Name = newActivities[nbrActivities]["name"],
                    Time = newActivities[nbrActivities]["moving_time"] / 60,
                    HeartMax = newActivities[nbrActivities]["max_heartrate"],
                    HeartRate = newActivities[nbrActivities]["average_heartrate"],
                };
                
                excelCarnet.SetNewActivity(activityDateTime, activity);
                nbrActivities++;
            }

        }
    }
}