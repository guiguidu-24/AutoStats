using System;
using System.IO;
using OfficeOpenXml;

namespace AutoStats
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            var yearJson = (string) new JsonReader(
                    @"C:\Users\guill\Programmation\dotNET_doc\projets\Console\AutoStats\AutoStats\appsettings.json").Content["Config"]["Year"];
            var beginningYear = int.Parse(yearJson.Split('-')[0]);
            var carnetFileInfo = new FileInfo((string)new JsonReader(@"C:\Users\guill\Programmation\dotNET_doc\projets\Console\AutoStats\AutoStats\appsettings.json").Content["Config"]["TrainingSheetPath"]);
            
            var excelCarnet = new ExcelCarnet(carnetFileInfo, beginningYear);
            
            var newActivityDate = excelCarnet.LastActivity.AddDays(1);
            var newActivities = Strava.ActivitiesAfter(UnixTimestamp.ToEpochTime(newActivityDate));
            
            var nbrActivities = 0;
            while (true)
            {
                string[] localDateTime;
                try
                {
                    localDateTime = ((string) newActivities[nbrActivities]["start_date_local"]).Split();
                }
                catch (ArgumentOutOfRangeException)
                {
                    Console.WriteLine($"Done, Writing {nbrActivities} Activities");
                    excelCarnet.SaveAndCloseExcel();
                    return;
                }
                
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