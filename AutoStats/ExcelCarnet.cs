using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace AutoStats
{
    public static class ExcelCarnet
    {
        private static ExcelPackage _package;
        private static ExcelWorksheet _carnetSheet;
        private static int _firstJanuaryRow;
        private static int _beginningYear;

        public static int LastActivity
        {
            get
            {
                var row = 0;
                for (var i = 350; i > 0; i--)
                {
                    if(_carnetSheet.GetValue(i, 4) == null) continue;
                    row = i;
                    break;
                }

                var dayOfYear = row - _firstJanuaryRow;

                if (dayOfYear < 0)
                    throw new NotImplementedException();
                else
                {
                    var dateTime = new DateTime(_beginningYear + 1, 1, 1) + new TimeSpan(dayOfYear+1,0,0,0);
                    return UnixTimestamp.ToEpochTime(dateTime);
                }

            }
        }

        static ExcelCarnet()
        {
            var fileInfo = new FileInfo((string)new JsonReader(@"C:\Users\guill\Programmation\dotNET_doc\projets\Console\AutoStats\AutoStats\appsettings.json").Content["Config"]["TrainingSheetPath"]);
            _package = new ExcelPackage(fileInfo);
            _carnetSheet = _package.Workbook.Worksheets["Carnet "];
            //get the 1st January row
            var firstWeek = WeekToSheet(1);
            for (var i = 0; i < 7; i++)
            {
                if (_carnetSheet.GetValue<int>(firstWeek + i, 3) != 1) continue;
                _firstJanuaryRow = firstWeek + i;
                break;
            }

            var yearJson = (string) new JsonReader(
                    @"C:\Users\guill\Programmation\dotNET_doc\projets\Console\AutoStats\AutoStats\appsettings.json")
                .Content["Config"]["Year"];
            _beginningYear = Convert.ToInt32(yearJson.Split('-')[0]);
        }
        
        public static void SetNewActivity(DateTime dateTime, Activity activity)
        {
            
            if (dateTime.Year == _beginningYear)
                throw new NotImplementedException("Not allowed to Write an activity for this year");
            else if (dateTime.Year == _beginningYear + 1)
            {
                var row = _firstJanuaryRow + dateTime.DayOfYear - 1;
                WriteActivity(row, activity);
            }
        }

        public static void SaveAndCloseExcel()
        {
            _package.Save();
            _package.Dispose();
        }
        
        
        private static int WeekToSheet(int week)
        {
            for (var i = 2; i <= 366; i++)
            {
                var content = _carnetSheet.GetValue<int>(i, 2);
                if (content == week)
                    return i;
            }

            throw new ExcelErrorValueException("The week n°" + week + " has not been found in the excel file", null);
        }
        private static void WriteActivity(int row, Activity activity)
        {
            _carnetSheet.SetValue(row, 4, activity.Name);
            _carnetSheet.SetValue(row, 5, activity.Time);
            _carnetSheet.SetValue(row, 6, activity.Distance);
            _carnetSheet.SetValue(row, 7, Math.Round(activity.AvrSpeed, 1));
            _carnetSheet.SetValue(row, 9, activity.HeartRate);
            _carnetSheet.SetValue(row, 10, activity.HeartMax);
        }

    }
}