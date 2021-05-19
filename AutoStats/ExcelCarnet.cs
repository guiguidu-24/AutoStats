using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace AutoStats
{
    public class ExcelCarnet
    {
        private readonly ExcelPackage package;
        private readonly ExcelWorksheet carnetSheet;
        private readonly int firstJanuaryRow;
        private readonly int beginningYear;
        
        /// <summary>
        /// Return the <see cref="DateTime"/> of the last activity written in the file
        /// </summary>
        public DateTime LastActivity
        {
            get
            {
                //first, get the row of the last activity
                var row = 0;
                for (var i = 365; i > 0; i--) //start reading by the end...
                {
                    if (carnetSheet.GetValue(i, (int)Column.Theme) != null) // ...and stop when we get a title for an activity
                    {
                        row = i;
                        break;
                    }
                }
                var dayOfYear = row - firstJanuaryRow;

                return new DateTime(beginningYear + 1, 1, 1).AddDays(dayOfYear);
            }
        }
        
        
        /// <summary>
        /// The default constructor
        /// </summary>
        /// <param name="carnetFileInfo">The path of the carnet d'entrainement file</param>
        /// <param name="firstYear">The beginning year for the carnet</param>
        public ExcelCarnet(FileInfo carnetFileInfo, int firstYear)
        {
            package = new ExcelPackage(carnetFileInfo);
            carnetSheet = package.Workbook.Worksheets["Carnet "];
            beginningYear = firstYear;
            
            //get the 1st January row
            var firstWeek = WeekToSheet(1);
            for (var i = 0; i < 7; i++)
            {
                if (carnetSheet.GetValue<int>(firstWeek + i, (int)Column.Day) != 1) continue;
                firstJanuaryRow = firstWeek + i;
                break;
            }
        }
        
        /// <summary>
        /// Set an activity in the carnet
        /// </summary>
        /// <param name="activityDate">The date of the activity</param>
        /// <param name="activity">The activity to set</param>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="activityDate"/> have to be in the first or in the second year of this carnet/></exception>
        public void SetNewActivity(DateTime activityDate, Activity activity)
        {
            if (activityDate.Year == beginningYear || activityDate.Year == beginningYear + 1)
            {
                var rowRelativeTo1StJ = activityDate - new DateTime(beginningYear + 1, 1, 1);
                var row = firstJanuaryRow + rowRelativeTo1StJ.Days; 
                row += activityDate.Year == beginningYear ? 1 : 0;
                
                //write the activity on the file
                carnetSheet.SetValue(row, (int)Column.Theme, activity.Name);
                carnetSheet.SetValue(row, (int)Column.Time, activity.Time);
                carnetSheet.SetValue(row, (int)Column.Distance, activity.Distance);
                carnetSheet.SetValue(row, (int)Column.Speed, activity.AvrSpeed);
                carnetSheet.SetValue(row, (int)Column.AverageHeart, activity.HeartRate);
                carnetSheet.SetValue(row, (int)Column.MaxHeart, activity.HeartMax);
                return;
            }
            
            throw new ArgumentOutOfRangeException(nameof(activityDate));
        }
        
        /// <summary>
        /// Save and dispose of the Excel sheet
        /// </summary>
        public void SaveAndCloseExcel()
        {
            package.Save();
            package.Dispose();
        }
        
        /// <summary>
        /// Transform a week number into a row 
        /// </summary>
        /// <param name="week">The week number from the beginning of the year</param>
        /// <returns>The row relative to a week</returns>
        /// <exception cref="ArgumentOutOfRangeException"><paramref name="week"/> have to be under or equal to 53"/></exception>
        /// <exception cref="ExcelErrorValueException">This <paramref name="week"/> isn't in the excel file</exception>
        private int WeekToSheet(int week)
        {
            if (week > 53) throw new ArgumentOutOfRangeException(nameof(week));

            for (var i = 2; i <= 366; i++)
            {
                var content = carnetSheet.GetValue<int>(i, 2);
                if (content == week)
                    return i;
            }

            throw new ExcelErrorValueException("The week n°" + week + " has not been found in the excel file", null);
        }
    }

    public enum Column
    {
        Month = 1,
        Week,
        Day,
        Theme,
        Time,
        Distance,
        Speed,
        AveragePower,
        AverageHeart,
        MaxHeart,
        Sensation,
        Weather
    }
}