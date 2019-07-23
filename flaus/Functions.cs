using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;


namespace peteli.flaus
{

    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string HelloDna(string name)
        {
            return "Hello " + name;
        }

        #region Iso8601 Date Time
        // https://www.sourcecodestore.com/Article.aspx?ID=33
        // https://blogs.msdn.microsoft.com/shawnste/2006/01/24/iso-8601-week-of-year-format-in-microsoft-net/
        /// <summary>
        /// Week number is quite complex: ISO 8601[^] specifies that weeks start on Monday, 
        /// and that Jan 1st is in week one if it occurs on a Thursday, Friday, Saturday, or Sunday. 
        /// Otherwise it is in the last week of the previous year, which may be week 51, 52, or 53.
        /// </summary>
        /// <param name="inDate"></param>
        /// <returns></returns>

        [ExcelFunction(Description = "Date to Year-Week")]
        public static string Date2Week(DateTime inDate)
        {
            // Uses the default calendar of the InvariantCulture.
            Calendar myCal = CultureInfo.InvariantCulture.Calendar;
            int week = GetIso8601WeekOfYear(inDate);
            int year = ((week == 52 || week == 53) && inDate.Month == 1) ? inDate.Year - 1 : inDate.Year;
            year = ((week == 1 || week == 1) && inDate.Month == 12) ? inDate.Year + 1 : year;
            string result = year.ToString("D4") + "ww" + week.ToString("D2");
            return result;
        }
 
        private static int WeekofYearIso8601Parameter(DateTime inDate)
        {
            Calendar myCal = CultureInfo.InvariantCulture.Calendar;
            return myCal.GetWeekOfYear(inDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        private static int GetIso8601WeekOfYear(System.DateTime inDate)
        {
            DayOfWeek yDay = inDate.DayOfWeek;
            System.DateTime dtBuff = inDate;
            if (yDay >= DayOfWeek.Monday & yDay <= DayOfWeek.Wednesday)
            {
                dtBuff = dtBuff.AddDays(3);
            }
            Calendar myCal = CultureInfo.InvariantCulture.Calendar;
            return WeekofYearIso8601Parameter(dtBuff);
        }
        #endregion

    }
}
