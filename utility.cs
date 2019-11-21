using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NthDayofMonth
{
    static class utility
    {

        public static bool NthDayOfMonth(DateTime date, DayOfWeek dow, int n)
        {
            int d = date.Day;
            
            DateTime dt = DateTime.Now; //if you want check current date, use DateTime.Now
            // Is this the last day of the month
            bool isLastDayOfMonth = (dt.Day == DateTime.DaysInMonth(dt.Year, dt.Month));
            return date.DayOfWeek == dow && (d - 1) / 7 == (n - 1);
        }


        public static bool LastDayOfMonth(DateTime date)
        {


           // DateTime dt = DateTime.Now; //if you want check current date, use DateTime.Now
            // Is this the last day of the month
            bool isLastDayOfMonth = (date.Day == DateTime.DaysInMonth(date.Year, date.Month));
            return isLastDayOfMonth;
            
        }
        

        public static int NthOfMonth(int Occurrence, DayOfWeek Day)
        {
            //int Occurrence = 0;
            //var fday = new DateTime(CurDate.Year, CurDate.Month, 1);
            var fday = DateTime.Now;
            var fOc = fday.DayOfWeek == Day ? fday : fday.AddDays(Day - fday.DayOfWeek);
            // CurDate = 2011.10.1 Occurance = 1, Day = Friday >> 2011.09.30 FIX. 
            if (fOc.Month < fday.Month) Occurrence = Occurrence + 1;
            return Occurrence;
        }


        public static DateTime GetLastFridayOfTheMonth(DateTime date)
        {
            var lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));

            while (lastDayOfMonth.DayOfWeek != DayOfWeek.Friday)
                lastDayOfMonth = lastDayOfMonth.AddDays(-1);

            return lastDayOfMonth;
        }

        
          

        public static DateTime GetLastWeekdayOfTheMonth(DateTime date, DayOfWeek dw)
        {
            var lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));

            while (lastDayOfMonth.DayOfWeek != dw)
                lastDayOfMonth = lastDayOfMonth.AddDays(-1);

            return lastDayOfMonth;
        }


        public static DateTime FirstDayOfMonth(this DateTime dt) {
        return new DateTime(dt.Year, dt.Month, 1);
        }


        //public static DateTime LastDayOfMonth(this DateTime dt)
        //{
        //    return dt.FirstDayOfMonth().AddMonths(1).AddDays(-1);
        //}
    }
}
