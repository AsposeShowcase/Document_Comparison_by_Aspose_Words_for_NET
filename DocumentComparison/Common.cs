using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DocumentComparison
{
    public class Common
    {
        public static string DataDir = "~/UserFiles/";
        public static string tempDir = "~/Temp/";
        public static string Success = "success";
        public static string Error = "error";
        public static string[] separator = {"|**|"};

        /// <summary>
        /// Get the size of file in string format
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public static string DisplaySize(long? size)
        {
            if (size == null)
                return string.Empty;
            else
            {
                if (size < 1024)
                    return string.Format("{0:N0} bytes", size.Value);
                else
                    return String.Format("{0:N0} KB", size.Value / 1024);
            }
        }

        /// <summary>
        /// Get the date time in terms of days, weeks passed
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static string FormatDate(DateTime dateTime)
        {
            string result = dateTime.ToString();

            // If today, then just display the time
            if (dateTime.Date == DateTime.Now.Date)
                result = dateTime.ToString("t");
            // If this year, then display month name and day of month
            else if (dateTime.Year == DateTime.Now.Year)
                result = dateTime.ToString("MMM d");
            // For previous year and all other emails, display mm/dd/yy
            else
                result = dateTime.ToString("M/d/y");

            return result;
        }
    }
}