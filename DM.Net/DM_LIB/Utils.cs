/*
 * C#
 * User: CRuff
 * Date: 2/14/2019
 * Time: 11:11 AM
 * DM_Lib.Utils
 * 
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DM_Lib
{
    public static class Utils
    {
        public static string Left(string text, int charCount)
        {
            return text.Substring(text.Length - charCount, charCount);
        }

        public static string Right(string text, int charCount)
        {
            return text.Substring(0, charCount);
        }
        
        public static string Mid(string text, int charStart, int charCount)
        {	
        	string new_text;
        	try{
        		new_text = text.Substring(charStart, charCount);
        		return new_text;
        	}catch{
        		return text;
        	}
        }

        public static DateTime ConvertToDateTime(string str)
        {
            string pattern = @"(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})\.(\d{3})";
            if (Regex.IsMatch(str, pattern))
            {
                Match match = Regex.Match(str, pattern);
                int year = Convert.ToInt32(match.Groups[1].Value);
                int month = Convert.ToInt32(match.Groups[2].Value);
                int day = Convert.ToInt32(match.Groups[3].Value);
                int hour = Convert.ToInt32(match.Groups[4].Value);
                int minute = Convert.ToInt32(match.Groups[5].Value);
                int second = Convert.ToInt32(match.Groups[6].Value);
                int millisecond = Convert.ToInt32(match.Groups[7].Value);
                return new DateTime(year, month, day, hour, minute, second, millisecond);
            }
            else
            {
                throw new Exception("Unable to parse.");
            }
        }
    }
}