using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _38_39Conversion.Utils
{
    public class StringUtils
    {
        public static string FirstCharacterToLower(string str)
        {
            if (String.IsNullOrEmpty(str) || Char.IsLower(str, 0))
                return str;

            return Char.ToLowerInvariant(str[0]) + str.Substring(1);
        }

        public static string formatDateMMDDYYYY(DateTime date)
        {
            return date.Month + "/" + date.Day + "/" + date.Year;
        }

        public static string formatWithWords(DateTime date)
        {
            
            return date.ToString("MMMM") + " " + date.Day + ", " + date.Year;
        }

        public static string getInts(string var)
        {
            return new String(var.Where(Char.IsDigit).ToArray());
        }
    }
}
