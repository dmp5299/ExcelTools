using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ContentaDataExport.Utils
{
    public class CalculationUtils
    {
        public static double percent(int total, double count)
        {
            return ((count / total) * 100);
        }
    }
}
