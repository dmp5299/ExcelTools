using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _38_39Conversion.Interfaces
{
    interface ExcelUtils
    {
        List<double> getColWidths(object ws);
        void addBlankRows(object wb, object ws, int numberRowsToAdd, int startIndex, string file = "");
        void deleteRows(object wb, object ws, int rowsToBeDeleted, ref int rowIndex, string file = "");
    }
}
