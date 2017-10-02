using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using OfficeOpenXml;

namespace _38_39Conversion.Interfaces
{
    public interface I38Data
    {
        IDictionary<string, object> parseThirtyEightFile(string file, Boolean clean);
        int getItemNoIndex(object sheet);
        string getMergedValue(object sheet, int row);
    }
}
