using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace _38_39Conversion.ExcelStyles
{
    public class CellStyles
    {
        public static OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml GetBoldCenter(ExcelWorksheet ws1)
        {
            var boldCenter = ws1.Workbook.Styles.CreateNamedStyle("BoldCenter");
            boldCenter.Style.Font.Bold = true;
            boldCenter.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            boldCenter.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            boldCenter.Style.Border.Right.Style = ExcelBorderStyle.Medium;
            return boldCenter;
        }

        public static OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml GetBoldRight(ExcelWorksheet ws1)
        {
            var boldRight = ws1.Workbook.Styles.CreateNamedStyle("BoldRight");
            boldRight.Style.Font.Bold = true;
            boldRight.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            boldRight.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            return boldRight;
        }
    }
}
