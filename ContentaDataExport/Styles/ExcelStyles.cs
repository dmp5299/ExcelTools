using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ContentaDataExport.Styles
{
    public class ExcelStyles
    {
        public static OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml getHeaderStyle(ExcelPackage package)
        {
            OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml headerStyle = package.Workbook.Styles.CreateNamedStyle("header");
            headerStyle.Style.TextRotation = 90;
            headerStyle.Style.Font.Size = 11;
            headerStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerStyle.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#4E7FC1"));

            return headerStyle;
        }

        public static OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml getBoldHeaderStyle(ExcelPackage package)
        {
            OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml headerStyle = package.Workbook.Styles.CreateNamedStyle("headerBold");
            headerStyle.Style.TextRotation = 90;
            headerStyle.Style.Font.Size = 11;
            headerStyle.Style.Font.Bold = true;
            headerStyle.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            headerStyle.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#4E7FC1"));

            return headerStyle;
        }
    }
}
