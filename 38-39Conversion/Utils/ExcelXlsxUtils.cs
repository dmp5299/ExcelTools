using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _38_39Conversion.Interfaces;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace _38_39Conversion.Utils
{
    public class ExcelXlsxUtils : ExcelUtils
    {
        public List<double> getColWidths(object ws)
        {
            ExcelWorksheet XlsxWs = (ExcelWorksheet)ws;
            var start = XlsxWs.Dimension.Start;
            var end = XlsxWs.Dimension.End;
            List<double> colWidths = new List<double>(10);
            for (int i = 0; i < end.Column; ++i)
            {
                colWidths.Add(XlsxWs.Column(i + 1).Width);
            }
            return colWidths;
        }

        public void addBlankRows(object wb, object ws, int numberRowsToAdd, int startIndex, string file = "")
        {
            ExcelPackage xlsxWb = (ExcelPackage)wb;
            ExcelWorksheet xlsxWs = (ExcelWorksheet)ws;
            for (int i = 0; i < numberRowsToAdd; i++)
            {
                xlsxWs.InsertRow(startIndex, 1);
                xlsxWs.Cells["A" + startIndex].Value = "";
                xlsxWs.Cells["A" + startIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                xlsxWs.Cells["A" + startIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                xlsxWs.Cells["B" + startIndex + ":C" + startIndex].Value = "";
                xlsxWs.Cells["B" + startIndex + ":C" + startIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                xlsxWs.Cells["B" + startIndex + ":C" + startIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                xlsxWs.Cells["B" + startIndex + ":C" + startIndex].Merge = true;

                xlsxWs.Cells["D" + startIndex + ":G" + startIndex].Value = "";
                xlsxWs.Cells["D" + startIndex + ":G" + startIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                xlsxWs.Cells["D" + startIndex + ":G" + startIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                xlsxWs.Cells["D" + startIndex + ":G" + startIndex].Merge = true;

                xlsxWs.Cells["H" + startIndex].Value = "";
                xlsxWs.Cells["H" + startIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                xlsxWs.Cells["H" + startIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                xlsxWs.Cells["I" + startIndex].Value = "";
                xlsxWs.Cells["I" + startIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                xlsxWs.Cells["I" + startIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                startIndex++;
            }
            xlsxWb.Save();
        }

        public void deleteRows(object wb, object ws, int rowsToBeDeleted, ref int rowIndex, string file = "")
        {
            ExcelPackage xlsxWb = (ExcelPackage)wb;
            ExcelWorksheet xlsxWs = (ExcelWorksheet)ws;
            for (int i = 0; i < rowsToBeDeleted; i++)
            {
                xlsxWs.DeleteRow(rowIndex);
                rowIndex--;
                xlsxWb.Save();
            }
        }
    }
}
