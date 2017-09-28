using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using ExcelLibrary.SpreadSheet;
using _38_39Conversion.Interfaces;
using System.IO;
using System.Windows.Forms;
using NPOI.XSSF.UserModel;

namespace _38_39Conversion.Utils
{
    public class ExcelXlsUtils : ExcelUtils
    {
        public List<double> getColWidths(object ws)
        {
            try
            {
                ISheet sheet = (ISheet)ws;
                IRow r = sheet.GetRow(sheet.FirstRowNum);
                int colCount = r.Cells.Count;
                List<double> colWidths = new List<double>(colCount);
                for (int i = 0; i < colCount; i++)
                {
                    colWidths.Add(sheet.GetColumnWidth(i));
                }
                return colWidths;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
        }
        
        public void addBlankRows(object wb, object ws, int numberRowsToAdd, int startIndex, string file = "")
        {
            HSSFWorkbook xlsWb = (HSSFWorkbook)wb;
            HSSFSheet xlsWs = (HSSFSheet)ws;
            ICellStyle cellStyle = xlsWb.CreateCellStyle();
            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            for (int i = 1; i < numberRowsToAdd; i++)
            {
                xlsWs.ShiftRows(startIndex-1, xlsWs.LastRowNum, 1);
                IRow row = xlsWs.CreateRow(startIndex - 1);

                ICell cell1  = row.CreateCell(0);
                cell1.SetCellValue("");
                cell1.CellStyle = cellStyle;

                CellRangeAddress address1 = new CellRangeAddress(startIndex - 1, startIndex - 1, 1, 2);
                xlsWs.AddMergedRegion(address1);
                ICell cell2 = row.CreateCell(1);
                cell2.SetCellValue("");
                RegionUtil.SetBorderBottom(1, address1, xlsWs, xlsWb);
                RegionUtil.SetBorderRight(1, address1, xlsWs, xlsWb);

                CellRangeAddress address2 = new CellRangeAddress(startIndex - 1, startIndex - 1, 3, 6);
                xlsWs.AddMergedRegion(address2);
                ICell cell3 = row.CreateCell(3);
                cell3.SetCellValue("");
                RegionUtil.SetBorderBottom(1, address2, xlsWs, xlsWb);
                RegionUtil.SetBorderRight(1, address2, xlsWs, xlsWb);

                ICell cell4 = row.CreateCell(7);
                cell4.SetCellValue("");
                cell4.CellStyle = cellStyle;

                ICell cell5 = row.CreateCell(8);
                cell5.SetCellValue("");
                cell5.CellStyle = cellStyle;
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    xlsWb.Write(fs);
                }
                startIndex++;
            }
        }
        
        public void deleteRows(object wb, object ws, int rowsToBeDeleted, ref int rowIndex, string file = "")
        {
            HSSFWorkbook xlsWb = (HSSFWorkbook)wb;
            HSSFSheet xlsWs = (HSSFSheet)ws;
            for (int i = 0; i < rowsToBeDeleted; i++)
            {
                xlsWs.RemoveRow(xlsWs.GetRow(rowIndex-1));
                xlsWs.ShiftRows(rowIndex, xlsWs.LastRowNum, -1);
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    xlsWb.Write(fs);
                }
                rowIndex--;
            }
        }
    }
}
