using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Windows.Forms;
using _38_39Conversion.ExcelObjects;
using ExcelLibrary.SpreadSheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using _38_39Conversion.Utils;
using _38_39Conversion.Interfaces;
using NPOI.XSSF.UserModel;

namespace _38_39Conversion._38ConversionFiles
{
    public class Dash38Xls : I38Data
    {
        

        public IDictionary<string, object> parseThirtyEightFile(string file)
        {
            try
            {
                IDictionary<string, object> dict = new Dictionary<string, object>();
                HSSFWorkbook hssfwb;
                
                using (FileStream excelFile = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new HSSFWorkbook(excelFile);
                }
                ISheet sheet = hssfwb.GetSheetAt(0);
                ExcelXlsUtils xlsUtils = new ExcelXlsUtils();
                List<double> colWidths = xlsUtils.getColWidths(sheet);
                dict["file"] = file;
                dict["form"] = getCellReferenceValue("E1", sheet);
                dict["page"] = getCellReferenceValue("I1", sheet);
                dict["revision"] = getCellReferenceValue("E2", sheet);
                dict["revDate"] = getCellReferenceValue("I2", sheet);
                dict["model"] = getCellReferenceValue("B4", sheet);
                dict["deliverableNo"] = getCellReferenceValue("E4", sheet);
                dict["statDate"] = getCellReferenceValue("I4", sheet);
                dict["reviewedBy"] = getCellReferenceValue("C5", sheet);
                dict["date"] = getCellReferenceAsDate("I5", sheet);
                dict["author"] = getCellReferenceValue("C7", sheet);
                List<Item> items = new List<Item>();
                Boolean keepGoing = true;
                int i = getItemNoIndex(sheet)+1;
                int totalRowsDeleted = 0;
                while (keepGoing)
                {
                    int rowsToBeDeleted = 0;
                    string mergedValue = "";
                    string itemNo = "";
                    mergedValue = getMergedValue(sheet, i);
                    if (mergedValue == "")
                    {
                        if (getMergedValue(sheet, i + 1) == "")
                        {
                            break;
                        }
                    }
                    while (getCellReferenceValue(("A" + (i + 1)).ToString(), sheet).ToString() == "")
                    {
                        mergedValue += getMergedValue(sheet, (i + 1));
                        if (getMergedValue(sheet, (i + 1)) == "")
                        {
                            if (getMergedValue(sheet, i + 2) == "")
                                break;
                        }
                        i++;
                        rowsToBeDeleted++;
                    }
                    if (rowsToBeDeleted > 0)
                    {
                        double heightBefore = sheet.GetRow(1).Height;
                        totalRowsDeleted += rowsToBeDeleted;
                        xlsUtils.deleteRows(hssfwb, sheet, rowsToBeDeleted, ref i, file);
                        int lineCount = GenericExcelUtils.GetLineCount(mergedValue, (int)GenericExcelUtils.getRangeWidth(colWidths, 3, 6), "xls");
                        sheet.GetRow(i - 1).GetCell(3).CellStyle.WrapText = true;
                        sheet.GetRow(i - 1).Height = (short)((lineCount*heightBefore)+(100*(lineCount-1)));
                        sheet.GetRow(i-1).GetCell(3).SetCellValue(mergedValue);
                        sheet.GetRow(i - 1).GetCell(0).CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        sheet.GetRow(i - 1).GetCell(1).CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        sheet.GetRow(i - 1).GetCell(7).CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        sheet.GetRow(i - 1).GetCell(8).CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                        {
                            hssfwb.Write(fs);
                        }
                        
                    }
                    itemNo = getCellReferenceValue(("A" + i).ToString(), sheet).ToString();
                    Item item = new Item
                    {
                        ItemNo = itemNo,
                        Comment = mergedValue
                    };
                    items.Add(item);
                    i++;
                }
                if (totalRowsDeleted > 0)
                {
                    xlsUtils.addBlankRows(hssfwb, sheet, totalRowsDeleted, (getItemNoIndex(sheet) + 1) + items.Count, file);
                }
                dict.Add("items", items);
                return dict;
            }
            catch(Exception e)
            {
                MessageBox.Show("Exception in "+ file);
                return null;
            }
        }

        public void clean38File(string file, ExcelWorksheet sheet)
        {

        }

        public int getItemNoIndex(object sheet)
        {
            ISheet ws = (ISheet)sheet;
            DataFormatter formatter = new DataFormatter();
            for (var i = 1; i <= ws.LastRowNum; i++)
            {
                var row = ws.GetRow(i);
                for (var j = 0; j <= row.LastCellNum; j++)
                {
                    string val = formatter.FormatCellValue(row.GetCell(j));
                    if(val != "")
                    {
                        if (val.ToLower() == "item no.")
                        {
                            return i+1;
                        }
                    }
                }
                Console.WriteLine();
            }
            throw new ArgumentException("item no row index not found");
        }

        

        public string getMergedValue(object sheet, int row)
        {
            try
            {
                ISheet ws = (ISheet)sheet;
                string[] letters = { "D", "E", "F", "G" };
                string mergedValue = "";
                for (var index = 0; index < 4; index++)
                {
                    mergedValue += getCellReferenceValue((letters[index] + row).ToString(), ws);
                }
                return mergedValue;
            }
            catch(Exception)
            {
                MessageBox.Show("the exception is in getMergedValue");
                return null;
            }
        }

        

        public static string getCellReferenceValue(string cell, ISheet sheet)
        {
            try
            {
                DataFormatter formatter = new DataFormatter();
                var cr = new CellReference(cell);
                if (cr == null)
                {
                    return "";
                }
                else
                {
                    DateTime date;
                    var row = sheet.GetRow(cr.Row);
                    if(DateTime.TryParse(formatter.FormatCellValue(row.GetCell(cr.Col)), out date))
                    {
                        return StringUtils.formatWithWords(date);
                    }
                    else
                    {
                        return formatter.FormatCellValue(row.GetCell(cr.Col));
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("the exception is in getCellReferenceValue");
                return "";
            }
            
        }

        public static DateTime getCellReferenceAsDate(string cell, ISheet sheet)
        {
            var cr = new CellReference(cell);
            var row = sheet.GetRow(cr.Row);
            return row.GetCell(cr.Col).DateCellValue.Date;
        }

        public static double getCellReferenceAsNumber(string cell, ISheet sheet)
        {
            var cr = new CellReference(cell);
            var row = sheet.GetRow(cr.Row);
            return row.GetCell(cr.Col).NumericCellValue;
        }
    }
}
