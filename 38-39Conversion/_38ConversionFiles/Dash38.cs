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

namespace _38_39Conversion._38ConversionFiles
{
    public class Dash38
    {
        public static IDictionary<string, object> parseThirtyEightFile(string file)
        {
            IDictionary<string, object> dict = new Dictionary<string, object>();
            var package = new ExcelPackage(new FileInfo(file));
            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();

            dict["file"] = file;
            dict["form"] = workSheet.Cells["E1"].Value.ToString();
            dict["page"] = workSheet.Cells["I1"].Value.ToString();
            dict["revision"] = workSheet.Cells["E2"].Value.ToString();
            long dateNum = Int32.Parse(workSheet.Cells["I2"].Value.ToString());
            dict["revDate"] = DateTime.FromOADate(dateNum);
            dict["model"] = workSheet.Cells["B4:C4"].Value.ToString();
            dict["deliverableNo"] = workSheet.Cells["E4"].Value.ToString();
            dict["statDate"] = workSheet.Cells["I4"].Value.ToString();
            dict["reviewedBy"] = workSheet.Cells["C5:E5"].Value.ToString();
            dict["date"] = workSheet.Cells["I5"].Value.ToString();
            dict["author"] = workSheet.Cells["C7:E7"].Value.ToString();
            List<Item> items = new List<Item>();
            for(int i = 15;i <= 44;i++)
            {
                string mergedValue = "";
                string itemNo = "";
                var range = workSheet.Cells["D" + i + ":G" + i];

                foreach (var rangeBase in range)
                {
                    mergedValue += rangeBase.Value;
                }

                if (workSheet.Cells["A" + i].Value != null)
                {
                    itemNo = workSheet.Cells["A" + i].Value.ToString();
                }
                Item item = new Item
                {
                    ItemNo = itemNo,
                    Comment = mergedValue
                };
                items.Add(item);
            }
            dict.Add("items",items);
            return dict;
        }

        public static IDictionary<string, object> parseThirtyEightXlsFile(string file)
        {
            IDictionary<string, object> dict = new Dictionary<string, object>();
            HSSFWorkbook hssfwb;
            using (FileStream excelFile = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(excelFile);
            }

            ISheet sheet = hssfwb.GetSheetAt(0);

            dict["file"] = file;
            
            dict["form"] = getCellReferenceValue("E1", sheet);
            dict["page"] = getCellReferenceValue("I1", sheet);
            dict["revision"] = getCellReferenceValue("E2", sheet);
            dict["revDate"] = getCellReferenceAsDate("I2", sheet);
            dict["model"] = getCellReferenceValue("B4", sheet);
            dict["deliverableNo"] = getCellReferenceValue("E4", sheet);
            dict["statDate"] = getCellReferenceAsDate("I4", sheet);
            dict["reviewedBy"] = getCellReferenceValue("C5", sheet);
            dict["date"] = getCellReferenceAsDate("I5", sheet);
            dict["author"] = getCellReferenceAsDate("C7", sheet);
            List<Item> items = new List<Item>();
            for (int i = 15; i <= 44; i++)
            {
                string mergedValue = "";
                string itemNo = "";
                string[] letters = { "D", "E", "F", "G" };
                for (var index = 0; index < 4; index++)
                {
                    mergedValue += getCellReferenceValue((letters[index] + i).ToString(),sheet);
                }
                if (getCellReferenceAsNumber( ("A"+i).ToString(),sheet ) > 0)
                {
                    itemNo = getCellReferenceAsNumber(("A" + i).ToString(), sheet).ToString();
                }
                Item item = new Item
                {
                    ItemNo = itemNo,
                    Comment = mergedValue
                };
                items.Add(item);
            }
            dict.Add("items", items);
            return dict;
        }

        public static string getCellReferenceValue(string cell, ISheet sheet)
        {
            var cr = new CellReference(cell);
            var row = sheet.GetRow(cr.Row);
            return row.GetCell(cr.Col).StringCellValue;
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
