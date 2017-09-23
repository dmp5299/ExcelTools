using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Windows.Forms;
using _38_39Conversion.ExcelStyles;
using _38_39Conversion.ExcelObjects;

namespace _38_39Conversion._38ConversionFiles
{
    class Dash39
    {
        public static void create39Files(List<IDictionary<string, object>> _38data)
        {
            foreach(IDictionary<string, object> data in _38data)
            {
                build39File(data);
            }
        }

        public static void xlsBuild39File(IDictionary<string, object> data)
        {

        }

        public static void build39File(IDictionary<string,object> data)
        {
            ExcelPackage _package = new ExcelPackage(new MemoryStream());
            var ws1 = _package.Workbook.Worksheets.Add("Worksheet1");

            var boldCenter = CellStyles.GetBoldCenter(ws1);
            var boldRight = CellStyles.GetBoldRight(ws1);

            //first row

            ws1.Cells["A1"].Value = "S&T FORM";
            ws1.Cells["A1"].StyleName = "BoldCenter";

            ws1.Cells["B1"].Value = "Form:";
            ws1.Cells["B1"].StyleName = "BoldRight";

            ws1.Cells["C1"].Value = data["form"].ToString();
            ws1.Cells["C1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws1.Cells["C1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;

            ws1.Cells["D1"].Value = "Page:";
            ws1.Cells["D1"].StyleName = "BoldRight";

            ws1.Cells["E1"].Value = data["page"].ToString();
            ws1.Cells["E1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws1.Cells["E1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;

            //second row

            ws1.Cells["A2"].Value = "";
            ws1.Cells["A2"].StyleName = "BoldCenter";

            ws1.Cells["B2"].Value = "Revision:";
            ws1.Cells["B2"].StyleName = "BoldRight";

            ws1.Cells["C2"].Value = data["revision"].ToString();
            ws1.Cells["C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws1.Cells["C2"].Style.Border.Right.Style = ExcelBorderStyle.Medium;

            ws1.Cells["D2"].Value = "Revision Date:";
            ws1.Cells["D2"].StyleName = "BoldRight";

            ws1.Cells["E2"].Value = data["revDate"].ToString();
            ws1.Cells["E2"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws1.Cells["E2"].Style.Border.Right.Style = ExcelBorderStyle.Medium;

            ws1.Cells["B4:E4"].Value = "PRODUCT IMPROVEMENT RESPONSE WORKSHEET";
            ws1.Cells["B4:E4"].Merge = true;
            ws1.Cells["B4:E4"].Style.Font.Bold = true;
            ws1.Cells["B4:E4"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            ws1.Cells["B5:E5"].Value = "Disposition all Invalid comments from SA20038 form";
            ws1.Cells["B5:E5"].Merge = true;
            ws1.Cells["B5:E5"].Style.Font.Bold = true;
            ws1.Cells["B5:E5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            ws1.Cells["A7"].Value = "ITEM NO";
            ws1.Cells["A7"].StyleName = "BoldCenter";
            ws1.Cells["A7"].Style.Border.Top.Style = ExcelBorderStyle.Medium;

            ws1.Cells["B7:E7"].Value = "Disposition of Invalid Comment";
            ws1.Cells["B7:E7"].Merge = true;
            ws1.Cells["B7:E7"].StyleName = "BoldCenter";
            ws1.Cells["B7:E7"].Style.Border.Top.Style = ExcelBorderStyle.Medium;

            List<Item> items = (List<Item>)data["items"];

            int cellRowIndex = 8;
            foreach(Item item in items)
            {
                ws1.Cells["A"+ cellRowIndex].Value = item.ItemNo;
                if (cellRowIndex == 37)
                {
                    ws1.Cells["A" + cellRowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                }
                else
                {
                    ws1.Cells["A" + cellRowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                ws1.Cells["A" + cellRowIndex].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ws1.Cells["B" + cellRowIndex + ":E" + cellRowIndex].Value = item.Comment;
                ws1.Cells["B" + cellRowIndex + ":E" + cellRowIndex].Style.WrapText = true;
                ws1.Cells["B" + cellRowIndex + ":E" + cellRowIndex].Merge = true;
                if (cellRowIndex == 37)
                {
                    ws1.Cells["B" + cellRowIndex + ":E" + cellRowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                }
                else
                {
                    ws1.Cells["B" + cellRowIndex + ":E" + cellRowIndex].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                ws1.Cells["B" + cellRowIndex + ":E" + cellRowIndex].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                cellRowIndex++;
            }
            cellRowIndex += 1;

            ws1.Cells["A" + cellRowIndex + ":E" + cellRowIndex].Value = "VERIFY CURRRENT REVISION OF FORM PRIOR TO USE";
            ws1.Cells["A" + cellRowIndex + ":E" + cellRowIndex].Style.WrapText = true;
            ws1.Cells["A" + cellRowIndex + ":E" + cellRowIndex].Merge = true;
            ws1.Cells["A" + cellRowIndex + ":E" + cellRowIndex].Style.Font.Bold = true;
            ws1.Cells["A" + cellRowIndex + ":E" + cellRowIndex].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            ws1.Column(1).Width = 20;
            ws1.Column(2).Width = 25;
            ws1.Column(4).Width = 20;
            ws1.Column(5).Width = 22;
            string fileWithoutExtension = data["file"].ToString().Substring(0, data["file"].ToString().IndexOf('.'))+"-39"+".xlsx";
            try
            {
                _package.SaveAs(new FileInfo(fileWithoutExtension));
            }
            catch(Exception i)
            {
                MessageBox.Show(i.Message);
            }
        }
    }
}
