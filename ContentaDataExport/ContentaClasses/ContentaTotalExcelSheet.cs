using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using ContentaDataExport.ContentaObjects;
using System.Windows.Forms;
using ContentaDataExport.Utils;

namespace ContentaDataExport.ContentaClasses
{
    public class ContentaTotalExcelSheet
    {
        string[] totalHeaders = { "System", "Title", "Total DMs", "CSDB Creation", "CSDB Creation (Count)",
        "Writing (%)", "Writing (Count)", "IATR (%)","IATR (Count)", "LSA Review (%)", "LSA Review (Count)",
        "Engineering Review (%)","Engineering Review (Count)", "SIK Comments (%)","SIK Comments (Count)","Accepted (%)","Accepted (Count)","Percent Complete"};

        public ExcelWorksheet calculateTotals(List<List<Record>> recordList, ref ExcelPackage wb)
        {
            ExcelWorksheet totalWorksheet = wb.Workbook.Worksheets.Add("Overall");
            try
            {
                createTotalHeaders(totalWorksheet, wb);
            }
            catch(Exception e)
            {
                throw new Exception("Error in calculateTotals: " + e.Message);
            }
            int row = 2;
            foreach(List<Record> list in recordList)
            {
                int total = list.Count;
                totalWorksheet.Cells["A" + row.ToString()].Value = list[0].Project.Substring(list[0].Project.IndexOf(' ')+1);
                totalWorksheet.Cells["A" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["C" + row.ToString()].Value = total;
                totalWorksheet.Cells["C" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["D" + row.ToString()].Value = CalculationUtils.percent(total,list.Where(e => (e.CSDB_Creation != null) && (e.CSDB_Creation != "")).Count()) + "%";
                totalWorksheet.Cells["D" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["E" + row.ToString()].Value = list.Where(e => (e.CSDB_Creation != null) && (e.CSDB_Creation != "")).Count();
                totalWorksheet.Cells["E" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int writingTotal = list.SelectMany(i => i.RoutingTasks)
                                   .Where(round => round.TASK == "writing" && round.DONE_DATE != null && round.DONE_DATE != "")
                                   .Count();
                totalWorksheet.Cells["F" + row.ToString()].Value = CalculationUtils.percent(total,writingTotal) + "%";
                totalWorksheet.Cells["F" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["G" + row.ToString()].Value = writingTotal;
                totalWorksheet.Cells["G" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int IatrTotal = list.SelectMany(i => i.RoutingTasks)
                                   .Where(round => round.TASK == "RCM_ATR" && round.DONE_DATE != null && round.DONE_DATE != "")
                                   .Count();

                totalWorksheet.Cells["H" + row.ToString()].Value = CalculationUtils.percent(total, IatrTotal) + "%";
                totalWorksheet.Cells["H" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["I" + row.ToString()].Value = IatrTotal;
                totalWorksheet.Cells["I" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int SacTotal = list.SelectMany(i => i.RoutingTasks)
                                   .Where(round => round.TASK == "SAC_Review" && round.DONE_DATE != null && round.DONE_DATE != "")
                                   .Count();

                totalWorksheet.Cells["J" + row.ToString()].Value = CalculationUtils.percent(total, SacTotal) + "%";
                totalWorksheet.Cells["J" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["K" + row.ToString()].Value = SacTotal;
                totalWorksheet.Cells["K" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["L" + row.ToString()].Value = CalculationUtils.percent(total, SacTotal) + "%";
                totalWorksheet.Cells["L" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["M" + row.ToString()].Value = SacTotal;
                totalWorksheet.Cells["M" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int AirForceReviewTotal = list.SelectMany(i => i.RoutingTasks)
                                   .Where(round => round.TASK == "Air_Force_Review" && round.DONE_DATE != null && round.DONE_DATE != "")
                                   .Count();

                totalWorksheet.Cells["N" + row.ToString()].Value = CalculationUtils.percent(total, AirForceReviewTotal) + "%";
                totalWorksheet.Cells["N" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["O" + row.ToString()].Value = AirForceReviewTotal;
                totalWorksheet.Cells["O" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int endTotal = list.SelectMany(i => i.RoutingTasks)
                                   .Where(round => round.TASK == "End" && round.DONE_DATE != null && round.DONE_DATE != "")
                                   .Count();
                
                totalWorksheet.Cells["P" + row.ToString()].Value = CalculationUtils.percent(total, endTotal) + "%";
                totalWorksheet.Cells["P" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                totalWorksheet.Cells["Q" + row.ToString()].Value = endTotal;
                totalWorksheet.Cells["Q" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                double percentTotal = (list.Sum(e => Convert.ToDouble(e.Percent_Complete))/total);

                totalWorksheet.Cells["R" + row.ToString()].Value = percentTotal;
                totalWorksheet.Cells["R" + row.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                row++;
            }

            return totalWorksheet;
        }

        public void createTotalHeaders(ExcelWorksheet ws, ExcelPackage wb)
        {
            try
            {
                OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml headerStyle = ContentaDataExport.Styles.ExcelStyles.getBoldHeaderStyle(wb);
                char iterator = 'A';
                for (int i = 0; i < totalHeaders.Length; i++)
                {
                    ws.Cells[iterator + 1.ToString()].Value = totalHeaders[i];
                    ws.Cells[iterator + 1.ToString()].StyleName = headerStyle.Name;
                    iterator++;
                }
            }
            catch(Exception e)
            {
                throw new Exception("Error in createTotalHeaders: " + e.Message);
            }
        }
    }
}
