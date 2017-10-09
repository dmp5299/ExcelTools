using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ContentaDataExport.ContentaObjects;
using System.Reflection;
using ContentaDataExport.Styles;

namespace ContentaDataExport.ContentaClasses
{
    public class ContentaExcel
    {
        public static void BuilExcelFile(List<List<Record>> rows, string fileName)
        {
            string[] alphabet = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "k", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG" };
            try
            {
                ExcelPackage _package = new ExcelPackage(new MemoryStream());

                ContentaTotalExcelSheet totals = new ContentaTotalExcelSheet();

                OfficeOpenXml.Style.XmlAccess.ExcelNamedStyleXml headerStyle = ContentaDataExport.Styles.ExcelStyles.getHeaderStyle(_package);

                ExcelWorksheet totalsWorksheet = totals.calculateTotals(rows, ref _package);

                int i = 1;
                int taskCount = 0;
                foreach(List<Record> recordList in rows)
                {
                    var ws1 = _package.Workbook.Worksheets.Add(recordList[0].Project);
                    int count = rows.Count + 1;
                    int c1 = 0;
                    Record record1 = new Record();
                    foreach (PropertyInfo propertyInfo in record1.GetType().GetProperties().Where(p => p.GetCustomAttributes(typeof(SkipPropertyAttribute), true).Length == 0))
                    {
                        if(propertyInfo.Name == "RoutingTasks")
                        {
                            //routing task headers
                            foreach(RoutingRecord r in recordList[0].RoutingTasks)
                            {
                                
                                if(r.TASK == "SAC_Review")
                                {
                                    ws1.Cells[alphabet[c1] + 1.ToString()].Value = "LSA Review";
                                    ws1.Cells[alphabet[c1] + 1.ToString()].StyleName = headerStyle.Name;
                                    c1++;
                                    taskCount++;
                                    ws1.Cells[alphabet[c1] + 1.ToString()].Value = "Engineering Review";
                                    ws1.Cells[alphabet[c1] + 1.ToString()].StyleName = headerStyle.Name;
                                    c1++;
                                    taskCount++;
                                }
                                else
                                {
                                    ws1.Cells[alphabet[c1] + 1.ToString()].Value = r.TASK;
                                    ws1.Cells[alphabet[c1] + 1.ToString()].StyleName = headerStyle.Name;
                                    c1++;
                                    taskCount++;
                                }
                            }
                            //routing task review headers
                            foreach (RoutingRecord r in recordList[0].RoutingTasks)
                            {
                                if (r.TASK == "Air_Force_Review" || r.TASK == "writing")
                                {
                                    ws1.Cells[alphabet[c1] + 1.ToString()].Value = r.TASK + " Reviewer";
                                    ws1.Cells[alphabet[c1] + 1.ToString()].StyleName = headerStyle.Name;
                                    c1++;
                                    taskCount++;
                                }
                                else
                                {

                                }
                            }
                            
                        }
                        else
                        {
                            ws1.Cells[alphabet[c1] + 1.ToString()].Value = propertyInfo.Name;

                            ws1.Cells[alphabet[c1] + 1.ToString()].StyleName = headerStyle.Name;
                            c1++;
                        }
                       
                    }
                    int rowIndex = 2;
                    foreach (Record record in recordList)
                    {
                        int c = 0;
                        foreach (PropertyInfo propertyInfo in record.GetType().GetProperties().Where(p => p.GetCustomAttributes(typeof(SkipPropertyForExcelBodyAttribute),true).Length == 0))
                        {
                            if(propertyInfo.GetValue(record, null) != null)
                            {
                                if (propertyInfo.Name == "RoutingTasks")
                                {
                                    //done date cells
                                    foreach (RoutingRecord r in (List<RoutingRecord>)propertyInfo.GetValue(record, null))
                                    {
                                        if (r.TASK == "SAC_Review")
                                        {
                                            ws1.Cells[alphabet[c] + rowIndex.ToString()].Value = r.DONE_DATE;
                                            ws1.Cells[alphabet[c] + rowIndex.ToString()].Style.Font.Size = 11;
                                            c++;
                                        }
                                        ws1.Cells[alphabet[c] + rowIndex.ToString()].Value = r.DONE_DATE;
                                        ws1.Cells[alphabet[c] + rowIndex.ToString()].Style.Font.Size = 11;
                                        c++;
                                    }
                                    //review cells
                                    foreach (RoutingRecord r in (List<RoutingRecord>)propertyInfo.GetValue(record, null))
                                    {
                                        if (r.TASK == "Air_Force_Review" || r.TASK == "writing")
                                        {
                                            ws1.Cells[alphabet[c] + rowIndex.ToString()].Value = r.USER;
                                            ws1.Cells[alphabet[c] + rowIndex.ToString()].Style.Font.Size = 11;
                                            c++;
                                        }
                                            
                                    }
                                }
                                else
                                {
                                    ws1.Cells[alphabet[c] + rowIndex.ToString()].Value = propertyInfo.GetValue(record, null).ToString();
                                    ws1.Cells[alphabet[c] + rowIndex.ToString()].Style.Font.Size = 11;
                                    c++;
                                }
                            }
                        }
                        rowIndex++;
                    }
                    ws1.Column(1).Width = 7;
                    ws1.Column(2).Width = 5;
                    ws1.Column(3).Width = 5;
                    ws1.Column(4).Width = 5;
                    ws1.Column(5).Width = 7;
                    ws1.Column(6).Width = 5;
                    ws1.Column(7).Width = 5;
                    ws1.Column(8).Width = 5;
                    ws1.Column(9).Width = 5;
                    ws1.Column(10).Width = 5;
                    ws1.Column(11).Width = 5;
                    ws1.Column(12).Width = 35;
                    ws1.Column(13).Width = 40;
                    ws1.Column(14).Width = 40;
                    for(int ii = 15; ii < taskCount+15; ii++)
                    {
                        ws1.Column(ii).Width = 20;
                    }
                    i++;
                }
                /*
                foreach (string task in workflow)
                {
                    ws1.Cells[current + "1"].Value = task;
                    ws1.Cells[current + "1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ws1.Cells[current + "1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    current = (Char)(Convert.ToUInt16(current) + 1);
                }
                */
                _package.SaveAs(new FileInfo(fileName));
                System.Diagnostics.Process.Start(fileName);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
