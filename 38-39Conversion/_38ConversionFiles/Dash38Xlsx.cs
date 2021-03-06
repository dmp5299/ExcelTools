﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _38_39Conversion.Interfaces;
using OfficeOpenXml;
using _38_39Conversion.ExcelObjects;
using _38_39Conversion.Utils;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml.Style;

namespace _38_39Conversion._38ConversionFiles
{
    class Dash38Xlsx : I38Data
    {
        public IDictionary<string, object> parseThirtyEightFile(string file, Boolean clean)
        {
            IDictionary<string, object> dict = new Dictionary<string, object>();
            ExcelPackage package;
            try
            {
                package = new ExcelPackage(new FileInfo(file));
            }
            catch (Exception)
            {
                throw new Exception("Error opening file: " + file + ". Please check if this file is open.");
            }
            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
            ExcelUtils xlsxUtils = new ExcelXlsxUtils();
            List<double> colWidths = xlsxUtils.getColWidths(workSheet);
            dict["file"] = file;
            dict["form"] = getCell("E1", workSheet).Replace("38","39");
            dict["page"] = getCell("I1", workSheet);
            dict["revision"] = getCell("E2", workSheet);
            if (workSheet.Cells["I2"].Value != null)
            {
                long dateNum = Int32.Parse(workSheet.Cells["I2"].Value.ToString());
                dict["revDate"] = Utils.StringUtils.formatWithWords(DateTime.FromOADate(dateNum));
            }
            else
            {
                dict["revDate"] = "";
            }
            List<Item> items = new List<Item>();
            Boolean keepGoing = true;
            int i = getItemNoIndex(workSheet) + 1;
            int totalRowsDeleted = 0;
            while (keepGoing)
            {
                int rowsToBeDeleted = 0;
                string mergedValue = "";
                string itemNo = "";
                mergedValue = getMergedValue(workSheet, i);

                if (mergedValue == "")
                {
                    if (getMergedValue(workSheet, i + 1) == "")
                        break;
                }
                if (clean)
                {
                    while (workSheet.Cells["A" + (i + 1)].Value == null)
                    {

                        mergedValue += getMergedValue(workSheet, (i + 1));
                        if (getMergedValue(workSheet, (i + 1)) == "")
                        {
                            if (getMergedValue(workSheet, i + 2) == "")
                                break;
                        }
                        i++;
                        rowsToBeDeleted++;
                    }

                    if (rowsToBeDeleted > 0)
                    {
                        double heightBefore = workSheet.Row(1).Height;
                        totalRowsDeleted += rowsToBeDeleted;
                        xlsxUtils.deleteRows(package, workSheet, rowsToBeDeleted, ref i);
                        workSheet.Cells["D" + i + ":" + "G" + i].Style.WrapText = true;
                        int lineCount = GenericExcelUtils.GetLineCount(mergedValue, (int)GenericExcelUtils.getRangeWidth(colWidths, 3, 6));
                        workSheet.Row(i).Height = lineCount * heightBefore;
                        workSheet.Cells["D" + i + ":" + "G" + i].Value = mergedValue;
                        workSheet.Cells["A" + i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["B" + i + ":" + "C" + i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["I" + i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        workSheet.Cells["H" + i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
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
                i++;
            }
            if (totalRowsDeleted > 0)
            {
                xlsxUtils.addBlankRows(package, workSheet, totalRowsDeleted, (getItemNoIndex(workSheet) + 1) + items.Count);
            }
            dict.Add("items", items);
            return dict;
        }

        public string getCell(string cell, ExcelWorksheet worksheet)
        {
            if(worksheet.Cells[cell].Value == null)
            {
                return "";
            }
            else
            {
                return worksheet.Cells[cell].Value.ToString();
            }
        }
        public int getItemNoIndex(object sheet)
        {
            ExcelWorksheet workSheet = (ExcelWorksheet)sheet;
            var start = workSheet.Dimension.Start;
            var end = workSheet.Dimension.End;
            for (int row = start.Row; row <= end.Row; row++)
            { // Row by row...
                for (int col = start.Column; col <= end.Column; col++)
                { // ... Cell by cell...
                    string cellValue = workSheet.Cells[row, col].Text; // This got me the actual value I needed.
                    if (cellValue.ToLower() == "item no.")
                    {
                        return row;
                    }
                }
            }
            throw new ArgumentException("item no row index not found");
        }

        public string getMergedValue(object workSheet, int row)
        {
            ExcelWorksheet ws = (ExcelWorksheet)workSheet;
            var range = ws.Cells["D" + row + ":G" + row];
            string mergedValue = "";
            ExcelRangeBase prev = null;
            foreach (var rangeBase in range)
            {
                if(prev != null)
                {
                    if (rangeBase.Value == prev.Value)
                    {
                        break;
                    }
                }                
                mergedValue += rangeBase.Value;
                prev = rangeBase;
            }

            return mergedValue;
        }
    }
}
