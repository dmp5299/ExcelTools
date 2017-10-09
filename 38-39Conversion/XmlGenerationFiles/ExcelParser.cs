using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using _38_39Conversion.ExcelObjects;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using _38_39Conversion.Utils;
using System.ComponentModel;

namespace _38_39Conversion.XmlGenerationFiles
{
    public class ExcelParser
    {
        ExcelPackage package;
        ExcelWorksheet workSheet;

        public List<_411Module> _411s { get; set; }

        public ExcelParser(string excelFile)
        {
            try
            {
                package = new ExcelPackage(new FileInfo(excelFile));
            }
            catch(Exception)
            {
                throw new Exception("Error opening " + excelFile + ". Check if this file is already open.");
            }
        }

        public void getExcelData(string filePath)
        {
            workSheet = package.Workbook.Worksheets.FirstOrDefault();

            List<_38_39Conversion.ExcelObjects.ExcelRow> newcollection = workSheet.ConvertSheetToObjects<_38_39Conversion.ExcelObjects.ExcelRow>().OrderBy(o => o._411DmcTitle).ToList();
            _411s = build411Modules(newcollection, filePath);
        }

        public static void get411Dms(BackgroundWorker worker, List<_411Module> _411s)
        {
            _411Xml.build411Dms(_411s, worker);
        }

        public List<_411Module> build411Modules(List<_38_39Conversion.ExcelObjects.ExcelRow> rows, string filePath)
        {
            List<_411Module> _411s = new List<_411Module>();
            for (int i=0;i<rows.Count;i++)
            {
                string _411Dmc = rows[i]._411DMC;
                string _411DmcTitle = rows[i]._411DmcTitle;
                if(!string.IsNullOrEmpty(_411DmcTitle))
                {
                    List<FaultIsolation> faultIsolation = new List<FaultIsolation>();
                    FaultIsolation f = new FaultIsolation()
                    {
                        FaultCode = rows[i].Id,
                        FailureName = rows[i].FailureName,
                        MaintenanceTaskName = rows[i].MaintenanceTaskName,
                        FaultIsolationProcedureId = rows[i].FaultIsolationProcedureId,
                        _920DmcTitle = rows[i]._920DmcTitle,
                        _920DMC = rows[i]._920DMC,
                        Name = rows[i].Name
                    };
                    _920Module _920 = new _920Module()
                    {
                        _920DmcTitle = rows[i]._920DmcTitle,
                        _920DMC = rows[i]._920DMC
                    };
                    faultIsolation.Add(buildFaultIsolationObject(rows[i].Name, rows[i].Id, rows[i].FailureName, rows[i].MaintenanceTaskName, rows[i].FaultIsolationProcedureId, rows[i]._920DmcTitle, rows[i]._920DMC));
                    int oldRow = i;
                    while ((i+1 < rows.Count) && (rows[i + 1]._411DmcTitle == rows[oldRow]._411DmcTitle))
                    {
                        faultIsolation.Add(buildFaultIsolationObject(rows[i + 1].Name,rows[i + 1].Id, rows[i + 1].FailureName, rows[i + 1].MaintenanceTaskName, rows[i + 1].FaultIsolationProcedureId,
                            rows[i + 1]._920DmcTitle, rows[i + 1]._920DMC));
                        if ((i + 1) == rows.Count)
                            break;
                        else
                            i++;
                    }
                    _411s.Add(build411Object(faultIsolation, _411DmcTitle, _411Dmc,_920, filePath));
                }
               
            }
            return _411s;
        }

        public FaultIsolation buildFaultIsolationObject(string name, string faultCode, string failureName, string maintenanceTaskName, string faultIsolationProcedureId, string _920DmcTitle, string _920DMC)
        {
            return new FaultIsolation()
            {
                Name = name,
                FailureName = failureName,
                FaultCode = faultCode,
                MaintenanceTaskName = maintenanceTaskName,
                FaultIsolationProcedureId = faultIsolationProcedureId,
                _920DmcTitle = _920DmcTitle,
                _920DMC = _920DMC
            };
        }

        public _411Module build411Object(List<FaultIsolation> faultIsolationList, string _411DmcTitle, string _411Dmc, _920Module _920, string filePath)
        {
            return new _411Module()
            {
                FaultIsolationElements = faultIsolationList,
                _920Element = _920,
                _411DmcTitle = _411DmcTitle,
                _411DMC = _411Dmc,
                excelPath = Path.GetDirectoryName(filePath)
            };
        }
    }
}
