using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PCMClientLib;
using PCMPortalLib;
using ContentaDataExport.Utils;
using System.Xml;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;
using ContentaDataExport.ContentaObjects;
using System.Collections;

namespace ContentaDataExport.ContentaClasses
{
    public class ContentaModule
    {
        public List<List<Record>> recordData = new List<List<Record>>();

        public void getContentaObjects(IPCMcommand command)
        {
            command.Select("/#1/#2/#6");
            IPCMdata S1000Dcontainers = command.ListChildren();
            string wipId = ContentaUtils.getWhip(S1000Dcontainers);
            getAllModules(command, wipId);
            ContentaExcel.BuilExcelFile(recordData, "c:/temp/contentaTaskList.xlsx");
        }

        public void getAllModules(IPCMcommand cmd, string wipId)
        {
            try
            {
                cmd.Select(wipId);

                IPCMdata children = cmd.ListChildren();
                for (int i = 0; i < children.RecordCount; i++)
                {
                    string objId = children.GetValueByLabel(i, "OBJECT_ID");
                    string projectName = children.GetValueByLabel(i, "NAME");

                    string id = (wipId + ("/" + objId));

                    cmd.Select(id);
                    IPCMdata children2 = cmd.ListChildren();

                    List<Record> records = new List<Record>();

                    using (System.IO.StreamWriter file =
            new System.IO.StreamWriter(@"C:\temp\tempText.txt"))
                    {
                        file.Write(cmd.ExecCmd("list history").XMLData);
                    }
                    MessageBox.Show("done");


                    for (int ii=1;ii < children2.RecordCount;ii++)
                    {
                        try
                        {
                            DmoduleRoot result;

                            string objId2 = children2.GetValueByLabel(ii, "OBJECT_ID");
                            
                            result = XmlUtils.SerializeXml(new DmoduleRoot(), cmd.Select(id + "/" + objId2).XMLData);

                            result.Record.Project = projectName;

                            string historyXml = cmd.ListHistory().XMLData;
                            
                            result.Record.CSDB_Creation = processCsdbCreation(historyXml);

                            double percentComplete = 5;

                            List<HistoryRecord> historyRecords = getObjectHistory(historyXml);
                            
                            result.Record.RoutingTasks = processRoutingHistory(cmd.ListRoutingHistory().XMLData, historyRecords, ref percentComplete, id + "/" + objId2);

                            result.Record.Percent_Complete = percentComplete.ToString();

                            records.Add(result.Record);
                        }
                        catch(Exception e)
                        {
                            MessageBox.Show(e.Message);
                        }
                    }
                    if (records.Count > 0)
                    {
                        recordData.Add(records);
                    }
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }
        
        public List<RoutingRecord> processRoutingHistory(string routingTasks, List<HistoryRecord> historyRecords, ref double percentComplete, string id)
        {
            RoutingHistoryRoot result =null;
            try
            {
                result = XmlUtils.SerializeXml(new RoutingHistoryRoot(), routingTasks);
            }
            catch(Exception e)
            {
                MessageBox.Show("Error in processRoutingHistory: " + e.Message);
            }
            return result.getAllRoutinStages(ref percentComplete, historyRecords, id);
        }

        public List<HistoryRecord> getObjectHistory(string historyData)
        {
            HistoryRoot result = XmlUtils.SerializeXml(new HistoryRoot(), historyData);
            return result.Records;
        }

        public string processCsdbCreation(string history)
        {
            HistoryRoot result = null;
            try
            {
                result = XmlUtils.SerializeXml(new HistoryRoot(), history);

            }
            catch (Exception e)
            {
                MessageBox.Show("Error in getCsdbCreation: " + e.Message);
            }
            return result.getCsdbCreation();
        }
    }
}
