using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Serialization;

namespace ContentaDataExport.ContentaObjects
{
    [XmlRoot("Data")]
    public class RoutingHistoryRoot
    {
        [XmlElement("Record")]
        public List<RoutingRecord> Records { get; set; }

        //get routing task information for single module to put in excel row
        public List<RoutingRecord> getAllRoutinStages(ref double percentComplete, List<HistoryRecord> historyRecords, string id)
        {
            List<RoutingRecord> stageValues = new List<RoutingRecord>();
            Dictionary<string, string> routingTaskPairs = new Dictionary<string, string>()
            {
                { "writing", "RCM_ATR" },
                { "RCM_ATR", "SAC_Review" },
                { "SAC_Review", "Air_Force_Review" },
                { "Air_Force_Review", "End" },
                { "End", "this is the end" }
            };
            Dictionary<string, double> taskPercentCompleteValues = new Dictionary<string, double>()
            {
                { "writing", 30},
                { "RCM_ATR", 7.5 },
                { "SAC_Review", 6},
                { "Air_Force_Review", 2 },
                { "End",  1},
            };
            try
            {
                //iterate through routing pairs and check if there is a done date
                foreach (KeyValuePair<string, string> entry in routingTaskPairs)
                {
                    //check if routing task is in dmodule routing history
                    if (Records.Where(p => p.TASK == entry.Key).Count() > 0)
                    {
                        int lastOccuranceIndex = Records.IndexOf(Records.Where(p => p.TASK == entry.Key).Last());
                        RoutingRecord lastOccuranceRecord = Records[lastOccuranceIndex];
                        if (entry.Key == "End")
                        {
                            percentComplete += Convert.ToDouble(taskPercentCompleteValues[Records[lastOccuranceIndex].TASK]);
                            //HistoryRecord lastTransfer = getLastForward(lastOccuranceIndex, historyRecords);
                            lastOccuranceRecord.DONE_DATE = stageValues.Last().DONE_DATE;
                            lastOccuranceRecord.USER = stageValues.Last().USER;
                            stageValues.Add(lastOccuranceRecord);
                        }
                        else if (lastOccuranceIndex + 1 != Records.Count)
                        {
                            if (Records[lastOccuranceIndex + 1].TASK == entry.Value)
                            {
                                percentComplete += Convert.ToDouble(taskPercentCompleteValues[Records[lastOccuranceIndex].TASK]);
                                lastOccuranceRecord.USER = historyRecords.Where(x => (x.DATETIME.ToString() == lastOccuranceRecord.DONE_DATE.ToString()) && 
                                x.OPERATION == "Forward").FirstOrDefault() == null ? "" : historyRecords.Where(x => (x.DATETIME.ToString() == lastOccuranceRecord.DONE_DATE.ToString()) &&
                                x.OPERATION == "Forward").FirstOrDefault().USER;
                                stageValues.Add(lastOccuranceRecord);
                            }
                            else
                            {
                                stageValues.Add(new RoutingRecord()
                                {
                                    TASK = entry.Key,
                                    USER = "",
                                    ROLE = "",
                                    DONE_DATE = ""
                                });
                            }
                        }
                        else
                        {
                            stageValues.Add(new RoutingRecord()
                            {
                                TASK = entry.Key,
                                USER = "",
                                ROLE = "",
                                DONE_DATE = ""
                            });
                        }
                    }
                    else
                    {
                        stageValues.Add(
                            new RoutingRecord()
                            {
                                TASK = entry.Key,
                                USER = "",
                                ROLE = "",
                                DONE_DATE = ""
                            }
                        );
                    }
                }
                
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception in getAllRoutinStages " + id + ": " + e.Message);
            }

            return stageValues;
        }

        private HistoryRecord getLastForward(int index, List<HistoryRecord> historyRecords)
        {
            for(int i = index-1;index > 0;index--)
            {
                if(historyRecords[index].OPERATION == "Forward")
                {
                    return historyRecords[index];
                }
            }
            throw new Exception("Error in getLastForward: \"Accepted date could not be found\"");
        }

    }
}
