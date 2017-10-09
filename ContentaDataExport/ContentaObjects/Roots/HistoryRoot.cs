using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ContentaDataExport.ContentaObjects
{
    [XmlRoot("Data")]
    public class HistoryRoot
    {
        [XmlElement("Record")]
        public List<HistoryRecord> Records { get; set; }

        public string getCsdbCreation()
        {
            try
            {
                return Records.Where(x => x.OPERATION == "Transfer").First().DATETIME;
            }
            catch(Exception e)
            {
                throw new Exception("Error in getCsdbCreation: " + e.Message);
            }
        }
    }
}
