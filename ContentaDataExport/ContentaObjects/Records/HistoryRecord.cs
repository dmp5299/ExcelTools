using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ContentaDataExport.ContentaObjects
{
    public class HistoryRecord
    {
        DateTime? doneDate;

        [XmlElement("OPERATION")]
        public string OPERATION { get; set; }

        [XmlElement("DATETIME")]
        public string DATETIME
        {
            get
            {
                return doneDate.ToString();
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    doneDate = DateTime.ParseExact(value, "MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                }
                else
                {
                    doneDate = null;
                }
            }
        }

        [XmlElement("USER")]
        public string USER { get; set; }
    }
}
