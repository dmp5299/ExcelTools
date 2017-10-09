using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Windows.Forms;

namespace ContentaDataExport.ContentaObjects
{
    public class RoutingRecord
    {
        DateTime? doneDate;
        string user;

        [XmlElement("USER")]
        public string USER
        {
            get
            {
                if (string.IsNullOrEmpty(user))
                {
                    return "";
                }
                else
                {
                    return user;
                }
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    user = value;
                }
                else
                {
                    user = "";
                }
            }
        }

        [XmlElement("TASK")]
        public string TASK { get; set; }

        [XmlElement("ROLE")]
        public string ROLE { get; set; }

        [XmlElement("DONE_DATE")]
        public string DONE_DATE
        {
            get
            {
                return doneDate.ToString();
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    DateTime dDate;
                    if (DateTime.TryParse(value, out dDate))
                    {
                        doneDate = dDate;
                    }
                    else
                    {
                        doneDate = DateTime.ParseExact(value, "MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                    }    
                }
                else
                {
                    doneDate = null;
                }
            }
        }
    }
}
