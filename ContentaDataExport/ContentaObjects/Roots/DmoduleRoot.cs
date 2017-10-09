using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ContentaDataExport.ContentaObjects
{
    [XmlRoot("Data")]
    public class DmoduleRoot
    {
        [XmlElement("Record")]
        public Record Record { get; set; }
    }
}
