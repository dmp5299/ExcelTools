using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.Windows.Forms;
using System.Globalization;

namespace ContentaDataExport.ContentaObjects
{
    public class Record
    {
        string percent_complete;

        [SkipProperty]
        [SkipPropertyForExcelBody]
        public string Project { get; set; }

        [XmlElement("Model_Identification_Code")]
        public string Model_Identification_Code { get; set; }

        [XmlElement("System_Difference_Code")]
        public string System_Difference_Code { get; set; }

        [XmlElement("System_Code")]
        public string System_Code { get; set; }

        [XmlElement("Subsystem_Code")]
        public string Subsystem_Code { get; set; }

        [XmlElement("SubSubsystem_Code")]
        public string SubSubsystem_Code { get; set; }

        [XmlElement("Unit_or_Assembly_Code")]
        public string Unit_or_Assembly_Code { get; set; }

        [XmlElement("Disassembly_Code")]
        public string Disassembly_Code { get; set; }

        [XmlElement("Disassembly_Code_Variant")]
        public string Disassembly_Code_Variant { get; set; }

        [XmlElement("Information_Code")]
        public string Information_Code { get; set; }

        [XmlElement("Information_Code_Variant")]
        public string Information_Code_Variant { get; set; }

        [XmlElement("Item_Location_Code")]
        public string Item_Location_Code { get; set; }

        [XmlElement("NAME")]
        public string Name { get; set; }

        [XmlElement("Technical_Name")]
        public string Technical_Name { get; set; }

        [XmlElement("Information_Name")]
        public string Information_Name { get; set; }

        public string CSDB_Creation { get; set; }

        public List<RoutingRecord> RoutingTasks { get; set; }

        public string Percent_Complete {
            get
            {
                return percent_complete;
            }
            set
            {
                percent_complete = value;
            }
        }

        [SkipPropertyForExcelBody]
        public string Comments { get; set; }

        [SkipPropertyForExcelBody]
        public string _40_Percent_Effort { get; set; }

        [SkipPropertyForExcelBody]
        public string _80_Percent_Effort { get; set; }

        [SkipPropertyForExcelBody]
        public string DM_Reworked { get; set; }
    }

    public class SkipPropertyAttribute : Attribute
    {
    }
    public class SkipPropertyForExcelBodyAttribute : Attribute
    {
    }
    public class DoubleAttribute : Attribute
    {
    }
}