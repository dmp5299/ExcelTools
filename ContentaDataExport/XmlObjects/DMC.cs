using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ContentaDataExport.XmlObjects
{
    public class DMC
    {
        public string Model_Identification_Code { get; set; }
        public string System_Difference_Code { get; set; }
        public string System_Code { get; set; }
        public string Subsystem_Code { get; set; }
        public string SubSubsystem_Code { get; set; }
        public string Unit_or_Assembly_Code { get; set; }
        public string Disassembly_Code { get; set; }
        public string Disassembly_Code_Variant { get; set; }
        public string Information_Code { get; set; }
        public string Information_Code_Variant { get; set; }
        public string Item_Location_Code { get; set; }
    }
}
