using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using ContentaDataExport.XmlObjects;

namespace ContentaDataExport.Utils
{
    public class XmlUtils
    {
        public static T SerializeXml<T>(T type, string xml)
        {
            string xmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using (TextReader reader = new StringReader(xmlHeader + xml))
            {
                type = (T)serializer.Deserialize(reader);
            }
            return type;
        }

        //build and return dmc object
        public static DMC getDMC(string dmcName)
        {
            string[] splitArray = dmcName.Split('-');
            try
            {
                DMC dmc = new DMC()
                {
                    Model_Identification_Code = splitArray[0],
                    System_Difference_Code = splitArray[1],
                    System_Code = splitArray[2],
                    Subsystem_Code = splitArray[3].Substring(0, 1),
                    SubSubsystem_Code = splitArray[3].Substring(1),
                    Unit_or_Assembly_Code = splitArray[4],
                    Disassembly_Code = splitArray[5].Substring(0,2),
                    Disassembly_Code_Variant = splitArray[5].Substring(2),
                    Information_Code = splitArray[6].Substring(0,3),
                    Information_Code_Variant = splitArray[6].Substring(3),
                    Item_Location_Code = splitArray[7]
                };
                return dmc;
            }
            catch (Exception e)
            {
                throw new Exception("Exception in getDMC " + e.Message);
            }
        }
    }
}
