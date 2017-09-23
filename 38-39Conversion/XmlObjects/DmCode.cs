using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Xml;
using _38_39Conversion.Utils;
using System.Windows.Forms;

namespace _38_39Conversion.XmlObjects
{
    public class DmCode
    {
        public string ModelIdentCode { get; set; }
        public string SystemDiffCode { get; set; }
        public string SystemCode { get; set; }
        public string SubSystemCode { get; set; }
        public string SubSubSystemCode { get; set; }
        public string AssyCode { get; set; }
        public string DisassyCode { get; set; }
        public string DisassyCodeVariant { get; set; }
        public string InfoCode { get; set; }
        public string InfoCodeVariant { get; set; }
        public string ItemLocationCode { get; set; }

        public XmlNode buildDmCode(XmlDocument doc)
        {
            XmlNode dmcode = doc.CreateElement("dmCode");
            foreach (PropertyInfo prop in this.GetType().GetProperties())
            {
                XmlAttribute attribute = doc.CreateAttribute(StringUtils.FirstCharacterToLower(prop.Name));
                attribute.Value = (string)prop.GetValue(this,null);
                dmcode.Attributes.Append(attribute);
            }
            return dmcode;
        }
    }
}
