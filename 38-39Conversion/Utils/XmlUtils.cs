using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using _38_39Conversion.XmlObjects;
using System.Reflection;
using System.Windows.Forms;

namespace _38_39Conversion.Utils
{
    public class XmlUtils
    {
        public static XmlNode BuildDmRef(string dmc, XmlDocument doc)
        {
            List<string> dmcAtts = dmc.Split('-').ToList();

            string subSystemCode = dmcAtts[3].Substring(0, 1);
            string SubSubSystemCode = dmcAtts[3].Substring(1,1);

            dmcAtts.RemoveAt(3);
            dmcAtts.Insert(3, subSystemCode);
            dmcAtts.Insert(4, SubSubSystemCode);

            string disassyCode = dmcAtts[6].Substring(0, 2);
            string disassyCodeVariant = dmcAtts[6].Substring(2, 3);

            dmcAtts.RemoveAt(6);
            dmcAtts.Insert(6, disassyCode);
            dmcAtts.Insert(7, disassyCodeVariant);

            string InfoCode = dmcAtts[8].Substring(0, 3);
            string InfoCodeVariant = dmcAtts[8].Substring(3, 1);

            dmcAtts.RemoveAt(8);
            dmcAtts.Insert(8, InfoCode);
            dmcAtts.Insert(9, InfoCodeVariant);

            int i = 0;
            Type t = typeof(DmCode);
            DmCode dmCode = new DmCode();
            foreach (PropertyInfo propertyInfo in t.GetProperties())
            {
                propertyInfo.SetValue(dmCode, dmcAtts[i]);
                i++;
            }
            return dmCode.buildDmCode(doc);
        }
    }
}
