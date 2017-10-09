using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ContentaDataExport.Utils;

namespace ContentaDataExport.ContentaClasses
{
    public class ContentaSearch
    {
        public XmlDocument searchByFilePath(PCMClientLib.IPCMcommand cmd, string id, string paths)
        {
            string command = "find -relpath ";
            string[] pathList = paths.Split('|');
            foreach (string path in pathList)
            {
                string configId = FileUtls.fileParse(path);
                command += "\"" + configId + "\"|\"" + id + "\" ";
            }
            PCMClientLib.IPCMdata data = cmd.ExecCmd(command);

            PCMClientLib.IPCMdata names = cmd.ExecCmd("translate " + paths.Replace('|', ' '));

            int records = data.RecordCount;
            List<string> rows = new List<string>();
            for (int i = 0; i < records; i++)
            {
                string namePath = data.GetValueByLabel(i, "NAME_PATH");
                string idPath = data.GetValueByLabel(i, "ID_PATH");

                ComposeNameIdPath(ref idPath, ref namePath, names);
                string config = FileUtls.fileParse(idPath);
                string object1 = FileUtls.fileParse(namePath);
                string row = "||" + config + "|" + object1 + "||" + idPath + "|" + namePath + "|";
                rows.Add(row);
            }
            return LoadDataFromArray("SCORE|OBJECT_ID|NAME|TYPE|ID_PATH|NAME_PATH", rows);
        }

        public XmlDocument LoadDataFromArray(string labels, List<string> rows)
        {
            string record = "|" + labels + "| ";

            for (int i = 0; i < rows.Count; i++)
            {
                record += rows[i] + " ";
            }

            string string1 = " |" + string.Format("{0:D6}", rows.Count + 2) + "|" + string.Format("{0:D6}", record.Length + 9) + "| |" + string.Format("{0:D6}", 6) + "| " + record;
            string response = setPortalData(string1);
            response = "<Data>" + response + "</Data>";
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(response);
            return doc;
        }

        public void ComposeNameIdPath(ref string idPath, ref string namePath, PCMClientLib.IPCMdata names)
        {
            string id = idPath.Substring(0, idPath.IndexOf('/'));
            int count = names.RecordCount;

            for (int i = 0; i < count; i++)
            {
                string tempIdPath = names.GetValueByLabel(i, "ID_PATH");
                int pos = tempIdPath.LastIndexOf('/');
                string tempId = tempIdPath.Substring(pos + 1);

                if (tempId == id)
                {
                    idPath = tempIdPath.Substring(0, pos + 1) + idPath;

                    string tempNamePath = names.GetValueByLabel(i, "NAME_PATH");
                    pos = tempNamePath.LastIndexOf('/');
                    namePath = tempNamePath.Substring(0, pos + 1) + namePath;
                }
            }
        }

        public string setPortalData(string string1)
        {
            string xml_str = "";

            string[] record = string1.Split('|');

            int size = Int32.Parse(record[4].Substring(0, 6));

            //string[] keys = new List<string>(record).GetRange(6, (size+5)).ToArray();

            string[] keys = record.ToList().GetRange(6, 6).ToArray();

            string[] values = new List<string>(record).GetRange((size + 7), (record.Length - (size + 7))).ToArray();

            int j = 0;

            while (j < values.Length)
            {
                xml_str += "<Record>\n";

                for (int i = 0; i < keys.Length; i++)
                {
                    xml_str += '<' + keys[i] + '>' + values[i + j] + @"</" + keys[i] + @">";
                }
                j += (size + 1);
                xml_str += "</Record>";
            }
            return xml_str;
        }
    }
}
