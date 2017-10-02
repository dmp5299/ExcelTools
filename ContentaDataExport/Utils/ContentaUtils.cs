using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ContentaDataExport.Utils
{
    public class ContentaUtils
    {
        public static void setCookie(string host, string socket, string database)
        {
            Uri cookieUri = new Uri(@"C:\temp");
            string expires = "expires=Sat, 10-Oct-2018 00:00:00 GMT";
            string cookieString = "serverName=" + host + ",socket=" + socket + ",database=" + database + ";" + expires;
            try
            {
                System.Windows.Application.SetCookie(cookieUri, cookieString);
            }
            catch (Exception e)
            {
                throw new Exception("Error setting cookie: " + e.Message);
            }
        }

        public static Dictionary<string,string> getCookie()
        {
            Uri cookieUri = new Uri(@"C:\temp");
            Dictionary<string, string> connectionObjectDict = new Dictionary<string, string>();
            string cookie = "";
            try
            {
                cookie = System.Windows.Application.GetCookie(cookieUri);
            }
            catch (Exception) { return null; }
            if (!String.IsNullOrEmpty(cookie))
            {
                string[] cookieVars = cookie.Split(',');
                foreach (string var in cookieVars)
                {

                    string[] valueDef = var.Split('=');
                    switch (valueDef[0])
                    {
                        case "serverName":
                            connectionObjectDict["host"] = valueDef[1];
                            break;
                        case "socket":
                            connectionObjectDict["socket"] = valueDef[1];
                            break;
                        case "database":
                            connectionObjectDict["database"] = valueDef[1];
                            break;
                    }
                }
                return connectionObjectDict;
            }
            return null;
        }

        public static string getWhip(PCMClientLib.IPCMdata containers)
        {
            string wipId = "";
            for (int i = 0; i < containers.RecordCount; i++)
            {
                string name = containers.GetValueByLabel(i, "NAME");
                if (name == "WIP")
                {
                    string desktop = containers.GetValueByLabel(i, "DESKTOP");
                    string config = containers.GetValueByLabel(i, "CONFIGURATION_ID");
                    string objId = containers.GetValueByLabel(i, "OBJECT_ID");
                    wipId = desktop + @"/" + config + "/" + objId;
                }
            }
            return wipId;
        }
    }
}
