using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _38_39Conversion.ContentaObjects
{
    public class DataConnection
    {
        string host;
        string socket;
        string database;

        public DataConnection (Dictionary<string,string> connDictionary)
        {
            Host = connDictionary == null ? "" : connDictionary["host"].ToString();
            Socket = connDictionary == null ? "" : connDictionary["socket"].ToString();
            Database = connDictionary == null ? "" : connDictionary["database"].ToString();
        }

        public string Host
        {
            get { return host; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    host = value;
                }
                else
                {
                    host = "";
                }
            }
        }

        public string Socket
        {
            get { return socket; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    socket = value;
                }
                else
                {
                    socket = "";
                }
            }
        }

        public string Database
        {
            get { return database; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    database = value;
                }
                else
                {
                    database = "";
                }
            }
        }

    }
}
