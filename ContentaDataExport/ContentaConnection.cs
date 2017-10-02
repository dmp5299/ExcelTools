using System;
using PCMClientLib;
using System.Windows.Forms;
using System.Xaml;
using ContentaDataExport;

namespace ContentaDataExport
{
    public class ContentaConnection
    {
        public static IPCMcommand getCommandObject(string host, string socket, string database)
        {
            IPCMcommand command = null;
            PCMClientLib.PCMConnection conn = new PCMConnection();
            try
            {
                command = conn.ConnectGetCommand((short)Int32.Parse(socket),host,database,"sysadmin","manager","XyGACTest",0);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error connecting to database: "+e.Message);
            }
            return command;
        }
    }
}
