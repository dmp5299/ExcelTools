using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using _38_39Conversion.ContentaObjects;
using ContentaDataExport;
using ContentaDataExport.Utils;

namespace _38_39Conversion
{
    /// <summary>
    /// Interaction logic for ContentaOptions.xaml
    /// </summary>
    public partial class ContentaOptions : Window
    {
        public DataConnection dataConn;

        public ContentaOptions()
        {
            InitializeComponent();
            trySetCookieValues();
        }

        private void trySetCookieValues()
        {
            dataConn = new DataConnection(ContentaUtils.getCookie());
            HostText.Text = dataConn.Host;
            SocketText.Text = dataConn.Socket;
            DatabaseText.Text = dataConn.Database;
        }
        
        private void Done_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                checkConnectionValues();
                if((HostText.Text != dataConn.Host) || (SocketText.Text != dataConn.Socket) || (DatabaseText.Text != dataConn.Database))
                {
                    ContentaUtils.setCookie(HostText.Text,SocketText.Text,DatabaseText.Text);
                }
                this.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void checkConnectionValues()
        {
            if(HostText.Text == "")
            {
                throw new Exception("Enter host.");
            }
            else if (SocketText.Text == "")
            {
                throw new Exception("Enter socket.");
            }
            else if (DatabaseText.Text == "")
            {
                throw new Exception("Enter database.");
            }
            return;
        }
    }
}
