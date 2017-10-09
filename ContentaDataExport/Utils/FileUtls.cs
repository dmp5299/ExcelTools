using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ContentaDataExport.Utils
{
    public class FileUtls
    {
        public static string fileParse(string path)
        {
            if (path.Contains('/'))
            {
                string[] subPaths = path.Split('/');
                return subPaths[subPaths.Length - 1];
            }
            else
                return path;
        }
    }
}
