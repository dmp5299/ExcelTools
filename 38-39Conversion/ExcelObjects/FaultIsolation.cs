using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _38_39Conversion.ExcelObjects
{
    public class FaultIsolation
    {
        public string FaultCode { get; set; }
        public string MaintenanceTaskName { get; set; }
        public string FaultIsolationProcedureId { get; set; }
        public string FailureName { get; set; }
        public string _920DmcTitle { get; set; }
        public string _920DMC { get; set; }
        public string Name { get; set; }
    }
}
