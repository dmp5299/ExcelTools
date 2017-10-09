using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _38_39Conversion.ExcelObjects
{
    public class ExcelRow
    {
        [Column(2)]
        [Required]
        public string Id { get; set; }

        [Column(3)]
        [Required]
        public string FailureName { get; set; }

        [Column(4)]
        [Required]
        public string MaintenanceTaskName { get; set; }

        [Column(5)]
        [Required]
        public string FaultIsolationProcedureId { get; set; }

        [Column(6)]
        [Required]
        public string _411DmcTitle { get; set; }

        [Column(7)]
        [Required]
        public string _411DMC { get; set; }

        [Column(8)]
        [Required]
        public string _920DmcTitle { get; set; }

        [Column(9)]
        [Required]
        public string _920DMC{ get; set; }

        [Column(10)]
        [Required]
        public string Name { get; set; }
    }
}
