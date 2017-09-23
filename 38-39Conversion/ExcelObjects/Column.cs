using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace _38_39Conversion.ExcelObjects
{
    [AttributeUsage(System.AttributeTargets.All)]
    public class Column : System.Attribute
    {
        public int ColumnIndex { get; set; }

        public Column(int column)
        {
            ColumnIndex = column;
        }
    }
}
