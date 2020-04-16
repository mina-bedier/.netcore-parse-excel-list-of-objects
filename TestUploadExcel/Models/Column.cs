using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TestUploadExcel.Models
{

    [AttributeUsage(AttributeTargets.All)]
    public class Column : System.Attribute
    {
        public int ColumnIndex { get; set; }


        public Column(int column)
        {
            ColumnIndex = column;
        }
    }
}
