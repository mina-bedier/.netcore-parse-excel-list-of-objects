using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace TestUploadExcel.Models
{
    public class Drivers
    {
        [Column(1)]
        [DisplayName("Driver Name Here")]
        [Required(ErrorMessage = "esm l swa2 matlob")]
        public string DriverName { get; set; }
        [Required(ErrorMessage = "rkm l swa2 mtlob")]
        [DisplayName("Driver Number Here")]
        [Column(2)]
        [RegularExpression(@"\d+", ErrorMessage = "UPRN must be numeric")]
        public string DriverNumber { get; set; }
        [DisplayName("A7mos Here")]
        [Column(3)]
        public string a7mos { get; set; }
    }
}
