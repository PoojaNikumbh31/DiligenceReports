using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{
    public class Referencetable
    {
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ref_full_name { get; set; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ref_position { get; set; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ref_location { get; set; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ref_employer { get; set; }
    }
}
