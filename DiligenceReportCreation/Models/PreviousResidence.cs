using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace DiligenceReportCreation.Models
{
    public class PreviousResidence
    {
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PreviousStreet { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PreviousCity { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PreviousCountry { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PreviousZipcode { set; get; }
    }
}
