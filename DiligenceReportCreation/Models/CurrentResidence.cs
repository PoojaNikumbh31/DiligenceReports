using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{
    public class CurrentResidence
    {
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentStreet { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentCity { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentCountry { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentZipcode { set; get; }
    }

}
