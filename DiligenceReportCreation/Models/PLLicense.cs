using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{
    public class PLLicense
    {                        
        public string record_Id { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string General_PL_License { set; get; }        
        public string PL_License_Type { set; get; }
        public string PL_Organization { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_Location { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_Number { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_Start_Date { set; get; }        
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_StartDateMonth { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_StartDateDay { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_StartDateYear { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_End_Date { set; get; }                
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_EndDateDay { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_EndDateYear { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_EndDateMonth { set; get; }
        public string PL_Confirmed { set; get; }
    }
}
