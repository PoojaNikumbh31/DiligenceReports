using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{    
    public class Education
    {        
        public string record_Id { set; get; }        
        public string Edu_History { set; get; }        
        public string Edu_Degree { set; get; }        
        public string Edu_School { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_Major { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_AdditionalInfo { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_Graddate { set; get; }       
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_GradDateMonth { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_GradDateDay { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_GradDateYear { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_StartDate { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_StartDateMonth { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_StartDateDay { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_StartDateYear { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_EndDateMonth { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_EndDateDay { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_EndDateYear { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_StartDate1 { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_EndDate { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_EndDate1 { set; get; }        
        public string Edu_Location { set; get; }        
        public string Edu_Confirmed { set; get; }        
    }
}
