using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{   
    public class Employer
    {
        public string record_Id { set; get; }     
        public string Emp_Status { set; get; }        
        public string Emp_Position { set; get; }        
        public string Emp_Employer { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_StartDate { set; get; }               
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_EndDate { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_StartDateMonth { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_StartDateDay { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_StartDateYear { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_EndDateMonth { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_EndDateDay { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_EndDateYear { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        
        public string Emp_Location { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_State { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_City { set; get; }
        public string Emp_Confirmed { set; get; }            
    }
}
