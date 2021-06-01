using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "DiligencePLLicense")]
    public class PLLicenseModel
    {
        [Key]
        [Column(name: "id")]
        public int id { get; set; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "general_PL_License")]
        public string General_PL_License { set; get; }
        [Column(name: "pl_license_type")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_License_Type { set; get; }
        [Column(name: "pl_organization")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_Organization { set; get; }
        [Column(name: "pl_location")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_Location { set; get; }
        [Column(name: "pl_number")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_Number { set; get; }
        //[Column(name: "pl_start_Date")]
        //public string PL_Start_Date { set; get; }
     
        [Column(name: "pl_confirmed")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_Confirmed { set; get; }
        [Column(name: "CreatedBy")]
        public string CreatedBy { set; get; }
        [Column(name: "pl_startdatemonth")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_StartDateMonth { set; get; }                
        [Column(name: "PL_StartDateDay")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_StartDateDay { set; get; }
        [Column(name: "PL_StartDateYear")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_StartDateYear { set; get; }
        [Column(name: "PL_EndDateDay")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_EndDateDay { set; get; }
        [Column(name: "PL_EndDateYear")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_EndDateYear { set; get; }
        [Column(name: "PL_EndDateMonth")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PL_EndDateMonth { set; get; }
    }
}
