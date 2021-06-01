using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "DiligenceEducation")]
    public class EducationModel
    {
        [Key]
        [Column(name: "id")]
        public int id { get; set; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "edu_history")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_History { set; get; }
        [Column(name: "degree")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_Degree { set; get; }
        [Column(name: "school")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_School { set; get; }
        [Column(name: "major")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_Major { set; get; }
        [Column(name: "additionalInfo")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_AdditionalInfo { set; get; }

        [Column(name: "edu_location")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_Location { set; get; }
        [Column(name: "edu_confirmed")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_Confirmed { set; get; }
        [Column(name: "CreatedBy")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CreatedBy { set; get; }
        [Column(name: "Edu_GradDateMonth")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_GradDateMonth { set; get; }
        [Column(name: "Edu_GradDateDay")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_GradDateDay { set; get; }
        [Column(name: "Edu_GradDateYear")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]

        public string Edu_GradDateYear { set; get; }
        [Column(name: "Edu_StartDateMonth")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_StartDateMonth { set; get; }
        [Column(name: "Edu_StartDateDay")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_StartDateDay { set; get; }
        [Column(name: "Edu_StartDateYear")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_StartDateYear { set; get; }
        [Column(name: "Edu_EndDateMonth")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_EndDateMonth { set; get; }
        [Column(name: "Edu_EndDateDay")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_EndDateDay { set; get; }
        [Column(name: "Edu_EndDateYear")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Edu_EndDateYear { set; get; }
    }
}
