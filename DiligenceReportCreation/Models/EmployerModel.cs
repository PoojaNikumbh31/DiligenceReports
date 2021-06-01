using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "DiligenceEmployee")]
    public class EmployerModel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column(name: "id")]        
        public int id { get; set; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "emp_status")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_Status { set; get; }
        [Column(name: "emp_position")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_Position { set; get; }
        [Column(name: "employer")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_Employer { set; get; }                        
        [Column(name: "emp_location")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_Location { set; get; }
        [Column(name: "emp_confirmed")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_Confirmed { set; get; }
        [Column(name: "CreatedBy")]        
        public string CreatedBy { set; get; }
        [Column(name: "Emp_StartDateMonth")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_StartDateMonth { set; get; }
        [Column(name: "Emp_StartDateDay")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_StartDateDay { set; get; }
        [Column(name: "Emp_StartDateYear")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_StartDateYear { set; get; }
        [Column(name: "Emp_EndDateMonth")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_EndDateMonth { set; get; }
        [Column(name: "Emp_EndDateDay")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_EndDateDay { set; get; }
        [Column(name: "Emp_EndDateYear")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_EndDateYear { set; get; }
        [Column(name: "Emp_State")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_State { set; get; }
        [Column(name: "Emp_City")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Emp_City { set; get; }
    }
}
