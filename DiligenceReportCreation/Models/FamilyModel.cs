using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{

    [Table(name: "diligence_Familytree")]
    public class FamilyModel
    {

        [Key]
        [Column(name: "Family_record_id")]
        public string Family_record_id { set; get; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "first_name")]
        public string first_name { set; get; }
        [Column(name: "middle_name")]
        public string middle_name { set; get; }
        [Column(name: "last_name")]
        public string last_name { set; get; }
        [Column(name: "applicant_type")]
        public string applicant_type { set; get; }
        [Column(name: "adult_minor")]
        public string adult_minor { set; get; }
        [Column(name: "case_number")]
        public string case_number { set; get; }
    }
}
