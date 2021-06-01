using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_References")]
    public class ReferencetableModel
    {
        [Key]
        [Column(name: "id")]
        public int id { get; set; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "ref_full_name")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ref_full_name { get; set; }
        [Column(name: "ref_position")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ref_position { get; set; }
        [Column(name: "ref_location")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ref_location { get; set; }
        [Column(name: "ref_employer")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ref_employer { get; set; }
    }
}
