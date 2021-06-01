using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_previousResidences")]
    public class PreviousResidenceModel
    {
        [Key]
        [Column(name: "id")]
        public int id { get; set; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "PreviousStreet")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PreviousStreet { set; get; }
        [Column(name: "PreviousCity")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PreviousCity { set; get; }
        [Column(name: "PreviousCountry")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PreviousCountry { set; get; }
        [Column(name: "PreviousZipcode")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string PreviousZipcode { set; get; }
    }
}
