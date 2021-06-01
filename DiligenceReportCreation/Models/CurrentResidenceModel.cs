using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_currentResidences")]
    public class CurrentResidenceModel
    {
        [Key]
        [Column(name: "id")]
        public int id { get; set; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "CurrentStreet")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentStreet { set; get; }
        [Column(name: "CurrentCity")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentCity { set; get; }
        [Column(name: "CurrentCountry")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentCountry { set; get; }
        [Column(name: "CurrentZipcode")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentZipcode { set; get; }
    }
}
