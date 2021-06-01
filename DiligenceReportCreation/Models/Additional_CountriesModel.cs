using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{

    [Table(name: "diligence_additionalCountries")]
    public class Additional_CountriesModel
    {
        [Key]
        [Column(name: "id")]
        public int id { get; set; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "additionalCountries")]
        public string additionalCountries { set; get; }
    }
}
