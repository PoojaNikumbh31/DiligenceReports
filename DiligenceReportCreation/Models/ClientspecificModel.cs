using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_clientspecific")]
    public class ClientspecificModel
    {
        [Key]
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "clientname")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string clientname { set; get; }
        [Column(name: "Military_Service")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Military_Service { set; get; }
        [Column(name: "source_of_wealth")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string source_of_wealth { set; get; }
        [Column(name: "discreet_reputational_inquiries")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string discreet_reputational_inquiries { set; get; }
        [Column(name: "Character_Reference")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Character_Reference { set; get; }
    }
}

