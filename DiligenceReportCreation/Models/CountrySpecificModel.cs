using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_countryspecific")]
    public class CountrySpecificModel
    {
        [Key]
        [Column(name: "id")]
        public Int32 id { set; get; }
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "country_specific_reg_hits")]
        public string country_specific_reg_hits { get; set; }
        [Column(name: "legal_record_hits")]
        public string legal_record_hits { get; set; }
        [Column(name: "insolvency_hits")]
        public string insolvency_hits { get; set; }
        [Column(name: "civil_record_hits")]
        public string civil_record_hits { get; set; }
        [Column(name: "criminal_record_hits")]
        public string criminal_record_hits { get; set; }
        [Column(name: "reg_trust_hits")]
        public string reg_trust_hits { get; set; }
        [Column(name: "driving_hits_doc_required")]
        public string driving_hits_doc_required { get; set; }
        [Column(name: "driving_hits")]
        public string driving_hits { get; set; }
        [Column(name: "credit_hits")]
        public string credit_hits { get; set; }
        [Column(name: "pprrecordhits")]
        public string pprrecordhits { get; set; }     
        [Column(name: "registeredwithhksfc")]
        public string registeredwithhksfc { get; set; }
        [Column(name: "hasnameonlybanruptymatch")]
        public string hasnameonlybanruptymatch { get; set; }
        [Column(name: "hasnameonlymatch")]
        public string hasnameonlymatch { get; set; }
        [Column(name: "hasnameonlycriminalmatch")] 
        public string hasnameonlycriminalmatch { get; set; }
        [Column(name: "germany_regsearches")]
        public string germany_regsearches { get; set; }
        [Column(name: "india_corpregistry")]
        public string india_corpregistry { get; set; }
        [Column(name: "country")]
        public string country { get; set; }
    }
}
