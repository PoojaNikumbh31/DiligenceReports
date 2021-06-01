using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_statewide_footnote")]
    public class StateWiseFootnoteModel
    {
        [Key]
        [Column(name: "id")]
        public int id { get; set; }
        [Column(name: "states")]
        public string states { get; set; }
        [Column(name: "state_specific_district_courts")]
        public string state_specific_district_courts { get; set; }
        [Column(name: "statewide_criminal_language_dd_report")]
        public string statewide_criminal_language_dd_report { get; set; }
        [Column(name: "statewide_criminal_language_pe_report")]
        public string statewide_criminal_language_pe_report { get; set; }
        [Column(name: "abbreviation")]
        public string abbreviation { get; set; }
    }
}
