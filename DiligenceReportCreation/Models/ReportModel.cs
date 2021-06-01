using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_reports")]
    public class ReportModel
    {
        [Key]
        [Column(name: "record_id")]
        public string record_Id { set; get; }
      
        [Column(name: "last_name")]
        public string lastname { set; get; }        
        [Column(name: "case_number")]
        public string casenumber { set; get; }       
        [Column(name: "template_type")]
        public string TemplateType { set; get; }
        [Column(name: "created_by")]
        public string createdby { set; get; }

    }
}
