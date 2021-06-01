using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{

    [Table(name: "diligence_comments")]
    public class CommentModel
    {
        [Key]
        [Column(name: "id")]
        public int id { get; set; }
        [Column(name: "comment_type")]
        public string Comment_type { get; set; }
        [Column(name: "confirmed_comment")]
        public string confirmed_comment { get; set; }
        [Column(name: "unconfirmed_comment")]
        public string unconfirmed_comment { get; set; }
        [Column(name: "other_comment")]
        public string other_comment { get; set; }

    }
}
