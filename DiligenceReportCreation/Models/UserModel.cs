using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_users")]
    public class UserModel
    {
        [Key]
        [Column(name: "username")]
        public string Username { get; set; }
        [Column(name: "role")]
        public string Role { get; set; }
        [Column(name: "password")]
        public string Password { get; set; }
        [Column(name: "email_id")]
        public string Email { get; set; }
        [Column(name: "permissions")]
        public string Permissions { get; set; }
        [Column(name: "last_login")]
        public DateTime LastLogin { get; set; }
    }
}
