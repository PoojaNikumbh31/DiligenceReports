using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    
    [Table(name: "DiligencePersonalInfo")]
    public class DiligenceInputModel
    {       
        [Key]
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "client_name")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ClientName { set; get; }
        [Column(name: "case_number")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CaseNumber { set; get; }
        [Column(name: "Marital_Status")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Marital_Status { set; get; }
        [Column(name: "Children")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Children { set; get; }
        [Column(name: "first_name")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string FirstName { set; get; }
        [Column(name: "middle_name")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string MiddleName { set; get; }
        [Column(name: "middle_initial")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string MiddleInitial { set; get; }
        [Column(name: "last_name")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string LastName { set; get; }
        [Column(name: "madian_name")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string MaidenName { set; get; }
        [Column(name: "full_alias_name")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string FullaliasName { set; get; }
        [Column(name: "foreignlanguage")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string foreignlanguage { get; set; }
        [Column(name: "dob")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Dob { set; get; }
        [Column(name: "pob")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Pob { set; get; }
        [Column(name: "nationality")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Nationality { set; get; }
        [Column(name: "national_id_number")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Nationalidnumber { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        [Column(name: "nationalid_issue_date")]
        public string Natlidissuedate { set; get; }
        [Column(name: "nationalid_expiration_date")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Natlidexpirationdate { set; get; }
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
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        [Column(name: "city")]
        public string City { set; get; }
        [Column(name: "country")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Country { set; get; }        
        [Column(name: "non_scopecountry1")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Nonscopecountry1 { set; get; }                        
        [Column(name: "SSNInitials")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]        
        public string SSNInitials { set; get; }
        [Column(name: "Employer1City")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]        
        public string Employer1City { set; get; }
        [Column(name: "Employer1State")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Employer1State { set; get; }
        [Column(name: "CommonNameSubject")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public bool CommonNameSubject { set; get; }
        [Column(name: "CurrentState")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentState { set; get; }      
        ////France
        //[Column(name: "france_spereghits")]
        //public string France_SpeRegHits { set; get; }
        //[Column(name: "france_legalrechits")]
        //public string France_LegalRecHits { set; get; }
        ////UK
        //[Column(name: "uk_insolvencyhits")]
        //public string UK_InsolvencyHits { set; get; }
        //[Column(name: "uk_civilrechit")]
        //public string UK_CivilRecHit { set; get; }
        //[Column(name: "uk_regtrusthit")]
        //public string UK_RegTrustHit { set; get; }
        //[Column(name: "uk_criminalrecord")]
        //public string UK_CriminalRecord { set; get; }
        //[Column(name: "uk_drivinghistory")]
        //public string UK_DrivingHistory { set; get; }
        //[Column(name: "uk_credithistory")]
        //public string UK_CreditHistory { set; get; }
        ////Canada
        //[Column(name: "can_drivinghistory")]
        //public string Can_DrivingHistory { set; get; }
        //[Column(name: "can_credithistory")]
        //public string Can_CreditHistory { set; get; }
        //[Column(name: "can_specreghits")]
        //public string Can_SpecRegHits { set; get; }
        //[Column(name: "can_insolvencyhits")]
        //public string Can_InsolvencyHits { set; get; }
        //[Column(name: "can_civilrechit")]
        //public string Can_CivilRecHit { set; get; }
        //[Column(name: "can_criminalrecord")]
        //public string Can_CriminalRecord { set; get; }
        //[Column(name: "can_pprrecordhits")]
        //public string Can_PPRRecordHits { set; get; }
        ////Australia
        //[Column(name: "aus_drivinghistory")]
        //public string Aus_DrivingHistory { set; get; }
        //[Column(name: "aus_credithistory")]
        //public string Aus_CreditHistory { set; get; }
        //[Column(name: "aus_asicreg")]
        //public string Aus_ASICReg { set; get; }
        //[Column(name: "aus_insolvencyhits")]
        //public string Aus_InsolvencyHits { set; get; }
        //[Column(name: "aus_civilrechit")]
        //public string Aus_CivilRecHit { set; get; }
        //[Column(name: "aus_criminalrecord")]
        //public string Aus_CriminalRecord { set; get; }
        //[Column(name: "aus_ppsrrecordhits")]
        //public string Aus_PPSRRecordHits { set; get; }

        //[Column(name: "CreatedBy")]
        //public string CreatedBy { set; get; }
    }
}
