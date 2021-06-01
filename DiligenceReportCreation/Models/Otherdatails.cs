using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligenceOthersInfo")]
    public class Otherdatails
    {
        [Key]
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "has_property_records")]
        public string Has_Property_Records { set; get; }
        [Column(name: "has_legal_records_hits")]
        public string Has_Legal_Records_Hits { set; get; }
        [Column(name: "has_regulatory_hits")]
        public string Has_Regulatory_Hits { set; get; }
        [Column(name: "has_hits_above")]
        public string Has_Hits_Above { set; get; }
        [Column(name: "has_companion_report")]
        public string Has_Companion_Report { set; get; }
        [Column(name: "has_business_affiliations")]
        public string Has_Business_Affiliations { set; get; }
        [Column(name: "has_intellectual_hits")]
        public string Has_Intellectual_Hits { set; get; }
        [Column(name: "ApplicantType")]
        public string ApplicantType { set; get; }
        [Column(name: "PatentType")]
        public string PatentType { set; get; }
        [Column(name: "has_reg_us_sec")]
        public string Has_Reg_US_SEC { set; get; }
        [Column(name: "has_reg_uk_fca")]
        public string Has_Reg_UK_FCA { set; get; }
        [Column(name: "has_reg_finra")]
        public string Has_Reg_FINRA { set; get; }
        [Column(name: "has_reg_us_nfa")]
        public string Has_Reg_US_NFA { set; get; }
        [Column(name: "holds_any_license")]
        public string HoldsAnyLicense { set; get; }
        [Column(name: "regulatory_flag")]
        public string RegulatoryFlag { set; get; }
        [Column(name: "media_based_hits")]
        public string Media_Based_Hits { set; get; }
        [Column(name: "uslegal_record_hits")]
        public string USLegal_Record_Hits { set; get; }
        [Column(name: "global_security_hits")]
        public string Global_Security_Hits { set; get; }
        [Column(name: "icij_hits")]
        public string ICIJ_Hits { set; get; }
        [Column(name: "press_media")]
        public string Press_Media { set; get; }
        [Column(name: "CurrentResidentialProperty")]
        public string CurrentResidentialProperty { set; get; }
        [Column(name: "OtherCurrentResidentialProperty")]
        public string OtherCurrentResidentialProperty { set; get; }
        [Column(name: "OtherPropertyOwnershipinfo")]
        public string OtherPropertyOwnershipinfo { set; get; }
        [Column(name: "PrevPropertyOwnershipRes")]
        public string PrevPropertyOwnershipRes { set; get; }        
        [Column(name: "HasBankruptcyRecHits")]
        public string HasBankruptcyRecHits { set; get; }
        [Column(name: "Registered_with_HKSFC")]
        public string Registered_with_HKSFC { set; get; }
        [Column(name: "Has_CriminalRecHit")]
        public string Has_CriminalRecHit { set; get; }
        [Column(name: "Has_Bureau_PrisonHit")]
        public string Has_Bureau_PrisonHit { set; get; }
        [Column(name: "Has_Sex_Offender_RegHit")]
        public string Has_Sex_Offender_RegHit { set; get; }
        [Column(name: "Has_Civil_Records")]
        public string Has_Civil_Records { set; get; }        
        [Column(name: "Has_US_Tax_Court_Hit")]
        public string Has_US_Tax_Court_Hit { set; get; }
        [Column(name: "Has_Tax_Liens")]
        public string Has_Tax_Liens { set; get; }
        
        [Column(name: "Has_Driving_Hits")]
        public string Has_Driving_Hits { set; get; }        
        [Column(name: "Was_credited_obtained")]
        public string Was_credited_obtained { set; get; }
        [Column(name: "Has_Name_Only")]
        public bool Has_Name_Only { set; get; }
        [Column(name: "Has_Name_Only_Tax_Lien")]
        public bool Has_Name_Only_Tax_Lien { set; get; }
        [Column(name: "Has_Name_only_driving_incidents")]
        public bool Has_Name_only_driving_incidents { set; get; }
        public string has_ucc_fillings { set; get; }
        [Column(name: "has_ucc_fillings1")]
        public bool has_ucc_fillings1 { set; get; }
        [Column(name: "has_civil_resultpending")]
        public bool has_civil_resultpending { set; get; }
        [Column(name: "HasBankruptcyRecHits1")]
        public bool HasBankruptcyRecHits1 { set; get; }
        [Column(name: "HasBankruptcyRecHits_resultpending")]
        public bool HasBankruptcyRecHits_resultpending { set; get; }
        [Column(name: "Has_CriminalRecHit1")]
        public bool Has_CriminalRecHit1 { set; get; }
        [Column(name: "Has_CriminalRecHit_resultpending")]
        public bool Has_CriminalRecHit_resultpending { set; get; }
        [Column(name: "Executivesummary")]
        public string Executivesummary { set; get; }
        [Column(name: "worldcheck_discloseBA")]
        public string worldcheck_discloseBA { set; get; }
        [Column(name: "undisclosedBA")]
        public string undisclosedBA { set; get; }
        [Column(name: "worldcheck_undiscloseBA")]
        public string worldcheck_undiscloseBA { set; get; }
        [Column(name: "Global_Sec_Family_Hits")]
        public string Global_Sec_Family_Hits { set; get; }
        [Column(name: "PEP_Hits")]
        public string PEP_Hits { set; get; }
        [Column(name: "Crim_Clearance_Certifi")]
        public string Crim_Clearance_Certifi { set; get; }
        [Column(name: "Fama")]
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Fama { set; get; }



    }
}
