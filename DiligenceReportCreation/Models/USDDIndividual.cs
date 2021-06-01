using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{
    public class USDDIndividual
    {
        public string record_Id { set; get; }
        public string ClientName { set; get; }
        public string CaseNumber { set; get; }
        public string Has_Property_Records { set; get; }
        public string FirstName { set; get; }
        public string MiddleName { set; get; }
        public string MiddleInitial { set; get; }
        public string LastName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string MaidenName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string FullaliasName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Dob { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string SSNInitials { set; get; }
        public string Nationality { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Nationalidnumber { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Natlidissuedate { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Natlidexpirationdate { set; get; }
        //public string Currentfulladdress { set; get; } 
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentStreet { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentCity { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentState { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentZipcode { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string City { set; get; }
        public string Country { set; get; }
        //public string Recommend_additional_countries { set; get; }    
        //[DisplayFormat(ConvertEmptyStringToNull = false)]
        //public string AdditinalStates { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string AdditionalJurisdictions { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Employer1City { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Employer1State { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public bool CommonNameSubject { set; get; }
        public string CurrentResidentialProperty { set; get; }
        public string OtherPropertyOwnershipinfo { set; get; }
        public string PrevPropertyOwnershipRes { set; get; }
        public string HasBankruptcyRecHits { set; get; }
        public string Has_CriminalRecHit { set; get; }
        public string Has_Bureau_PrisonHit { set; get; }
        public string Has_Sex_Offender_RegHit { set; get; }
        public string Has_Civil_Records { set; get; }
        public bool Has_Name_Only { set; get; }
        public string Has_US_Tax_Court_Hit { set; get; }
        public string Has_Tax_Liens { set; get; }
        public bool Has_Name_Only_Tax_Lien { set; get; }
        public string Has_Driving_Hits { set; get; }
        public bool Has_Name_only_driving_incidents { set; get; }
        public string Was_credited_obtained { set; get; }
        public string Has_Legal_Records_Hits { set; get; }
        public string Has_Regulatory_Hits { set; get; }
        public string Has_Hits_Above { set; get; }
        public string Has_Companion_Report { set; get; }
        public string Has_Business_Affiliations { set; get; }
        public string Has_Intellectual_Hits { set; get; }
        public string Has_Reg_US_SEC { set; get; }
        public string Has_Reg_UK_FCA { set; get; }
        public string Has_Reg_FINRA { set; get; }
        public string Has_Reg_US_NFA { set; get; }
        public string HoldsAnyLicense { set; get; }
        public string RegulatoryFlag { set; get; }
        public string Registered_with_HKSFC { set; get; }
        public string USLegal_Record_Hits { set; get; }
        public string Global_Security_Hits { set; get; }
        public string ICIJ_Hits { set; get; }
        public string Press_Media { set; get; }
        public string has_ucc_fillings { set; get; }        
        public bool has_ucc_fillings1 { set; get; }        
        public bool has_civil_resultpending { set; get; }        
        public bool HasBankruptcyRecHits1 { set; get; }        
        public bool HasBankruptcyRecHits_resultpending { set; get; }        
        public bool Has_CriminalRecHit1 { set; get; }        
        public bool Has_CriminalRecHit_resultpending { set; get; }
        //Additional states
        public List<Additional_states> additionalstates { set; get; }
        //Employeer
        public List<Employer> employerList { get; set; }
        //Education
        public List<Education> educationList { get; set; }
        //PLLicense
        public List<PLLicense> plLicenseList { get; set; }
        public SummaryResulttableModel summary { get; set; }
    }
}