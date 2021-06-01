using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DiligenceReportCreation.Models
{
    // public class person
    //  {
    //     public string firstName { set; get; }
    //     public string lastName { set; get; }
    //}    
    public class DiligenceInput
    {                    
        public string record_Id { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Foreignlanguage { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string ClientName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CaseNumber { set; get; } 
        public string Marital_Status { set; get; }
        public string Children { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Military_Service { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string source_of_wealth { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Character_Reference { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string worldcheck_discloseBA { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string undisclosedBA { set; get; }
        public string worldcheck_undiscloseBA { set; get; }
        public string discreet_reputational_inquiries { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string FirstName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string MiddleName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string MiddleInitial { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string LastName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string MaidenName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string FullaliasName { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Dob { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Pob { set; get; }        
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
        public string CurrentCountry { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string CurrentZipcode { set; get; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string City { set; get; }
        public string Country { set; get; }
        //public string Recommend_additional_countries { set; get; }    
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string Nonscopecountry1 { set; get; }        
        public string Has_Property_Records { set; get; }        
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
        public string Media_Based_Hits { set; get; }        
        public string USLegal_Record_Hits { set; get; }        
        public string Global_Security_Hits { set; get; }
        public string Global_Sec_Family_Hits { set; get; }
        public string PEP_Hits { set; get; }
        public string ICIJ_Hits { set; get; }        
        public string Press_Media { set; get; }
        public string Crim_Clearance_Certifi { set; get; }
        //France
        public string France_SpeRegHits { set; get; }
        public string France_LegalRecHits { set; get; }
        //UK
        public string UK_InsolvencyHits { set; get; }
        public string UK_CivilRecHit { set; get; }
        public string UK_RegTrustHit { set; get; }
        public string UK_CriminalRecord { set; get; }
        public string UK_DrivingHistory { set; get; }
        public string UK_CreditHistory { set; get; }
        //Canada
        public string Can_DrivingHistory { set; get; }
        public string Can_CreditHistory { set; get; }
        public string Can_SpecRegHits { set; get; }
        public string Can_InsolvencyHits { set; get; }
        public string Can_CivilRecHit { set; get; }
        public string Can_CriminalRecord { set; get; }
        public string Can_PPRRecordHits { set; get; }
        //Australia
        public string Aus_DrivingHistory { set; get; }
        public string Aus_CreditHistory { set; get; }
        public string Aus_ASICReg { set; get; }
        public string Aus_InsolvencyHits { set; get; }
        public string Aus_CivilRecHit { set; get; }
        public string Aus_CriminalRecord { set; get; }
        public string Aus_PPSRRecordHits { set; get; }        
        //United Arab Emirates
        public string RegisteredWithHKSFC { set; get; }
        public string HasNameOnlyMatch { set; get; }
        public string HasNameOnlyBanruptyMatch { set; get; }
        public string HasNameOnlyCriminalMatch { set; get; }
        public string UnitedArabEmi_DrivingHistory { set; get; }
        public string UnitedArabEmi_CreditHistory { set; get; }        
        //Others
        public string others_SpeRegHits { get; set; }
        public string others_LegalRecHits { set; get; }
        public string others_DrivingHistory { set; get; }
        public string others_CreditHistory { set; get; }
        //Germany
        public string Germany_CreditHistory { set; get; }        
        public string Germany_regsearches { set; get; }
        //Switzerland
        public string Switzerland_CreditHistory { set; get; }        
        //India        
        public string India_DrivingHistory { set; get; }
        public string India_CreditHistory { set; get; }
        public string India_Corpregistry { set; get; }
        public SummaryTable summary { set; get; }
        //Employeer
        public List<Employer> employerList { get; set; }
        //Education
        public List<Education> educationList { get; set; }
        public List<PLLicense> plLicenseList { get; set; }   
        //Refernce details
        public List<Referencetable> reftable { get; set; }
        //current residence
        public List<CurrentResidence> currentResidences { get; set; }
        //previous residence
        public List<PreviousResidence> previousResidences { get; set; }
        public List<Additional_states> additional_States { set; get; }
        public List<Additional_countries> additional_Countries { set; get; }
        public List<Family> family { set; get; }
    }
}
