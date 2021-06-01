using System;
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations.Schema;

namespace DiligenceReportCreation.Models
{
    [Table(name: "diligence_summaryresult")]
    public class SummaryResulttableModel
    {
        [Key]
        [Column(name: "record_id")]
        public string record_Id { set; get; }
        [Column(name: "personal_Identification")]
        public string personal_Identification { get; set; }
        [Column(name: "Summarytable")]
        public string Summarytable { get; set; }
        [Column(name: "name_add_Ver_History")]
        public string name_add_Ver_History { get; set; }
        [Column(name: "company_Directorship")]
        public string company_Directorship { get; set; }
        [Column(name: "bankruptcy_filings")]
        public string bankruptcy_filings { get; set; }
        [Column(name: "bankruptcy_filings1")]
        public bool bankruptcy_filings1 { get; set; }
        [Column(name: "civil_court_Litigation")]
        public string civil_court_Litigation { get; set; }
        [Column(name: "civil_court_Litigation1")]
        public bool civil_court_Litigation1 { get; set; }
        [Column(name: "civil_judge_Liens")]
        public string civil_judge_Liens { get; set; }
        [Column(name: "civil_judge_Liens1")]
        public bool civil_judge_Liens1 { get; set; }
        [Column(name: "criminal_records")]
        public string criminal_records { get; set; }
        [Column(name: "criminal_records1")]
        public bool criminal_records1 { get; set; }
        [Column(name: "social_securitytrace")]
        public string social_securitytrace { get; set; }
        [Column(name: "real_estate_prop")]
        public string real_estate_prop { get; set; }
        [Column(name: "secretary_state_director")]
        public string secretary_state_director { get; set; }
        [Column(name: "driving_history")]
        public string driving_history { get; set; }
        [Column(name: "credit_history")]
        public string credit_history { get; set; }
        [Column(name: "uniform_commercial")]
        public string uniform_commercial { get; set; }
        [Column(name: "uniform_commercial1")]
        public string uniform_commercial1 { get; set; }
        [Column(name: "news_media_searches")]
        public string news_media_searches { get; set; }
        [Column(name: "department_foreign")]
        public string department_foreign { get; set; }
        [Column(name: "european_union")]
        public string european_union { get; set; }
        [Column(name: "HM_treasury")]
        public string HM_treasury { get; set; }
        [Column(name: "US_bureau")]
        public string US_bureau { get; set; }
        [Column(name: "US_department")]
        public string US_department { get; set; }
        [Column(name: "US_Directorate")]
        public string US_Directorate { get; set; }
        [Column(name: "US_general")]
        public string US_general { get; set; }
        [Column(name: "US_office")]
        public string US_office { get; set; }
        [Column(name: "UN_consolidated")]
        public string UN_consolidated { get; set; }
        [Column(name: "world_bank_list")]
        public string world_bank_list { get; set; }
        [Column(name: "city_london_police")]
        public string city_london_police { get; set; }
        [Column(name: "constabularies_cheshire")]
        public string constabularies_cheshire { get; set; }
        [Column(name: "hampshire_police")]
        public string hampshire_police { get; set; }
        [Column(name: "hong_kong_police")]
        public string hong_kong_police { get; set; }
        [Column(name: "interpol")]
        public string interpol { get; set; }
        [Column(name: "metropolitan_police")]
        public string metropolitan_police { get; set; }
        [Column(name: "national_crime")]
        public string national_crime { get; set; }
        [Column(name: "north_yorkshire_polic")]
        public string north_yorkshire_polic { get; set; }
        [Column(name: "nottinghamshire_police")]
        public string nottinghamshire_police { get; set; }
        [Column(name: "surrey_police")]
        public string surrey_police { get; set; }
        [Column(name: "thames_valley_police")]
        public string thames_valley_police { get; set; }
        [Column(name: "US_Federal")]
        public string US_Federal { get; set; }
        [Column(name: "US_secret_service")]
        public string US_secret_service { get; set; }
        [Column(name: "warwickshire_police")]
        public string warwickshire_police { get; set; }
        [Column(name: "alberta_securities_commission")]
        public string alberta_securities_commission { get; set; }
        [Column(name: "asset_recovery_agency")]
        public string asset_recovery_agency { get; set; }
        [Column(name: "australian_prudential")]
        public string australian_prudential { get; set; }
        [Column(name: "australian_securities")]
        public string australian_securities { get; set; }
        [Column(name: "banque_de_CECEI")]
        public string banque_de_CECEI { get; set; }
        [Column(name: "banque_de_commission")]
        public string banque_de_commission { get; set; }
        [Column(name: "british_virgin_islands")]
        public string british_virgin_islands { get; set; }
        [Column(name: "cayman_islands_monetary")]
        public string cayman_islands_monetary { get; set; }
        [Column(name: "commission_de_surveillance")]
        public string commission_de_surveillance { get; set; }
        [Column(name: "commodity_futures")]
        public string commodity_futures { get; set; }
        [Column(name: "council_financial_activities")]
        public string council_financial_activities { get; set; }
        [Column(name: "departamento_de_investigacoes")]
        public string departamento_de_investigacoes { get; set; }
        [Column(name: "department_labour_inspection")]
        public string department_labour_inspection { get; set; }
        [Column(name: "federal_deposit")]
        public string federal_deposit { get; set; }
        [Column(name: "financial_action_task")]
        public string financial_action_task { get; set; }
        [Column(name: "federal_reserve")]
        public string federal_reserve { get; set; }
        [Column(name: "financial_crimes")]
        public string financial_crimes { get; set; }
        [Column(name: "financial_industry")]
        public string financial_industry { get; set; }
        [Column(name: "international_consortium")]
        public string international_consortium { get; set; }
        [Column(name: "financial_regulator_ireland")]
        public string financial_regulator_ireland { get; set; }
        [Column(name: "hongkong_monetary_authority")]
        public string hongkong_monetary_authority { get; set; }
        [Column(name: "hongkong_securities_futures")]
        public string hongkong_securities_futures { get; set; }
        [Column(name: "investment_association_Canada")]
        public string investment_association_Canada { get; set; }
        [Column(name: "investment_management_regulatory")]
        public string investment_management_regulatory { get; set; }
        [Column(name: "isle_financial_supervision")]
        public string isle_financial_supervision { get; set; }
        [Column(name: "jersey_financial_commission")]
        public string jersey_financial_commission { get; set; }
        [Column(name: "lloyd_insurance_arimbolaet")]
        public string lloyd_insurance_arimbolaet { get; set; }
        [Column(name: "monetary_authority_singapore")]
        public string monetary_authority_singapore { get; set; }
        [Column(name: "national_credit")]
        public string national_credit { get; set; }
        [Column(name: "new_york_stock")]
        public string new_york_stock { get; set; }
        [Column(name: "Office_of_comptroller")]
        public string Office_of_comptroller { get; set; }
        [Column(name: "Office_of_superintendent")]
        public string Office_of_superintendent { get; set; }
        [Column(name: "resolution_trust")]
        public string resolution_trust { get; set; }
        [Column(name: "securities_exchange")]
        public string securities_exchange { get; set; }
        [Column(name: "securities_exchange_commission")]
        public string securities_exchange_commission { get; set; }
        [Column(name: "securities_futuresauthority")]
        public string securities_futuresauthority { get; set; }
        [Column(name: "swedish_financial_supervisory")]
        public string swedish_financial_supervisory { get; set; }
        [Column(name: "swiss_federal_banking")]
        public string swiss_federal_banking { get; set; }
        [Column(name: "U_K_companies_disqualified")]
        public string U_K_companies_disqualified { get; set; }
        [Column(name: "U_K_financial_conduct_authority")]
        public string U_K_financial_conduct_authority { get; set; }
        [Column(name: "US_court")]
        public string US_court { get; set; }
        [Column(name: "US_department_justice")]
        public string US_department_justice { get; set; }
        [Column(name: "US_federal_trade")]
        public string US_federal_trade { get; set; }
        [Column(name: "US_national")]
        public string US_national { get; set; }
        [Column(name: "US_office_thrifts")]
        public string US_office_thrifts { get; set; }
        [Column(name: "central_intelligence")]
        public string central_intelligence { get; set; }
        [Column(name: "international_consortium_investigative")]
        public string international_consortium_investigative { get; set; }

    }
}
