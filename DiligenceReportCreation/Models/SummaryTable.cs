using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{
    public class SummaryTable
    {
        public string personal_Identification { get; set; }
        public string name_add_Ver_History { get; set; }
        public string company_Directorship { get; set; }
        public string bankruptcy_filings { get; set; }
        public bool bankruptcy_filings1 { get; set; }
        public string civil_court_Litigation { get; set; }
        public bool civil_court_Litigation1 { get; set; }
        public string civil_judge_Liens { get; set; }
        public bool civil_judge_Liens1 { get; set; }
        public string criminal_records { get; set; }
        public bool criminal_records1 { get; set; }
        public string social_securitytrace { get; set; }
        public string real_estate_prop { get; set; }
        public string secretary_state_director { get; set; }
        public string driving_history { get; set; }
        public string credit_history { get; set; }
        public string uniform_commercial { get; set; }
        public string uniform_commercial1 { get; set; }
        public string news_media_searches { get; set; }
        public string department_foreign { get; set; }
        public string european_union { get; set; }
        public string HM_treasury { get; set; }
        public string US_bureau { get; set; }
        public string US_department { get; set; }
        public string US_Directorate { get; set; }
        public string US_general { get; set; }
        [DisplayFormat(ConvertEmptyStringToNull = false)]
        public string US_office { get; set; }
        public string UN_consolidated { get; set; }
        public string world_bank_list { get; set; }
        public string city_london_police { get; set; }
        public string constabularies_cheshire { get; set; }
        public string hampshire_police { get; set; }
        public string hong_kong_police { get; set; }
        public string interpol { get; set; }
        public string metropolitan_police { get; set; }
        public string national_crime { get; set; }
        public string north_yorkshire_polic { get; set; }
        public string nottinghamshire_police { get; set; }
        public string surrey_police { get; set; }
        public string thames_valley_police { get; set; }
        public string US_Federal { get; set; }
        public string US_secret_service { get; set; }
        public string warwickshire_police { get; set; }
        public string alberta_securities_commission { get; set; }
        public string asset_recovery_agency { get; set; }
        public string australian_prudential { get; set; }
        public string australian_securities { get; set; }
        public string banque_de_CECEI { get; set; }
        public string banque_de_commission { get; set; }
        public string british_virgin_islands { get; set; }
        public string cayman_islands_monetary { get; set; }
        public string commission_de_surveillance { get; set; }
        public string commodity_futures { get; set; }
        public string council_financial_activities { get; set; }
        public string departamento_de_investigacoes { get;set;}
        public string department_labour_inspection { get;set;}        
        public string federal_deposit { get; set; }
        public string financial_action_task { get; set; }
        public string federal_reserve { get; set; }
        public string financial_crimes { get; set; }
        public string financial_industry { get; set; }
        public string international_consortium { get; set; }
        public string financial_regulator_ireland { get;set;}
        public string hongkong_monetary_authority { get;set;}
        public string hongkong_securities_futures { get;set;}
        public string investment_association_Canada { get;set;}
        public string investment_management_regulatory { get;set;}        
        public string isle_financial_supervision { get;set;}
        public string jersey_financial_commission { get;set;}
        public string lloyd_insurance_arimbolaet { get;set;}
        public string monetary_authority_singapore { get;set;}
        public string national_credit { get; set; }
        public string new_york_stock { get; set; }
        public string Office_of_comptroller { get; set; }
        public string Office_of_superintendent { get; set; }
        public string resolution_trust { get; set; }
        public string securities_exchange { get; set; }
        public string securities_exchange_commission { get; set; }
        public string securities_futuresauthority { get; set; }
        public string swedish_financial_supervisory { get; set; }
        public string swiss_federal_banking { get; set; }
        public string U_K_companies_disqualified { get; set; }
        public string U_K_financial_conduct_authority { get; set; }
        public string US_court { get; set; }
        public string US_department_justice { get; set; }
        public string US_federal_trade { get; set; }
        public string US_national { get; set; }
        public string US_office_thrifts { get; set; }
        public string central_intelligence { get; set; }
        public string international_consortium_investigative { get; set; }
    }
}
