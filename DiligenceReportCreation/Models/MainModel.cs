using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DiligenceReportCreation.Models
{
    public class MainModel
    {
        public DiligenceInputModel diligenceInputModel { set; get; }
        public List<EmployerModel> EmployerModel { set; get;}
        public List<EducationModel> educationModels { set; get; }
        public List<PLLicenseModel> pllicenseModels { set; get; }        
        public Otherdatails otherdetails { set; get; }
        public CountrySpecificModel csModel { set; get; }
        public SummaryResulttableModel summarymodel { set; get; }
        public List<Additional_statesModel> additional_States { set; get; }
        public List<Additional_CountriesModel> additional_Countries { set; get; }
        public ClientspecificModel clientspecific { set; get; }
        public List<ReferencetableModel> referencetableModels { set; get; }
        public List<CurrentResidenceModel> currentResidenceModels { set; get; }
        public List<PreviousResidenceModel> PreviousResidenceModels { set; get; }
        public List<FamilyModel> familyModels { set; get; }
        public DiligenceInput diligence { get; set; }
    }
}
