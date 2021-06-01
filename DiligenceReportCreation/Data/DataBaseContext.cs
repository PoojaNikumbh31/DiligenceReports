using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DiligenceReportCreation.Models;
using Microsoft.EntityFrameworkCore;

namespace DiligenceReportCreation.Data
{
    public class DataBaseContext : DbContext
    {
        public DataBaseContext(DbContextOptions<DataBaseContext> options) : base(options)
        {

        }        
        public DbSet<UserModel> DbUser { get; set; }
        public DbSet<DiligenceInputModel> DbPersonalInfo { get; set; }
        public DbSet<EducationModel> DbEducation { get; set; }
        public DbSet<EmployerModel> DbEmployer { get; set; }
        public DbSet<PLLicenseModel> DbPLLicense { get; set; }
        public DbSet<CommentModel> DbComment { get; set; }
        public DbSet<CountrySpecificModel> CSComment { get; set; }
        public DbSet<ReportModel> reportModel { get; set; }
        public DbSet<StateWiseFootnoteModel> stateModel { get; set; }
        public DbSet<Additional_statesModel> Dbadditionalstates { get; set; }
        public DbSet<Additional_CountriesModel> DbadditionalCountries { get; set; }
        public DbSet<Otherdatails>othersModel { get; set; }
        public DbSet<SummaryResulttableModel> summaryResulttableModels { set; get; }
        public DbSet<ClientspecificModel> DbclientspecificModels { get; set; }
        public DbSet<ReferencetableModel> DbReferencetableModel { get; set; }
        public DbSet<PreviousResidenceModel> DbpreviousResidenceModels { get; set; }
        public DbSet<CurrentResidenceModel> DbcurrentResidenceModels { get; set; }
        public DbSet<FamilyModel> familyModels { set; get; }
    }

}