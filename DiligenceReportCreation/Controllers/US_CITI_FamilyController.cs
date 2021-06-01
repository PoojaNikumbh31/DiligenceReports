using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using DiligenceReportCreation.Models;
using DiligenceReportCreation.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System.Globalization;
using System.ComponentModel.DataAnnotations.Schema;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using System.IO;
using Spire.Doc.Fields;
using System.Text.RegularExpressions;

namespace DiligenceReportCreation.Controllers
{
    public class US_CITI_FamilyController : Controller
    {
        private readonly DataBaseContext _context;
        private readonly IConfiguration _config;
        public US_CITI_FamilyController(IConfiguration config, DataBaseContext context)
        {
            _config = config;
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult US_Citizen_Family_Newpage()
        {
            string recordid = HttpContext.Session.GetString("recordid");
            MainModel mainModel = new MainModel();
            mainModel.familyModels = _context.familyModels
                        .Where(u => u.record_Id == recordid)
                        .ToList();
            return View(mainModel);
        }
        [HttpPost]
        public IActionResult US_Citizen_Family_Newpage(MainModel main, string SaveData, string Submit)
        {
            string recordid = HttpContext.Session.GetString("recordid");
            MainModel mainModel = new MainModel();
            try
            {
                if (SaveData.StartsWith("Delete_"))
                {
                    string[] str = new string[3];
                    str = SaveData.Split("_");
                    string id_c = str[1];

                    FamilyModel family = _context.familyModels
                              .Where(a => a.Family_record_id == id_c).FirstOrDefault();
                    _context.familyModels.Remove(family);
                    _context.SaveChanges();
                    mainModel.familyModels = _context.familyModels
                    .Where(u => u.record_Id == recordid)
                    .ToList();
                }
            }
            catch { }
            if (Submit == "Submit Data")
            {
                MainModel mainmodel = new MainModel();
                DiligenceInputModel dinput = new DiligenceInputModel();
                Otherdatails otherdatails = new Otherdatails();                
                EmployerModel employer = new EmployerModel();
                EducationModel education = new EducationModel();
                PLLicenseModel pL = new PLLicenseModel();
                ClientspecificModel clientspecific = new ClientspecificModel();
                SummaryResulttableModel summary = new SummaryResulttableModel();
                CurrentResidenceModel currentResidence = new CurrentResidenceModel();
                // PreviousResidenceModel previous = new PreviousResidenceModel();
                ReferencetableModel referencetable = new ReferencetableModel();
                mainmodel.familyModels = _context.familyModels
                     .Where(a => a.record_Id == recordid)
                     .ToList();
                for (int i = 0; i < mainmodel.familyModels.Count; i++)
                {
                    if (mainmodel.familyModels[i].adult_minor == "Adult")
                    {
                        if (mainmodel.familyModels[i].applicant_type == "Main Applicant")
                        {
                            try
                            {
                                clientspecific = _context.DbclientspecificModels
                                                  .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                                  .FirstOrDefault();
                                if (clientspecific == null)
                                {
                                    TempData["message"] = "Please enter the details in client specific section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                                }
                                else
                                {
                                    if (clientspecific.clientname.ToString().Equals("Dominica"))
                                    {
                                        if (clientspecific.Character_Reference.ToString().Equals("Provided but Unsuccessful") || clientspecific.Character_Reference.ToString().Equals("Provided and Contacted"))
                                        {
                                            try
                                            {
                                                referencetable = _context.DbReferencetableModel
                                                            .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                                            .FirstOrDefault();
                                                if (referencetable == null)
                                                {
                                                    TempData["message"] = "Please enter the details in Reference Table details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                                                }
                                            }
                                            catch
                                            {
                                                TempData["message"] = "Please enter the details in Reference Table details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                TempData["message"] = "Please enter the details in client specific section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                            }
                        }
                        try
                        {
                            otherdatails = _context.othersModel
                                       .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                       .FirstOrDefault();
                            if (otherdatails == null)
                            {
                                TempData["message"] = "Please enter the details in other details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                            }
                        }
                        catch
                        {
                            TempData["message"] = "Please enter the details in other details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                        }                       
                        try
                        {
                            employer = _context.DbEmployer
                                        .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                        .FirstOrDefault();
                            if (employer == null)
                            {
                                TempData["message"] = "Please enter the details in Employee details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                            }
                        }
                        catch
                        {
                            TempData["message"] = "Please enter the details in Employee details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                        }
                        try
                        {
                            education = _context.DbEducation
                                       .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                       .FirstOrDefault();
                            if (education == null)
                            {
                                TempData["message"] = "Please enter the details in Education details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                            }
                        }
                        catch
                        {
                            TempData["message"] = "Please enter the details in Education details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                        }
                        try
                        {
                            pL = _context.DbPLLicense
                                       .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                        .FirstOrDefault();
                            if (pL == null)
                            {
                                TempData["message"] = "Please enter the details in Pllicence details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                            }
                        }
                        catch
                        {
                            TempData["message"] = "Please enter the details in Pllicence details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                        }
                        try
                        {
                            summary = _context.summaryResulttableModels
                                              .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                              .FirstOrDefault();
                            if (summary == null)
                            {
                                TempData["message"] = "Please enter the details in summary result table section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                            }
                        }
                        catch
                        {
                            TempData["message"] = "Please enter the details in summary result table section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                        }

                    }
                    else { }
                    try
                    {
                        currentResidence = _context.DbcurrentResidenceModels
                                    .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                    .FirstOrDefault();
                        if (currentResidence == null)
                        {
                            TempData["message"] = "Please enter the details in Current Residence details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                        }
                    }
                    catch
                    {
                        TempData["message"] = "Please enter the details in Current Residence details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                    }
                    try
                    {
                        dinput = _context.DbPersonalInfo
                                          .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                          .FirstOrDefault();
                        if (dinput == null)
                        {
                            TempData["message"] = "Please enter the details in personal details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                        }
                    }
                    catch
                    {
                        TempData["message"] = "Please enter the details in personal details section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                    }

                }
                if (TempData["message"] == null)
                {
                    HttpContext.Session.SetString("recordid", recordid);
                    return Save_Page1(mainmodel);
                }
            }
            try
            {
                switch (SaveData)
                {
                    case "Save Data":
                        try
                        {
                            if (recordid == null)
                            {
                                TempData["message"] = "Session is Expired";
                                return RedirectToAction("LoginFormView", "Home");
                            }
                            else
                            {
                                try
                                {
                                    ReportModel report = _context.reportModel
                                              .Where(a => a.record_Id == recordid)
                                              .FirstOrDefault();
                                    List<FamilyModel> familyModels = _context.familyModels
                                              .Where(a => a.record_Id == recordid)
                                              .ToList();
                                    if (familyModels == null || familyModels.Count == 0)
                                    {
                                        for (int i = 0; i < main.diligence.family.Count; i++)
                                        {
                                            FamilyModel familyModel = new FamilyModel();
                                            familyModel.record_Id = recordid;
                                            string family_record_id = Guid.NewGuid().ToString();
                                            familyModel.Family_record_id = family_record_id;
                                            familyModel.first_name = main.diligence.family[i].first_name;
                                            familyModel.middle_name = main.diligence.family[i].middle_name;
                                            familyModel.last_name = main.diligence.family[i].last_name;
                                            familyModel.applicant_type = main.diligence.family[i].applicant_type;
                                            familyModel.adult_minor = main.diligence.family[i].adult_minor;
                                            familyModel.case_number = report.casenumber;
                                            _context.familyModels.Add(familyModel);
                                            _context.SaveChanges();
                                            DiligenceInputModel diligenceInput = new DiligenceInputModel();
                                            diligenceInput.record_Id = family_record_id;
                                            diligenceInput.FirstName = main.diligence.family[i].first_name;
                                            diligenceInput.MiddleName = main.diligence.family[i].middle_name;
                                            diligenceInput.LastName = main.diligence.family[i].last_name;
                                            diligenceInput.CaseNumber = report.casenumber;
                                            _context.DbPersonalInfo.Add(diligenceInput);
                                            _context.SaveChanges();
                                        }
                                    }
                                    else
                                    {

                                        for (int i = 0; i < familyModels.Count; i++)
                                        {
                                            familyModels[i].record_Id = recordid;
                                            familyModels[i].first_name = main.familyModels[i].first_name;
                                            // familyModels[i].Family_record_id = main.familyModels[i].Family_record_id;
                                            familyModels[i].middle_name = main.familyModels[i].middle_name;
                                            familyModels[i].last_name = main.familyModels[i].last_name;
                                            familyModels[i].applicant_type = main.familyModels[i].applicant_type;
                                            familyModels[i].adult_minor = main.familyModels[i].adult_minor;
                                            familyModels[i].case_number = report.casenumber;
                                            _context.familyModels.Update(familyModels[i]);
                                            _context.SaveChanges();

                                            DiligenceInputModel diligenceInput = _context.DbPersonalInfo
                               .Where(u => u.record_Id == familyModels[i].Family_record_id)
                                              .FirstOrDefault();
                                            if (diligenceInput == null) { diligenceInput = new DiligenceInputModel(); }
                                            diligenceInput.record_Id = familyModels[i].Family_record_id;
                                            diligenceInput.FirstName = main.familyModels[i].first_name;
                                            diligenceInput.MiddleName = main.familyModels[i].middle_name;
                                            diligenceInput.LastName = main.familyModels[i].last_name;
                                            diligenceInput.CaseNumber = report.casenumber;
                                            try
                                            {
                                                _context.DbPersonalInfo.Update(diligenceInput);
                                                _context.SaveChanges();
                                            }
                                            catch
                                            {
                                                _context.DbPersonalInfo.Add(diligenceInput);
                                                _context.SaveChanges();
                                            }

                                        }
                                        try
                                        {
                                            for (int i = 0; i < main.diligence.family.Count; i++)
                                            {
                                                if (main.diligence.family[i].first_name.ToString().Equals("")) { }
                                                else
                                                {
                                                    FamilyModel familyModel = new FamilyModel();
                                                    familyModel.record_Id = recordid;
                                                    string family_record_id = Guid.NewGuid().ToString();
                                                    familyModel.Family_record_id = family_record_id;
                                                    familyModel.first_name = main.diligence.family[i].first_name;
                                                    familyModel.middle_name = main.diligence.family[i].middle_name;
                                                    familyModel.last_name = main.diligence.family[i].last_name;
                                                    familyModel.applicant_type = main.diligence.family[i].applicant_type;
                                                    familyModel.adult_minor = main.diligence.family[i].adult_minor;
                                                    familyModel.case_number = report.casenumber;
                                                    _context.familyModels.Add(familyModel);
                                                    _context.SaveChanges();
                                                    DiligenceInputModel diligenceInput = new DiligenceInputModel();
                                                    diligenceInput.record_Id = family_record_id;
                                                    diligenceInput.FirstName = main.diligence.family[i].first_name;
                                                    diligenceInput.MiddleName = main.diligence.family[i].middle_name;
                                                    diligenceInput.LastName = main.diligence.family[i].last_name;
                                                    diligenceInput.CaseNumber = report.casenumber;
                                                    _context.DbPersonalInfo.Add(diligenceInput);
                                                    _context.SaveChanges();
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                catch
                                {
                                    ReportModel report = _context.reportModel
                                             .Where(a => a.record_Id == recordid)
                                             .FirstOrDefault();
                                    for (int i = 0; i < main.diligence.family.Count; i++)
                                    {
                                        try
                                        {
                                            if (main.diligence.family[i].first_name.ToString().Equals("")) { }
                                            else
                                            {
                                                FamilyModel familyModel = new FamilyModel();
                                                familyModel.record_Id = recordid;
                                                string family_record_id = Guid.NewGuid().ToString();
                                                familyModel.Family_record_id = family_record_id;
                                                familyModel.first_name = main.diligence.family[i].first_name;
                                                familyModel.middle_name = main.diligence.family[i].middle_name;
                                                familyModel.last_name = main.diligence.family[i].last_name;
                                                familyModel.applicant_type = main.diligence.family[i].applicant_type;
                                                familyModel.adult_minor = main.diligence.family[i].adult_minor;
                                                familyModel.case_number = report.casenumber;
                                                _context.familyModels.Add(familyModel);
                                                _context.SaveChanges();
                                                DiligenceInputModel diligenceInput = new DiligenceInputModel();
                                                diligenceInput.record_Id = family_record_id;
                                                diligenceInput.FirstName = main.diligence.family[i].first_name;
                                                diligenceInput.MiddleName = main.diligence.family[i].middle_name;
                                                diligenceInput.LastName = main.diligence.family[i].last_name;
                                                diligenceInput.CaseNumber = report.casenumber;
                                                _context.DbPersonalInfo.Add(diligenceInput);
                                                _context.SaveChanges();
                                            }
                                        }
                                        catch { }
                                    }
                                }
                            }
                        }
                        catch (IOException e)
                        {
                            TempData["message"] = e;
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        break;
                }
            }
            catch { }
            mainModel.familyModels = _context.familyModels
                       .Where(u => u.record_Id == recordid)
                       .ToList();
            HttpContext.Session.SetString("recordid", recordid);
            return View(mainModel);
        }
        [HttpGet]
        public IActionResult Edit(string id)
        {
            string recordid = HttpContext.Session.GetString("recordid");
            MainModel mainModel = new MainModel();
            mainModel.familyModels = _context.familyModels
                  .Where(u => u.Family_record_id == id.ToString())
                  .ToList();
            HttpContext.Session.SetString("recordid", recordid);
            ViewBag.firstName = mainModel.familyModels[0].first_name.ToUpper();
            ViewBag.lastName = mainModel.familyModels[0].last_name.ToUpper();
            ViewBag.caseNumber = mainModel.familyModels[0].case_number;
            ViewBag.adultMinor = mainModel.familyModels[0].adult_minor;
            ViewBag.applicanttype = mainModel.familyModels[0].applicant_type;
            try
            {
                ViewBag.middleName = mainModel.familyModels[0].middle_name;
            }
            catch { }
            try
            {
                mainModel.diligenceInputModel = _context.DbPersonalInfo
                 .Where(u => u.record_Id == id.ToString())
                 .FirstOrDefault();
            }
            catch { }
            try
            {
                mainModel.additional_States = _context.Dbadditionalstates
                            .Where(u => u.record_Id == id.ToString())
                            .ToList();
            }
            catch { }
            try
            {
                mainModel.educationModels = _context.DbEducation
                 .Where(u => u.record_Id == id.ToString())
                 .ToList();
            }
            catch { }
            try
            {
                mainModel.EmployerModel = _context.DbEmployer
                 .Where(u => u.record_Id == id.ToString())
                 .ToList();
            }
            catch { }
            try
            {
                mainModel.pllicenseModels = _context.DbPLLicense
                    .Where(u => u.record_Id == id.ToString())
                    .ToList();
            }
            catch { }
            try
            {
                mainModel.PreviousResidenceModels = _context.DbpreviousResidenceModels
                    .Where(u => u.record_Id == id.ToString())
                    .ToList();
            }
            catch { }
            try
            {
                mainModel.currentResidenceModels = _context.DbcurrentResidenceModels
                    .Where(u => u.record_Id == id.ToString())
                    .ToList();
            }
            catch { }
            try
            {
                mainModel.referencetableModels = _context.DbReferencetableModel
                    .Where(u => u.record_Id == id.ToString())
                    .ToList();
            }
            catch { }
            try
            {
                mainModel.summarymodel = _context.summaryResulttableModels
                    .Where(u => u.record_Id == id.ToString())
                    .FirstOrDefault();
            }
            catch { }
            try
            {
                mainModel.clientspecific = _context.DbclientspecificModels
                    .Where(u => u.record_Id == id.ToString())
                    .FirstOrDefault();
            }
            catch { }
            
            List<string> CountryList = new List<string>();
            CultureInfo[] CInfoList = CultureInfo.GetCultures(CultureTypes.SpecificCultures);
            foreach (CultureInfo CInfo in CInfoList)
            {
                RegionInfo R = new RegionInfo(CInfo.LCID);
                if (!(CountryList.Contains(R.EnglishName)))
                {
                    CountryList.Add(R.EnglishName);
                }
            }
            CountryList.Sort();
            ViewBag.CountryList = CountryList;
            HttpContext.Session.SetString("firstName", mainModel.familyModels[0].first_name);
            HttpContext.Session.SetString("lastName", mainModel.familyModels[0].last_name);
            HttpContext.Session.SetString("caseNumber", mainModel.familyModels[0].case_number);
            HttpContext.Session.SetString("adultMinor", mainModel.familyModels[0].adult_minor);
            HttpContext.Session.SetString("applicanttype", mainModel.familyModels[0].applicant_type);
            try
            {
                HttpContext.Session.SetString("middleName", mainModel.familyModels[0].middle_name);
            }
            catch { }
            HttpContext.Session.SetString("family_record_id", id);
            return View(mainModel);
        }
        [HttpPost]
        public IActionResult Edit(MainModel mainModel, string SaveData)
        {
            ViewBag.firstName = HttpContext.Session.GetString("firstName");
            ViewBag.lastName = HttpContext.Session.GetString("lastName");
            ViewBag.caseNumber = HttpContext.Session.GetString("caseNumber");
            ViewBag.adultMinor = HttpContext.Session.GetString("adultMinor");
            ViewBag.middleName = HttpContext.Session.GetString("middleName");
            ViewBag.applicanttype = HttpContext.Session.GetString("applicanttype");
            string recordid = HttpContext.Session.GetString("recordid");
            string family_record_id = HttpContext.Session.GetString("family_record_id");
            HttpContext.Session.SetString("recordid", recordid);
            ReportModel report = _context.reportModel
               .Where(u => u.record_Id == recordid)
                                 .FirstOrDefault();
            string lastname = HttpContext.Session.GetString("lastname");
            string casenumber = HttpContext.Session.GetString("casenumber");
            List<string> CountryList = new List<string>();
            CultureInfo[] CInfoList = CultureInfo.GetCultures(CultureTypes.SpecificCultures);
            foreach (CultureInfo CInfo in CInfoList)
            {
                RegionInfo R = new RegionInfo(CInfo.LCID);
                if (!(CountryList.Contains(R.EnglishName)))
                {
                    CountryList.Add(R.EnglishName);
                }
            }
            CountryList.Sort();
            ViewBag.CountryList = CountryList;
            if (SaveData == "Save")
            {
                mainModel.familyModels = _context.familyModels
                       .Where(u => u.record_Id == recordid)
                       .ToList();
                return RedirectToAction("US_Citizen_Family_Newpage", "US_CITI_Family");
            }
            switch (SaveData)
            {
                case "SavePIData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            try
                            {
                                DiligenceInputModel inputModel = _context.DbPersonalInfo
                           .Where(u => u.record_Id == family_record_id)
                                          .FirstOrDefault();
                                inputModel.record_Id = family_record_id;
                                inputModel.FirstName = ViewBag.firstName;
                                inputModel.LastName = ViewBag.lastName;
                                inputModel.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                                inputModel.MiddleName = ViewBag.middleName;
                                inputModel.CaseNumber = report.casenumber;
                                inputModel.MaidenName = mainModel.diligenceInputModel.MaidenName;
                                inputModel.FullaliasName = mainModel.diligenceInputModel.FullaliasName;
                                inputModel.foreignlanguage = mainModel.diligenceInputModel.foreignlanguage;
                                inputModel.Dob = mainModel.diligenceInputModel.Dob;
                                inputModel.Pob = mainModel.diligenceInputModel.Pob;
                                inputModel.SSNInitials = mainModel.diligenceInputModel.SSNInitials;
                                inputModel.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
                                inputModel.CurrentState = mainModel.diligenceInputModel.CurrentState;
                                //try
                                //{
                                //    inputModel.Employer1City = mainModel.diligenceInputModel.Employer1City;
                                //    inputModel.Employer1State = mainModel.diligenceInputModel.Employer1State;
                                //}
                                //catch { }
                                inputModel.CommonNameSubject = mainModel.diligenceInputModel.CommonNameSubject;                                
                                try
                                {
                                    inputModel.Marital_Status = mainModel.diligenceInputModel.Marital_Status;
                                    inputModel.Children = mainModel.diligenceInputModel.Children;
                                }
                                catch { }
                                _context.Entry(inputModel).State = EntityState.Modified;
                                _context.SaveChanges();
                            }
                            catch
                            {
                                DiligenceInputModel inputModel = new DiligenceInputModel();
                                inputModel.record_Id = family_record_id;
                                inputModel.FirstName = ViewBag.firstName;
                                inputModel.LastName = ViewBag.lastName;
                                inputModel.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                                inputModel.MiddleName = ViewBag.middleName;
                                inputModel.CaseNumber = report.casenumber;
                                inputModel.MaidenName = mainModel.diligenceInputModel.MaidenName;
                                inputModel.FullaliasName = mainModel.diligenceInputModel.FullaliasName;
                                inputModel.foreignlanguage = mainModel.diligenceInputModel.foreignlanguage;
                                inputModel.Dob = mainModel.diligenceInputModel.Dob;
                                inputModel.Pob = mainModel.diligenceInputModel.Pob;
                                inputModel.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
                                inputModel.CurrentState = mainModel.diligenceInputModel.CurrentState;
                                //try
                                //{
                                //    inputModel.Employer1City = mainModel.diligenceInputModel.Employer1City;
                                //    inputModel.Employer1State = mainModel.diligenceInputModel.Employer1State;
                                //}
                                //catch { }
                                inputModel.CommonNameSubject = mainModel.diligenceInputModel.CommonNameSubject;
                                try
                                {
                                    inputModel.Marital_Status = mainModel.diligenceInputModel.Marital_Status;
                                    inputModel.Children = mainModel.diligenceInputModel.Children;
                                }
                                catch { }
                                _context.DbPersonalInfo.Add(inputModel);
                                _context.SaveChanges();
                            }
                            TempData["PI"] = "Done";
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SaveOD":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            try
                            {
                                Otherdatails strupdate = _context.othersModel
                               .Where(u => u.record_Id == family_record_id)
                                              .FirstOrDefault();
                                strupdate.record_Id = family_record_id;
                                strupdate.Press_Media = mainModel.otherdetails.Press_Media;                                
                                strupdate.RegulatoryFlag = mainModel.otherdetails.RegulatoryFlag;
                                strupdate.Has_Reg_FINRA = mainModel.otherdetails.Has_Reg_FINRA;
                                strupdate.Has_Reg_UK_FCA = mainModel.otherdetails.Has_Reg_UK_FCA;
                                strupdate.Has_Reg_US_NFA = mainModel.otherdetails.Has_Reg_US_NFA;
                                strupdate.Has_Reg_US_SEC = mainModel.otherdetails.Has_Reg_US_SEC;
                                strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;
                                strupdate.ICIJ_Hits = mainModel.otherdetails.ICIJ_Hits;
                                strupdate.Global_Sec_Family_Hits = mainModel.otherdetails.Global_Sec_Family_Hits;
                                strupdate.Has_Bureau_PrisonHit = mainModel.otherdetails.Has_Bureau_PrisonHit;
                                strupdate.PEP_Hits = mainModel.otherdetails.PEP_Hits;
                                strupdate.undisclosedBA = mainModel.otherdetails.undisclosedBA;
                                strupdate.worldcheck_discloseBA = mainModel.otherdetails.worldcheck_discloseBA;
                                strupdate.worldcheck_undiscloseBA = mainModel.otherdetails.worldcheck_undiscloseBA;
                                strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                                strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                                strupdate.Has_Companion_Report = mainModel.otherdetails.Has_Companion_Report;
                                strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
                                strupdate.Media_Based_Hits = mainModel.otherdetails.Media_Based_Hits;
                                strupdate.Crim_Clearance_Certifi = mainModel.otherdetails.Crim_Clearance_Certifi;
                                strupdate.CurrentResidentialProperty = mainModel.otherdetails.CurrentResidentialProperty;
                                strupdate.OtherPropertyOwnershipinfo = mainModel.otherdetails.OtherPropertyOwnershipinfo;
                                strupdate.OtherCurrentResidentialProperty = mainModel.otherdetails.OtherCurrentResidentialProperty;
                                strupdate.PrevPropertyOwnershipRes = mainModel.otherdetails.PrevPropertyOwnershipRes;
                                strupdate.HasBankruptcyRecHits = mainModel.otherdetails.HasBankruptcyRecHits;
                                strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;
                                strupdate.Has_CriminalRecHit = mainModel.otherdetails.Has_CriminalRecHit;
                                strupdate.Has_Bureau_PrisonHit = mainModel.otherdetails.Has_Bureau_PrisonHit;
                                strupdate.Has_Sex_Offender_RegHit = mainModel.otherdetails.Has_Sex_Offender_RegHit;
                                strupdate.Has_Civil_Records = mainModel.otherdetails.Has_Civil_Records;
                                strupdate.Has_Name_Only = mainModel.otherdetails.Has_Name_Only;
                                strupdate.Has_US_Tax_Court_Hit = mainModel.otherdetails.Has_US_Tax_Court_Hit;
                                strupdate.Has_Tax_Liens = mainModel.otherdetails.Has_Tax_Liens;
                                strupdate.Has_Name_Only_Tax_Lien = mainModel.otherdetails.Has_Name_Only_Tax_Lien;
                                strupdate.Has_Driving_Hits = mainModel.otherdetails.Has_Driving_Hits;
                                var checkBox3 = Request.Form["Has_Name_only_driving_incidents"];
                                if (checkBox3 == "on")
                                {
                                    strupdate.Has_Name_only_driving_incidents = true;
                                }
                                else
                                {
                                    strupdate.Has_Name_only_driving_incidents = mainModel.otherdetails.Has_Name_only_driving_incidents;
                                }
                                strupdate.Was_credited_obtained = mainModel.otherdetails.Was_credited_obtained;                                                                                                
                                strupdate.Has_Property_Records = mainModel.otherdetails.Has_Property_Records;
                                var checkBox4 = Request.Form["HasBankruptcyRecHits1"];
                                if (checkBox4 == "on")
                                {
                                    strupdate.HasBankruptcyRecHits1 = true;
                                }
                                else
                                {
                                    strupdate.HasBankruptcyRecHits1 = mainModel.otherdetails.HasBankruptcyRecHits1;
                                }
                                var checkBox5 = Request.Form["HasBankruptcyRecHits_resultpending"];
                                if (checkBox5 == "on")
                                {
                                    strupdate.HasBankruptcyRecHits_resultpending = true;
                                }
                                else
                                {
                                    strupdate.HasBankruptcyRecHits_resultpending = mainModel.otherdetails.HasBankruptcyRecHits_resultpending;
                                }
                                var checkBox6 = Request.Form["has_civil_resultpending"];
                                if (checkBox6 == "on")
                                {
                                    strupdate.has_civil_resultpending = true;
                                }
                                else
                                {
                                    strupdate.has_civil_resultpending = mainModel.otherdetails.has_civil_resultpending;
                                }
                                var checkBox7 = Request.Form["Has_CriminalRecHit1"];
                                if (checkBox7 == "on")
                                {
                                    strupdate.Has_CriminalRecHit1 = true;
                                }
                                else
                                {
                                    strupdate.Has_CriminalRecHit1 = mainModel.otherdetails.Has_CriminalRecHit1;
                                }
                                var checkBox8 = Request.Form["Has_CriminalRecHit_resultpending"];
                                if (checkBox8 == "on")
                                {
                                    strupdate.Has_CriminalRecHit_resultpending = true;
                                }
                                else
                                {
                                    strupdate.Has_CriminalRecHit_resultpending = mainModel.otherdetails.Has_CriminalRecHit_resultpending;
                                }
                                strupdate.has_ucc_fillings = mainModel.otherdetails.has_ucc_fillings;
                                var checkBox10 = Request.Form["has_ucc_fillings1"];
                                if (checkBox10 == "on")
                                {
                                    strupdate.has_ucc_fillings1 = true;
                                }
                                else
                                {
                                    strupdate.has_ucc_fillings1 = mainModel.otherdetails.has_ucc_fillings1;
                                }
                                _context.Entry(strupdate).State = EntityState.Modified;
                                _context.SaveChanges();
                            }
                            catch
                            {
                                Otherdatails strupdate = new Otherdatails();
                                strupdate.record_Id = family_record_id;
                                strupdate.Has_Property_Records = mainModel.otherdetails.Has_Property_Records;
                                strupdate.Press_Media = mainModel.otherdetails.Press_Media;                                
                                strupdate.RegulatoryFlag = mainModel.otherdetails.RegulatoryFlag;
                                strupdate.Has_Reg_FINRA = mainModel.otherdetails.Has_Reg_FINRA;
                                strupdate.Has_Reg_UK_FCA = mainModel.otherdetails.Has_Reg_UK_FCA;
                                strupdate.Has_Reg_US_NFA = mainModel.otherdetails.Has_Reg_US_NFA;
                                strupdate.Has_Reg_US_SEC = mainModel.otherdetails.Has_Reg_US_SEC;
                                strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;                                
                                strupdate.ICIJ_Hits = mainModel.otherdetails.ICIJ_Hits;
                                strupdate.Global_Sec_Family_Hits = mainModel.otherdetails.Global_Sec_Family_Hits;
                                strupdate.Has_Bureau_PrisonHit = mainModel.otherdetails.Has_Bureau_PrisonHit;
                                strupdate.PEP_Hits = mainModel.otherdetails.PEP_Hits;
                                strupdate.undisclosedBA = mainModel.otherdetails.undisclosedBA;
                                strupdate.worldcheck_discloseBA = mainModel.otherdetails.worldcheck_discloseBA;
                                strupdate.worldcheck_undiscloseBA = mainModel.otherdetails.worldcheck_undiscloseBA;
                                strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                                strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                                strupdate.Has_Companion_Report = mainModel.otherdetails.Has_Companion_Report;
                                strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
                                strupdate.Media_Based_Hits = mainModel.otherdetails.Media_Based_Hits;
                                strupdate.Crim_Clearance_Certifi = mainModel.otherdetails.Crim_Clearance_Certifi;
                                strupdate.CurrentResidentialProperty = mainModel.otherdetails.CurrentResidentialProperty;
                                strupdate.OtherPropertyOwnershipinfo = mainModel.otherdetails.OtherPropertyOwnershipinfo;
                                strupdate.OtherCurrentResidentialProperty = mainModel.otherdetails.OtherCurrentResidentialProperty;
                                strupdate.PrevPropertyOwnershipRes = mainModel.otherdetails.PrevPropertyOwnershipRes;
                                strupdate.HasBankruptcyRecHits = mainModel.otherdetails.HasBankruptcyRecHits;
                                strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;
                                strupdate.Has_CriminalRecHit = mainModel.otherdetails.Has_CriminalRecHit;
                                strupdate.Has_Bureau_PrisonHit = mainModel.otherdetails.Has_Bureau_PrisonHit;
                                strupdate.Has_Sex_Offender_RegHit = mainModel.otherdetails.Has_Sex_Offender_RegHit;
                                strupdate.Has_Civil_Records = mainModel.otherdetails.Has_Civil_Records;
                                strupdate.Has_Name_Only = mainModel.otherdetails.Has_Name_Only;
                                strupdate.Has_US_Tax_Court_Hit = mainModel.otherdetails.Has_US_Tax_Court_Hit;
                                strupdate.Has_Tax_Liens = mainModel.otherdetails.Has_Tax_Liens;
                                strupdate.Has_Name_Only_Tax_Lien = mainModel.otherdetails.Has_Name_Only_Tax_Lien;
                                strupdate.Has_Driving_Hits = mainModel.otherdetails.Has_Driving_Hits;
                                var checkBox3 = Request.Form["Has_Name_only_driving_incidents"];
                                if (checkBox3 == "on")
                                {
                                    strupdate.Has_Name_only_driving_incidents = true;
                                }
                                else
                                {
                                    strupdate.Has_Name_only_driving_incidents = mainModel.otherdetails.Has_Name_only_driving_incidents;
                                }
                                strupdate.Was_credited_obtained = mainModel.otherdetails.Was_credited_obtained;                                                                
                                
                                strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                                                        
                                var checkBox4 = Request.Form["HasBankruptcyRecHits1"];
                                if (checkBox4 == "on")
                                {
                                    strupdate.HasBankruptcyRecHits1 = true;
                                }
                                else
                                {
                                    strupdate.HasBankruptcyRecHits1 = mainModel.otherdetails.HasBankruptcyRecHits1;
                                }
                                var checkBox5 = Request.Form["HasBankruptcyRecHits_resultpending"];
                                if (checkBox5 == "on")
                                {
                                    strupdate.HasBankruptcyRecHits_resultpending = true;
                                }
                                else
                                {
                                    strupdate.HasBankruptcyRecHits_resultpending = mainModel.otherdetails.HasBankruptcyRecHits_resultpending;
                                }
                                var checkBox6 = Request.Form["has_civil_resultpending"];
                                if (checkBox6 == "on")
                                {
                                    strupdate.has_civil_resultpending = true;
                                }
                                else
                                {
                                    strupdate.has_civil_resultpending = mainModel.otherdetails.has_civil_resultpending;
                                }
                                var checkBox7 = Request.Form["Has_CriminalRecHit1"];
                                if (checkBox7 == "on")
                                {
                                    strupdate.Has_CriminalRecHit1 = true;
                                }
                                else
                                {
                                    strupdate.Has_CriminalRecHit1 = mainModel.otherdetails.Has_CriminalRecHit1;
                                }
                                var checkBox8 = Request.Form["Has_CriminalRecHit_resultpending"];
                                if (checkBox8 == "on")
                                {
                                    strupdate.Has_CriminalRecHit_resultpending = true;
                                }
                                else
                                {
                                    strupdate.Has_CriminalRecHit_resultpending = mainModel.otherdetails.Has_CriminalRecHit_resultpending;
                                }

                                strupdate.has_ucc_fillings = mainModel.otherdetails.has_ucc_fillings;
                                var checkBox10 = Request.Form["has_ucc_fillings1"];
                                if (checkBox10 == "on")
                                {
                                    strupdate.has_ucc_fillings1 = true;
                                }
                                else
                                {
                                    strupdate.has_ucc_fillings1 = mainModel.otherdetails.has_ucc_fillings1;
                                }
                                _context.othersModel.Add(strupdate);
                                _context.SaveChanges();
                            }
                            TempData["OD"] = "Done";
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SaveEMPData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            try
                            {
                                List<EmployerModel> employerModel1 = _context.DbEmployer
                                          .Where(a => a.record_Id == family_record_id)
                                          .ToList();
                                if (employerModel1 == null || employerModel1.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.employerList.Count; i++)
                                    {
                                        EmployerModel employerModel = new EmployerModel();
                                        employerModel.record_Id = family_record_id;
                                        employerModel.Emp_Employer = mainModel.diligence.employerList[i].Emp_Employer;
                                        employerModel.Emp_Location = mainModel.diligence.employerList[i].Emp_Location;
                                        employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                                        employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                                        employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                                        try
                                        {
                                            if (mainModel.diligence.employerList[i].Emp_StartDateDay.ToString().Equals(""))
                                            {
                                                employerModel.Emp_StartDateDay = "Day";
                                            }
                                            else
                                            {
                                                employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel.Emp_StartDateDay = "Day";
                                        }
                                        employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                                        try
                                        {
                                            if (mainModel.diligence.employerList[i].Emp_StartDateYear.ToString().Equals(""))
                                            {
                                                employerModel.Emp_StartDateYear = "Year";
                                            }
                                            else
                                            {
                                                employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel.Emp_StartDateYear = "Year";
                                        }
                                        try
                                        {
                                            if (mainModel.diligence.employerList[i].Emp_EndDateDay.ToString().Equals(""))
                                            {
                                                employerModel.Emp_EndDateDay = "Day";
                                            }
                                            else
                                            {
                                                employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel.Emp_EndDateDay = "Day";
                                        }
                                        employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                                        try
                                        {
                                            if (mainModel.diligence.employerList[i].Emp_EndDateYear.ToString().Equals(""))
                                            {
                                                employerModel.Emp_EndDateYear = "Year";
                                            }
                                            else
                                            {
                                                employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel.Emp_EndDateYear = "Year";
                                        }
                                        employerModel.CreatedBy = Environment.UserName;
                                        _context.DbEmployer.Add(employerModel);
                                        _context.SaveChanges();
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i < employerModel1.Count; i++)
                                    {
                                        employerModel1[i].record_Id = family_record_id;
                                        employerModel1[i].Emp_Employer = mainModel.EmployerModel[i].Emp_Employer;
                                        employerModel1[i].Emp_Location = mainModel.EmployerModel[i].Emp_Location;
                                        employerModel1[i].Emp_Position = mainModel.EmployerModel[i].Emp_Position;
                                        employerModel1[i].Emp_Confirmed = mainModel.EmployerModel[i].Emp_Confirmed;
                                        employerModel1[i].Emp_Status = mainModel.EmployerModel[i].Emp_Status;
                                        employerModel1[i].Emp_StartDateDay = mainModel.EmployerModel[i].Emp_StartDateDay;
                                        employerModel1[i].Emp_StartDateMonth = mainModel.EmployerModel[i].Emp_StartDateMonth;
                                        employerModel1[i].Emp_StartDateYear = mainModel.EmployerModel[i].Emp_StartDateYear;
                                        employerModel1[i].Emp_EndDateDay = mainModel.EmployerModel[i].Emp_EndDateDay;
                                        employerModel1[i].Emp_EndDateMonth = mainModel.EmployerModel[i].Emp_EndDateMonth;
                                        employerModel1[i].Emp_EndDateYear = mainModel.EmployerModel[i].Emp_EndDateYear;
                                        employerModel1[i].CreatedBy = Environment.UserName;
                                        _context.DbEmployer.Update(employerModel1[i]);
                                        _context.SaveChanges();
                                    }
                                    try
                                    {
                                        for (int i = 0; i < mainModel.diligence.employerList.Count; i++)
                                        {
                                            if (mainModel.diligence.employerList[i].Emp_Status.ToString().Equals("Current")) { }
                                            else
                                            {
                                                if (mainModel.diligence.employerList[i].Emp_Status.ToString().Equals("Previous") && mainModel.diligence.employerList[i].Emp_Employer.ToString().Equals("") && mainModel.diligence.employerList[i].Emp_Position.ToString().Equals("")) { }
                                                else
                                                {
                                                    EmployerModel employerModel = new EmployerModel();
                                                    employerModel.record_Id = family_record_id;
                                                    employerModel.Emp_Employer = mainModel.diligence.employerList[i].Emp_Employer;
                                                    employerModel.Emp_Location = mainModel.diligence.employerList[i].Emp_Location;
                                                    employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                                                    employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                                                    employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                                                    try
                                                    {
                                                        if (mainModel.diligence.employerList[i].Emp_StartDateDay.ToString().Equals(""))
                                                        {
                                                            employerModel.Emp_StartDateDay = "Day";
                                                        }
                                                        else
                                                        {
                                                            employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        employerModel.Emp_StartDateDay = "Day";
                                                    }
                                                    employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                                                    try
                                                    {
                                                        if (mainModel.diligence.employerList[i].Emp_StartDateYear.ToString().Equals(""))
                                                        {
                                                            employerModel.Emp_StartDateYear = "Year";
                                                        }
                                                        else
                                                        {
                                                            employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        employerModel.Emp_StartDateYear = "Year";
                                                    }
                                                    try
                                                    {
                                                        if (mainModel.diligence.employerList[i].Emp_EndDateDay.ToString().Equals(""))
                                                        {
                                                            employerModel.Emp_EndDateDay = "Day";
                                                        }
                                                        else
                                                        {
                                                            employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        employerModel.Emp_EndDateDay = "Day";
                                                    }
                                                    employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                                                    try
                                                    {
                                                        if (mainModel.diligence.employerList[i].Emp_EndDateYear.ToString().Equals(""))
                                                        {
                                                            employerModel.Emp_EndDateYear = "Year";
                                                        }
                                                        else
                                                        {
                                                            employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
                                                        }
                                                    }
                                                    catch
                                                    {
                                                        employerModel.Emp_EndDateYear = "Year";
                                                    }
                                                    employerModel.CreatedBy = Environment.UserName;
                                                    _context.DbEmployer.Add(employerModel);
                                                    _context.SaveChanges();
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.employerList.Count; i++)
                                {
                                    EmployerModel employerModel = new EmployerModel();
                                    employerModel.record_Id = family_record_id;
                                    employerModel.Emp_Employer = mainModel.diligence.employerList[i].Emp_Employer;
                                    employerModel.Emp_Location = mainModel.diligence.employerList[i].Emp_Location;
                                    employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                                    employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                                    employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                                    try
                                    {
                                        if (mainModel.diligence.employerList[i].Emp_StartDateDay.ToString().Equals(""))
                                        {
                                            employerModel.Emp_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                                        }
                                    }
                                    catch
                                    {
                                        employerModel.Emp_StartDateDay = "Day";
                                    }
                                    employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                                    try
                                    {
                                        if (mainModel.diligence.employerList[i].Emp_StartDateYear.ToString().Equals(""))
                                        {
                                            employerModel.Emp_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                                        }
                                    }
                                    catch
                                    {
                                        employerModel.Emp_StartDateYear = "Year";
                                    }
                                    try
                                    {
                                        if (mainModel.diligence.employerList[i].Emp_EndDateDay.ToString().Equals(""))
                                        {
                                            employerModel.Emp_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                                        }
                                    }
                                    catch
                                    {
                                        employerModel.Emp_EndDateDay = "Day";
                                    }
                                    employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                                    try
                                    {
                                        if (mainModel.diligence.employerList[i].Emp_EndDateYear.ToString().Equals(""))
                                        {
                                            employerModel.Emp_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
                                        }
                                    }
                                    catch
                                    {
                                        employerModel.Emp_EndDateYear = "Year";
                                    }
                                    employerModel.CreatedBy = Environment.UserName;
                                    _context.DbEmployer.Add(employerModel);
                                    _context.SaveChanges();
                                }
                            }
                            mainModel.EmployerModel = _context.DbEmployer
                    .Where(u => u.record_Id == family_record_id)
                    .ToList();
                            TempData["EMP"] = "Done";
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SaveEduData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            try
                            {
                                List<EducationModel> educationModel1 = _context.DbEducation
                                   .Where(a => a.record_Id == family_record_id)
                                   .ToList();
                                if (educationModel1 == null || educationModel1.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                    {
                                        EducationModel educationModel = new EducationModel();
                                        educationModel.record_Id = family_record_id;
                                        educationModel.Edu_Degree = mainModel.diligence.educationList[i].Edu_Degree;
                                        educationModel.Edu_Location = mainModel.diligence.educationList[i].Edu_Location;
                                        educationModel.Edu_School = mainModel.diligence.educationList[i].Edu_School;
                                        educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                        educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                        educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                        educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                        educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                        educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay;
                                        educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                        educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                        educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                        educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                        educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                        educationModel.CreatedBy = Environment.UserName;
                                        _context.DbEducation.Add(educationModel);
                                        _context.SaveChanges();
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i < educationModel1.Count; i++)
                                    {
                                        educationModel1[i].record_Id = family_record_id;
                                        educationModel1[i].Edu_Degree = mainModel.educationModels[i].Edu_Degree;
                                        educationModel1[i].Edu_Location = mainModel.educationModels[i].Edu_Location;
                                        educationModel1[i].Edu_School = mainModel.educationModels[i].Edu_School;
                                        educationModel1[i].Edu_History = mainModel.educationModels[i].Edu_History;
                                        educationModel1[i].Edu_Confirmed = mainModel.educationModels[i].Edu_Confirmed;
                                        educationModel1[i].Edu_GradDateMonth = mainModel.educationModels[i].Edu_GradDateMonth;
                                        educationModel1[i].Edu_GradDateDay = mainModel.educationModels[i].Edu_GradDateDay;
                                        educationModel1[i].Edu_GradDateYear = mainModel.educationModels[i].Edu_GradDateYear;
                                        educationModel1[i].Edu_StartDateDay = mainModel.educationModels[i].Edu_StartDateDay;
                                        educationModel1[i].Edu_StartDateMonth = mainModel.educationModels[i].Edu_StartDateMonth;
                                        educationModel1[i].Edu_StartDateYear = mainModel.educationModels[i].Edu_StartDateYear;
                                        educationModel1[i].Edu_EndDateDay = mainModel.educationModels[i].Edu_EndDateDay;
                                        educationModel1[i].Edu_EndDateMonth = mainModel.educationModels[i].Edu_EndDateMonth;
                                        educationModel1[i].Edu_EndDateYear = mainModel.educationModels[i].Edu_EndDateYear;
                                        educationModel1[i].CreatedBy = Environment.UserName;
                                        _context.DbEducation.Update(educationModel1[i]);
                                        _context.SaveChanges();
                                    }
                                    try
                                    {
                                        for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                        {
                                            if (mainModel.diligence.educationList[i].Edu_History.ToString().Equals("No")) { }
                                            else
                                            {
                                                EducationModel educationModel = new EducationModel();
                                                educationModel.record_Id = family_record_id;
                                                educationModel.Edu_Degree = mainModel.diligence.educationList[i].Edu_Degree;
                                                educationModel.Edu_Location = mainModel.diligence.educationList[i].Edu_Location;
                                                educationModel.Edu_School = mainModel.diligence.educationList[i].Edu_School;
                                                educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                                educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                                educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                                educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                                educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                                educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay;
                                                educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                                educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                                educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                                educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                                educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                                educationModel.CreatedBy = Environment.UserName;
                                                _context.DbEducation.Add(educationModel);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                {
                                    EducationModel educationModel = new EducationModel();
                                    educationModel.record_Id = family_record_id;
                                    educationModel.Edu_Degree = mainModel.diligence.educationList[i].Edu_Degree;
                                    educationModel.Edu_Location = mainModel.diligence.educationList[i].Edu_Location;
                                    educationModel.Edu_School = mainModel.diligence.educationList[i].Edu_School;
                                    educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                    educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                    educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                    educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                    educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                    educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay;
                                    educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                    educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                    educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                    educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                    educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                    educationModel.CreatedBy = Environment.UserName;
                                    _context.DbEducation.Add(educationModel);
                                    _context.SaveChanges();
                                }
                            }
                            mainModel.educationModels = _context.DbEducation
                              .Where(u => u.record_Id == family_record_id)
                              .ToList();
                            TempData["EDU"] = "Done";
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SavePLData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            try
                            {
                                List<PLLicenseModel> pLLicenseModel = _context.DbPLLicense
                                   .Where(a => a.record_Id == family_record_id)
                                   .ToList();
                                if (pLLicenseModel == null || pLLicenseModel.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.plLicenseList.Count; i++)
                                    {
                                        PLLicenseModel pLLicenseModel1 = new PLLicenseModel();
                                        pLLicenseModel1.General_PL_License = mainModel.diligence.plLicenseList[i].General_PL_License;
                                        pLLicenseModel1.PL_Confirmed = mainModel.diligence.plLicenseList[i].PL_Confirmed;
                                        pLLicenseModel1.PL_License_Type = mainModel.diligence.plLicenseList[i].PL_License_Type;
                                        pLLicenseModel1.PL_Location = mainModel.diligence.plLicenseList[i].PL_Location;
                                        pLLicenseModel1.PL_Number = mainModel.diligence.plLicenseList[i].PL_Number;
                                        pLLicenseModel1.PL_Organization = mainModel.diligence.plLicenseList[i].PL_Organization;
                                        pLLicenseModel1.record_Id = family_record_id;
                                        pLLicenseModel1.CreatedBy = Environment.UserName;
                                        pLLicenseModel1.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                        pLLicenseModel1.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                        pLLicenseModel1.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                        pLLicenseModel1.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                        pLLicenseModel1.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                        pLLicenseModel1.PL_EndDateMonth = mainModel.diligence.plLicenseList[i].PL_EndDateMonth;
                                        _context.DbPLLicense.Add(pLLicenseModel1);
                                        _context.SaveChanges();
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i < pLLicenseModel.Count; i++)
                                    {
                                        pLLicenseModel[i].General_PL_License = mainModel.pllicenseModels[i].General_PL_License;
                                        pLLicenseModel[i].PL_Confirmed = mainModel.pllicenseModels[i].PL_Confirmed;
                                        pLLicenseModel[i].PL_License_Type = mainModel.pllicenseModels[i].PL_License_Type;
                                        pLLicenseModel[i].PL_Location = mainModel.pllicenseModels[i].PL_Location;
                                        pLLicenseModel[i].PL_Number = mainModel.pllicenseModels[i].PL_Number;
                                        pLLicenseModel[i].PL_Organization = mainModel.pllicenseModels[i].PL_Organization;
                                        pLLicenseModel[i].CreatedBy = Environment.UserName;
                                        pLLicenseModel[i].PL_StartDateDay = mainModel.pllicenseModels[i].PL_StartDateDay;
                                        pLLicenseModel[i].PL_StartDateMonth = mainModel.pllicenseModels[i].PL_StartDateMonth;
                                        pLLicenseModel[i].PL_StartDateYear = mainModel.pllicenseModels[i].PL_StartDateYear;
                                        pLLicenseModel[i].PL_EndDateDay = mainModel.pllicenseModels[i].PL_EndDateDay;
                                        pLLicenseModel[i].PL_EndDateYear = mainModel.pllicenseModels[i].PL_EndDateYear;
                                        pLLicenseModel[i].PL_EndDateMonth = mainModel.pllicenseModels[i].PL_EndDateMonth;
                                        pLLicenseModel[i].record_Id = family_record_id;
                                        _context.DbPLLicense.Update(pLLicenseModel[i]);
                                        _context.SaveChanges();
                                    }
                                    try
                                    {
                                        for (int i = 0; i < mainModel.diligence.plLicenseList.Count; i++)
                                        {
                                            if (mainModel.diligence.plLicenseList[i].General_PL_License.ToString().Equals("No")) { }
                                            else
                                            {
                                                PLLicenseModel pLLicenseModel1 = new PLLicenseModel();
                                                pLLicenseModel1.General_PL_License = mainModel.diligence.plLicenseList[i].General_PL_License;
                                                pLLicenseModel1.PL_Confirmed = mainModel.diligence.plLicenseList[i].PL_Confirmed;
                                                pLLicenseModel1.PL_License_Type = mainModel.diligence.plLicenseList[i].PL_License_Type;
                                                pLLicenseModel1.PL_Location = mainModel.diligence.plLicenseList[i].PL_Location;
                                                pLLicenseModel1.PL_Number = mainModel.diligence.plLicenseList[i].PL_Number;
                                                pLLicenseModel1.PL_Organization = mainModel.diligence.plLicenseList[i].PL_Organization;
                                                pLLicenseModel1.record_Id = family_record_id;
                                                pLLicenseModel1.CreatedBy = Environment.UserName;
                                                pLLicenseModel1.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                                pLLicenseModel1.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                                pLLicenseModel1.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                                pLLicenseModel1.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                                pLLicenseModel1.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                                pLLicenseModel1.PL_EndDateMonth = mainModel.diligence.plLicenseList[i].PL_EndDateMonth;
                                                _context.DbPLLicense.Add(pLLicenseModel1);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                    catch { }
                                }
                                mainModel.pllicenseModels = _context.DbPLLicense
                                   .Where(u => u.record_Id == family_record_id)
                                   .ToList();
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.plLicenseList.Count; i++)
                                {
                                    PLLicenseModel pLLicenseModel = new PLLicenseModel();
                                    pLLicenseModel.General_PL_License = mainModel.diligence.plLicenseList[i].General_PL_License;
                                    pLLicenseModel.PL_Confirmed = mainModel.diligence.plLicenseList[i].PL_Confirmed;
                                    pLLicenseModel.PL_License_Type = mainModel.diligence.plLicenseList[i].PL_License_Type;
                                    pLLicenseModel.PL_Location = mainModel.diligence.plLicenseList[i].PL_Location;
                                    pLLicenseModel.PL_Number = mainModel.diligence.plLicenseList[i].PL_Number;
                                    pLLicenseModel.PL_Organization = mainModel.diligence.plLicenseList[i].PL_Organization;
                                    pLLicenseModel.record_Id = family_record_id;
                                    pLLicenseModel.CreatedBy = Environment.UserName;
                                    pLLicenseModel.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                    pLLicenseModel.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                    pLLicenseModel.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                    pLLicenseModel.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                    pLLicenseModel.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                    pLLicenseModel.PL_EndDateMonth = mainModel.diligence.plLicenseList[i].PL_EndDateMonth;
                                    _context.DbPLLicense.Add(pLLicenseModel);
                                    _context.SaveChanges();
                                }
                            }
                            mainModel.pllicenseModels = _context.DbPLLicense
                                  .Where(u => u.record_Id == family_record_id)
                                  .ToList();
                            TempData["PL"] = "Done";
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SaveCSData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            try
                            {
                                List<Additional_statesModel> additionaL = _context.Dbadditionalstates
                                    .Where(a => a.record_Id == recordid)
                                    .ToList();
                                if (additionaL == null || additionaL.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.additional_States.Count; i++)
                                    {
                                        if (mainModel.diligence.additional_States[i].additionalstate.ToString().Equals("Select State")) { }
                                        else
                                        {
                                            Additional_statesModel additional = new Additional_statesModel();
                                            additional.record_Id = recordid;
                                            additional.additionalstate = mainModel.diligence.additional_States[i].additionalstate;
                                            _context.Dbadditionalstates.Add(additional);
                                            _context.SaveChanges();
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i < additionaL.Count; i++)
                                    {
                                        additionaL[i].additionalstate = mainModel.additional_States[i].additionalstate;
                                        additionaL[i].record_Id = recordid;
                                        _context.Dbadditionalstates.Update(additionaL[i]);
                                        _context.SaveChanges();
                                    }
                                    try
                                    {
                                        for (int i = 0; i < mainModel.diligence.additional_States.Count; i++)
                                        {
                                            if (mainModel.diligence.additional_States[i].additionalstate.ToString().Equals("Select State")) { }
                                            else
                                            {
                                                Additional_statesModel additional = new Additional_statesModel();
                                                additional.record_Id = recordid;
                                                additional.additionalstate = mainModel.diligence.additional_States[i].additionalstate;
                                                _context.Dbadditionalstates.Add(additional);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.additional_States.Count; i++)
                                {
                                    if (mainModel.diligence.additional_States[i].additionalstate.ToString().Equals("Select State")) { }
                                    else
                                    {
                                        Additional_statesModel additional = new Additional_statesModel();
                                        additional.record_Id = recordid;
                                        additional.additionalstate = mainModel.diligence.additional_States[i].additionalstate;
                                        _context.Dbadditionalstates.Add(additional);
                                        _context.SaveChanges();
                                    }
                                }
                            }
                            mainModel.additional_States = _context.Dbadditionalstates
                       .Where(u => u.record_Id == recordid)
                       .ToList();
                            TempData["CS"] = "Done";
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SaveCLSData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            ClientspecificModel clientspecific = new ClientspecificModel();
                            try
                            {
                                clientspecific = _context.DbclientspecificModels
                           .Where(u => u.record_Id == family_record_id)
                                          .FirstOrDefault();
                                if (clientspecific == null)
                                {
                                    clientspecific = new ClientspecificModel();
                                }
                            }
                            catch
                            {
                                clientspecific = new ClientspecificModel();
                            }
                            clientspecific.record_Id = family_record_id;
                            clientspecific.clientname = mainModel.clientspecific.clientname;
                            clientspecific.Military_Service = mainModel.clientspecific.Military_Service;
                            clientspecific.source_of_wealth = mainModel.clientspecific.source_of_wealth;
                            clientspecific.discreet_reputational_inquiries = mainModel.clientspecific.discreet_reputational_inquiries;
                            clientspecific.Character_Reference = mainModel.clientspecific.Character_Reference;
                            try
                            {
                                _context.Entry(clientspecific).State = EntityState.Modified;
                                _context.SaveChanges();
                            }
                            catch
                            {
                                _context.DbclientspecificModels.Add(clientspecific);
                                _context.SaveChanges();
                            }
                            if (mainModel.clientspecific.clientname.ToString().Equals("Dominica"))
                            {
                                try
                                {
                                    List<ReferencetableModel> referencetable1 = _context.DbReferencetableModel
                                  .Where(a => a.record_Id == family_record_id)
                                  .ToList();
                                    if (referencetable1 == null || referencetable1.Count == 0)
                                    {
                                        for (int i = 0; i < mainModel.diligence.reftable.Count; i++)
                                        {
                                            if (mainModel.diligence.reftable[i].ref_position == "" && mainModel.diligence.reftable[i].ref_full_name == "" && mainModel.diligence.reftable[i].ref_employer == "" && mainModel.diligence.reftable[i].ref_location == "") { }
                                            else
                                            {
                                                ReferencetableModel referencetable = new ReferencetableModel();
                                                referencetable.ref_position = mainModel.diligence.reftable[i].ref_position;
                                                referencetable.ref_full_name = mainModel.diligence.reftable[i].ref_full_name;
                                                referencetable.ref_employer = mainModel.diligence.reftable[i].ref_employer;
                                                referencetable.ref_location = mainModel.diligence.reftable[i].ref_location;
                                                referencetable.record_Id = family_record_id;
                                                _context.DbReferencetableModel.Add(referencetable);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int i = 0; i < referencetable1.Count; i++)
                                        {
                                            referencetable1[i].ref_position = mainModel.referencetableModels[i].ref_position;
                                            referencetable1[i].ref_full_name = mainModel.referencetableModels[i].ref_full_name;
                                            referencetable1[i].ref_employer = mainModel.referencetableModels[i].ref_employer;
                                            referencetable1[i].ref_location = mainModel.referencetableModels[i].ref_location;
                                            referencetable1[i].record_Id = family_record_id;
                                            _context.DbReferencetableModel.Update(referencetable1[i]);
                                            _context.SaveChanges();
                                        }
                                        try
                                        {
                                            for (int i = 0; i < mainModel.diligence.reftable.Count; i++)
                                            {
                                                if (mainModel.diligence.reftable[i].ref_position == "" && mainModel.diligence.reftable[i].ref_full_name == "" && mainModel.diligence.reftable[i].ref_employer == "" && mainModel.diligence.reftable[i].ref_location == "") { }
                                                else
                                                {
                                                    ReferencetableModel referencetable = new ReferencetableModel();
                                                    referencetable.ref_position = mainModel.diligence.reftable[i].ref_position;
                                                    referencetable.ref_full_name = mainModel.diligence.reftable[i].ref_full_name;
                                                    referencetable.ref_employer = mainModel.diligence.reftable[i].ref_employer;
                                                    referencetable.ref_location = mainModel.diligence.reftable[i].ref_location;
                                                    referencetable.record_Id = family_record_id;
                                                    _context.DbReferencetableModel.Add(referencetable);
                                                    _context.SaveChanges();
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                catch
                                {
                                    for (int i = 0; i < mainModel.diligence.reftable.Count; i++)
                                    {
                                        if (mainModel.diligence.reftable[i].ref_position == "" && mainModel.diligence.reftable[i].ref_full_name == "" && mainModel.diligence.reftable[i].ref_employer == "" && mainModel.diligence.reftable[i].ref_location == "") { }
                                        else
                                        {
                                            ReferencetableModel referencetable = new ReferencetableModel();
                                            referencetable.ref_position = mainModel.diligence.reftable[i].ref_position;
                                            referencetable.ref_full_name = mainModel.diligence.reftable[i].ref_full_name;
                                            referencetable.ref_employer = mainModel.diligence.reftable[i].ref_employer;
                                            referencetable.ref_location = mainModel.diligence.reftable[i].ref_location;
                                            referencetable.record_Id = family_record_id;
                                            _context.DbReferencetableModel.Add(referencetable);
                                            _context.SaveChanges();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SaveCRData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            try
                            {
                                List<CurrentResidenceModel> currentResidence1 = _context.DbcurrentResidenceModels
                                 .Where(a => a.record_Id == family_record_id)
                                 .ToList();
                                if (currentResidence1 == null || currentResidence1.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.currentResidences.Count; i++)
                                    {
                                        CurrentResidenceModel currentResidence = new CurrentResidenceModel();
                                        currentResidence.record_Id = family_record_id;
                                        currentResidence.CurrentStreet = mainModel.diligence.currentResidences[i].CurrentStreet;
                                        currentResidence.CurrentZipcode = mainModel.diligence.currentResidences[i].CurrentZipcode;
                                        currentResidence.CurrentCountry = mainModel.diligence.currentResidences[i].CurrentCountry;
                                        currentResidence.CurrentCity = mainModel.diligence.currentResidences[i].CurrentCity;
                                        _context.DbcurrentResidenceModels.Add(currentResidence);
                                        _context.SaveChanges();
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i < currentResidence1.Count; i++)
                                    {
                                        currentResidence1[i].record_Id = family_record_id;
                                        currentResidence1[i].CurrentStreet = mainModel.currentResidenceModels[i].CurrentStreet;
                                        currentResidence1[i].CurrentZipcode = mainModel.currentResidenceModels[i].CurrentZipcode;
                                        currentResidence1[i].CurrentCountry = mainModel.currentResidenceModels[i].CurrentCountry;
                                        currentResidence1[i].CurrentCity = mainModel.currentResidenceModels[i].CurrentCity;
                                        _context.DbcurrentResidenceModels.Update(currentResidence1[i]);
                                        _context.SaveChanges();
                                    }
                                    try
                                    {
                                        for (int i = 0; i < mainModel.diligence.currentResidences.Count; i++)
                                        {
                                            if (mainModel.diligence.currentResidences[i].CurrentStreet == "" && mainModel.diligence.currentResidences[i].CurrentCity == "" && mainModel.diligence.currentResidences[i].CurrentCountry == "Select State" && mainModel.diligence.currentResidences[i].CurrentZipcode == "") { }
                                            else
                                            {
                                                CurrentResidenceModel currentResidence = new CurrentResidenceModel();
                                                currentResidence.record_Id = family_record_id;
                                                currentResidence.CurrentStreet = mainModel.diligence.currentResidences[i].CurrentStreet;
                                                currentResidence.CurrentZipcode = mainModel.diligence.currentResidences[i].CurrentZipcode;
                                                currentResidence.CurrentCountry = mainModel.diligence.currentResidences[i].CurrentCountry;
                                                currentResidence.CurrentCity = mainModel.diligence.currentResidences[i].CurrentCity;
                                                _context.DbcurrentResidenceModels.Add(currentResidence);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.currentResidences.Count; i++)
                                {
                                    CurrentResidenceModel currentResidence = new CurrentResidenceModel();
                                    currentResidence.record_Id = family_record_id;
                                    currentResidence.CurrentStreet = mainModel.diligence.currentResidences[i].CurrentStreet;
                                    currentResidence.CurrentZipcode = mainModel.diligence.currentResidences[i].CurrentZipcode;
                                    currentResidence.CurrentCountry = mainModel.diligence.currentResidences[i].CurrentCountry;
                                    currentResidence.CurrentCity = mainModel.diligence.currentResidences[i].CurrentCity;
                                    _context.DbcurrentResidenceModels.Add(currentResidence);
                                    _context.SaveChanges();
                                }
                            }
                            mainModel.currentResidenceModels = _context.DbcurrentResidenceModels
                                 .Where(u => u.record_Id == family_record_id)
                                 .ToList();
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SavePRData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            try
                            {
                                List<PreviousResidenceModel> currentResidence1 = _context.DbpreviousResidenceModels
                                .Where(a => a.record_Id == family_record_id)
                                .ToList();
                                if (currentResidence1 == null || currentResidence1.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.previousResidences.Count; i++)
                                    {
                                        PreviousResidenceModel currentResidence = new PreviousResidenceModel();
                                        currentResidence.record_Id = family_record_id;
                                        currentResidence.PreviousStreet = mainModel.diligence.previousResidences[i].PreviousStreet;
                                        currentResidence.PreviousCity = mainModel.diligence.previousResidences[i].PreviousCity;
                                        currentResidence.PreviousCountry = mainModel.diligence.previousResidences[i].PreviousCountry;
                                        currentResidence.PreviousZipcode = mainModel.diligence.previousResidences[i].PreviousZipcode;
                                        _context.DbpreviousResidenceModels.Add(currentResidence);
                                        _context.SaveChanges();
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i < currentResidence1.Count; i++)
                                    {
                                        currentResidence1[i].record_Id = family_record_id;
                                        currentResidence1[i].PreviousStreet = mainModel.PreviousResidenceModels[i].PreviousStreet;
                                        currentResidence1[i].PreviousCity = mainModel.PreviousResidenceModels[i].PreviousCity;
                                        currentResidence1[i].PreviousCountry = mainModel.PreviousResidenceModels[i].PreviousCountry;
                                        currentResidence1[i].PreviousZipcode = mainModel.PreviousResidenceModels[i].PreviousZipcode;
                                    }
                                    try
                                    {
                                        for (int i = 0; i < mainModel.diligence.previousResidences.Count; i++)
                                        {
                                            if (mainModel.diligence.previousResidences[i].PreviousStreet == "" && mainModel.diligence.previousResidences[i].PreviousCity == "" && mainModel.diligence.previousResidences[i].PreviousCountry == "" && mainModel.diligence.previousResidences[i].PreviousZipcode == "") { }
                                            else
                                            {
                                                PreviousResidenceModel currentResidence = new PreviousResidenceModel();
                                                currentResidence.record_Id = family_record_id;
                                                currentResidence.PreviousStreet = mainModel.diligence.previousResidences[i].PreviousStreet;
                                                currentResidence.PreviousCity = mainModel.diligence.previousResidences[i].PreviousCity;
                                                currentResidence.PreviousCountry = mainModel.diligence.previousResidences[i].PreviousCountry;
                                                currentResidence.PreviousZipcode = mainModel.diligence.previousResidences[i].PreviousZipcode;
                                                _context.DbpreviousResidenceModels.Add(currentResidence);
                                                _context.SaveChanges();
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.previousResidences.Count; i++)
                                {
                                    PreviousResidenceModel currentResidence = new PreviousResidenceModel();
                                    currentResidence.record_Id = family_record_id;
                                    currentResidence.PreviousStreet = mainModel.diligence.previousResidences[i].PreviousStreet;
                                    currentResidence.PreviousCity = mainModel.diligence.previousResidences[i].PreviousCity;
                                    currentResidence.PreviousCountry = mainModel.diligence.previousResidences[i].PreviousCountry;
                                    currentResidence.PreviousZipcode = mainModel.diligence.previousResidences[i].PreviousZipcode;
                                    _context.DbpreviousResidenceModels.Add(currentResidence);
                                    _context.SaveChanges();
                                }

                            }
                            mainModel.PreviousResidenceModels = _context.DbpreviousResidenceModels
                                  .Where(u => u.record_Id == family_record_id)
                                  .ToList();
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SaveSummaryData":
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            SummaryResulttableModel resulttableModel = new SummaryResulttableModel();
                            try
                            {
                                resulttableModel = _context.summaryResulttableModels
                                    .Where(a => a.record_Id == family_record_id)
                                    .FirstOrDefault();
                                if (resulttableModel == null)
                                {
                                    resulttableModel = new SummaryResulttableModel();
                                }
                            }
                            catch
                            {
                                resulttableModel = new SummaryResulttableModel();
                            }
                            resulttableModel.record_Id = family_record_id;
                            resulttableModel.name_add_Ver_History = mainModel.summarymodel.name_add_Ver_History;
                            resulttableModel.personal_Identification = mainModel.summarymodel.personal_Identification;
                            resulttableModel.social_securitytrace = mainModel.summarymodel.social_securitytrace;
                            resulttableModel.real_estate_prop = mainModel.summarymodel.real_estate_prop;
                            resulttableModel.secretary_state_director = mainModel.summarymodel.secretary_state_director;
                            resulttableModel.news_media_searches = mainModel.summarymodel.news_media_searches;
                            resulttableModel.department_foreign = mainModel.summarymodel.department_foreign;
                            resulttableModel.european_union = mainModel.summarymodel.european_union;
                            resulttableModel.HM_treasury = mainModel.summarymodel.HM_treasury;
                            resulttableModel.US_bureau = mainModel.summarymodel.US_bureau;
                            resulttableModel.US_department = mainModel.summarymodel.US_department;
                            resulttableModel.US_Directorate = mainModel.summarymodel.US_Directorate;
                            resulttableModel.US_general = mainModel.summarymodel.US_general;
                            resulttableModel.US_office = mainModel.summarymodel.US_office;
                            resulttableModel.UN_consolidated = mainModel.summarymodel.UN_consolidated;
                            resulttableModel.world_bank_list = mainModel.summarymodel.world_bank_list;
                            resulttableModel.city_london_police = mainModel.summarymodel.city_london_police;
                            resulttableModel.constabularies_cheshire = mainModel.summarymodel.constabularies_cheshire;
                            resulttableModel.hampshire_police = mainModel.summarymodel.hampshire_police;
                            resulttableModel.hong_kong_police = mainModel.summarymodel.hong_kong_police;
                            resulttableModel.interpol = mainModel.summarymodel.interpol;
                            resulttableModel.metropolitan_police = mainModel.summarymodel.metropolitan_police;
                            resulttableModel.national_crime = mainModel.summarymodel.national_crime;
                            resulttableModel.north_yorkshire_polic = mainModel.summarymodel.north_yorkshire_polic;
                            resulttableModel.nottinghamshire_police = mainModel.summarymodel.nottinghamshire_police;
                            resulttableModel.surrey_police = mainModel.summarymodel.surrey_police;
                            resulttableModel.thames_valley_police = mainModel.summarymodel.thames_valley_police;
                            resulttableModel.US_Federal = mainModel.summarymodel.US_Federal;
                            resulttableModel.US_secret_service = mainModel.summarymodel.US_secret_service;
                            resulttableModel.warwickshire_police = mainModel.summarymodel.warwickshire_police;
                            resulttableModel.alberta_securities_commission = mainModel.summarymodel.alberta_securities_commission;
                            resulttableModel.asset_recovery_agency = mainModel.summarymodel.asset_recovery_agency;
                            resulttableModel.australian_prudential = mainModel.summarymodel.australian_prudential;
                            resulttableModel.australian_securities = mainModel.summarymodel.australian_securities;
                            resulttableModel.banque_de_CECEI = mainModel.summarymodel.banque_de_CECEI;
                            resulttableModel.banque_de_commission = mainModel.summarymodel.banque_de_commission;
                            resulttableModel.british_virgin_islands = mainModel.summarymodel.british_virgin_islands;
                            resulttableModel.cayman_islands_monetary = mainModel.summarymodel.cayman_islands_monetary;
                            resulttableModel.commission_de_surveillance = mainModel.summarymodel.commission_de_surveillance;
                            resulttableModel.commodity_futures = mainModel.summarymodel.commodity_futures;
                            resulttableModel.council_financial_activities = mainModel.summarymodel.council_financial_activities;
                            resulttableModel.departamento_de_investigacoes = mainModel.summarymodel.departamento_de_investigacoes;
                            resulttableModel.department_labour_inspection = mainModel.summarymodel.department_labour_inspection;
                            resulttableModel.federal_deposit = mainModel.summarymodel.federal_deposit;
                            resulttableModel.financial_action_task = mainModel.summarymodel.financial_action_task;
                            resulttableModel.federal_reserve = mainModel.summarymodel.federal_reserve;
                            resulttableModel.financial_crimes = mainModel.summarymodel.financial_crimes;
                            resulttableModel.financial_industry = mainModel.summarymodel.financial_industry;
                            resulttableModel.financial_regulator_ireland = mainModel.summarymodel.financial_regulator_ireland;
                            resulttableModel.hongkong_monetary_authority = mainModel.summarymodel.hongkong_monetary_authority;
                            resulttableModel.hongkong_securities_futures = mainModel.summarymodel.hongkong_securities_futures;
                            resulttableModel.investment_association_Canada = mainModel.summarymodel.banque_de_commission;
                            resulttableModel.investment_management_regulatory = mainModel.summarymodel.investment_management_regulatory;
                            resulttableModel.isle_financial_supervision = mainModel.summarymodel.isle_financial_supervision;
                            resulttableModel.jersey_financial_commission = mainModel.summarymodel.jersey_financial_commission;
                            resulttableModel.lloyd_insurance_arimbolaet = mainModel.summarymodel.lloyd_insurance_arimbolaet;
                            resulttableModel.monetary_authority_singapore = mainModel.summarymodel.monetary_authority_singapore;
                            resulttableModel.new_york_stock = mainModel.summarymodel.new_york_stock;
                            resulttableModel.national_credit = mainModel.summarymodel.national_credit;
                            resulttableModel.Office_of_comptroller = mainModel.summarymodel.Office_of_comptroller;
                            resulttableModel.Office_of_superintendent = mainModel.summarymodel.Office_of_superintendent;
                            resulttableModel.resolution_trust = mainModel.summarymodel.resolution_trust;
                            resulttableModel.securities_exchange = mainModel.summarymodel.securities_exchange;
                            resulttableModel.securities_exchange_commission = mainModel.summarymodel.securities_exchange_commission;
                            resulttableModel.securities_futuresauthority = mainModel.summarymodel.securities_futuresauthority;
                            resulttableModel.swedish_financial_supervisory = mainModel.summarymodel.swedish_financial_supervisory;
                            resulttableModel.swiss_federal_banking = mainModel.summarymodel.swiss_federal_banking;
                            resulttableModel.U_K_companies_disqualified = mainModel.summarymodel.U_K_companies_disqualified;
                            resulttableModel.U_K_financial_conduct_authority = mainModel.summarymodel.U_K_financial_conduct_authority;
                            resulttableModel.US_court = mainModel.summarymodel.US_court;
                            resulttableModel.US_department_justice = mainModel.summarymodel.US_department_justice;
                            resulttableModel.US_federal_trade = mainModel.summarymodel.US_federal_trade;
                            resulttableModel.US_national = mainModel.summarymodel.US_national;
                            resulttableModel.US_office_thrifts = mainModel.summarymodel.US_office_thrifts;
                            resulttableModel.central_intelligence = mainModel.summarymodel.central_intelligence;
                            try
                            {
                                if (resulttableModel == null)
                                {
                                    _context.summaryResulttableModels.Add(resulttableModel);
                                    _context.SaveChanges();
                                }
                                else
                                {
                                    _context.summaryResulttableModels.Update(resulttableModel);
                                    _context.SaveChanges();
                                }
                            }
                            catch
                            {
                                _context.summaryResulttableModels.Add(resulttableModel);
                                _context.SaveChanges();
                            }
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                default:
                    try
                    {
                        if (recordid == null)
                        {
                            TempData["message"] = "Session is Expired";
                            return RedirectToAction("LoginFormView", "Home");
                        }
                        else
                        {
                            SummaryResulttableModel resulttableModel = new SummaryResulttableModel();
                            try
                            {
                                resulttableModel = _context.summaryResulttableModels
                                    .Where(a => a.record_Id == recordid)
                                    .FirstOrDefault();
                                if (resulttableModel == null)
                                {
                                    resulttableModel = new SummaryResulttableModel();
                                }
                            }
                            catch
                            {
                                resulttableModel = new SummaryResulttableModel();
                            }
                            resulttableModel.record_Id = recordid;
                            resulttableModel.name_add_Ver_History = mainModel.summarymodel.name_add_Ver_History;
                            resulttableModel.personal_Identification = mainModel.summarymodel.personal_Identification;
                            resulttableModel.social_securitytrace = mainModel.summarymodel.social_securitytrace;
                            resulttableModel.real_estate_prop = mainModel.summarymodel.real_estate_prop;
                            resulttableModel.secretary_state_director = mainModel.summarymodel.secretary_state_director;
                            resulttableModel.news_media_searches = mainModel.summarymodel.news_media_searches;
                            resulttableModel.department_foreign = mainModel.summarymodel.department_foreign;
                            resulttableModel.european_union = mainModel.summarymodel.european_union;
                            resulttableModel.HM_treasury = mainModel.summarymodel.HM_treasury;
                            resulttableModel.US_bureau = mainModel.summarymodel.US_bureau;
                            resulttableModel.US_department = mainModel.summarymodel.US_department;
                            resulttableModel.US_Directorate = mainModel.summarymodel.US_Directorate;
                            resulttableModel.US_general = mainModel.summarymodel.US_general;
                            resulttableModel.US_office = mainModel.summarymodel.US_office;
                            resulttableModel.UN_consolidated = mainModel.summarymodel.UN_consolidated;
                            resulttableModel.world_bank_list = mainModel.summarymodel.world_bank_list;
                            resulttableModel.city_london_police = mainModel.summarymodel.city_london_police;
                            resulttableModel.constabularies_cheshire = mainModel.summarymodel.constabularies_cheshire;
                            resulttableModel.hampshire_police = mainModel.summarymodel.hampshire_police;
                            resulttableModel.hong_kong_police = mainModel.summarymodel.hong_kong_police;
                            resulttableModel.interpol = mainModel.summarymodel.interpol;
                            resulttableModel.metropolitan_police = mainModel.summarymodel.metropolitan_police;
                            resulttableModel.national_crime = mainModel.summarymodel.national_crime;
                            resulttableModel.north_yorkshire_polic = mainModel.summarymodel.north_yorkshire_polic;
                            resulttableModel.nottinghamshire_police = mainModel.summarymodel.nottinghamshire_police;
                            resulttableModel.surrey_police = mainModel.summarymodel.surrey_police;
                            resulttableModel.thames_valley_police = mainModel.summarymodel.thames_valley_police;
                            resulttableModel.US_Federal = mainModel.summarymodel.US_Federal;
                            resulttableModel.US_secret_service = mainModel.summarymodel.US_secret_service;
                            resulttableModel.warwickshire_police = mainModel.summarymodel.warwickshire_police;
                            resulttableModel.alberta_securities_commission = mainModel.summarymodel.alberta_securities_commission;
                            resulttableModel.asset_recovery_agency = mainModel.summarymodel.asset_recovery_agency;
                            resulttableModel.australian_prudential = mainModel.summarymodel.australian_prudential;
                            resulttableModel.australian_securities = mainModel.summarymodel.australian_securities;
                            resulttableModel.banque_de_CECEI = mainModel.summarymodel.banque_de_CECEI;
                            resulttableModel.banque_de_commission = mainModel.summarymodel.banque_de_commission;
                            resulttableModel.british_virgin_islands = mainModel.summarymodel.british_virgin_islands;
                            resulttableModel.cayman_islands_monetary = mainModel.summarymodel.cayman_islands_monetary;
                            resulttableModel.commission_de_surveillance = mainModel.summarymodel.commission_de_surveillance;
                            resulttableModel.commodity_futures = mainModel.summarymodel.commodity_futures;
                            resulttableModel.council_financial_activities = mainModel.summarymodel.council_financial_activities;
                            resulttableModel.departamento_de_investigacoes = mainModel.summarymodel.departamento_de_investigacoes;
                            resulttableModel.department_labour_inspection = mainModel.summarymodel.department_labour_inspection;
                            resulttableModel.federal_deposit = mainModel.summarymodel.federal_deposit;
                            resulttableModel.financial_action_task = mainModel.summarymodel.financial_action_task;
                            resulttableModel.federal_reserve = mainModel.summarymodel.federal_reserve;
                            resulttableModel.financial_crimes = mainModel.summarymodel.financial_crimes;
                            resulttableModel.financial_industry = mainModel.summarymodel.financial_industry;
                            resulttableModel.financial_regulator_ireland = mainModel.summarymodel.financial_regulator_ireland;
                            resulttableModel.hongkong_monetary_authority = mainModel.summarymodel.hongkong_monetary_authority;
                            resulttableModel.hongkong_securities_futures = mainModel.summarymodel.hongkong_securities_futures;
                            resulttableModel.investment_association_Canada = mainModel.summarymodel.banque_de_commission;
                            resulttableModel.investment_management_regulatory = mainModel.summarymodel.investment_management_regulatory;
                            resulttableModel.isle_financial_supervision = mainModel.summarymodel.isle_financial_supervision;
                            resulttableModel.jersey_financial_commission = mainModel.summarymodel.jersey_financial_commission;
                            resulttableModel.lloyd_insurance_arimbolaet = mainModel.summarymodel.lloyd_insurance_arimbolaet;
                            resulttableModel.monetary_authority_singapore = mainModel.summarymodel.monetary_authority_singapore;
                            resulttableModel.new_york_stock = mainModel.summarymodel.new_york_stock;
                            resulttableModel.national_credit = mainModel.summarymodel.national_credit;
                            resulttableModel.Office_of_comptroller = mainModel.summarymodel.Office_of_comptroller;
                            resulttableModel.Office_of_superintendent = mainModel.summarymodel.Office_of_superintendent;
                            resulttableModel.resolution_trust = mainModel.summarymodel.resolution_trust;
                            resulttableModel.securities_exchange = mainModel.summarymodel.securities_exchange;
                            resulttableModel.securities_exchange_commission = mainModel.summarymodel.securities_exchange_commission;
                            resulttableModel.securities_futuresauthority = mainModel.summarymodel.securities_futuresauthority;
                            resulttableModel.swedish_financial_supervisory = mainModel.summarymodel.swedish_financial_supervisory;
                            resulttableModel.swiss_federal_banking = mainModel.summarymodel.swiss_federal_banking;
                            resulttableModel.U_K_companies_disqualified = mainModel.summarymodel.U_K_companies_disqualified;
                            resulttableModel.U_K_financial_conduct_authority = mainModel.summarymodel.U_K_financial_conduct_authority;
                            resulttableModel.US_court = mainModel.summarymodel.US_court;
                            resulttableModel.US_department_justice = mainModel.summarymodel.US_department_justice;
                            resulttableModel.US_federal_trade = mainModel.summarymodel.US_federal_trade;
                            resulttableModel.US_national = mainModel.summarymodel.US_national;
                            resulttableModel.US_office_thrifts = mainModel.summarymodel.US_office_thrifts;
                            resulttableModel.central_intelligence = mainModel.summarymodel.central_intelligence;
                            try
                            {
                                if (resulttableModel == null)
                                {
                                    _context.summaryResulttableModels.Add(resulttableModel);
                                    _context.SaveChanges();
                                }
                                else
                                {
                                    _context.summaryResulttableModels.Update(resulttableModel);
                                    _context.SaveChanges();
                                }
                            }
                            catch
                            {
                                _context.summaryResulttableModels.Add(resulttableModel);
                                _context.SaveChanges();
                            }
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
            }            
            return View(mainModel);
        }
        public static string GetNewFileExtension(string[] fileNames)
        {
            int maxIndex = 0;
            foreach (var fileName in fileNames)
            {
                // get the file extension and remove the "."

                string extension = Path.GetExtension(fileName).Substring(1);
                int parsedIndex;
                // try to parse the file index and do a one pass max element search
                if (int.TryParse(extension, out parsedIndex))
                {
                    if (parsedIndex > maxIndex)
                    {
                        maxIndex = parsedIndex;
                    }
                }
            }
            // increment max index by 1
            return string.Format(".{0}", maxIndex + 1);
        }
        public IActionResult Save_Page1(MainModel uSDDIndividuals)
        {
            try
            {
                string regflag = "";
                string holdresult = "";
                string templatePath;
                string savePath = _config.GetValue<string>("ReportPath");
                MainModel uSDDIndividual = new MainModel();
                string recordid = HttpContext.Session.GetString("recordid");
                uSDDIndividual.familyModels = _context.familyModels
                     .Where(u => u.Family_record_id == uSDDIndividuals.familyModels[0].Family_record_id)
                     .ToList();
                uSDDIndividual.clientspecific = _context.DbclientspecificModels
                     .Where(u => u.record_Id == uSDDIndividuals.familyModels[0].Family_record_id)
                     .FirstOrDefault();
                uSDDIndividual.diligenceInputModel = _context.DbPersonalInfo
                    .Where(u => u.record_Id == uSDDIndividuals.familyModels[0].Family_record_id)
                    .FirstOrDefault();
                if (uSDDIndividual.familyModels.Count == 1)
                {
                    templatePath = _config.GetValue<string>("USCITIFamilytemplatePath");
                }
                else
                {
                    templatePath = _config.GetValue<string>("USCITIFamilytemplatePath");
                }
                switch (uSDDIndividual.clientspecific.clientname)
                {
                    case "Malta":
                        templatePath = string.Concat(templatePath, "US_CITI_Individual_SterlingDiligenceReport(XXXXX)_DRAFT(MALTA).docx");
                        break;
                    case "Dominica":
                        templatePath = string.Concat(templatePath, "US_CITI_Individual_SterlingDiligenceReport(XXXXX)_DRAFT(Dominica).docx");
                        break;
                    case "St. Kitts":
                        templatePath = string.Concat(templatePath, "US_CITI_Individual_SterlingDiligenceReport(XXXXX)_DRAFT(St_Kitts).docx");
                        break;
                    default:
                        templatePath = string.Concat(templatePath, "US_CITI_Individual_SterlingDiligenceReport(XXXXX)_DRAFT(Cyprus).docx");
                        break;
                }
                string pathTo = _config.GetValue<string>("OlderReport"); // the destination file name would be appended later
                savePath = string.Concat(savePath, uSDDIndividual.diligenceInputModel.LastName.ToString(), "_SterlingDiligenceReport(", uSDDIndividual.diligenceInputModel.CaseNumber.ToString(), ")_DRAFT.docx");
                string sourcePath = savePath;
                string filename = Path.GetFileName(sourcePath);
                var fileInfo = new FileInfo(sourcePath);
                if (!fileInfo.Exists)
                {
                }
                else
                {
                    // Get all files by mask "test.txt.*"
                    var files = Directory.GetFiles(pathTo, string.Format("{0}.*", filename)).ToArray();
                    var newExtension = GetNewFileExtension(files); // will return .1, .2, ... .N     
                    try
                    {
                        fileInfo.MoveTo(Path.Combine(pathTo, string.Format("{0}{1}", filename, newExtension)));
                    }
                    catch
                    {

                    }
                }
                Document doc = new Document(templatePath);
                string strblnres = "";                                
                string Client_Name = uSDDIndividual.clientspecific.clientname;
                Table table1 = doc.Sections[3].Tables[0] as Table;

                for (int i = 1; i < uSDDIndividual.familyModels.Count(); i++)
                {
                    Table table2 = table1.Clone();
                    doc.Sections[3].Tables.Add(table2);
                    doc.SaveToFile(savePath);
                }
                doc.Replace("[Month Generated] [Year Generated]", string.Concat(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month).ToString(), " ", DateTime.Now.Year.ToString()), true, true);
                doc.SaveToFile(savePath);
                string current_full_address = "";
                for (int j = 0; j < uSDDIndividuals.familyModels.Count(); j++)
                {
                    string strpreviousresidence = "";
                    string strcurrentresidence = "";
                    int currowcount = 5;
                    uSDDIndividual.diligenceInputModel = _context.DbPersonalInfo
                  .Where(u => u.record_Id == uSDDIndividuals.familyModels[j].Family_record_id)
                  .FirstOrDefault();
                    uSDDIndividual.currentResidenceModels = _context.DbcurrentResidenceModels
                    .Where(u => u.record_Id == uSDDIndividuals.familyModels[j].Family_record_id)
                    .ToList();
                    uSDDIndividual.PreviousResidenceModels = _context.DbpreviousResidenceModels
                    .Where(u => u.record_Id == uSDDIndividuals.familyModels[j].Family_record_id)
                    .ToList();
                    uSDDIndividual.additional_Countries = _context.DbadditionalCountries
                   .Where(u => u.record_Id == uSDDIndividuals.familyModels[j].Family_record_id)
                   .ToList();
                    Table table3 = doc.Sections[3].Tables[j] as Table;
                    TableCell cell11 = table3.Rows[0].Cells[0];
                    Paragraph p1 = cell11.Paragraphs[0];
                    try
                    {
                        p1.Text = string.Concat("\n", uSDDIndividual.diligenceInputModel.FirstName.ToUpper(), " ", uSDDIndividual.diligenceInputModel.MiddleName.ToUpper(), " ", uSDDIndividual.diligenceInputModel.LastName.ToUpper(), " (", uSDDIndividual.diligenceInputModel.foreignlanguage, ") (", uSDDIndividuals.familyModels[j].applicant_type, ")");
                    }
                    catch
                    {
                        p1.Text = string.Concat("\n", uSDDIndividual.diligenceInputModel.FirstName.ToUpper(), " ", uSDDIndividual.diligenceInputModel.LastName.ToUpper(), " (", uSDDIndividual.diligenceInputModel.foreignlanguage, ") (", uSDDIndividuals.familyModels[j].applicant_type, ")");
                    }
                    TableCell cel2 = table3.Rows[1].Cells[1];
                    Paragraph p2 = cel2.Paragraphs[0];
                    try
                    {
                        p2.Text = uSDDIndividual.diligenceInputModel.FullaliasName;
                    }
                    catch
                    {
                        p2.Text = "<not provided>";
                    }
                    TableCell cel3 = table3.Rows[2].Cells[1];
                    Paragraph p3 = cel3.Paragraphs[0];
                    try
                    {
                        p3.Text = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Month) + " " + Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Day + ", " + Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Year;
                    }
                    catch
                    {
                        p3.Text = "<not provided>";
                    }
                    TableCell cel4 = table3.Rows[3].Cells[1];
                    Paragraph p4 = cel4.Paragraphs[0];
                    try
                    {
                        if (uSDDIndividual.diligenceInputModel.Pob.ToString().Equals("") || uSDDIndividual.diligenceInputModel.Pob.ToString().Equals("NA") || uSDDIndividual.diligenceInputModel.Pob.ToString().Equals("N/A"))
                        {
                            p4.Text = "<not provided>";
                        }
                        else
                        {
                            p4.Text = uSDDIndividual.diligenceInputModel.Pob;
                        }
                    }
                    catch
                    {
                        p4.Text = "<not provided>";
                    }
                    TableCell cel5 = table3.Rows[4].Cells[1];
                    Paragraph p5 = cel5.Paragraphs[0];
                    try
                    {
                        p5.Text = uSDDIndividual.diligenceInputModel.Nationality;
                    }
                    catch
                    {
                        p5.Text = "<not provided>";
                    }
                    TableCell cel6 = table3.Rows[5].Cells[1];
                    Paragraph p6 = cel6.Paragraphs[0];

                    for (int i = 0; i < uSDDIndividual.currentResidenceModels.Count(); i++)
                    {
                        if (uSDDIndividual.currentResidenceModels[i].CurrentStreet.ToString().Equals("")) { }
                        else
                        {
                            strcurrentresidence = uSDDIndividual.currentResidenceModels[i].CurrentStreet;
                        }
                        if (uSDDIndividual.currentResidenceModels[i].CurrentCity.ToString().Equals("")) { }
                        else
                        {
                            strcurrentresidence = string.Concat(strcurrentresidence, ", ", uSDDIndividual.currentResidenceModels[i].CurrentCity);
                        }
                        if (uSDDIndividual.currentResidenceModels[i].CurrentCountry.ToString().Equals("")) { }
                        else
                        {
                            strcurrentresidence = string.Concat(strcurrentresidence, ", ", uSDDIndividual.currentResidenceModels[i].CurrentCountry);
                        }
                        if (uSDDIndividual.currentResidenceModels[i].CurrentZipcode.ToString().Equals("")) { }
                        else
                        {
                            strcurrentresidence = string.Concat(strcurrentresidence, ", ", uSDDIndividual.currentResidenceModels[i].CurrentZipcode);
                        }
                        if (i == 0)
                        {
                            
                            if (strcurrentresidence.ToString().Equals(""))
                            {
                                current_full_address = "<not provided>";
                            }
                            else
                            {
                                p6.Text = string.Concat(strcurrentresidence, ", reportedly since <date>");
                                current_full_address = strcurrentresidence;
                            }
                            if (uSDDIndividual.currentResidenceModels.Count() == 1)
                            {
                                currowcount = 5;
                                table3.Rows.RemoveAt(6);
                                table3.Rows.RemoveAt(6);
                            }
                        }
                        else
                        {
                            if (i == 1)
                            {
                                TableCell cel7 = table3.Rows[6].Cells[1];
                                Paragraph p7 = cel7.Paragraphs[0];
                                p7.Text = string.Concat(strcurrentresidence, ", reportedly since <date>");
                                if (uSDDIndividual.currentResidenceModels.Count() == 2)
                                {
                                    currowcount = 6;
                                    table3.Rows.RemoveAt(7);
                                }
                            }
                            else
                            {
                                TableCell cel8 = table3.Rows[7].Cells[1];
                                Paragraph p8 = cel8.Paragraphs[0];
                                currowcount = 7;
                                p8.Text = string.Concat(strcurrentresidence, ", reportedly since <date>");
                            }
                        }

                    }
                    doc.SaveToFile(savePath);
                    try
                    {
                        if (uSDDIndividual.currentResidenceModels.Count() == 1)
                        {
                            doc.Replace("Current Residences", "Current Residence", true, true);
                        }
                    }
                    catch { }
                    //Previous Residence                
                    for (int i = 0; i < uSDDIndividual.PreviousResidenceModels.Count(); i++)
                    {
                        TableCell cel10 = table3.Rows[currowcount + 1].Cells[1];
                        Paragraph p10 = cel10.Paragraphs[0];
                        if (uSDDIndividual.PreviousResidenceModels[i].PreviousStreet.ToString().Equals("")) { }
                        else
                        {
                            strpreviousresidence = uSDDIndividual.PreviousResidenceModels[i].PreviousStreet;
                        }
                        if (uSDDIndividual.PreviousResidenceModels[i].PreviousCity.ToString().Equals("")) { }
                        else
                        {
                            strpreviousresidence = string.Concat(strpreviousresidence, ", ", uSDDIndividual.PreviousResidenceModels[i].PreviousCity);
                        }
                        if (uSDDIndividual.PreviousResidenceModels[i].PreviousCountry.ToString().Equals("")) { }
                        else
                        {
                            strpreviousresidence = string.Concat(strpreviousresidence, ", ", uSDDIndividual.PreviousResidenceModels[i].PreviousCountry);
                        }
                        if (uSDDIndividual.PreviousResidenceModels[i].PreviousZipcode.ToString().Equals("")) { }
                        else
                        {
                            strpreviousresidence = string.Concat(strpreviousresidence, ", ", uSDDIndividual.PreviousResidenceModels[i].PreviousZipcode);
                        }
                        if (i == 0)
                        {
                            if (strpreviousresidence.ToString().Equals(""))
                            {
                                doc.Replace("[PREVIOUSFULLADDRESSDESC]1, reportedly from <date> to <date>", "", true, true);
                            }
                            else
                            {
                                p10.Text = string.Concat(strpreviousresidence, ", reportedly from <date> to <date>");
                            }
                            if (uSDDIndividual.PreviousResidenceModels.Count() == 1)
                            {
                                table3.Rows.RemoveAt(currowcount + 2);
                                table3.Rows.RemoveAt(currowcount + 2);
                                if (strpreviousresidence.ToString().Equals(""))
                                {
                                    table3.Rows.RemoveAt(currowcount + 1);
                                }
                            }
                        }
                        else
                        {
                            if (i == 1)
                            {
                                p10.Text = string.Concat(strpreviousresidence, ", reportedly from <date> to <date>");
                                if (uSDDIndividual.PreviousResidenceModels.Count() == 2)
                                {
                                    table3.Rows.RemoveAt(currowcount + 3);
                                }
                            }
                            else
                            {
                                p10.Text = string.Concat(strpreviousresidence, ", reportedly from <date> to <date>");
                            }
                        }
                    }
                    try
                    {
                        if (uSDDIndividual.PreviousResidenceModels.Count() == 1)
                        {
                            doc.Replace("Previous Residences", "Previous Residence", true, true);
                        }
                        else
                        {
                            if (uSDDIndividual.PreviousResidenceModels.Count() == 0)
                            {
                                table3.Rows.RemoveAt(currowcount + 1);
                                table3.Rows.RemoveAt(currowcount + 1);
                                table3.Rows.RemoveAt(currowcount + 1);
                            }
                        }
                    }
                    catch { }
                    doc.SaveToFile(savePath);
                    currowcount = currowcount + 1;
                    TableCell cel9 = table3.Rows[currowcount].Cells[1];
                    Paragraph p9 = cel9.Paragraphs[0];
                    if (uSDDIndividuals.familyModels[j].adult_minor == "Adult")
                    {
                        //Marital_Status
                        switch (uSDDIndividual.diligenceInputModel.Marital_Status)
                        {
                            case "Married":
                                p9.Text = string.Concat("According to their Application Form  (a copy of which was provided), ", uSDDIndividual.diligenceInputModel.LastName, " married <spouse> on <date> in <location> (a copy of their Marriage Certificate, issued by <authority> number <number>, was provided and was authenticated).[MARRITALSTATUSDESC]");
                                break;
                            case "Never Married":
                                p9.Text = string.Concat("According to their Application Form  (a copy of which was provided), ", uSDDIndividual.diligenceInputModel.LastName, " has never been married.[MARRITALSTATUSDESC]");
                                break;
                            case "Divorced":
                                p9.Text = string.Concat("According to their Application Form  (a copy of which was provided), ", uSDDIndividual.diligenceInputModel.LastName, " previously married <spouse> on <date> in <location> (a copy of their Marriage Certificate, issued by <authority> number <number>, was provided and was authenticated).  They divorced on <date> in <location>  (a copy of their Divorce Certificate, issued by <authority> number <number>, was provided and was authenticated).[MARRITALSTATUSDESC]");
                                break;

                        }
                        doc.SaveToFile(savePath);
                        switch (uSDDIndividual.diligenceInputModel.Children)
                        {
                            case "Children":
                                doc.Replace("[MARRITALSTATUSDESC]", string.Concat("\n\nThey have <number> children together, who reside with their parents at the aforenoted residential address in [Country]: <Name>, born on <date>; and <Name> born on <date>.  As noted above, ", uSDDIndividual.diligenceInputModel.LastName, "’s spouse and children are not included in this application."), true, false);
                                break;
                            case "No Children":
                                doc.Replace("[MARRITALSTATUSDESC]", string.Concat("\n\nAccording to their Application Form, ", uSDDIndividual.diligenceInputModel.LastName, " has no children."), true, false);
                                break;
                        }
                        //if (j == 0)
                        //{
                        //Military_Service
                        if (uSDDIndividual.clientspecific.clientname.ToString().Equals("Malta") || uSDDIndividual.clientspecific.clientname.ToString().Equals("Dominica") || uSDDIndividual.clientspecific.clientname.ToString().Equals("St. Kitts"))
                        {
                            currowcount = currowcount + 1;
                            TableCell cel11 = table3.Rows[currowcount].Cells[1];
                            Paragraph p11 = cel11.Paragraphs[0];
                            switch (uSDDIndividual.clientspecific.Military_Service)
                            {
                                case "Service Confirmed":
                                    p11.Text = string.Concat("As confirmed, ", uSDDIndividual.diligenceInputModel.LastName, " served as a <rank/position> in the <Country> military from <year> to <year>.  <investigator to specify whether military certificate was provided and/or authenticated>");
                                    doc.Replace("MILITARYRES", "Confirmed", true, true);
                                    doc.Replace("MILITARYCOMMENT", "As confirmed, [Last Name] served as a <rank/position> in the <Country> military from <year> to <year>", true, true);
                                    break;
                                case "Service Unconfirmed":
                                    p11.Text = string.Concat(uSDDIndividual.diligenceInputModel.LastName, " reportedly served as a <rank/position> in the <Country> military from <year> to <year>, however, <investigator to specify why not confirmed and whether certificate was provided>.");
                                    doc.Replace("MILITARYRES", "Unconfirmed", true, true);
                                    doc.Replace("MILITARYCOMMENT", "[Last Name] reportedly served as a <rank/position> in the <Country> military from <year> to <year>, however, <investigator to specify why not confirmed>", true, true);
                                    break;
                                case "No Military Service":
                                    p11.Text = string.Concat("The subject reportedly has not served in the military <investigator to specify whether Exemption certificate was provided, adjust confirmation language if confirmed>.");
                                    doc.Replace("MILITARYRES", "N/A", true, true);
                                    doc.Replace("MILITARYCOMMENT", "The subject reportedly has not served in the military <investigator to specify whether Exemption certificate was provided, adjust confirmation language if confirmed>", true, true);
                                    break;
                            }
                        }
                        //}
                    }
                    else
                    {
                        //table3.Rows.RemoveAt(table3.Rows.Count - 1);
                        table3.Rows.RemoveAt(table3.Rows.Count - 1);
                        doc.SaveToFile(savePath);
                        table3.Rows.RemoveAt(table3.Rows.Count - 1);
                        doc.SaveToFile(savePath);
                    }
                    doc.SaveToFile(savePath);
                    TextSelection[] text30 = doc.FindAllString(string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName), false, false);
                    if (text30 != null)
                    {
                        foreach (TextSelection seletion in text30)
                        {
                            seletion.GetAsOneRange().CharacterFormat.FontSize = 11;
                            seletion.GetAsOneRange().CharacterFormat.FontName = "Calibri (Body)";
                            seletion.GetAsOneRange().CharacterFormat.Bold = true;
                        }
                    }
                    doc.SaveToFile(savePath);
                    try
                    {
                        //doc.Replace(string.Concat(uSDDIndividual.diligenceInputModel.MiddleName, "_", j), uSDDIndividual.diligenceInputModel.MiddleName, true, true);
                    }
                    catch
                    {
                        //doc.Replace(string.Concat("_", j), "", true, true);
                    }
                    doc.Replace(" ()", "", false, false);
                    doc.SaveToFile(savePath);
                    doc.SaveToFile(savePath);
                }
                string familyname = "";
                string family_applicant = "";
                string family_middle = "";
                string Family_namesub = "";
                int minor_count = 0;
                int adult_count = 0;
                int age = 0;
                for (int i = 0; i < uSDDIndividuals.familyModels.Count; i++)
                {
                    uSDDIndividual.diligenceInputModel = _context.DbPersonalInfo
                    .Where(u => u.record_Id == uSDDIndividuals.familyModels[i].Family_record_id)
                    .FirstOrDefault();
                    if (uSDDIndividuals.familyModels.Count == 1)
                    {
                        Family_namesub = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                        try { if (uSDDIndividual.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.foreignlanguage); } } catch { }
                        Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.Country);
                        familyname = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                        family_middle = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName);
                        family_applicant = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName);
                    }
                    else
                    {
                        if (uSDDIndividuals.familyModels.Count == 2)
                        {
                            if (uSDDIndividuals.familyModels[i].adult_minor == "Adult")
                            {
                                if (i == 0)
                                {
                                    Family_namesub = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                                    try { if (uSDDIndividual.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.foreignlanguage); } } catch { }
                                    Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.Country);
                                    familyname = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName, " and ");
                                    family_middle = string.Concat(family_middle, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, " and ");
                                    family_applicant = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type, " and ");
                                }
                                else
                                {
                                    Family_namesub = string.Concat(Family_namesub, "\n\n", uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                                    try { if (uSDDIndividual.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.foreignlanguage); } } catch { }
                                    Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.Country);
                                    familyname = string.Concat(familyname, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                                    family_middle = string.Concat(family_middle, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName);
                                    family_applicant = string.Concat(family_applicant, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type);
                                }
                            }
                            else
                            {
                                familyname = string.Concat(familyname, "child");
                                family_middle = string.Concat(family_middle, "child");
                                age = DateTime.Now.Year - Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Year;
                                family_applicant = string.Concat(family_applicant, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type, ", who is ", age, " years old, ");
                            }
                        }
                        else
                        {
                            if (uSDDIndividuals.familyModels[i].adult_minor == "Adult")
                            {
                                adult_count = adult_count + 1;
                            }
                            else
                            {
                                minor_count = minor_count + 1;
                            }
                        }
                    }
                }
                if (adult_count == 0) { }
                else
                {
                    for (int i = 0; i < uSDDIndividuals.familyModels.Count; i++)
                    {
                        uSDDIndividual.diligenceInputModel = _context.DbPersonalInfo
                                       .Where(u => u.record_Id == uSDDIndividuals.familyModels[i].Family_record_id)
                                       .FirstOrDefault();
                        if (minor_count == 0)
                        {
                            if (i == 0)
                            {
                                Family_namesub = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                                try { if (uSDDIndividual.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.foreignlanguage); } } catch { }
                                Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.Country);
                            }
                            else
                            {
                                Family_namesub = string.Concat(Family_namesub, "\n\n", uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                                try { if (uSDDIndividual.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.foreignlanguage); } } catch { }
                                Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.Country);
                            }
                            if (i == uSDDIndividuals.familyModels.Count - 1)
                            {
                                familyname = string.Concat(familyname, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                                family_middle = string.Concat(family_middle, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName);
                                family_applicant = string.Concat(family_applicant, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type, ", ");
                            }
                            else
                            {
                                if (i == uSDDIndividuals.familyModels.Count - 2)
                                {
                                    familyname = string.Concat(familyname, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName, " and ");
                                    family_middle = string.Concat(family_middle, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, " and ");
                                    family_applicant = string.Concat(family_applicant, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type, " and ");
                                }
                                else
                                {
                                    familyname = string.Concat(familyname, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName, ", ");
                                    family_middle = string.Concat(family_middle, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ");
                                    family_applicant = string.Concat(family_applicant, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type, ", ");
                                }
                            }
                        }
                        else
                        {
                            if (uSDDIndividuals.familyModels[i].adult_minor == "Adult")
                            {
                                if (i == 0)
                                {
                                    Family_namesub = string.Concat(uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                                    try { if (uSDDIndividual.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.foreignlanguage); } } catch { }
                                    Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.Country);
                                }
                                else
                                {
                                    Family_namesub = string.Concat(Family_namesub, "\n\n", uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName);
                                    try { if (uSDDIndividual.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.foreignlanguage); } } catch { }
                                    Family_namesub = string.Concat(Family_namesub, "\n", uSDDIndividual.diligenceInputModel.Country);
                                }
                                if (i == uSDDIndividuals.familyModels.Count - 2)
                                {
                                    familyname = string.Concat(familyname, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName, " and ");
                                    family_middle = string.Concat(family_middle, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, " and ");
                                    family_applicant = string.Concat(family_applicant, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type, " and ");
                                }
                                else
                                {
                                    familyname = string.Concat(familyname, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleInitial, " ", uSDDIndividual.diligenceInputModel.LastName, ", ");
                                    family_middle = string.Concat(family_middle, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ");
                                    family_applicant = string.Concat(family_applicant, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type, ", ");
                                }
                            }
                            else
                            {
                                if (i == uSDDIndividuals.familyModels.Count - 2)
                                {
                                    familyname = string.Concat(familyname, " and ");
                                    family_middle = string.Concat(family_middle, " and ");
                                }
                                if (i == uSDDIndividuals.familyModels.Count - 1)
                                {
                                    family_applicant = string.Concat(family_applicant, " and ");
                                }
                                age = DateTime.Now.Year - Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Year;
                                family_applicant = string.Concat(family_applicant, uSDDIndividual.diligenceInputModel.FirstName, " ", uSDDIndividual.diligenceInputModel.MiddleName, " ", uSDDIndividual.diligenceInputModel.LastName, ", ", uSDDIndividuals.familyModels[i].applicant_type, ", who is ", age, " years old, ");
                            }
                        }
                    }
                    if (minor_count == 1)
                    {
                        familyname = string.Concat(familyname, "child");
                        family_middle = string.Concat(family_middle, "child");
                    }
                    else
                    {
                        familyname = string.Concat(familyname, "FAMILY");
                        family_middle = string.Concat(family_middle, "FAMILY");
                    }
                }
                familyname = familyname.ToUpper();
                doc.Replace("[FamilyWithsubandcoun]", Family_namesub, true, true);
                doc.Replace("[FamilyNames]", familyname, true, true);
                doc.Replace("[Family_middle_name_applicant]", family_applicant, true, true);
                doc.Replace("[Family_Middle_Name]", family_middle.ToUpper(), true, true);
                doc.SaveToFile(savePath);
                if (uSDDIndividual.diligenceInputModel.FullaliasName.ToString().Equals("") || uSDDIndividual.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("NA") || uSDDIndividual.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("N/A"))
                {
                    doc.Replace("(also known as [FullAliasName])", "", true, false);
                }
                else
                {
                    string strfn = "";
                    if (uSDDIndividual.diligenceInputModel.CommonNameSubject == true)
                    {
                        strfn = "It is noted that the subject was identified in connection with the alternate name variations of “[FullAliasName]” and investigative efforts were undertaken in connection with the same, as appropriate. Additionally, due to the commonality of [LastName]'s name, research efforts were focused to the subject's known jurisdictions and timeframe of [LastName]'s affiliation with the same.";
                    }
                    else
                    {
                        strfn = "It is noted that the subject was identified in connection with the name variation “[FullAliasName]” and investigative efforts were undertaken in connection with the same, as appropriate.";
                    }
                    doc.Replace("[FullAliasName])", "[FullAliasName])", true, false);
                    for (int j = 2; j < 5; j++)
                    {
                        Section section = doc.Sections[j];
                        foreach (Paragraph para in section.Paragraphs)
                        {
                            DocumentObject obj = null;
                            for (int i = 0; i < para.ChildObjects.Count; i++)
                            {
                                obj = para.ChildObjects[i];
                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                {
                                    TextRange textRange = obj as TextRange;
                                    string abc = textRange.Text;
                                    // Find the word "Spire.Doc" in paragraph1
                                    if (abc.ToString().Equals("[FullAliasName]") || abc.ToString().EndsWith("[FullAliasName])") || abc.ToString().Contains("FullAliasName"))
                                    {
                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                        //Insert footnote1 after the word "Spire.Doc"
                                        para.ChildObjects.Insert(i + 1, footnote1);
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strfn);
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        doc.SaveToFile(savePath);
                                        strblnres = "true";
                                        break;
                                    }
                                }
                            }
                            if (strblnres == "true")
                            {
                                break;
                            }
                        }
                        if (strblnres == "true")
                        {
                            break;
                        }
                    }
                    doc.Replace("[FullAliasName]", uSDDIndividual.diligenceInputModel.FullaliasName.TrimEnd(), true, false);
                }

                doc.Replace("First_Name", uSDDIndividual.diligenceInputModel.FirstName.ToUpper().TrimEnd().ToString(), true, true);
                try
                {
                    if (uSDDIndividual.diligenceInputModel.MiddleName.ToUpper().TrimEnd().ToString() == "") {
                        doc.Replace(" Middle_Name", "", false, false);
                    }
                    else
                    {
                        doc.Replace("Middle_Name", uSDDIndividual.diligenceInputModel.MiddleName.ToUpper().TrimEnd().ToString(), true, true);
                    }
                }
                catch
                {
                    doc.Replace(" Middle_Name", "", true, false);
                }
                doc.Replace("Last_Name", uSDDIndividual.diligenceInputModel.LastName.TrimEnd().ToUpper().ToString(), true, true);
                try
                {
                    doc.Replace("ClientName", uSDDIndividual.diligenceInputModel.ClientName.TrimEnd(), true, true);
                }
                catch
                {
                    doc.Replace("ClientName", "<not provided>", true, false);
                }
                try
                {
                    doc.Replace("Position1", uSDDIndividual.EmployerModel[0].Emp_Position.TrimEnd(), true, true);
                }
                catch
                {
                    doc.Replace("Position1", "", true, true);
                }
                try
                {
                    doc.Replace("Employer1", uSDDIndividual.EmployerModel[0].Emp_Employer.TrimEnd(), true, true);
                }
                catch
                {
                    doc.Replace("Employer1", "", true, true);
                }
                try
                {
                    doc.Replace("ClientName", uSDDIndividual.diligenceInputModel.ClientName.TrimEnd(), true, true);
                }
                catch
                {
                    doc.Replace("ClientName", "<not provided>", true, true);
                }
                try
                {
                    doc.Replace("[DateofBirth]", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Month) + " " + Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Year, true, true);
                }
                catch
                {
                    doc.Replace("[DateofBirth]", "<not provided>", true, true);
                }
                try
                {
                    if (uSDDIndividual.diligenceInputModel.SSNInitials.TrimEnd().ToString() == "")
                    {
                        doc.Replace("[Initial_digits_of_SSN]", uSDDIndividual.diligenceInputModel.SSNInitials.TrimEnd().ToString(), true, true);
                    }
                    else
                    {
                        doc.Replace("[Initial_digits_of_SSN]", "", true, true);
                    }
                }
                catch
                {
                    doc.Replace("[Initial_digits_of_SSN]", "", true, true);
                }
                doc.Replace("[CaseNumber]", uSDDIndividual.diligenceInputModel.CaseNumber.TrimEnd(), true, true);
                
                              
                doc.SaveToFile(savePath);
                for (int famcount = 0; famcount < uSDDIndividuals.familyModels.Count; famcount++)
                {
                    uSDDIndividual.familyModels = _context.familyModels
               .Where(u => u.Family_record_id == uSDDIndividuals.familyModels[famcount].Family_record_id)
               .ToList();
                    uSDDIndividual.clientspecific = _context.DbclientspecificModels
                    .Where(u => u.record_Id == uSDDIndividuals.familyModels[0].Family_record_id)
                    .FirstOrDefault();
                    uSDDIndividual.pllicenseModels = _context.DbPLLicense
                        .Where(u => u.record_Id == uSDDIndividuals.familyModels[famcount].Family_record_id)
                        .ToList();
                    uSDDIndividual.EmployerModel = _context.DbEmployer
                        .Where(u => u.record_Id == uSDDIndividuals.familyModels[famcount].Family_record_id)
                        .ToList();
                    uSDDIndividual.educationModels = _context.DbEducation
                        .Where(u => u.record_Id == uSDDIndividuals.familyModels[famcount].Family_record_id)
                        .ToList();
                    uSDDIndividual.referencetableModels = _context.DbReferencetableModel
                       .Where(u => u.record_Id == uSDDIndividuals.familyModels[famcount].Family_record_id)
                       .ToList();
                    uSDDIndividual.summarymodel = _context.summaryResulttableModels
                         .Where(u => u.record_Id == uSDDIndividuals.familyModels[famcount].Family_record_id)
                         .FirstOrDefault();
                    uSDDIndividual.otherdetails = _context.othersModel
                        .Where(u => u.record_Id == uSDDIndividuals.familyModels[famcount].Family_record_id)
                        .FirstOrDefault();
                    uSDDIndividual.diligenceInputModel = _context.DbPersonalInfo
                        .Where(u => u.record_Id == uSDDIndividuals.familyModels[famcount].Family_record_id)
                        .FirstOrDefault();
                    uSDDIndividual.additional_States = _context.Dbadditionalstates
                        .Where(u => u.record_Id == uSDDIndividuals.familyModels[famcount].Family_record_id)
                        .ToList();


                    //Business Affiliation              
                    if (uSDDIndividual.otherdetails.Has_Business_Affiliations.ToString().Equals("Yes"))
                    {
                        doc.Replace("SODCOMMENT", "[LastName] is a [Position1] of [Employer1] in United States, since [EmpStartDate1][SODCOMMENT]\n\nFurther, [LastName] was identified in connection with additional business entities", true, true);
                        doc.Replace("SODRESULT", "Records", true, true);
                        doc.Replace("[BUSINESSAFFILIATIONSIDENTIFIED]", "Moreover, pursuant to <his/her> express written authorization, research efforts were conducted utilizing the subject’s Social Security Number through an automated third-party verification system that houses personnel records for numerous large companies in the United States, to identify any other employment history, and <Investigator to add relevant results, note that “no additional employment was identified” OR remove paragraph if non-consented investigation.>\n\nIn addition to the above, research of records maintained by the Secretary of State’s Office, as well as other sources, revealed the subject in connection with the following business entities:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>[COMADDBUSINESSAFTER]", true, true);
                    }
                    else
                    {
                        doc.Replace("SODCOMMENT", "[LastName] is a [Position1] of [Employer1] in United States, since [EmpStartDate1]", true, true);
                        doc.Replace("SODRESULT", "Clear", true, true);
                        doc.Replace("[BUSINESSAFFILIATIONSIDENTIFIED]", "Moreover, pursuant to <his/her> express written authorization, research efforts were conducted utilizing the subject’s Social Security Number through an automated third-party verification system that houses personnel records for numerous large companies in the United States, to identify any other employment history, and <Investigator to add relevant results, note that “no additional employment was identified” OR remove paragraph if non-consented investigation.>\n\nIn addition to the above, research of records maintained by the Secretary of State’s Office, as well as other sources, did not identify the subject in connection with any business entities.[COMADDBUSINESSAFTER]", true, true);
                    }
                    try
                    {
                        if (uSDDIndividual.EmployerModel[0].Emp_Position.Equals("")) { doc.Replace("[Position1]", "<not provided>", true, true); }
                        else { doc.Replace("[Position1]", uSDDIndividual.EmployerModel[0].Emp_Position, true, true); }
                    }
                    catch
                    {
                        doc.Replace("[Position1]", "<not provided>", true, true);
                    }
                    try
                    {
                        if (uSDDIndividual.EmployerModel[0].Emp_Employer.Equals("")) { doc.Replace("[Employer1]", "<not provided>", true, true); }
                        else { doc.Replace("[Employer1]", uSDDIndividual.EmployerModel[0].Emp_Employer, true, true); }
                    }
                    catch
                    {
                        doc.Replace("[Employer1]", "<not provided>", true, true);
                    }
                    try
                    {
                        if (uSDDIndividual.EmployerModel[0].Emp_StartDateYear.ToString().Equals(""))
                        {
                            doc.Replace("[EmpStartDate1]", "<not provided>", true, true);
                        }
                        else
                        {
                            doc.Replace("[EmpStartDate1]", uSDDIndividual.EmployerModel[0].Emp_StartDateYear, true, true);
                        }
                    }
                    catch
                    {
                        doc.Replace("[EmpStartDate1]", "<not provided>", true, true);
                    }
                    doc.SaveToFile(savePath);
                    string strcommentconcurrent = "";
                    for (int i = 1; uSDDIndividual.EmployerModel.Count < i; i++)
                    {
                        if (uSDDIndividual.EmployerModel[i].Emp_Status == "Concurrent")
                        {
                            strcommentconcurrent = string.Concat(strcommentconcurrent, "\n\n[LastName] is a [Position1] of [Employer1] in [Employer1Location], since [EmpStartDate1]");
                            strcommentconcurrent = strcommentconcurrent.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName);
                            try
                            {
                                if (uSDDIndividual.EmployerModel[i].Emp_Position.Equals("")) { strcommentconcurrent = strcommentconcurrent.Replace("[Position1]", "<not provided>"); }
                                else { strcommentconcurrent = strcommentconcurrent.Replace("[Position1]", uSDDIndividual.EmployerModel[i].Emp_Position); }
                            }
                            catch
                            {
                                strcommentconcurrent = strcommentconcurrent.Replace("[Position1]", "<not provided>");
                            }
                            try
                            {
                                if (uSDDIndividual.EmployerModel[i].Emp_Employer.Equals("")) { strcommentconcurrent = strcommentconcurrent.Replace("[Employer1]", "<not provided>"); }
                                else { strcommentconcurrent = strcommentconcurrent.Replace("[Employer1]", uSDDIndividual.EmployerModel[i].Emp_Employer); }
                            }
                            catch
                            {
                                strcommentconcurrent = strcommentconcurrent.Replace("[Employer1]", "<not provided>");
                            }
                            try
                            {
                                if (uSDDIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals(""))
                                {
                                    strcommentconcurrent = strcommentconcurrent.Replace("[EmpStartDate1]", "<not provided>");
                                }
                                else
                                {
                                    strcommentconcurrent = strcommentconcurrent.Replace("[EmpStartDate1]", uSDDIndividual.EmployerModel[i].Emp_StartDateYear);
                                }
                            }
                            catch
                            {
                                strcommentconcurrent = strcommentconcurrent.Replace("[EmpStartDate1]", "<not provided>");
                            }
                            try
                            {
                                if (uSDDIndividual.EmployerModel[i].Emp_Location.ToString().Equals(""))
                                {
                                    strcommentconcurrent = strcommentconcurrent.Replace("[Employer1Location]", "<not provided>");
                                }
                                else
                                {
                                    strcommentconcurrent = strcommentconcurrent.Replace("[Employer1Location]", uSDDIndividual.EmployerModel[i].Emp_Location);
                                }
                            }
                            catch
                            {
                                strcommentconcurrent = strcommentconcurrent.Replace("[Employer1Location]", "<not provided>");
                            }
                        }
                    }
                    doc.Replace("[SODCOMMENT]", strcommentconcurrent, true, false);
                    doc.SaveToFile(savePath);
                    if (uSDDIndividual.diligenceInputModel.CommonNameSubject == true)
                    {
                        doc.Replace("[COMADDBUSINESSAFTER]", "\n\nIt should be noted that numerous individuals known only as “[First Name] [Last Name]” were identified in corporate records throughout the United States, and significant additional research efforts would be required to determine whether any of the same relate to the subject of this investigation, which can be undertaken upon request -- if an expanded scope is warranted.", true, true);
                    }
                    else
                    {
                        doc.Replace("[COMADDBUSINESSAFTER]", "", true, true);
                    }
                    doc.SaveToFile(savePath);
                    //Intellectual Hits
                    string intellectual_comment;
                    CommentModel intellec_comment2 = _context.DbComment
                                       .Where(u => u.Comment_type == "Intellectual_hits")
                                       .FirstOrDefault();
                    if (uSDDIndividual.otherdetails.Has_Intellectual_Hits.ToString().Equals("Yes"))
                    {
                        intellectual_comment = intellec_comment2.confirmed_comment.ToString();
                        intellectual_comment = string.Concat("\n", intellectual_comment.ToString(), "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n");
                    }
                    else
                    {
                        intellectual_comment = "";
                    }
                    doc.Replace("INTELLECTUALPROPERTYIDENTIFIED", intellectual_comment, true, true);
                    doc.SaveToFile(savePath);

                    //worldcheck_discloseBA
                    switch (uSDDIndividual.otherdetails.worldcheck_discloseBA)
                    {
                        case "Yes":
                            doc.Replace("[WORLDCHECKDISCLOSEBADESC]", "\nIt should be noted that government compliance, regulatory agencies, anti-terrorist and anti-money laundering searches were conducted in connection with the afore-mentioned business affiliations, which revealed that the following information:[DISCLOSEFONTCHANGE]\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n", true, false);
                            break;
                        case "No":
                            doc.Replace("[WORLDCHECKDISCLOSEBADESC]", "\nIt should be noted that government compliance, regulatory agencies, anti-terrorist and anti-money laundering searches were conducted in connection with the afore-mentioned business affiliations, and no such records were identified.[DISCLOSEFONTCHANGE]\n", true, false);
                            break;
                        case "N/A":
                            doc.Replace("[WORLDCHECKDISCLOSEBADESC]", "", true, false);
                            break;
                    }
                    doc.SaveToFile(savePath);
                    if (uSDDIndividual.otherdetails.worldcheck_discloseBA.ToString().Equals("Yes") || uSDDIndividual.otherdetails.worldcheck_discloseBA.ToString().Equals("No"))
                    {
                        string disbnres = "";
                        try
                        {
                            for (int j = 1; j < 7; j++)
                            {
                                Section section = doc.Sections[j];
                                foreach (Paragraph para in section.Paragraphs)
                                {
                                    DocumentObject obj = null;
                                    for (int k = 0; k < para.ChildObjects.Count; k++)
                                    {
                                        obj = para.ChildObjects[k];
                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                        {
                                            TextRange textRange = obj as TextRange;
                                            string abc = textRange.Text;
                                            // Find the word "Civil" in paragraph1
                                            if (abc.ToString().Contains("[DISCLOSEFONTCHANGE]"))
                                            {
                                                textRange.CharacterFormat.Italic = true;
                                                doc.SaveToFile(savePath);
                                                doc.Replace("[DISCLOSEFONTCHANGE]", "", true, false);
                                                doc.SaveToFile(savePath);
                                                disbnres = "true";
                                                break;
                                            }
                                        }
                                    }
                                    if (disbnres.Equals("true")) { break; }
                                }
                                if (disbnres.Equals("true")) { break; }
                            }
                        }
                        catch { }
                    }
                    //undisclosedBA
                    switch (uSDDIndividual.otherdetails.undisclosedBA)
                    {
                        case "Yes":
                            doc.Replace("[UNDISCLOSEDBADESC]", "While not reported by the subject, research of records maintained by the [Corp Registry], as well as other sources, revealed the subject in connection with the following additional business entities:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>", true, false);
                            break;
                        case "No":
                            doc.Replace("[UNDISCLOSEDBADESC]", "In addition to the above, research of records maintained by the [Corp Registry], as well as other sources, did not identify the subject in connection with any unreported business entities.", true, false);
                            break;
                    }
                    //worldcheck_undiscloseBA

                    if (famcount == uSDDIndividual.familyModels.Count - minor_count - 1 || famcount == uSDDIndividual.familyModels.Count - 1)
                    {
                        switch (uSDDIndividual.otherdetails.worldcheck_undiscloseBA)
                        {
                            case "Yes":
                                doc.Replace("[WORLDCHECKUNDISCLOSEDBADESC]", "\nIt should be noted that government compliance, regulatory agencies, anti-terrorist and anti-money laundering searches were conducted in connection with the afore-mentioned business affiliations, which revealed that the following information:[UNDISCLOSEFONTCHANGE]\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n", true, false);
                                break;
                            case "No":
                                doc.Replace("[WORLDCHECKUNDISCLOSEDBADESC]", "\nIt should be noted that government compliance, regulatory agencies, anti-terrorist and anti-money laundering searches were conducted in connection with the afore-mentioned business affiliations, and no such records were identified.[UNDISCLOSEFONTCHANGE]\n", true, false);
                                break;
                            case "N/A":
                                doc.Replace("[WORLDCHECKUNDISCLOSEDBADESC]", "", true, false);
                                break;
                        }
                    }
                    else
                    {
                        switch (uSDDIndividual.otherdetails.worldcheck_undiscloseBA)
                        {
                            case "Yes":
                                doc.Replace("[WORLDCHECKUNDISCLOSEDBADESC]", "\nIt should be noted that government compliance, regulatory agencies, anti-terrorist and anti-money laundering searches were conducted in connection with the afore-mentioned business affiliations, which revealed that the following information:[UNDISCLOSEFONTCHANGE]\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\nFirst_Name Middle_Name Last_Name\n\nREPORTED BUSINESS AFFILIATIONS AND EMPLOYMENT HISTORY\n\nEMPLOYEEDESCRIPTION\n[BUSINESSAFFILIATIONSIDENTIFIED]\nINTELLECTUALPROPERTYIDENTIFIED[WORLDCHECKDISCLOSEBADESC]\nUNDISCLOSED AFFILIATIONS\n\n[UNDISCLOSEDBADESC]\n[WORLDCHECKUNDISCLOSEDBADESC]", true, false);
                                break;
                            case "No":
                                doc.Replace("[WORLDCHECKUNDISCLOSEDBADESC]", "\nIt should be noted that government compliance, regulatory agencies, anti-terrorist and anti-money laundering searches were conducted in connection with the afore-mentioned business affiliations, and no such records were identified.[UNDISCLOSEFONTCHANGE]\n\nFirst_Name Middle_Name Last_Name\n\nREPORTED BUSINESS AFFILIATIONS AND EMPLOYMENT HISTORY\n\nEMPLOYEEDESCRIPTION\n[BUSINESSAFFILIATIONSIDENTIFIED]\nINTELLECTUALPROPERTYIDENTIFIED[WORLDCHECKDISCLOSEBADESC]\nUNDISCLOSED AFFILIATIONS\n\n[UNDISCLOSEDBADESC]\n[WORLDCHECKUNDISCLOSEDBADESC]", true, false);
                                break;
                            case "N/A":
                                doc.Replace("[WORLDCHECKUNDISCLOSEDBADESC]", "\n\nFirst_Name Middle_Name Last_Name\n\nREPORTED BUSINESS AFFILIATIONS AND EMPLOYMENT HISTORY\n\nEMPLOYEEDESCRIPTION\n[BUSINESSAFFILIATIONSIDENTIFIED]\nINTELLECTUALPROPERTYIDENTIFIED[WORLDCHECKDISCLOSEBADESC]\nUNDISCLOSED AFFILIATIONS\n\n[UNDISCLOSEDBADESC]\n[WORLDCHECKUNDISCLOSEDBADESC]", true, false);
                                break;
                        }
                    }
                    doc.SaveToFile(savePath);




                    //Global security hits
                    string globalhit_comment;
                    //CommentModel globalhit_comment2 = _context.DbComment
                    //                  .Where(u => u.Comment_type == "Global_Security")
                    //                  .FirstOrDefault();
                    if (uSDDIndividual.otherdetails.Global_Security_Hits.ToString().Equals("Yes"))
                    {
                        //globalhit_comment = globalhit_comment2.confirmed_comment.ToString();
                        globalhit_comment = "Research was undertaken in connection with various governmental, quasi-governmental and private sector list repositories maintained for purposes of international security, fraud prevention, anti-terrorism and anti-money laundering, including the List of Specially Designated Nationals and Blocked Persons of the United States Office of Foreign Assets Control (“OFAC”), the United States Bureau of Industry and Security, and other agencies,";
                        globalhit_comment = string.Concat(globalhit_comment, "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here> \n");
                    }
                    else
                    {
                        //globalhit_comment = globalhit_comment2.unconfirmed_comment.ToString();
                        globalhit_comment = "Research was undertaken in connection with various governmental, quasi-governmental and private sector list repositories maintained for purposes of international security, fraud prevention, anti-terrorism and anti-money laundering, including the List of Specially Designated Nationals and Blocked Persons of the United States Office of Foreign Assets Control (“OFAC”), the United States Bureau of Industry and Security, and other agencies,";
                        globalhit_comment = string.Concat(globalhit_comment, "\n");
                    }
                    doc.Replace("GLOBALSECURITYHITSDESCRIPTION", globalhit_comment, true, true);
                    doc.SaveToFile(savePath);

                    //Global_Sec_Family_Hits
                    if (uSDDIndividual.otherdetails.Global_Sec_Family_Hits.ToString().Equals("Yes"))
                    {
                        doc.Replace("GLOBALSECFAMILYHITSDESCRIPTION", "Searches of the same sources were also conducted in connection with the subject’s <family members> (utilizing their names and provided in their Application Forms) and the following records were identified:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n", true, true);
                    }
                    else
                    {
                        doc.Replace("GLOBALSECFAMILYHITSDESCRIPTION", "Searches of the same sources were also conducted in connection with the subject’s <family members> (utilizing their names and provided in their Application Forms) and no such records were identified.\n", true, true);
                    }
                    doc.SaveToFile(savePath);
                    //PEP_Hits
                    if (uSDDIndividual.otherdetails.PEP_Hits.ToString().Equals("Yes"))
                    {
                        doc.Replace("PEPHITSDESCRIPTION", "Further, research efforts conducted through electronically-available lists of “politically-exposed” persons and entities identified the following information in connection with [Last Name]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n", true, true);
                        doc.Replace("POLITICRESULT", "Record", true, true);
                        doc.Replace("POLITICCOMMENT", "<investigator to add details>", true, true);
                    }
                    else
                    {
                        doc.Replace("PEPHITSDESCRIPTION", "Further, research efforts conducted through electronically-available lists of “politically-exposed” persons and entities did not identify any such records in connection with [Last Name].\n", true, true);
                        doc.Replace("POLITICRESULT", "Clear", true, true);
                        doc.Replace("POLITICCOMMENT", "", true, true);
                    }
                    doc.SaveToFile(savePath);

                    //ICIJ hits
                    string icij_comment;
                    CommentModel icij_comment2 = _context.DbComment
                                      .Where(u => u.Comment_type == "ICIJ")
                                      .FirstOrDefault();
                    string icijsummcomment = "";
                    string icijsummresult = "";
                    if (uSDDIndividual.otherdetails.ICIJ_Hits.ToString().Equals("Yes"))
                    {
                        icij_comment = icij_comment2.confirmed_comment.ToString();
                        icij_comment = string.Concat(icij_comment, "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here> \n");
                        icijsummcomment = "<investigator to insert summary here>";
                        icijsummresult = "Records";
                    }
                    else
                    {
                        icij_comment = icij_comment2.unconfirmed_comment.ToString();
                        icij_comment = string.Concat(icij_comment, "\n");
                        icijsummcomment = "";
                        icijsummresult = "Clear";
                    }
                    doc.Replace("ICIJHITSDESCRIPTION", icij_comment, true, true);
                    doc.Replace("INCOJOCOMMENT", icijsummcomment, true, true);
                    doc.Replace("INCOJORESULT", icijsummresult, true, true);
                    doc.SaveToFile(savePath);
                    //US_SEC
                    string usseccommentmodel = "";
                    CommentModel usseccommentModel1 = _context.DbComment
                                     .Where(u => u.Comment_type == "Reg_US_SEC")
                                     .FirstOrDefault();
                    if (uSDDIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered"))
                    {
                        usseccommentmodel = "";
                        doc.Replace("[SECCOMMENT]", "", true, true);
                        doc.Replace("[SECRESULT]", "Clear", true, true);
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Yes-With Adverse"))
                        {
                            usseccommentmodel = usseccommentModel1.confirmed_comment.ToString();
                            doc.Replace("[SECCOMMENT]", "<investigator to insert summary here>", true, true);
                            doc.Replace("[SECRESULT]", "Records", true, true);
                        }
                        else
                        {
                            usseccommentmodel = usseccommentModel1.unconfirmed_comment.ToString();
                            doc.Replace("[SECCOMMENT]", "", true, true);
                            doc.Replace("[SECRESULT]", "Clear", true, true);
                        }
                        usseccommentmodel = usseccommentmodel.Replace("*n ", "\n");
                        usseccommentmodel = usseccommentmodel.Replace("*n", "\n");
                        usseccommentmodel = usseccommentmodel.Replace("*t", "\t");
                        // doc.Replace("US_SECHEADER", " \nUnited States Securities and Exchange Commission \n", true, true);
                        usseccommentmodel = string.Concat("\nUnited States Securities and Exchange Commission\n\n", usseccommentmodel, "\n");
                    }
                    //UK_FCA
                    string ukfcacommentmodel = "";
                    CommentModel ukfcacommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "Reg_UK_FCA")
                                    .FirstOrDefault();
                    if (uSDDIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered"))
                    {
                        ukfcacommentmodel = "";
                        doc.Replace("UKFICOCOMMENT", "", true, true);
                        doc.Replace("UKFICORESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Yes-With Adverse"))
                        {
                            ukfcacommentmodel = ukfcacommentmodel1.confirmed_comment.ToString();
                            doc.Replace("UKFICOCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("UKFICORESULT", "Records", true, true);
                        }
                        else
                        {
                            ukfcacommentmodel = ukfcacommentmodel1.unconfirmed_comment.ToString();
                            doc.Replace("UKFICOCOMMENT", "", true, true);
                            doc.Replace("UKFICORESULT", "Clear", true, true);
                        }
                        ukfcacommentmodel = ukfcacommentmodel.Replace("*n ", "\n");
                        ukfcacommentmodel = ukfcacommentmodel.Replace("*n", "\n");
                        ukfcacommentmodel = ukfcacommentmodel.Replace("*t", "\t");
                        // doc.Replace("UK_FCAHEADER", "\n United Kingdom’s Financial Conduct Authority \n", true, true);
                        ukfcacommentmodel = string.Concat("\nUnited Kingdom’s Financial Conduct Authority\n\n", ukfcacommentmodel, "\n");
                    }
                    //FINRA
                    string finracommentmodel = "";
                    CommentModel finracommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "FINRA")
                                    .FirstOrDefault();
                    if (uSDDIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered"))
                    {
                        finracommentmodel = "";
                        doc.Replace("FININDCOMMENT", "", true, true);
                        doc.Replace("FININDRESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Yes-With Adverse"))
                        {
                            finracommentmodel = finracommentmodel1.confirmed_comment.ToString();
                            doc.Replace("FININDCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("FININDRESULT", "Records", true, true);
                        }
                        else
                        {
                            finracommentmodel = finracommentmodel1.unconfirmed_comment.ToString();
                            doc.Replace("FININDCOMMENT", "", true, true);
                            doc.Replace("FININDRESULT", "Clear", true, true);
                        }
                        finracommentmodel = finracommentmodel.Replace("*n ", "\n");
                        finracommentmodel = finracommentmodel.Replace("*n", "\n");
                        finracommentmodel = finracommentmodel.Replace("*t", "\t");
                        //doc.Replace("FINRAHEADER", "\n United States Financial Industry Regulatory Authority \n", true, true);
                        finracommentmodel = string.Concat("\nUnited States Financial Industry Regulatory Authority\n\n", finracommentmodel, "\n");
                    }

                    //NFA
                    string nfacommentmodel = "";
                    CommentModel nfacommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "NFA")
                                    .FirstOrDefault();
                    if (uSDDIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered"))
                    {
                        nfacommentmodel = "";
                        doc.Replace("USNFCOMMENT", "", true, true);
                        doc.Replace("USNFRESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Yes-With Adverse"))
                        {
                            nfacommentmodel = nfacommentmodel1.confirmed_comment.ToString();
                            doc.Replace("USNFCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USNFRESULT", "Records", true, true);
                        }
                        else
                        {
                            doc.Replace("USNFCOMMENT", "", true, true);
                            doc.Replace("USNFRESULT", "Clear", true, true);
                            nfacommentmodel = nfacommentmodel1.unconfirmed_comment.ToString();
                        }
                        nfacommentmodel = nfacommentmodel.Replace("*n ", "\n");
                        nfacommentmodel = nfacommentmodel.Replace("*n", "\n");
                        nfacommentmodel = nfacommentmodel.Replace("*t", "\t");
                        //doc.Replace("US_NFAHEADER", "\n United States National Futures Association  \n", true, true);
                        nfacommentmodel = string.Concat("\nUnited States National Futures Association\n\n", nfacommentmodel, "\n");
                    }
                    string hksfccommentmodel = "";
                    string holdslicensecommentmodel = "";
                    //string regflag = "";
                    //HKFSC
                    if (uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Not Registered"))
                    {
                        hksfccommentmodel = "";
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes - Without Adverse"))
                        {
                            hksfccommentmodel = "[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>)), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nThere are no public disciplinary actions on file against the subject within the past five years.";
                        }
                        else
                        {
                            hksfccommentmodel = "[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>)), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nAdditionally, the following public disciplinary actions were on file in connection with the subject:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                        }
                        hksfccommentmodel = string.Concat("\nHong Kong Securities and Futures Commission\n\n", hksfccommentmodel, "\n");
                    }

                    //Holds Any License                  
                    CommentModel holdslicensecommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "Other_License_Language")
                                    .FirstOrDefault();
                    if (uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse"))
                    {
                        regflag = "Records";
                    }
                    else
                    {
                        regflag = "Clear";

                    }

                    if (uSDDIndividual.otherdetails.Registered_with_HKSFC.StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_FINRA.StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.StartsWith("Yes") || uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                    {
                        holdslicensecommentmodel = "\nOther Professional Licensures and/or Designations\n\nInvestigative efforts did not reveal any additional professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction and license-type basis to confirm the same.\n";
                    }
                    else
                    {
                        holdslicensecommentmodel = "\nInvestigative efforts did not reveal any professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction and license-type basis to confirm the same.";
                    }
                    if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().ToUpper().Equals("YES")) { doc.Replace("[PLLICENSEDESC]", string.Concat("\n", usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false); }
                    else
                    {
                        doc.Replace("[PLLICENSEDESC]", string.Concat(usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false);
                    }
                    doc.SaveToFile(savePath);

                    if (uSDDIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered")) { }
                    else
                    {
                        string blnresult = "";
                        for (int j = 1; j < 6; j++)
                        {
                            Section section = doc.Sections[j];
                            foreach (Paragraph para in section.Paragraphs)
                            {
                                DocumentObject obj = null;
                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                {
                                    obj = para.ChildObjects[i];
                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                    {
                                        TextRange textRange = obj as TextRange;
                                        string abc = textRange.Text;
                                        if (abc.ToString().Equals("United States Securities and Exchange Commission"))
                                        {
                                            textRange.CharacterFormat.Italic = true;
                                            textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                            doc.SaveToFile(savePath);
                                            blnresult = "true";
                                            break;
                                        }
                                    }
                                }
                                if (blnresult.Equals("true")) { break; }
                            }
                            if (blnresult.Equals("true")) { break; }
                        }
                    }

                    if (uSDDIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered")) { }
                    else
                    {
                        string blnresult = "";
                        for (int j = 1; j < 6; j++)
                        {
                            Section section = doc.Sections[j];
                            foreach (Paragraph para in section.Paragraphs)
                            {
                                DocumentObject obj = null;
                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                {
                                    obj = para.ChildObjects[i];
                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                    {
                                        TextRange textRange = obj as TextRange;
                                        string abc = textRange.Text;
                                        if (abc.ToString().Equals("United Kingdom’s Financial Conduct Authority"))
                                        {
                                            textRange.CharacterFormat.Italic = true;
                                            textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                            doc.SaveToFile(savePath);
                                            blnresult = "true";
                                            break;
                                        }
                                    }
                                }
                                if (blnresult.Equals("true")) { break; }
                            }
                            if (blnresult.Equals("true")) { break; }
                        }
                    }

                    if (uSDDIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered")) { }
                    else
                    {
                        string blnresult = "";
                        for (int j = 1; j < 6; j++)
                        {
                            Section section = doc.Sections[j];
                            foreach (Paragraph para in section.Paragraphs)
                            {
                                DocumentObject obj = null;
                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                {
                                    obj = para.ChildObjects[i];
                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                    {
                                        TextRange textRange = obj as TextRange;
                                        string abc = textRange.Text;
                                        if (abc.ToString().Equals("United States Financial Industry Regulatory Authority"))
                                        {
                                            textRange.CharacterFormat.Italic = true;
                                            textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                            doc.SaveToFile(savePath);
                                            blnresult = "true";
                                            break;
                                        }
                                    }
                                }
                                if (blnresult.Equals("true")) { break; }
                            }
                            if (blnresult.Equals("true")) { break; }
                        }
                    }

                    if (uSDDIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered")) { }
                    else
                    {
                        string blnresult = "";
                        for (int j = 1; j < 6; j++)
                        {
                            Section section = doc.Sections[j];
                            foreach (Paragraph para in section.Paragraphs)
                            {
                                DocumentObject obj = null;
                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                {
                                    obj = para.ChildObjects[i];
                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                    {
                                        TextRange textRange = obj as TextRange;
                                        string abc = textRange.Text;
                                        if (abc.ToString().Equals("United States National Futures Association"))
                                        {
                                            textRange.CharacterFormat.Italic = true;
                                            textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                            doc.SaveToFile(savePath);
                                            blnresult = "true";
                                            break;
                                        }
                                    }
                                }
                                if (blnresult.Equals("true")) { break; }
                            }
                            if (blnresult.Equals("true")) { break; }
                        }
                    }
                    holdresult = "";
                    //string holdresult = "";
                    for (int j = 1; j < 6; j++)
                    {
                        Section section = doc.Sections[j];
                        foreach (Paragraph para in section.Paragraphs)
                        {
                            DocumentObject obj = null;
                            for (int i = 0; i < para.ChildObjects.Count; i++)
                            {
                                obj = para.ChildObjects[i];
                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                {
                                    TextRange textRange = obj as TextRange;
                                    string abc = textRange.Text;
                                    if (abc.ToString().Equals("Other Professional Licensures and/or Designations") || abc.ToString().Equals("Licensure and/or Professional Designations"))
                                    {
                                        textRange.CharacterFormat.Italic = true;
                                        textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                        doc.SaveToFile(savePath);
                                        holdresult = "true";
                                        break;
                                    }
                                }
                            }
                            if (holdresult.Equals("true")) { break; }
                        }
                        if (holdresult.Equals("true")) { break; }
                    }
                    holdresult = "";
                    if (hksfccommentmodel.ToString().Equals("")) { }
                    else
                    {
                        for (int j = 1; j < 6; j++)
                        {
                            Section section = doc.Sections[j];
                            foreach (Paragraph para in section.Paragraphs)
                            {
                                DocumentObject obj = null;
                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                {
                                    obj = para.ChildObjects[i];
                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                    {
                                        TextRange textRange = obj as TextRange;
                                        string abc = textRange.Text;
                                        if (abc.ToString().Equals("Hong Kong Securities and Futures Commission"))
                                        {
                                            textRange.CharacterFormat.Italic = true;
                                            textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                            doc.SaveToFile(savePath);
                                            holdresult = "true";
                                            break;
                                        }
                                    }
                                }
                                if (holdresult.Equals("true")) { break; }
                            }
                        }
                    }

                    //Regulatory_Red_Flag                    
                    if (uSDDIndividual.otherdetails.RegulatoryFlag.ToString().Equals("Yes"))
                    {
                        doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", "In addition to the above, national financial institutions sanctions and legal actions were searched covering  the banking, mortgage and securities industries,", true, true);
                    }
                    else
                    {
                        doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", "In addition to the above, national financial institutions sanctions and legal actions were searched covering the banking, mortgage and securities industries,", true, true);
                    }
                    doc.SaveToFile(savePath);
                    string blnredresultfound = "";
                    for (int j = 2; j < 8; j++)
                    {
                        Section section = doc.Sections[j];
                        foreach (Paragraph para in section.Paragraphs)
                        {
                            DocumentObject obj = null;
                            for (int i = 0; i < para.ChildObjects.Count; i++)
                            {
                                obj = para.ChildObjects[i];
                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                {
                                    TextRange textRange = obj as TextRange;
                                    string abc = textRange.Text;
                                    // Find the word "Spire.Doc" in paragraph1
                                    if (abc.ToString().EndsWith("mortgage and securities industries,"))
                                    {
                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                        //Insert footnote1 after the word "Spire.Doc"
                                        para.ChildObjects.Insert(i + 1, footnote1);
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("Searches included sanctions and/or legal actions of the following federal agencies: the Securities and Exchange Commission; the Financial Industry Regulatory Authority (formerly the National Association of Securities Dealers) and the New York Stock Exchange’s member regulation, enforcement and arbitration functions (prior to the consolidation of the two in July 2007); the Federal Deposit Insurance Corporation; Resolution Trust Corporation; Federal Reserve Board; National Credit Union Administrative Actions; Office of the Comptroller of the Currency; Department of Justice; Department of Housing and Urban Development; and other federal agencies (through the General Services Administration).");
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        //Append the line
                                        string strredflagtextappended = "";
                                        if (uSDDIndividual.otherdetails.RegulatoryFlag.Equals("Yes"))
                                        {
                                            strredflagtextappended = " and the following information was identified in connection with [LastName]:   <Investigator to insert results here>";
                                        }
                                        else
                                        {
                                            strredflagtextappended = " and it is noted that [LastName] was not identified in any of these records.";
                                        }
                                        strredflagtextappended = strredflagtextappended.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName);
                                        TextRange tr = para.AppendText(strredflagtextappended);
                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                        tr.CharacterFormat.FontSize = 11;
                                        doc.SaveToFile(savePath);
                                        blnredresultfound = "true";
                                        break;
                                    }

                                }
                            }
                            if (blnredresultfound == "true")
                            {
                                break;
                            }
                        }
                        if (blnredresultfound == "true")
                        {
                            break;
                        }
                    }
                    //AdditionalStates
                    string add_states = "";
                    string add_states2 = "";
                    string strAdditionalStatesComment = "";
                    string arr = "";
                    string state1 = uSDDIndividual.diligenceInputModel.CurrentState.ToString();
                    //string state2 = uSDDIndividual.diligenceInputModel.Employer1State.ToString();
                    int statecount = 0;
                    try
                    {
                        if (uSDDIndividual.additional_States.Count > 0)
                        {
                            if (uSDDIndividual.additional_States[0].additionalstate.ToUpper().ToString().Equals("Select State".ToUpper()) && uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                            {
                                doc.Replace("[ADDITIONALSTATES]", "", true, true);
                                add_states = "<not provides>";
                            }

                            if (state1 == "Select State") { }
                            else
                            {
                                add_states = string.Concat("<investigator to add relevant county/counties>, ", uSDDIndividual.diligenceInputModel.CurrentState.ToString());
                                arr = uSDDIndividual.diligenceInputModel.CurrentState.ToString();
                                statecount = statecount + 1;
                            }
                            //if (state2 == "Select State" || state2 == state1) { }
                            //else
                            //{
                            //    add_states2 = string.Concat("<investigator to add relevant county/counties>, ", uSDDIndividual.diligenceInputModel.Employer1State.ToString());
                            //    arr = string.Concat(arr, ", ", uSDDIndividual.diligenceInputModel.Employer1State.ToString());
                            //    statecount = statecount + 1;
                            //}

                            if (uSDDIndividual.additional_States[0].additionalstate.ToUpper().ToString().Equals("Select State".ToUpper()))
                            {
                                if (add_states2 == "") { }
                                else
                                {
                                    add_states = string.Concat(add_states, " and ", add_states2);
                                }
                                doc.Replace("[ADDITIONALSTATES]", add_states, true, true);
                            }
                            else
                            {
                                if (add_states2 == "") { }
                                else
                                {
                                    add_states = string.Concat(add_states, ", ", add_states2);
                                }
                                strAdditionalStatesComment = add_states;
                            }
                            if (uSDDIndividual.additional_States[0].additionalstate.ToUpper().ToString().Equals("Select State".ToUpper()))
                            {
                            }
                            else
                            {
                                if (uSDDIndividual.additional_States.Count > 1)
                                {
                                    int count = uSDDIndividual.additional_States.Count();
                                    int i = 1;
                                    for (int j = 0; j < count; j++)
                                    {
                                        if (uSDDIndividual.additional_States[j].additionalstate.ToUpper().ToString().Equals("Select State".ToUpper()))
                                        {
                                        }
                                        else
                                        {
                                            if (j == 0)
                                            {
                                                if (strAdditionalStatesComment == "") { } else { strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, ", "); }
                                            }
                                            if (i == count - 1)
                                            {
                                                strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, "<investigator to add relevant county/counties>, ", uSDDIndividual.additional_States[j].additionalstate, " and ");
                                                add_states = strAdditionalStatesComment;
                                            }
                                            else
                                            {
                                                strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, "<investigator to add relevant county/counties>, ", uSDDIndividual.additional_States[j].additionalstate, ", ");
                                                if (i == count) { add_states = string.Concat(add_states, "<investigator to add relevant county/counties>, ", uSDDIndividual.additional_States[j].additionalstate); }
                                                else
                                                {
                                                    add_states = strAdditionalStatesComment;
                                                }
                                            }
                                            arr = string.Concat(arr, ", ", uSDDIndividual.additional_States[j].additionalstate.ToString());
                                            statecount = statecount + 1;
                                        }
                                        i++;
                                    }
                                }
                                else
                                {
                                    if (add_states == "") { } else { add_states = string.Concat(add_states, " and "); }
                                    strAdditionalStatesComment = string.Concat(add_states, "<investigator to add relevant county/counties>, ", uSDDIndividual.additional_States[0].additionalstate.ToString());
                                    add_states = strAdditionalStatesComment;
                                    arr = string.Concat(arr, ", ", uSDDIndividual.additional_States[0].additionalstate.ToString());
                                    statecount = statecount + 1;
                                }
                                doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, false);
                            }
                        }
                        else
                        {
                            if (uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                            {
                                doc.Replace("[ADDITIONALSTATES]", "", true, true);
                                add_states = "<not provides>";
                            }

                            if (state1 == "Select State") { }
                            else
                            {
                                add_states = string.Concat("<investigator to add relevant county/counties>, ", uSDDIndividual.diligenceInputModel.CurrentState.ToString());
                                arr = uSDDIndividual.diligenceInputModel.CurrentState.ToString();
                                statecount = statecount + 1;
                            }
                            //if (state2 == "Select State" || state2 == state1) { }
                            //else
                            //{
                            //    add_states2 = string.Concat("<investigator to add relevant county/counties>, ", uSDDIndividual.diligenceInputModel.Employer1State.ToString());
                            //    arr = string.Concat(arr, ", ", uSDDIndividual.diligenceInputModel.Employer1State.ToString());
                            //    statecount = statecount + 1;
                            //}

                            if (add_states2 == "") { }
                            else
                            {
                                add_states = string.Concat(add_states, " and ", add_states2);
                            }
                            doc.Replace("[ADDITIONALSTATES]", add_states, true, true);
                            strAdditionalStatesComment = add_states;
                        }
                    }
                    catch
                    {

                    }
                    String[] arrlist = arr.Split(", ");
                    //ADDITIONALJURIDICTIONS
                    if (uSDDIndividual.diligenceInputModel.Nonscopecountry1.ToString().ToUpper().Equals("NA") || uSDDIndividual.diligenceInputModel.Nonscopecountry1.ToString().ToUpper().Equals("") || uSDDIndividual.diligenceInputModel.Nonscopecountry1.ToString().ToUpper().Equals("N/A"))
                    {
                        doc.Replace("ADDITIONALJURIDICTIONS", "", true, true);
                    }
                    else
                    {
                        if (uSDDIndividual.diligenceInputModel.Nonscopecountry1.Contains(", "))
                        {
                            doc.Replace("ADDITIONALJURIDICTIONS", "\nAdditionally, the subject has historical and/or possible ties to [ADDITIONALJURISDESC], and additional research would be required in these jurisdictions, which can be undertaken upon request -- if an expanded scope is warranted. <investigator to alter language to account for only one jurisdiction -- as needed>\n", true, true);
                        }
                        else
                        {
                            doc.Replace("ADDITIONALJURIDICTIONS", "\nAdditionally, the subject has historical and/or possible ties to [ADDITIONALJURISDESC], and additional research would be required in this jurisdiction, which can be undertaken upon request -- if an expanded scope is warranted. <investigator to alter language to account for multiple jurisdictions as needed>\n", true, true);
                        }
                        doc.SaveToFile(savePath);
                        doc.Replace("[ADDITIONALJURISDESC]", uSDDIndividual.diligenceInputModel.Nonscopecountry1, true, true);
                    }
                    doc.SaveToFile(savePath);

                    //Real Estate
                    if (uSDDIndividual.otherdetails.CurrentResidentialProperty.Equals("No") && uSDDIndividual.otherdetails.OtherPropertyOwnershipinfo.Equals("No") && uSDDIndividual.otherdetails.PrevPropertyOwnershipRes.Equals("No") && uSDDIndividual.otherdetails.OtherCurrentResidentialProperty.Equals("No"))
                    {
                        doc.Replace("REALCOMMENT", "", true, true);
                        doc.Replace("REALRESULT", "No Records", true, true);
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.OtherCurrentResidentialProperty.Equals("Yes, multiple") || uSDDIndividual.otherdetails.PrevPropertyOwnershipRes.Equals("Yes, multiple") || uSDDIndividual.otherdetails.CurrentResidentialProperty.ToString().Equals("Yes") && uSDDIndividual.otherdetails.OtherPropertyOwnershipinfo.ToString().Equals("Yes") || uSDDIndividual.otherdetails.CurrentResidentialProperty.ToString().Equals("Yes") && uSDDIndividual.otherdetails.PrevPropertyOwnershipRes.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.OtherPropertyOwnershipinfo.ToString().Equals("Yes") && uSDDIndividual.otherdetails.PrevPropertyOwnershipRes.ToString().StartsWith("Yes"))
                        {
                            doc.Replace("REALCOMMENT", "[LastName] was identified as the owner of properties located in <Investigator to insert counties and states>", true, true);
                            doc.Replace("REALRESULT", "Records", true, true);
                        }
                        else
                        {
                            doc.Replace("REALCOMMENT", "[LastName] was identified as the owner of a property located in <investigator to insert county, state>", true, true);
                            doc.Replace("REALRESULT", "Record", true, true);
                        }
                    }


                    //CURRENTRESIDENTIALPROPERTYDESC
                    if (uSDDIndividual.otherdetails.CurrentResidentialProperty.ToString().Equals("Yes"))
                    {
                        doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "Current ownership records are identified in connection with a certain property located at [CurrentFullAddress].  According to <County> County, [CurrentState1] property records, the subject purchased this property for <purchase price> from <seller names> on <date>, <Investigator to insert other mortgage information or UCC details here>. This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>.[ADDOOFOOTNOTE]\n", true, true);
                    }
                    else
                    {

                        doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "While [LastName] currently resides at [CurrentFullAddress], the subject was not identified as the owner of the same. <Investigator to insert property record or rental information>.CURRENTRESIDENTIALPROPERTYDESC", true, true);
                    }
                    doc.SaveToFile(savePath);
                    if (uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                    {
                        doc.Replace("[CurrentState1]", "<State>", true, true);
                    }
                    else
                    {
                        doc.Replace("[CurrentState1]", uSDDIndividual.diligenceInputModel.CurrentState.ToString(), true, true);
                    }
                    doc.SaveToFile(savePath);
                    blnredresultfound = "";
                    for (int j = 2; j < 6; j++)
                    {
                        Section section = doc.Sections[j];
                        foreach (Paragraph para in section.Paragraphs)
                        {
                            DocumentObject obj = null;
                            for (int i = 0; i < para.ChildObjects.Count; i++)
                            {
                                obj = para.ChildObjects[i];
                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                {
                                    TextRange textRange = obj as TextRange;
                                    string abc = textRange.Text;
                                    // Find the word "Spire.Doc" in paragraph1
                                    if (abc.ToString().EndsWith("[ADDOOFOOTNOTE]") || abc.ToString().Contains("[ADDOOFOOTNOTE]"))
                                    {
                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                        //footnote1.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                                        //Insert footnote1 after the word "Spire.Doc"
                                        para.ChildObjects.Insert(i + 1, footnote1);
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("It is noted that assessment and/or market values may be reflected as a percentage, or on a ratio basis, of the total value of the property, and are not necessarily reflective of true market value.");
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        TextRange tr = para.AppendText("CURRENTRESIDENTIALPROPERTYDESC");
                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                        tr.CharacterFormat.FontSize = 11;
                                        blnredresultfound = "true";

                                        doc.SaveToFile(savePath);
                                        break;
                                    }

                                }
                            }
                            if (blnredresultfound == "true")
                            {
                                break;
                            }
                        }
                        if (blnredresultfound == "true")
                        {
                            break;
                        }
                    }
                    doc.SaveToFile(savePath);
                    //OtherCurrentResidentialProperty
                    try
                    {
                        if (uSDDIndividual.otherdetails.OtherCurrentResidentialProperty.ToString().Equals("Yes, only one"))
                        {
                            doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "\n\nCurrent ownership records are also identified in connection with a certain property located at <Investigator to add other current property address>. According to <County> County, <State> property records, the subject purchased this property for <purchase price> from <seller names> on <date>, <Investigator to insert other mortgage information or UCC details here>. This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>.[ADDOOFOOT1NOTE]\n", true, false);
                        }
                        else
                        {
                            if (uSDDIndividual.otherdetails.OtherCurrentResidentialProperty.ToString().Equals("Yes, multiple"))
                            {
                                doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "\n\nIn addition, current ownership records also identified [LastName] in connection with the following properties, which are located in <Investigator to add all relevant County, State(s)>:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>", true, false);
                            }
                            else
                            {
                                doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "", true, false);
                            }
                        }
                        doc.SaveToFile(savePath);
                        blnredresultfound = "";
                        for (int j = 2; j < 6; j++)
                        {
                            Section section = doc.Sections[j];
                            foreach (Paragraph para in section.Paragraphs)
                            {
                                DocumentObject obj = null;
                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                {
                                    obj = para.ChildObjects[i];
                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                    {
                                        TextRange textRange = obj as TextRange;
                                        string abc = textRange.Text;
                                        // Find the word "Spire.Doc" in paragraph1
                                        if (abc.ToString().EndsWith("[ADDOOFOOT1NOTE]") || abc.ToString().Contains("[ADDOOFOOT1NOTE]"))
                                        {
                                            Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                            footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                            footnote1.MarkerCharacterFormat.FontSize = 11;
                                            //footnote1.MarkerCharacterFormat.SubSuperScript = SubSuperScript.SuperScript;
                                            //Insert footnote1 after the word "Spire.Doc"
                                            para.ChildObjects.Insert(i + 1, footnote1);
                                            TextRange _text = footnote1.TextBody.AddParagraph().AppendText("It is noted that assessment and/or market values may be reflected as a percentage, or on a ratio basis, of the total value of the property, and are not necessarily reflective of true market value.");
                                            footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                            _text.CharacterFormat.FontName = "Calibri (Body)";
                                            _text.CharacterFormat.FontSize = 9;
                                            blnredresultfound = "true";

                                            doc.SaveToFile(savePath);
                                            doc.Replace("[ADDOOFOOT1NOTE]", "", false, false);
                                            doc.SaveToFile(savePath);
                                            break;
                                        }

                                    }
                                }
                                if (blnredresultfound == "true")
                                {
                                    break;
                                }
                            }
                            if (blnredresultfound == "true")
                            {
                                break;
                            }
                        }
                        doc.SaveToFile(savePath);
                    }
                    catch
                    {
                        doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "", true, false);
                        doc.SaveToFile(savePath);
                    }
                    //OTHERPROPOWNERSHIPDESC
                    if (uSDDIndividual.otherdetails.OtherPropertyOwnershipinfo.ToString().Equals("Yes"))
                    {
                        doc.Replace("OTHERPROPOWNERSHIPDESC", "\nCurrent ownership records are also identified in connection with a certain property located at <address>.  According to <County> County, <State> property records, the subject purchased this property for <purchase price> from <seller names> on <date>. <Investigator to insert other mortgage information or UCC details here>. This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>.\n", true, true);
                    }
                    else
                    {
                        doc.Replace("OTHERPROPOWNERSHIPDESC", "", true, true);
                    }
                    //PREVIOUSPROPERTYOWNERSHIPDESC
                    if (uSDDIndividual.otherdetails.PrevPropertyOwnershipRes.ToString().Equals("No"))
                    {
                        doc.Replace("PREVIOUSPROPERTYOWNERSHIPDESC", "", true, true);
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.PrevPropertyOwnershipRes.ToString().Equals("Yes, only one"))
                        {
                            doc.Replace("PREVIOUSPROPERTYOWNERSHIPDESC", "\nPrevious ownership records were identified in connection with a certain property located at <address>. According to <County> County, <State> property records, the subject sold this property for <sale price> to <buyer names> on <date>.\n", true, true);
                        }
                        else
                        {
                            doc.Replace("PREVIOUSPROPERTYOWNERSHIPDESC", "\nFurther, previous ownership records were identified in connection with the following properties, which are located in <add Counties/States>:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n", true, true);
                        }
                    }
                    //REGISTEREDWITHHKSFCDESC
                    if (uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Not Registered"))
                    {
                        doc.Replace("[HKSFCHEAD]", "", true, false);
                        doc.Replace("[REGISTEREDWITHHKSFCDESC]", "", true, false);
                        doc.Replace("KONSECCOMMENT", "", true, true);
                        doc.Replace("KONSECRESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes - With Adverse"))
                        {
                            doc.Replace("KONSECCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("KONSECRESULT", "Records", true, true);
                            doc.Replace("[HKSFCHEAD]", "\nHong Kong Securities and Futures Commission", true, false);
                            doc.Replace("[REGISTEREDWITHHKSFCDESC]", "\n\n[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nAdditionally, the following public disciplinary actions were on file in connection with the subject:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>", true, false);
                        }
                        else
                        {
                            doc.Replace("KONSECCOMMENT", "", true, true);
                            doc.Replace("KONSECRESULT", "Clear", true, true);
                            doc.Replace("[HKSFCHEAD]", "\nHong Kong Securities and Futures Commission", true, false);
                            doc.Replace("[REGISTEREDWITHHKSFCDESC]", "\n\n[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nThere are no public disciplinary actions on file against the subject within the past five years.\n\n", true, false);
                        }
                    }
                    //COMMONNAMESUBDESC
                    if (uSDDIndividual.diligenceInputModel.CommonNameSubject == false)
                    {
                        doc.Replace("COMMONNAMESUBDESC", " Additionally, it is noted that searches of the United States District and Bankruptcy Courts were nearly national in scope.", true, true);
                    }
                    else
                    {
                        doc.Replace("COMMONNAMESUBDESC", "", true, true);
                    }
                    //HASBANKRUPTYRECHITDESC
                    string strbankrupfootnt1 = "";
                    string strconcatBankrupty = "";
                    strAdditionalStatesComment = "";
                    string strbank1 = "";
                    string strbanksumcomm = "";
                    string strbankressumm = "";
                    if (uSDDIndividual.otherdetails.HasBankruptcyRecHits_resultpending == true)
                    {
                        doc.Replace("HASBANKRUPTYRECHITDESC", "Bankruptcy court records searches are currently pending in <jurisdiction> the results of which will be provided under separate cover upon receipt.[BANKFONTCHANGE]\n\nHASBANKRUPTYRECHITDESC", true, true);
                        doc.SaveToFile(savePath);
                    }
                    if (uSDDIndividual.otherdetails.HasBankruptcyRecHits.ToString().Equals("Yes, Multiple Records"))
                    {
                        strbank1 = "Records";
                        doc.Replace("HASBANKRUPTYRECHITDESC", "Efforts pursued in [ADDITIONALBANKSTATES]ADDFOOTNOTE", true, true);
                        strconcatBankrupty = " identified the subject, personally, as a <party type> in connection with the following bankruptcy filings:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n";
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.HasBankruptcyRecHits.ToString().Equals("No"))
                        {
                            strbank1 = "Clear";
                            doc.Replace("HASBANKRUPTYRECHITDESC", "Efforts pursued in [ADDITIONALBANKSTATES]ADDFOOTNOTE", true, true);
                            strconcatBankrupty = " did not identify the subject, personally, in connection with any bankruptcy filings.\n";
                        }
                        else
                        {
                            strbank1 = "Record";
                            doc.Replace("HASBANKRUPTYRECHITDESC", "Efforts pursued in [ADDITIONALBANKSTATES]ADDFOOTNOTE", true, true);
                            strconcatBankrupty = " identified the subject, personally, as a <party type> in connection with the following bankruptcy filings:\n\n\t•\t<Investigator to insert bulleted list of results here>\n";
                        }
                    }
                    if (uSDDIndividual.diligenceInputModel.CommonNameSubject == true)
                    {
                        strconcatBankrupty = string.Concat(strconcatBankrupty, "\nIn light of the commonality of [LastName]’s name, this portion of the investigation was conducted utilizing the subject’s Social Security number, in order to identify only those records conclusively relating to the subject of interest. <Investigator to amend or remove if applicable>\n");
                    }
                    if (uSDDIndividual.otherdetails.HasBankruptcyRecHits1 == true)
                    {
                        strconcatBankrupty = string.Concat(strconcatBankrupty, "\nIt should be noted that one or more individuals known only as “[First Name] [Last Name]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>. <Investigator to add other relevant details (i.e. relevant bank cases and/or dispositions)>\n\nManual bankruptcy court records research efforts -- as available -- would be required to determine whether the same relate to the subject of this investigation, which can be undertaken upon request -- if an expanded scope is warranted.\n");
                    }

                    if (arrlist.Count() == 0)
                    {
                        doc.Replace("[ADDITIONALBANKSTATES]", "(<not provided>)", true, false);
                    }
                    else
                    {
                        if (arrlist.Count() > 1)
                        {
                            int count = arrlist.Count();
                            int i = 1;
                            strbankrupfootnt1 = "Searches were pursued through the United States Bankruptcy Courts for the [STATESPECIFICDISTRICTCOURT].";
                            string strstatecourt = "";
                            for (int j = 0; j < arrlist.Count(); j++)
                            {
                                if (arrlist[j].ToString().Equals("Select State".ToUpper()))
                                {
                                }
                                else
                                {
                                    StateWiseFootnoteModel comment1 = _context.stateModel
                                     .Where(u => u.states.ToUpper().TrimEnd() == arrlist[j].TrimEnd().ToUpper())
                                     .FirstOrDefault();
                                    string strstate = "";
                                    try
                                    {
                                        if (comment1.state_specific_district_courts.ToString().Equals(""))
                                        {
                                            strstate = "<not provided>";
                                        }
                                        else
                                        {
                                            strstate = comment1.state_specific_district_courts;
                                        }
                                    }
                                    catch
                                    {
                                        strstate = "<not available>";
                                    }
                                    if (i == count - 1)
                                    {
                                        strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, arrlist[j], " and ");
                                        strstatecourt = string.Concat(strstatecourt, strstate, " and ");
                                    }
                                    else
                                    {
                                        if (i == count)
                                        {
                                            strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, arrlist[j]);
                                            strstatecourt = string.Concat(strstatecourt, strstate);
                                        }
                                        else
                                        {
                                            strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, arrlist[j], ", ");
                                            strstatecourt = string.Concat(strstatecourt, strstate, ", ");
                                        }
                                    }
                                }
                                i++;
                            }
                            strstatecourt = strstatecourt.Replace("Courts for ", "");
                            strstatecourt = strstatecourt.Replace("Court for ", "");
                            strbankrupfootnt1 = strbankrupfootnt1.Replace("[STATESPECIFICDISTRICTCOURT]", strstatecourt);
                        }
                        else
                        {
                            string strstatecourt = "";
                            StateWiseFootnoteModel comment1 = _context.stateModel
                                     .Where(u => u.states.ToUpper().TrimEnd() == arrlist[0].ToString().ToUpper().TrimEnd())
                                     .FirstOrDefault();
                            try
                            {
                                if (comment1.state_specific_district_courts.ToString().Equals("")) { strstatecourt = "<not provided>"; }
                                else
                                {
                                    strstatecourt = comment1.state_specific_district_courts;
                                }
                            }
                            catch
                            {
                                strstatecourt = "<not available>";
                            }
                            strbankrupfootnt1 = "Searches were pursued through the United States Bankruptcy [STATESPECIFICDISTRICTCOURT].";
                            strbankrupfootnt1 = strbankrupfootnt1.Replace("[STATESPECIFICDISTRICTCOURT]", strstatecourt);
                            strAdditionalStatesComment = arrlist[0].ToString();
                        }
                    }

                    doc.Replace("[ADDITIONALBANKSTATES]", strAdditionalStatesComment, true, false);
                    doc.SaveToFile(savePath);
                    strblnres = "";
                    for (int j = 1; j < 7; j++)
                    {
                        Section section = doc.Sections[j];
                        foreach (Paragraph para in section.Paragraphs)
                        {
                            DocumentObject obj = null;
                            for (int i = 0; i < para.ChildObjects.Count; i++)
                            {
                                obj = para.ChildObjects[i];
                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                {
                                    TextRange textRange = obj as TextRange;
                                    string abc = textRange.Text;
                                    // Find the word "Spire.Doc" in paragraph1
                                    if (abc.ToString().Equals("ADDFOOTNOTE") || abc.ToString().EndsWith("ADDFOOTNOTE") || abc.ToString().Contains("ADDFOOTNOTE"))
                                    {
                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                        //Insert footnote1 after the word "Spire.Doc"
                                        para.ChildObjects.Insert(i + 1, footnote1);
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strbankrupfootnt1);
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        TextRange tr = para.AppendText(strconcatBankrupty);
                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                        tr.CharacterFormat.FontSize = 11;
                                        doc.SaveToFile(savePath);
                                        strblnres = "true";
                                        break;
                                    }
                                }
                            }
                            if (strblnres == "true")
                            {
                                break;
                            }
                        }
                        if (strblnres == "true")
                        {
                            break;
                        }
                    }
                    doc.SaveToFile(savePath);
                    doc.Replace("ADDFOOTNOTE", "", true, false);
                    doc.Replace("[ADDITIONALSTATES]", "", true, false);
                    doc.SaveToFile(savePath);
                    string strbankrs = "";
                    string strbankpr = "";
                    string strbankresrs = "";
                    string strbankrespr = "";
                    if (uSDDIndividual.otherdetails.HasBankruptcyRecHits_resultpending == true)
                    {
                        strbankresrs = "Bankruptcy court records searches are currently pending in <jurisdiction> the results of which will be provided under separate cover upon receipt[BANKRSITALIC]\n\n";
                        strbankrs = "Results Pending\n\n";
                    }
                    if (uSDDIndividual.otherdetails.HasBankruptcyRecHits1 == true)
                    {
                        strbankpr = "\n\nOne or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> <case type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                        strbankrespr = "\n\nPossible Records";
                    }
                    switch (strbank1)
                    {
                        case "Clear":
                            strbankressumm = "Clear";
                            strbanksumcomm = "";
                            break;
                        case "Record":
                            strbanksumcomm = "[LastName] was identified as having been a <party type> in connection with a <case type>, which was recorded in <State> in <YYYY>, and is currently <status>";
                            strbankressumm = "Record";
                            break;
                        case "Records":
                            strbanksumcomm = "[LastName] was identified as a <party type> in connection with at least <number> <case type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                            strbankressumm = "Records";
                            break;
                    }
                    strbanksumcomm = string.Concat(strbankresrs, strbanksumcomm.TrimEnd(), strbankpr);
                    strbankressumm = string.Concat(strbankrs, strbankressumm, strbankrespr);
                    doc.Replace("BANKRUPCOMMENT", strbanksumcomm, true, true);
                    doc.Replace("BANKRUPRESULT", strbankressumm, true, true);
                    doc.SaveToFile(savePath);
                    Table table2 = doc.Sections[1].Tables[0] as Table;
                    TableCell cell3 = table2.Rows[12].Cells[2];
                    TableCell cell4 = table2.Rows[12].Cells[1];
                    foreach (Paragraph p1 in cell3.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1
                                if (abc.ToString().EndsWith("[BANKRSITALIC]"))
                                {
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                    doc.SaveToFile(savePath);
                                    doc.Replace("[BANKRSITALIC]", "", true, false);
                                    doc.SaveToFile(savePath);
                                    break;
                                }
                            }
                        }
                    }

                    foreach (Paragraph p1 in cell4.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1
                                if (abc.ToString().EndsWith("Result Pending"))
                                {
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    doc.SaveToFile(savePath);
                                    break;
                                }
                                else
                                {
                                    if (abc.ToString().EndsWith("Possible Records"))
                                    {
                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                        textRange.CharacterFormat.FontSize = 11;
                                        textRange.CharacterFormat.Italic = true;
                                        doc.SaveToFile(savePath);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //HASBUREAUPRISONHITDESC
                    if (uSDDIndividual.otherdetails.Has_Bureau_PrisonHit.ToString().Equals("Yes"))
                    {
                        doc.Replace("HASBUREAUPRISONHITDESC", "\nIn addition, research efforts are conducted of Federal Bureau of Prisons (“BOP”) incarceration records, which reports felony inmate records on a nearly-nationwide basis.  In this regard, the following information was identified in connection with [LastName]:  <Investigator to insert results here>\n", true, true);
                    }
                    else
                    {
                        doc.Replace("HASBUREAUPRISONHITDESC", "\nIn addition, research efforts are conducted of Federal Bureau of Prisons (“BOP”) incarceration records, which reports felony inmate records on a nearly-nationwide basis.  In this regard, no such records were revealed relating to the subject.\n", true, true);
                    }
                    //HASSEXOFFENDERREGHITDESC
                    if (uSDDIndividual.otherdetails.Has_Sex_Offender_RegHit.ToString().Equals("Yes"))
                    {
                        doc.Replace("HASSEXOFFENDERREGHITDESC", "Research efforts were further conducted through the United States Department of Justice’s National Sex Offender Registry, which contains sex offender registry information from the 50 states, the District of Columbia and the territories of Guam and Puerto Rico. The public availability of sex offender-related information varies by state. In this regard, the following information was identified in connection with [LastName]:  <Investigator to insert results here>\n", true, true);
                    }
                    else
                    {
                        doc.Replace("HASSEXOFFENDERREGHITDESC", "Research efforts were further conducted through the United States Department of Justice’s National Sex Offender Registry, which contains sex offender registry information from the 50 states, the District of Columbia and the territories of Guam and Puerto Rico. The public availability of sex offender-related information varies by state. No records were identified for the subject.\n", true, true);
                    }
                    //HASNAMECIVILLITIDESC 
                    if (uSDDIndividual.otherdetails.Has_Name_Only == false)
                    {
                        doc.Replace("[HASNAMECIVILLITIDESC]", "", true, false);
                    }
                    else
                    {
                        doc.Replace("[HASNAMECIVILLITIDESC]", "\n\nIt should be noted that one or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>[ADDNAMEONLYFN]\n\nManual civil court records research efforts – as available – would be required to determine whether the same relate to the subject of this investigation, which can be undertaken upon request – if an expanded scope is warranted.", true, false);
                    }
                    doc.SaveToFile(savePath);
                    strblnres = "";
                    for (int j = 2; j < 5; j++)
                    {
                        Section section = doc.Sections[j];
                        foreach (Paragraph para in section.Paragraphs)
                        {
                            DocumentObject obj = null;
                            for (int i = 0; i < para.ChildObjects.Count; i++)
                            {
                                obj = para.ChildObjects[i];
                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                {
                                    TextRange textRange = obj as TextRange;
                                    string abc = textRange.Text;
                                    // Find the word "Spire.Doc" in paragraph1
                                    if (abc.ToString().Equals("[ADDNAMEONLYFN]") || abc.ToString().EndsWith("[ADDNAMEONLYFN]") || abc.ToString().Contains("[ADDNAMEONLYFN]"))
                                    {
                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                        //Insert footnote1 after the word "Spire.Doc"
                                        para.ChildObjects.Insert(i + 1, footnote1);
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("<investigator to insert details here>");
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        doc.SaveToFile(savePath);
                                        strblnres = "true";
                                        break;
                                    }
                                }
                            }
                            if (strblnres == "true")
                            {
                                break;
                            }
                        }
                        if (strblnres == "true")
                        {
                            break;
                        }
                    }
                    doc.SaveToFile(savePath);

                    //USTAXCOURTHITDESC  [ADDNAMEONLYFN] 
                    if (uSDDIndividual.otherdetails.Has_US_Tax_Court_Hit.ToString().StartsWith("Yes"))
                    {
                        doc.Replace("[USTAXCOURTHITDESC]", "\n\nAdditionally, research efforts were conducted of filings through the United States Tax Court, and the following information was identified in connection with [LastName]:  <Investigator to insert results here>\n", true, false);
                    }
                    else
                    {
                        doc.Replace("[USTAXCOURTHITDESC]", "\n\nAdditionally, research efforts were conducted of filings through the United States Tax Court, and no records were identified for the subject.\n", true, false);
                    }
                    //HASTAXLIENSCIVILUCCDESC
                    string strtaxlien = "";
                    if (uSDDIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("No"))
                    {
                        strtaxlien = "Investigative efforts did not reveal any tax liens in connection with [LastName].";
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("Yes, Multiple Records"))
                        {
                            strtaxlien = "Investigative efforts revealed the subject as having been a debtor in connection with the following tax liens:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                        }
                        else
                        {
                            strtaxlien = "Investigative efforts revealed the subject as having been a debtor in connection with the following tax lien:\n\n\t•	<Investigator to insert bulleted list of results here>";
                        }
                    }
                    //HAS CIVIL_JUDGEMENT
                    string strciviljudge = "";
                    if (uSDDIndividual.otherdetails.Has_Property_Records.Equals("Yes, Single Record"))
                    {
                        strciviljudge = "\n\nAdditionally, investigative efforts revealed the subject as having been a <party type> in connection with the following civil judgment:\n\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_Property_Records.Equals("Yes, Multiple Records"))
                        {
                            strciviljudge = "\n\nAdditionally, investigative efforts revealed the subject in connection with the following civil judgments:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                        }
                        else
                        {
                            strciviljudge = "\n\nAdditionally, investigative efforts did not reveal any civil judgments in connection with [LastName]. <Investigator to combine/modify as needed with other filing types (i.e. any tax liens or judgments).>";
                        }
                    }
                    string strjudge = "";
                    if (uSDDIndividual.otherdetails.Has_Property_Records.Equals("Yes, Single Record") && uSDDIndividual.otherdetails.Has_Tax_Liens.Equals("Yes, Single Record") || uSDDIndividual.otherdetails.Has_Property_Records.Equals("Yes, Multiple Records") || uSDDIndividual.otherdetails.Has_Tax_Liens.Equals("Yes, Multiple Records"))
                    {
                        strjudge = "Records";
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_Property_Records.Equals("Yes, Single Record") || uSDDIndividual.otherdetails.Has_Tax_Liens.Equals("Yes, Single Record"))
                        {
                            strjudge = "Record";
                        }
                        else
                        {
                            strjudge = "Clear";
                        }
                    }
                    string strcommjudge = "";
                    string strresjudge = "";
                    switch (strjudge)
                    {
                        case "Clear":
                            strcommjudge = "";
                            strresjudge = "Clear";
                            break;
                        case "Record":
                            strcommjudge = "[LastName] was identified as a <party type> in connection with a <record type>, which was filed in <State> in <YYYY>, and is currently <status>";
                            strresjudge = "Record";
                            break;
                        case "Records":
                            strcommjudge = "[LastName] was identified as a <party type> in connection with at least <number> <record types>, which were filed in <States> between <YYYY> and <YYYY>, and are currently <status>";
                            strresjudge = "Records";
                            break;
                    }

                    if (uSDDIndividual.otherdetails.Has_Name_Only_Tax_Lien == true)
                    {
                        strresjudge = string.Concat(strresjudge, "\n\nPossible Records");
                        strcommjudge = string.Concat(strcommjudge, "\n\nOne or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>");
                    }

                    doc.Replace("TAXLIENCOMMENT", strcommjudge, true, true);
                    doc.Replace("TAXLIENRESULT", strresjudge, true, true);
                    doc.SaveToFile(savePath);
                    table1 = doc.Sections[1].Tables[0] as Table;
                    cell3 = table1.Rows[14].Cells[2];
                    cell4 = table1.Rows[14].Cells[1];

                    foreach (Paragraph p1 in cell4.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1                        
                                if (abc.ToString().EndsWith("Possible Records"))
                                {
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    doc.SaveToFile(savePath);
                                    break;
                                }
                            }
                        }
                    }
                    //HASNAMETAXLIENUCCDESC
                    string strtaxliennameonly = "";
                    if (uSDDIndividual.otherdetails.Has_Name_Only_Tax_Lien == true)
                    {
                        strtaxliennameonly = "\n\nIt should be noted that one or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.  Manual records research efforts -- as available -- would be required to determine whether the same relate to the subject of this investigation, which can be undertaken upon request -- if an expanded scope is warranted.";
                    }
                    //HASUCC
                    string strucccomment = "";
                    if (uSDDIndividual.otherdetails.has_ucc_fillings.ToString().Equals("Yes, Single Record"))
                    {
                        strucccomment = "\n\nFurther, investigative efforts revealed the subject as having been a <party type> in connection with the following Uniform Commercial Code filing:\n\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.has_ucc_fillings.ToString().Equals("Yes, Multiple Records"))
                        {
                            strucccomment = "\n\nFurther, investigative efforts revealed the subject in connection with the following Uniform Commercial Code filings:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                        }
                        else
                        {
                            strucccomment = "\n\nAdditionally, investigative efforts did not reveal [LastName] in connection with any Uniform Commercial Code filings. <Investigator to combine/modify as needed with other filing types (i.e. any tax liens or UCC filings)>.";
                        }
                    }
                    string commomstr = "";
                    if (uSDDIndividual.diligenceInputModel.CommonNameSubject == true)
                    {
                        commomstr = "\n\nIn light of the commonality of [LastName]’s name, this portion of the investigation was conducted utilizing the subject’s Social Security number, in order to identify only those records conclusively relating to the subject of interest. <Investigator to amend or remove if applicable>";
                    }
                    string strucc1 = "";
                    if (uSDDIndividual.otherdetails.has_ucc_fillings1 == true)
                    {
                        strucc1 = "\n\nIt should be noted that one or more individuals known only as “[First Name] [Last Name]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.  Manual records research efforts -- as available -- would be required to determine whether the same relate to the subject of this investigation, which can be undertaken upon request -- if an expanded scope is warranted.";
                    }
                    if (uSDDIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("No") && uSDDIndividual.otherdetails.Has_Property_Records.Equals("No") && uSDDIndividual.otherdetails.has_ucc_fillings.ToString().Equals("No"))
                    {
                        strciviljudge = "Investigative efforts did not reveal any tax liens, civil judgments or Uniform Commercial Code filings in connection with [LastName].";
                        strciviljudge = string.Concat(strciviljudge, strtaxliennameonly, commomstr, strucc1);
                        doc.Replace("HASNAMETAXLIENUCCDESC", strciviljudge, true, true);
                    }
                    else
                    {
                        strciviljudge = string.Concat(strtaxlien, strciviljudge, strtaxliennameonly, strucccomment, commomstr, strucc1);
                        doc.Replace("HASNAMETAXLIENUCCDESC", strciviljudge, true, true);
                    }
                    string struccsummres = "";
                    string struccsummcom = "";
                    switch (uSDDIndividual.otherdetails.has_ucc_fillings)
                    {
                        case "No":
                            struccsummcom = "";
                            struccsummres = "Clear";
                            break;
                        case "Yes, Single Record":
                            struccsummcom = "[LastName] was identified as a <party type> in connection with a UCC filing, which was recorded in <State> in <YYYY>, and is currently <status>";
                            struccsummres = "Record";
                            break;
                        case "Yes, Multiple Records":
                            struccsummcom = "[LastName] was identified as a <party type> in connection with at least <number> UCC filings, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                            struccsummres = "Records";
                            break;
                    }
                    if (uSDDIndividual.otherdetails.has_ucc_fillings1 == true)
                    {
                        struccsummres = string.Concat(struccsummres, "\n\nPossible Records");
                        struccsummcom = string.Concat(struccsummcom, "\n\nOne or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>");
                    }
                    doc.Replace("UNIFCOMMENT", struccsummcom, true, true);
                    doc.Replace("UNIFRESULT", struccsummres, true, true);
                    doc.SaveToFile(savePath);
                    table1 = doc.Sections[1].Tables[0] as Table;
                    cell3 = table1.Rows[15].Cells[2];
                    cell4 = table1.Rows[15].Cells[1];

                    foreach (Paragraph p1 in cell4.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1                        
                                if (abc.ToString().EndsWith("Possible Records"))
                                {
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    doc.SaveToFile(savePath);
                                    break;
                                }
                            }
                        }
                    }
                    //CREDITHISTORYDESCRIPTION             
                    string credit_sumresult = "";
                    string credit_summcomment = "";
                    if (uSDDIndividual.otherdetails.Was_credited_obtained.ToString().Equals("No"))
                    {
                        doc.Replace("CREDITHISTORYDESCRIPTION", "[LastName]’s express written authorization would be required in order to obtain the subject’s personal credit history information.", true, true);
                        credit_sumresult = "N/A";
                        credit_summcomment = "Credit history information cannot be obtained without the subject’s express written authorization";
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Was_credited_obtained.ToString().Equals("Credit with Adverse Hits"))
                        {
                            doc.Replace("CREDITHISTORYDESCRIPTION", "With [LastName]’s express written authorization, the subject’s personal credit history was retrieved, which revealed that <Investigator to insert summary of trade account information here>.\n\nFurther, <Investigator to insert summary of adverse results here>.\n\nThere were no additional payment delinquencies or collections accounts reported in the subject's credit file for the past seven years or so.", true, true);
                            credit_sumresult = "Records";
                            credit_summcomment = "With the subject’s express written authorization, [LastName]’s personal credit history was retrieved, which revealed <investigator to insert summary here>";
                        }
                        else
                        {
                            doc.Replace("CREDITHISTORYDESCRIPTION", "With [LastName]’s express written authorization, the subject’s personal credit history was retrieved, which revealed that <Investigator to insert summary of trade account information here>.\n\nThere were no payment delinquencies or collections accounts reported in the subject's credit file for the past seven years or so.", true, true);
                            credit_sumresult = "Clear";
                            credit_summcomment = "";
                        }
                    }
                    doc.Replace("PERCREDITRESULT", credit_sumresult, true, true);
                    doc.Replace("PERCREDITCOMMENT", credit_summcomment, true, true);
                    //PRESSANDMEDIASEARCHDESCRIPTION
                    switch (uSDDIndividual.otherdetails.Press_Media.ToString())
                    {
                        case "Common name with adverse Hits":
                            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the commonality of the subject's name, searches were structured and designed to identify specific reports relating to [LastName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. Further, the subject’s name was also searched using a variety of other terms, focusing on the subject’s employment history, and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);
                            break;
                        case "Common name without adverse Hits":
                            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the commonality of the subject's name, searches were structured and designed to identify specific reports relating to [LastName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. Further, the subject’s name was also searched using a variety of other terms, focusing on the subject’s employment history, and a thorough review of the same did not identify any adverse or materially-significant information in connection with [LastName].", true, true);
                            break;
                        case "High volume with adverse Hits":
                            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the volume of press and other media sources identified in connection with the subject, searches were structured and designed to identify specific reports relating to [LastName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. Further, the subject’s name was also searched using a variety of other terms, focusing on the subject’s employment history, and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);
                            break;
                        case "High volume without adverse Hits":
                            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the volume of press and other media sources identified in connection with the subject, searches were structured and designed to identify specific reports relating to [LastName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. Further, the subject’s name was also searched using a variety of other terms, focusing on the subject’s employment history, and a thorough review of the same did not identify any adverse or materially-significant information in connection with [LastName].", true, true);
                            break;
                        case "Standard search with adverse Hits":
                            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, identified various articles and/or media references in connection with [LastName], and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);
                            break;
                        case "Standard search without adverse Hits":
                            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, identified various articles and/or media references in connection with [LastName], and a thorough review of the same did not identify any adverse or materially-significant information.", true, true);
                            break;
                        case "No Hits":
                            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, did not identify any articles and/or media references in connection with [LastName].", true, true);
                            break;
                    }
                    doc.SaveToFile(savePath);
                    //DRIVINGHISTORYDESCRIPTION
                    //DRIVINGRESULT
                    string drivingcomm = "";
                    string divingresult = "";
                    switch (uSDDIndividual.otherdetails.Has_Driving_Hits.ToString())
                    {
                        case "No Records, No Consent":
                            drivingcomm = "\nIt is noted that [Last Name]’s express written authorization would be required to obtain the subject’s driving history information.";
                            divingresult = "N/A [No Consent or hits]";
                            break;
                        case "Clear, With Consent":
                            drivingcomm = "\nIt is noted that [Last Name] has a valid <State> driver’s license, which is set to expire in <expiration date (month and year only)>, unless renewed.  A driving history report was requested from the <State driving agency>, covering the past three years or so, which did not reveal any traffic violations, license suspensions or motor vehicle incidents involving the subject.";
                            divingresult = "Clear";
                            break;
                        case "Multiple Records, No Consent":
                            drivingcomm = "\nIt is noted that [Last Name]’s express written authorization would be required to obtain the subject’s driving history information.\n\nHowever, additional research efforts identified <Investigator to insert results here>.";
                            divingresult = "Records [Hit w/o Consent]";
                            break;
                        case "Single Record, No Consent":
                            drivingcomm = "\nIt is noted that [Last Name]’s express written authorization would be required to obtain the subject’s driving history information.\n\nHowever, additional research efforts identified <Investigator to insert results here>.";
                            divingresult = "Record [Hit w/o Consent]";
                            break;
                        case "Multiple Records, with Consent":
                            drivingcomm = "\nIt is noted that [Last Name] has a valid <State> driver’s license, which is set to expire in <expiration date (month and year only)>, unless renewed.  A driving history report was requested from the <State driving agency>, covering the past three years or so, which revealed that the subject was cited for the following:\n\n\t•\t<Investigator to insert results>\n\t•\t<Investigator to insert results>\n\nThere were no other traffic violations, license suspensions or motor vehicle incidents involving the subject.";
                            divingresult = "Records [Hits w/ Consent]";
                            break;
                        case "Single Record, with Consent":
                            drivingcomm = "\nIt is noted that [Last Name] has a valid <State> driver’s license, which is set to expire in <expiration date (month and year only)>, unless renewed.  A driving history report was requested from the <State driving agency>, covering the past three years or so, which revealed that the subject was cited for <Investigator to insert results>.\n\nThere were no other traffic violations, license suspensions or motor vehicle incidents involving the subject.";
                            divingresult = "Record [Hit w/ Consent]";
                            break;
                        case "Not Licensed":
                            drivingcomm = "\nIt shoud be noted that the subject reportedly does not hold a driver's license in the United States.";
                            divingresult = "Clear";
                            break;
                    }
                    doc.Replace("DRIVINGHISTORYDESCRIPTION", drivingcomm, true, true);
                    string strcomment = "";
                    switch (divingresult)
                    {
                        case "Clear":
                            doc.Replace("DRIVINGCOMMENT", "", true, true);
                            doc.Replace("DRIVINGRESULT", "Clear", true, true);
                            break;
                        case "Record [Hit w/ Consent]":
                            doc.Replace("DRIVINGCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("DRIVINGRESULT", "Record", true, true);
                            break;
                        case "Records [Hits w/ Consent]":
                            doc.Replace("DRIVINGCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("DRIVINGRESULT", "Records", true, true);
                            break;
                        case "Record [Hit w/o Consent]":
                            strcomment = "While [LastName]’s driving history record cannot be obtained without the subject’s express written authorization, research efforts conducted in relevant jurisdictions identified the subject <investigator to insert summary here>";
                            strcomment = strcomment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName);
                            doc.Replace("DRIVINGCOMMENT", strcomment, true, true);
                            doc.Replace("DRIVINGRESULT", "Record", true, true);
                            break;
                        case "Records [Hit w/o Consent]":
                            strcomment = "While [LastName]’s driving history record cannot be obtained without the subject’s express written authorization, research efforts conducted in relevant jurisdictions identified the subject <investigator to insert summary here>";
                            strcomment = strcomment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName);
                            doc.Replace("DRIVINGCOMMENT", strcomment, true, true);
                            doc.Replace("DRIVINGRESULT", "Records", true, true);
                            break;
                        case "N/A [No Consent or hits]":
                            doc.Replace("DRIVINGCOMMENT", "Driving history records cannot be obtained without the subject’s express written authorization", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            break;
                    }
                    //HASNAMEDRIVINGINCIDENTDESC
                    if (uSDDIndividual.otherdetails.Has_Name_only_driving_incidents == true)
                    {
                        doc.Replace("HASNAMEDRIVINGINCIDENTDESC", "\nIt should be noted that one or more individuals known only as “[FirstName] [LastName]” were identified in connection with at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.  Manual records research efforts -- as available -- would be required to determine whether the same relate to the subject of interest, which can be undertaken upon request -- if an expanded scope is warranted.\n", true, true);
                        switch (divingresult)
                        {
                            case "Clear":
                                doc.Replace("DRIVINGCOMMENT", "\n\n\nIt should be noted that one or more individuals known only as “[First Name] [Last Name]” were identified in connection with at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.", true, true);
                                doc.Replace("DRIVINGRESULT", "Clear\n\nPossible Records", true, true);
                                break;
                            case "Record [Hit w/ Consent]":
                                doc.Replace("DRIVINGCOMMENT", "<investigator to insert summary here>\n\nIt should be noted that one or more individuals known only as “[First Name] [Last Name]” were identified in connection with at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.", true, true);
                                doc.Replace("DRIVINGRESULT", "Record\n\nPossible Records", true, true);
                                break;
                            case "Records [Hits w/ Consent]":
                                doc.Replace("DRIVINGCOMMENT", "<investigator to insert summary here>\n\nIt should be noted that one or more individuals known only as “[First Name] [Last Name]” were identified in connection with at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.", true, true);
                                doc.Replace("DRIVINGRESULT", "Records\n\nPossible Records", true, true);
                                break;
                            case "Record [Hit w/o Consent]":
                                strcomment = "While [LastName]’s driving history record cannot be obtained without the subject’s express written authorization, research efforts conducted in relevant jurisdictions identified the subject <investigator to insert summary here>\n\nIt should be noted that one or more individuals known only as “[First Name] [Last Name]” were identified in connection with at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.";
                                strcomment = strcomment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName);
                                doc.Replace("DRIVINGCOMMENT", strcomment, true, true);
                                doc.Replace("DRIVINGRESULT", "Record\n\nPossible Records", true, true);
                                break;
                            case "Records [Hit w/o Consent]":
                                strcomment = "While [LastName]’s driving history record cannot be obtained without the subject’s express written authorization, research efforts conducted in relevant jurisdictions identified the subject <investigator to insert summary here>\n\nIt should be noted that one or more individuals known only as “[First Name] [Last Name]” were identified in connection with at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.";
                                strcomment = strcomment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName);
                                doc.Replace("DRIVINGCOMMENT", strcomment, true, true);
                                doc.Replace("DRIVINGRESULT", "Records\n\nPossible Records", true, true);
                                break;
                            case "N/A [No Consent or hits]":
                                doc.Replace("DRIVINGCOMMENT", "Driving history records cannot be obtained without the subject’s express written authorization.\n\nIt should be noted that one or more individuals known only as “[First Name] [Last Name]” were identified in connection with at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A\n\nPossible Records", true, true);
                                break;
                          
                        }
                    }
                    else
                    {
                        doc.Replace("HASNAMEDRIVINGINCIDENTDESC", "", true, true);
                    }
                    //CRIMINALRECDES
                    if (uSDDIndividual.otherdetails.Has_CriminalRecHit_resultpending == true)
                    {
                        doc.Replace("CRIMINALRECDES", "\n\nA criminal records search is currently pending through the <investigator to insert source/court/law enforcement agency>, the results of which will be provided under separate cover upon receipt.[CRIMFONTCHANGE]CRIMINALRECDES", true, false);
                        doc.SaveToFile(savePath);
                    }
                    string strconcatcrim = "";
                    string strcrimsumcom = "";
                    string strcrimsumres = "";
                    switch (uSDDIndividual.otherdetails.Has_CriminalRecHit.ToString())
                    {
                        case "Yes, Multiple Records":
                            strcrimsumres = "Records";
                            doc.Replace("CRIMINALRECDES", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                            strconcatcrim = " identified the subject as a Defendant in connection with the following criminal record:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                            break;
                        case "No":
                            strcrimsumres = "Clear";
                            doc.Replace("CRIMINALRECDES", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                            strconcatcrim = " did not identify the subject in connection with any criminal records.";
                            break;
                        case "Yes, Single Record":
                            strcrimsumres = "Record";
                            doc.Replace("CRIMINALRECDES", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                            strconcatcrim = " identified the subject as a Defendant in connection with the following criminal record:\n\n\t•\t<Investigator to insert bulleted list of results here>";
                            break;

                    }
                    if (uSDDIndividual.diligenceInputModel.CommonNameSubject == true)
                    {
                        strconcatcrim = string.Concat(strconcatcrim, "\n\nIn light of the commonality of [Last Name]’s name, this portion of the investigation was conducted utilizing the subject’s date of birth and Social Security number, in order to identify only those records conclusively relating to the subject of interest.  Records may exist that do not contain this type of identifying information, and an expanded scope would be required in this regard. <investigator to amend or remove if applicable>");
                    }
                    if (uSDDIndividual.otherdetails.Has_CriminalRecHit1 == true)
                    {
                        strconcatcrim = string.Concat(strconcatcrim, "\n\nIt should be noted that one or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>. <Investigator to add additional details as available>\n\nManual criminal court records research efforts – as available – would be required to determine whether the same relate to the subject of this investigation, which can be undertaken upon request – if an expanded scope is warranted.");
                    }
                    doc.SaveToFile(savePath);
                    strblnres = "";
                    if (arr == "")
                    {
                        strAdditionalStatesComment = string.Concat("(<not provided>)", strconcatcrim);
                        doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, false);
                        strAdditionalStatesComment = "";
                    }
                    else
                    {
                        if (arrlist.Count() > 1)
                        {
                            int count = arrlist.Count();
                            int k = 1;
                            for (int a = 0; a < arrlist.Count(); a++)
                            {
                                if (arrlist[a].ToUpper().ToString().Equals("Select State".ToUpper()))
                                {
                                }
                                else
                                {
                                    string strstatecourt = "";
                                    string strcriminallang = "";
                                    StateWiseFootnoteModel comment1 = _context.stateModel
                                         .Where(u => u.states.ToUpper().TrimEnd() == arrlist[a].ToUpper().TrimEnd())
                                         .FirstOrDefault();
                                    try
                                    {
                                        if (comment1.state_specific_district_courts.ToString().Equals("")) { strstatecourt = "<not available>"; }
                                        else { strstatecourt = comment1.state_specific_district_courts; }
                                    }
                                    catch (Exception) { strstatecourt = "<not available>"; }
                                    try
                                    {
                                        if (comment1.statewide_criminal_language_pe_report.ToString().Equals("")) { strcriminallang = ""; }
                                        else
                                        {
                                            strcriminallang = comment1.statewide_criminal_language_pe_report;
                                        }
                                    }
                                    catch { strcriminallang = ""; }
                                    if (k == count - 1)
                                    {
                                        strblnres = "";
                                        strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE]");
                                        doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                        doc.SaveToFile(savePath);
                                        for (int j = 2; j < 6; j++)
                                        {
                                            Section section = doc.Sections[j];
                                            foreach (Paragraph para in section.Paragraphs)
                                            {
                                                DocumentObject obj = null;
                                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                                {
                                                    obj = para.ChildObjects[i];
                                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                                    {
                                                        TextRange textRange = obj as TextRange;
                                                        string abc = textRange.Text;
                                                        // Find the word "Spire.Doc" in paragraph1
                                                        if (abc.ToString().Equals("[ADDITIONALSTATE]") || abc.ToString().EndsWith("[ADDITIONALSTATE]") || abc.ToString().Contains("[ADDITIONALSTATE]"))
                                                        {
                                                            Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                            footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                            footnote1.MarkerCharacterFormat.FontSize = 11;
                                                            //Insert footnote1 after the word "Spire.Doc"
                                                            para.ChildObjects.Insert(i + 1, footnote1);
                                                            string strapendfootnotetext = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert county courts searched, if applicable>.  [STATEWIDECRIMINALLANGUAGE]";
                                                            strapendfootnotetext = strapendfootnotetext.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                                                            strapendfootnotetext = strapendfootnotetext.Replace("[STATEWIDECRIMINALLANGUAGE]", strcriminallang);
                                                            TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strapendfootnotetext);
                                                            footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                            _text.CharacterFormat.FontName = "Calibri (Body)";
                                                            _text.CharacterFormat.FontSize = 9;
                                                            doc.Replace("[ADDITIONALSTATE]", "", true, false);
                                                            doc.SaveToFile(savePath);
                                                            TextRange tr = para.AppendText(" and [ADDITIONALSTATES]");
                                                            tr.CharacterFormat.FontName = "Calibri (Body)";
                                                            tr.CharacterFormat.FontSize = 11;
                                                            doc.SaveToFile(savePath);
                                                            strblnres = "true";
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (strblnres == "true")
                                                {
                                                    break;
                                                }
                                            }
                                            if (strblnres == "true")
                                            {
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (k == count)
                                        {
                                            strblnres = "";
                                            strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE]");
                                            doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                            doc.SaveToFile(savePath);
                                            for (int j = 2; j < 6; j++)
                                            {
                                                Section section = doc.Sections[j];
                                                foreach (Paragraph para in section.Paragraphs)
                                                {
                                                    DocumentObject obj = null;
                                                    for (int i = 0; i < para.ChildObjects.Count; i++)
                                                    {
                                                        obj = para.ChildObjects[i];
                                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                                        {
                                                            TextRange textRange = obj as TextRange;
                                                            string abc = textRange.Text;
                                                            // Find the word "Spire.Doc" in paragraph1
                                                            if (abc.ToString().Equals("[ADDITIONALSTATE]") || abc.ToString().EndsWith("[ADDITIONALSTATE]") || abc.ToString().Contains("[ADDITIONALSTATE]"))
                                                            {
                                                                Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                                footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                                footnote1.MarkerCharacterFormat.FontSize = 11;
                                                                //Insert footnote1 after the word "Spire.Doc"
                                                                para.ChildObjects.Insert(i + 1, footnote1);
                                                                string strapendfootnotetext = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert county courts searched, if applicable>. [STATEWIDECRIMINALLANGUAGE]";
                                                                strapendfootnotetext = strapendfootnotetext.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                                                                strapendfootnotetext = strapendfootnotetext.Replace("[STATEWIDECRIMINALLANGUAGE]", strcriminallang);
                                                                TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strapendfootnotetext);
                                                                footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                                _text.CharacterFormat.FontName = "Calibri (Body)";
                                                                _text.CharacterFormat.FontSize = 9;
                                                                doc.Replace("[ADDITIONALSTATE]", "", true, false);
                                                                doc.SaveToFile(savePath);
                                                                TextRange tr = para.AppendText(strconcatcrim);
                                                                tr.CharacterFormat.FontName = "Calibri (Body)";
                                                                tr.CharacterFormat.FontSize = 11;
                                                                doc.SaveToFile(savePath);
                                                                strblnres = "true";
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    if (strblnres == "true")
                                                    {
                                                        break;
                                                    }
                                                }
                                                if (strblnres == "true")
                                                {
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            strblnres = "";
                                            strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE],");
                                            doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                            doc.SaveToFile(savePath);
                                            for (int j = 2; j < 6; j++)
                                            {
                                                Section section = doc.Sections[j];
                                                foreach (Paragraph para in section.Paragraphs)
                                                {
                                                    DocumentObject obj = null;
                                                    for (int i = 0; i < para.ChildObjects.Count; i++)
                                                    {
                                                        obj = para.ChildObjects[i];
                                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                                        {
                                                            TextRange textRange = obj as TextRange;
                                                            string abc = textRange.Text;
                                                            // Find the word "Spire.Doc" in paragraph1
                                                            if (abc.ToString().Equals("[ADDITIONALSTATE]") || abc.ToString().EndsWith("[ADDITIONALSTATE]") || abc.ToString().Contains("[ADDITIONALSTATE]"))
                                                            {
                                                                Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                                footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                                footnote1.MarkerCharacterFormat.FontSize = 11;
                                                                //Insert footnote1 after the word "Spire.Doc"
                                                                para.ChildObjects.Insert(i + 1, footnote1);
                                                                string strapendfootnotetext = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert county courts searched, if applicable>.  [STATEWIDECRIMINALLANGUAGE]";
                                                                strapendfootnotetext = strapendfootnotetext.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                                                                strapendfootnotetext = strapendfootnotetext.Replace("[STATEWIDECRIMINALLANGUAGE]", strcriminallang);
                                                                TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strapendfootnotetext);
                                                                footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                                _text.CharacterFormat.FontName = "Calibri (Body)";
                                                                _text.CharacterFormat.FontSize = 9;
                                                                doc.Replace("[ADDITIONALSTATE]", "", true, false);
                                                                doc.SaveToFile(savePath);
                                                                TextRange tr = para.AppendText(" [ADDITIONALSTATES]");
                                                                tr.CharacterFormat.FontName = "Calibri (Body)";
                                                                tr.CharacterFormat.FontSize = 11;
                                                                doc.SaveToFile(savePath);
                                                                strblnres = "true";
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    if (strblnres == "true")
                                                    {
                                                        break;
                                                    }
                                                }
                                                if (strblnres == "true")
                                                {
                                                    break;
                                                }
                                            }

                                        }
                                    }
                                }
                                k++;
                            }
                        }
                        else
                        {
                            string strstatecourt = "";
                            string strcriminallang = "";
                            StateWiseFootnoteModel comment1 = _context.stateModel
                                 .Where(u => u.states.ToUpper().TrimEnd() == arrlist[0].ToString().ToUpper().TrimEnd())
                                 .FirstOrDefault();
                            try
                            {
                                if (comment1.state_specific_district_courts.ToString().Equals("")) { strstatecourt = "<not available>"; }
                                else { strstatecourt = comment1.state_specific_district_courts; }
                            }
                            catch
                            {
                                strstatecourt = "<not available>";
                            }
                            try
                            {
                                if (comment1.statewide_criminal_language_pe_report.ToString().Equals("")) { strcriminallang = ""; }
                                else
                                {
                                    strcriminallang = comment1.statewide_criminal_language_pe_report;
                                }
                            }
                            catch
                            {
                                strcriminallang = "";
                            }

                            strAdditionalStatesComment = string.Concat(arrlist[0].ToString(), "[ADDITIONALSTATE]");
                            doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, false);
                            doc.SaveToFile(savePath);
                            for (int j = 2; j < 6; j++)
                            {
                                Section section = doc.Sections[j];
                                foreach (Paragraph para in section.Paragraphs)
                                {
                                    DocumentObject obj = null;
                                    for (int i = 0; i < para.ChildObjects.Count; i++)
                                    {
                                        obj = para.ChildObjects[i];
                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                        {
                                            TextRange textRange = obj as TextRange;
                                            string abc = textRange.Text;
                                            // Find the word "Spire.Doc" in paragraph1
                                            if (abc.ToString().Equals("[ADDITIONALSTATE]") || abc.ToString().EndsWith("[ADDITIONALSTATE]") || abc.ToString().Contains("[ADDITIONALSTATE]"))
                                            {
                                                Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                footnote1.MarkerCharacterFormat.FontSize = 11;
                                                //Insert footnote1 after the word "Spire.Doc"
                                                para.ChildObjects.Insert(i + 1, footnote1);
                                                string strapendfootnotetext = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert county courts searched, if applicable>.  [STATEWIDECRIMINALLANGUAGE]";
                                                strapendfootnotetext = strapendfootnotetext.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                                                strapendfootnotetext = strapendfootnotetext.Replace("[STATEWIDECRIMINALLANGUAGE]", strcriminallang);
                                                TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strapendfootnotetext);
                                                footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                _text.CharacterFormat.FontName = "Calibri (Body)";
                                                _text.CharacterFormat.FontSize = 9;
                                                TextRange tr = para.AppendText(strconcatcrim.ToString());
                                                tr.CharacterFormat.FontName = "Calibri (Body)";
                                                tr.CharacterFormat.FontSize = 11;
                                                doc.Replace("[ADDITIONALSTATE]", "", true, false);
                                                doc.SaveToFile(savePath);
                                                strblnres = "true";
                                                break;
                                            }
                                        }
                                    }
                                    if (strblnres == "true")
                                    {
                                        break;
                                    }
                                }
                                if (strblnres == "true")
                                {
                                    break;
                                }
                            }
                        }
                    }
                    switch (strcrimsumres.ToString())
                    {
                        case "Clear":
                            if (uSDDIndividual.otherdetails.Has_CriminalRecHit_resultpending == true)
                            {
                                strcrimsumcom = "<investigator to specify which states came back clear>";
                            }
                            else { strcrimsumcom = ""; }
                            break;
                        case "Record":
                            strcrimsumcom = "[LastName] was identified as a Defendant in connection with a <case type>, which was recorded in <State> in <YYYY>, and is currently <status>";
                            break;
                        case "Records":
                            strcrimsumcom = "[LastName] was identified as a Defendant in connection with at least <record types>, which were recorded in <State> between <YYYY> and <YYYY>, and are currently <status>";
                            break;
                    }
                    if (uSDDIndividual.otherdetails.Has_CriminalRecHit_resultpending == true)
                    {
                        strcrimsumres = string.Concat("Result Pending\n\n", strcrimsumres);
                        strcrimsumcom = string.Concat("Criminal records searches are currently pending through the <investigator to insert source/court/law enforcement agency>, the results of which will be provided under separate cover upon receipt[CRIMRSITALIC]\n\n", strcrimsumcom);
                    }
                    if (uSDDIndividual.otherdetails.Has_CriminalRecHit1 == true)
                    {
                        strcrimsumres = string.Concat(strcrimsumres, "\n\nPossible Records");
                        strcrimsumcom = string.Concat(strcrimsumcom, "\n\nOne or more individuals known only as “[First Name] [Last Name]” were identified as having been Defendants in at least <number> criminal records in <States>, which were filed between <YYYY> and <YYYY>, and pertain to <type of charges> and are currently <status>");
                    }
                    doc.Replace("CRIMINALCOMMENT", strcrimsumcom, true, true);
                    doc.Replace("CRIMINALRESULT", strcrimsumres, true, true);
                    doc.SaveToFile(savePath);
                    Table table = doc.Sections[1].Tables[0] as Table;
                    TableCell cell1 = table.Rows[17].Cells[2];
                    TableCell cell2 = table.Rows[17].Cells[1];
                    foreach (Paragraph p1 in cell1.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1
                                if (abc.ToString().EndsWith("[CRIMRSITALIC]"))
                                {
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                    doc.SaveToFile(savePath);
                                    doc.Replace("[CRIMRSITALIC]", "", true, false);
                                    doc.SaveToFile(savePath);
                                    break;
                                }
                            }
                        }
                    }

                    foreach (Paragraph p1 in cell2.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1
                                if (abc.ToString().EndsWith("Result Pending"))
                                {
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    doc.SaveToFile(savePath);
                                    break;
                                }
                                else
                                {
                                    if (abc.ToString().EndsWith("Possible Records"))
                                    {
                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                        textRange.CharacterFormat.FontSize = 11;
                                        textRange.CharacterFormat.Italic = true;
                                        doc.SaveToFile(savePath);
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    //HASCIVILRECDESC      
                    if (uSDDIndividual.otherdetails.has_civil_resultpending == true)
                    {
                        doc.Replace("[HASCIVILRECDESC]", "\n\nA civil court records search is currently pending through the <enter source/court>, the results of which will be provided under separate cover upon receipt.[CIVILFONTCHANGE][HASCIVILRECDESC]", true, false);
                        doc.SaveToFile(savePath);
                    }
                    string strconcatcivilrec = "";
                    strAdditionalStatesComment = "";
                    string strsummcivilcom = "";
                    string strsummcivilres = "";
                    switch (uSDDIndividual.otherdetails.Has_Civil_Records.ToString())
                    {
                        case "No":
                            strsummcivilres = "Clear";
                            doc.Replace("[HASCIVILRECDESC]", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                            strconcatcivilrec = " did not identify the subject, personally, in connection with any civil litigation filings.";
                            break;
                        case "Yes, Multiple Records":
                            strsummcivilres = "Records";
                            doc.Replace("[HASCIVILRECDESC]", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                            strconcatcivilrec = " identified the subject, personally, in connection with the following civil litigation filings:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                            break;
                        case "Yes, Single Record":
                            strsummcivilres = "Record";
                            doc.Replace("[HASCIVILRECDESC]", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                            strconcatcivilrec = " identified the subject, personally, in connection with the following civil litigation filings:\n\n\t•\t<Investigator to insert bulleted list of results here>";
                            break;
                    }
                    doc.SaveToFile(savePath);
                    strblnres = "";
                    if (arr == "")
                    {
                        strAdditionalStatesComment = string.Concat("(<not provided>)", strconcatcrim);
                        doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, false);
                        strAdditionalStatesComment = "";
                    }
                    else
                    {
                        if (arrlist.Count() > 1)
                        {
                            int count = arrlist.Count();
                            int k = 1;
                            for (int a = 0; a < arrlist.Count(); a++)
                            {
                                if (arrlist[a].ToUpper().ToString().Equals("Select State".ToUpper()))
                                {
                                }
                                else
                                {
                                    string strstatecourt = "";
                                    StateWiseFootnoteModel comment1 = _context.stateModel
                                          .Where(u => u.states.ToUpper().TrimEnd() == arrlist[a].ToUpper().TrimEnd())
                                          .FirstOrDefault();
                                    try
                                    {
                                        if (comment1.state_specific_district_courts.ToString().Equals("")) { strstatecourt = "<not provided>"; }
                                        else
                                        {
                                            strstatecourt = comment1.state_specific_district_courts;
                                        }
                                    }
                                    catch
                                    {
                                        strstatecourt = "<not available>";
                                    }
                                    string strcivilfootnote = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert county courts searched>.";
                                    strcivilfootnote = strcivilfootnote.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                                    if (k == count - 1)
                                    {
                                        strblnres = "";
                                        strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE]");
                                        doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                        doc.SaveToFile(savePath);
                                        for (int j = 2; j < 7; j++)
                                        {
                                            Section section = doc.Sections[j];
                                            foreach (Paragraph para in section.Paragraphs)
                                            {
                                                DocumentObject obj = null;
                                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                                {
                                                    obj = para.ChildObjects[i];
                                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                                    {
                                                        TextRange textRange = obj as TextRange;
                                                        string abc = textRange.Text;
                                                        // Find the word "Spire.Doc" in paragraph1
                                                        if (abc.ToString().Equals("[ADDITIONALSTATE]") || abc.ToString().EndsWith("[ADDITIONALSTATE]") || abc.ToString().Contains("[ADDITIONALSTATE]"))
                                                        {
                                                            Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                            footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                            footnote1.MarkerCharacterFormat.FontSize = 11;
                                                            //Insert footnote1 after the word "Spire.Doc"
                                                            para.ChildObjects.Insert(i + 1, footnote1);
                                                            TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strcivilfootnote);
                                                            footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                            _text.CharacterFormat.FontName = "Calibri (Body)";
                                                            _text.CharacterFormat.FontSize = 9;
                                                            doc.Replace("[ADDITIONALSTATE]", "", true, false);
                                                            doc.SaveToFile(savePath);
                                                            TextRange tr = para.AppendText(" and [ADDITIONALSTATES]");
                                                            tr.CharacterFormat.FontName = "Calibri (Body)";
                                                            tr.CharacterFormat.FontSize = 11;
                                                            doc.SaveToFile(savePath);
                                                            strblnres = "true";
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (strblnres == "true")
                                                {
                                                    break;
                                                }
                                            }
                                            if (strblnres == "true")
                                            {
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (k == count)
                                        {
                                            strblnres = "";
                                            strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE]");
                                            doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                            doc.SaveToFile(savePath);
                                            for (int j = 2; j < 7; j++)
                                            {
                                                Section section = doc.Sections[j];
                                                foreach (Paragraph para in section.Paragraphs)
                                                {
                                                    DocumentObject obj = null;
                                                    for (int i = 0; i < para.ChildObjects.Count; i++)
                                                    {
                                                        obj = para.ChildObjects[i];
                                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                                        {
                                                            TextRange textRange = obj as TextRange;
                                                            string abc = textRange.Text;
                                                            // Find the word "Spire.Doc" in paragraph1
                                                            if (abc.ToString().Equals("[ADDITIONALSTATE]") || abc.ToString().EndsWith("[ADDITIONALSTATE]") || abc.ToString().Contains("[ADDITIONALSTATE]"))
                                                            {
                                                                Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                                footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                                footnote1.MarkerCharacterFormat.FontSize = 11;
                                                                //Insert footnote1 after the word "Spire.Doc"
                                                                para.ChildObjects.Insert(i + 1, footnote1);
                                                                TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strcivilfootnote);
                                                                footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                                _text.CharacterFormat.FontName = "Calibri (Body)";
                                                                _text.CharacterFormat.FontSize = 9;
                                                                doc.Replace("[ADDITIONALSTATE]", "", true, false);
                                                                doc.SaveToFile(savePath);
                                                                TextRange tr = para.AppendText(strconcatcivilrec);
                                                                tr.CharacterFormat.FontName = "Calibri (Body)";
                                                                tr.CharacterFormat.FontSize = 11;
                                                                doc.SaveToFile(savePath);
                                                                strblnres = "true";
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    if (strblnres == "true")
                                                    {
                                                        break;
                                                    }
                                                }
                                                if (strblnres == "true")
                                                {
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            strblnres = "";
                                            strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE],");
                                            doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                            doc.SaveToFile(savePath);
                                            for (int j = 2; j < 7; j++)
                                            {
                                                Section section = doc.Sections[j];
                                                foreach (Paragraph para in section.Paragraphs)
                                                {
                                                    DocumentObject obj = null;
                                                    for (int i = 0; i < para.ChildObjects.Count; i++)
                                                    {
                                                        obj = para.ChildObjects[i];
                                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                                        {
                                                            TextRange textRange = obj as TextRange;
                                                            string abc = textRange.Text;
                                                            // Find the word "Spire.Doc" in paragraph1
                                                            if (abc.ToString().Equals("[ADDITIONALSTATE]") || abc.ToString().EndsWith("[ADDITIONALSTATE]") || abc.ToString().Contains("[ADDITIONALSTATE]"))
                                                            {
                                                                Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                                footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                                footnote1.MarkerCharacterFormat.FontSize = 11;
                                                                //Insert footnote1 after the word "Spire.Doc"
                                                                para.ChildObjects.Insert(i + 1, footnote1);
                                                                TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strcivilfootnote);
                                                                footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                                _text.CharacterFormat.FontName = "Calibri (Body)";
                                                                _text.CharacterFormat.FontSize = 9;
                                                                doc.Replace("[ADDITIONALSTATE]", "", true, false);
                                                                doc.SaveToFile(savePath);
                                                                TextRange tr = para.AppendText(" [ADDITIONALSTATES]");
                                                                tr.CharacterFormat.FontName = "Calibri (Body)";
                                                                tr.CharacterFormat.FontSize = 11;
                                                                doc.SaveToFile(savePath);
                                                                strblnres = "true";
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    if (strblnres == "true")
                                                    {
                                                        break;
                                                    }
                                                }
                                                if (strblnres == "true")
                                                {
                                                    break;
                                                }
                                            }

                                        }
                                    }
                                }
                                k++;
                            }
                        }
                        else
                        {
                            strAdditionalStatesComment = string.Concat(arrlist[0].ToString(), "[ADDITIONALSTATE]");
                            doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, false);
                            StateWiseFootnoteModel comment1 = _context.stateModel
                                      .Where(u => u.states.ToUpper() == arrlist[0].ToString().ToUpper())
                                      .FirstOrDefault();
                            string strstatecourt = comment1.state_specific_district_courts;
                            string strcivilfootnote = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert county courts searched>.";
                            strcivilfootnote = strcivilfootnote.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                            doc.SaveToFile(savePath);
                            for (int j = 2; j < 7; j++)
                            {
                                Section section = doc.Sections[j];
                                foreach (Paragraph para in section.Paragraphs)
                                {
                                    DocumentObject obj = null;
                                    for (int i = 0; i < para.ChildObjects.Count; i++)
                                    {
                                        obj = para.ChildObjects[i];
                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                        {
                                            TextRange textRange = obj as TextRange;
                                            string abc = textRange.Text;
                                            // Find the word "Spire.Doc" in paragraph1
                                            if (abc.ToString().Equals("[ADDITIONALSTATE]") || abc.ToString().EndsWith("[ADDITIONALSTATE]") || abc.ToString().Contains("[ADDITIONALSTATE]"))
                                            {
                                                Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                footnote1.MarkerCharacterFormat.FontSize = 11;
                                                //Insert footnote1 after the word "Spire.Doc"
                                                para.ChildObjects.Insert(i + 1, footnote1);
                                                TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strcivilfootnote);
                                                footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                _text.CharacterFormat.FontName = "Calibri (Body)";
                                                _text.CharacterFormat.FontSize = 9;
                                                TextRange tr = para.AppendText(strconcatcivilrec.ToString());
                                                tr.CharacterFormat.FontName = "Calibri (Body)";
                                                tr.CharacterFormat.FontSize = 11;
                                                doc.Replace("[ADDITIONALSTATE]", "", true, false);
                                                doc.SaveToFile(savePath);
                                                strblnres = "true";
                                                break;
                                            }
                                        }
                                    }
                                    if (strblnres == "true")
                                    {
                                        break;
                                    }
                                }
                                if (strblnres == "true")
                                {
                                    break;
                                }
                            }
                        }
                    }

                    if (uSDDIndividual.otherdetails.Has_US_Tax_Court_Hit.Equals("Yes, Single Record") && strsummcivilres == "Clear")
                    {
                        strsummcivilres = "Record";
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_US_Tax_Court_Hit.Equals("Yes, Single Record") && strsummcivilres == "Records" || strsummcivilres == "Record")
                        {
                            strsummcivilres = "Records";
                        }
                        if (uSDDIndividual.otherdetails.Has_US_Tax_Court_Hit.Equals("Yes, Multiple Records"))
                        {
                            strsummcivilres = "Records";
                        }
                    }
                    switch (strsummcivilres)
                    {
                        case "Clear":
                            strsummcivilcom = "";
                            break;
                        case "Record":
                            strsummcivilcom = "[LastName] was identified as a <party type> in connection with a civil litigation filing, which were recorded in <State> in <YYYY>, and is currently <status>";
                            break;
                        case "Records":
                            strsummcivilcom = "[LastName] was identified as a <party type> in connection with at least <number> civil litigation filings, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                            break;
                    }
                    if (uSDDIndividual.otherdetails.has_civil_resultpending == true)
                    {
                        strsummcivilres = string.Concat("Result Pending\n\n", strsummcivilres);
                        strsummcivilcom = string.Concat("A civil court records search is currently pending through the <enter source/court>, the results of which will be provided under separate cover upon receipt[CIVILRSITALIC]\n\n", strsummcivilcom);
                    }
                    if (uSDDIndividual.otherdetails.Has_Name_Only == true)
                    {
                        strsummcivilres = string.Concat(strsummcivilres, "\n\nPossible Records");
                        strsummcivilcom = string.Concat(strsummcivilcom, "\n\nOne or more individuals known only as “[First Name] [Last Name]” were identified as <party type> to at least <number> <case type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>");
                    }
                    doc.Replace("CIVILRESULT", strsummcivilres, true, true);
                    doc.Replace("CIVILCOMMENT", strsummcivilcom, true, true);
                    doc.SaveToFile(savePath);

                    table = doc.Sections[1].Tables[0] as Table;
                    TableCell cell5 = table.Rows[13].Cells[2];
                    TableCell cell6 = table.Rows[13].Cells[1];
                    foreach (Paragraph p1 in cell5.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1
                                if (abc.ToString().EndsWith("[CIVILRSITALIC]"))
                                {
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                    doc.SaveToFile(savePath);
                                    doc.Replace("[CIVILRSITALIC]", "", true, false);
                                    doc.SaveToFile(savePath);
                                    break;
                                }
                            }
                        }
                    }

                    foreach (Paragraph p1 in cell6.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1
                                if (abc.ToString().EndsWith("Result Pending"))
                                {
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    doc.SaveToFile(savePath);
                                    break;
                                }
                                else
                                {
                                    if (abc.ToString().EndsWith("Possible Records"))
                                    {
                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                        textRange.CharacterFormat.FontSize = 11;
                                        textRange.CharacterFormat.Italic = true;
                                        doc.SaveToFile(savePath);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    //EmployeeDetails
                    string emp_Comment = "";
                    string strempstartdate = "";
                    string strempenddate = "";
                    for (int i = 0; i < uSDDIndividual.EmployerModel.Count; i++)
                    {
                        strempstartdate = "<not provided>";
                        strempenddate = "<not provided>";
                        if (!uSDDIndividual.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && !uSDDIndividual.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strempstartdate = uSDDIndividual.EmployerModel[i].Emp_StartDateMonth + " " + uSDDIndividual.EmployerModel[i].Emp_StartDateDay + ", " + uSDDIndividual.EmployerModel[i].Emp_StartDateYear;
                            // employerModel.Emp_StartDate = strempstartdate;
                        }
                        if (uSDDIndividual.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && uSDDIndividual.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && uSDDIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strempstartdate = "<not provided>";
                        }
                        if (uSDDIndividual.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && !uSDDIndividual.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strempstartdate = uSDDIndividual.EmployerModel[i].Emp_StartDateMonth + " " + uSDDIndividual.EmployerModel[i].Emp_StartDateYear;
                            //employerModel.Emp_StartDate = strempstartdate;
                        }
                        if (uSDDIndividual.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && uSDDIndividual.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strempstartdate = uSDDIndividual.EmployerModel[i].Emp_StartDateYear;
                            //employerModel.Emp_StartDate = strempstartdate;
                        }
                        if (i == 0 && uSDDIndividual.EmployerModel[0].Emp_Status.ToString().Equals("Current")) { }
                        else
                        {
                            if (uSDDIndividual.EmployerModel[i].Emp_Status.ToString().Equals("Concurrent")) { }
                            else
                            {
                                if (!uSDDIndividual.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && !uSDDIndividual.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                                {
                                    strempenddate = uSDDIndividual.EmployerModel[i].Emp_EndDateMonth + " " + uSDDIndividual.EmployerModel[i].Emp_EndDateDay + ", " + uSDDIndividual.EmployerModel[i].Emp_EndDateYear;
                                    //  employerModel.Emp_EndDate = strempenddate;
                                }
                                if (uSDDIndividual.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && uSDDIndividual.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && uSDDIndividual.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                                {
                                    strempenddate = "<not provided>";
                                }
                                if (uSDDIndividual.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && !uSDDIndividual.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                                {
                                    strempenddate = uSDDIndividual.EmployerModel[i].Emp_EndDateMonth + " " + uSDDIndividual.EmployerModel[i].Emp_EndDateYear;
                                    //employerModel.Emp_EndDate = strempenddate;
                                }
                                if (uSDDIndividual.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && uSDDIndividual.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                                {
                                    strempenddate = uSDDIndividual.EmployerModel[i].Emp_EndDateYear;
                                    //employerModel.Emp_EndDate = strempenddate;
                                }
                            }
                        }
                        if (i == 0)
                        {
                            CommentModel comment = _context.DbComment
                                  .Where(u => u.Comment_type == "Emp1")
                                  .FirstOrDefault();
                            if (uSDDIndividual.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                            {
                                emp_Comment = comment.confirmed_comment.ToString();
                                emp_Comment = emp_Comment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                                emp_Comment = emp_Comment.Replace("[Position1]", uSDDIndividual.EmployerModel[0].Emp_Position.ToString());
                                emp_Comment = emp_Comment.Replace("[Employer1]", uSDDIndividual.EmployerModel[0].Emp_Employer.ToString());
                                emp_Comment = emp_Comment.Replace("[EmpLocation1]", uSDDIndividual.EmployerModel[0].Emp_Location.ToString());
                                emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                                emp_Comment = string.Concat(emp_Comment, "\n");
                            }
                            else
                            {
                                emp_Comment = comment.unconfirmed_comment.ToString();
                                emp_Comment = emp_Comment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                                emp_Comment = emp_Comment.Replace("[Position1]", uSDDIndividual.EmployerModel[0].Emp_Position.ToString());
                                emp_Comment = emp_Comment.Replace("[Employer1]", uSDDIndividual.EmployerModel[0].Emp_Employer.ToString());
                                emp_Comment = emp_Comment.Replace("[EmpLocation1]", uSDDIndividual.EmployerModel[0].Emp_Location.ToString());
                                emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                                emp_Comment = string.Concat(emp_Comment, "\n");
                            }
                            doc.Replace("EMPLOYEEDESCRIPTION", emp_Comment, true, true);
                            doc.SaveToFile(savePath);
                            strblnres = "";
                            for (int j = 2; j < 8; j++)
                            {
                                Section section = doc.Sections[j];
                                foreach (Paragraph para in section.Paragraphs)
                                {
                                    DocumentObject obj = null;
                                    for (int k = 0; k < para.ChildObjects.Count; k++)
                                    {
                                        obj = para.ChildObjects[k];
                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                        {
                                            TextRange textRange = obj as TextRange;
                                            string abc = textRange.Text;
                                            // Find the word "Spire.Doc" in paragraph1
                                            if (abc.ToString().Contains("[Employer1Footnote]") || abc.ToString().EndsWith("[Employer1Footnote]"))
                                            {
                                                Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                footnote1.MarkerCharacterFormat.FontSize = 11;
                                                //Insert footnote1 after the word "Spire.Doc"
                                                para.ChildObjects.Insert(k + 1, footnote1);
                                                TextRange _text = footnote1.TextBody.AddParagraph().AppendText("<Investigator to insert company registry detail here>.");
                                                footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                _text.CharacterFormat.FontName = "Calibri (Body)";
                                                _text.CharacterFormat.FontSize = 9;
                                                //Append the line
                                                string stremp1textappended = "";
                                                if (uSDDIndividual.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                                                {
                                                    stremp1textappended = " in [EmpLocation1], since [EmpStartDate1]. EMPLOYEEDESCRIPTION";
                                                }
                                                else
                                                {
                                                    stremp1textappended = " in [EmpLocation1], since [EmpStartDate1]. EMPLOYEEDESCRIPTION";
                                                }
                                                doc.Replace("[Employer1Footnote]", uSDDIndividual.EmployerModel[0].Emp_Employer.ToString(), true, true);
                                                stremp1textappended = stremp1textappended.Replace("[EmpLocation1]", uSDDIndividual.EmployerModel[0].Emp_Location.ToString());
                                                stremp1textappended = stremp1textappended.Replace("[EmpStartDate1]", strempstartdate);
                                                TextRange tr = para.AppendText(stremp1textappended);
                                                tr.CharacterFormat.FontName = "Calibri (Body)";
                                                tr.CharacterFormat.FontSize = 11;
                                                if (uSDDIndividual.EmployerModel[0].Emp_Status.ToString().Equals("Current"))
                                                {
                                                    doc.Replace("EMPLOYEEDESCRIPTION", "<investigator to insert brief summary of company here> EMPLOYEEDESCRIPTION", true, true);
                                                }
                                                doc.SaveToFile(savePath);
                                                doc.SaveToFile(savePath);
                                                strblnres = "true";
                                                break;
                                            }
                                        }
                                    }
                                    if (strblnres == "true")
                                    {
                                        break;
                                    }
                                }
                                if (strblnres == "true")
                                {
                                    break;
                                }
                            }
                            doc.SaveToFile(savePath);
                            emp_Comment = "";
                        }
                        else
                        {
                            if (i == 1) { emp_Comment = "\n"; }
                            if (uSDDIndividual.EmployerModel[i].Emp_Status.ToString().Equals("Concurrent"))
                            {
                                if (uSDDIndividual.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                                {
                                    emp_Comment = "In addition to the above, [LastName] is also a [Position1] of [Employer1Footnote]";
                                    emp_Comment = emp_Comment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                                    emp_Comment = emp_Comment.Replace("[Position1]", uSDDIndividual.EmployerModel[i].Emp_Position.ToString());
                                    emp_Comment = emp_Comment.Replace("[Employer1]", uSDDIndividual.EmployerModel[i].Emp_Employer.ToString());
                                    emp_Comment = emp_Comment.Replace("[EmpLocation1]", uSDDIndividual.EmployerModel[i].Emp_Location.ToString());
                                    emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                                    emp_Comment = string.Concat("\n\n", emp_Comment);
                                }
                                else
                                {
                                    emp_Comment = "In addition to the above, [LastName] is also a [Position1] of [Employer1Footnote]";
                                    emp_Comment = emp_Comment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                                    emp_Comment = emp_Comment.Replace("[Position1]", uSDDIndividual.EmployerModel[i].Emp_Position.ToString());
                                    emp_Comment = emp_Comment.Replace("[Employer1]", uSDDIndividual.EmployerModel[i].Emp_Employer.ToString());
                                    emp_Comment = emp_Comment.Replace("[EmpLocation1]", uSDDIndividual.EmployerModel[i].Emp_Location.ToString());
                                    emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                                    emp_Comment = string.Concat("\n\n", emp_Comment);
                                }
                                doc.Replace("EMPLOYEEDESCRIPTION", emp_Comment, true, true);
                                doc.SaveToFile(savePath);
                                strblnres = "";
                                try
                                {
                                    for (int j = 2; j < 8; j++)
                                    {
                                        Section section = doc.Sections[j];
                                        foreach (Paragraph para in section.Paragraphs)
                                        {
                                            DocumentObject obj = null;
                                            for (int k = 0; k < para.ChildObjects.Count; k++)
                                            {
                                                obj = para.ChildObjects[k];
                                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                                {
                                                    TextRange textRange = obj as TextRange;
                                                    string abc = textRange.Text;
                                                    // Find the word "Spire.Doc" in paragraph1
                                                    if (abc.ToString().Contains("[Employer1Footnote]") || abc.ToString().EndsWith("[Employer1Footnote]"))
                                                    {
                                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                                        //Insert footnote1 after the word "Spire.Doc"
                                                        para.ChildObjects.Insert(k + 1, footnote1);
                                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("<Investigator to insert company registry detail here>.");
                                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                                        _text.CharacterFormat.FontSize = 9;
                                                        //Append the line
                                                        string stremp1textappended = "";
                                                        if (uSDDIndividual.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                                                        {
                                                            stremp1textappended = " in [EmpLocation1], since [EmpStartDate1]. EMPLOYEEDESCRIPTION";
                                                        }
                                                        else
                                                        {
                                                            stremp1textappended = " in [EmpLocation1], since [EmpStartDate1]. EMPLOYEEDESCRIPTION";
                                                        }
                                                        doc.Replace("[Employer1Footnote]", uSDDIndividual.EmployerModel[i].Emp_Employer.ToString(), true, true);
                                                        stremp1textappended = stremp1textappended.Replace("[EmpLocation1]", uSDDIndividual.EmployerModel[i].Emp_Location.ToString());
                                                        stremp1textappended = stremp1textappended.Replace("[EmpStartDate1]", strempstartdate);
                                                        TextRange tr = para.AppendText(stremp1textappended);
                                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                                        tr.CharacterFormat.FontSize = 11;
                                                        doc.SaveToFile(savePath);
                                                        //if (uSDDIndividual.EmployerModel[i].Emp_Status.ToString().Equals("Current"))
                                                        //{
                                                        //    doc.Replace("EMPLOYEEDESCRIPTION", "<investigator to insert brief summary of company here> EMPLOYEEDESCRIPTION", true, true);
                                                        //}
                                                        //doc.SaveToFile(savePath);
                                                        strblnres = "true";
                                                        break;
                                                    }
                                                }
                                            }
                                            if (strblnres == "true")
                                            {
                                                break;
                                            }
                                        }
                                        if (strblnres == "true")
                                        {
                                            break;
                                        }
                                    }
                                }
                                catch { }
                                doc.SaveToFile(savePath);
                                emp_Comment = "";
                            }
                            else
                            {
                                CommentModel comment = _context.DbComment
                                  .Where(u => u.Comment_type == "Emp2")
                                  .FirstOrDefault();
                                if (uSDDIndividual.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                                {
                                    emp_Comment = string.Concat(emp_Comment, "\n", comment.confirmed_comment.ToString());
                                }
                                else
                                {
                                    if (uSDDIndividual.EmployerModel[i].Emp_Confirmed.Equals("No"))
                                    {
                                        emp_Comment = string.Concat(emp_Comment, "\n", comment.unconfirmed_comment.ToString());
                                    }
                                    else
                                    {
                                        emp_Comment = string.Concat(emp_Comment, "\n", comment.other_comment.ToString(), " APPENDEMPTEXT", i.ToString());
                                    }
                                }
                                emp_Comment = emp_Comment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                                emp_Comment = emp_Comment.Replace("[Position]", uSDDIndividual.EmployerModel[i].Emp_Position.ToString());
                                emp_Comment = emp_Comment.Replace("[Employer]", uSDDIndividual.EmployerModel[i].Emp_Employer.ToString());
                                emp_Comment = emp_Comment.Replace("[EmpLocation]", uSDDIndividual.EmployerModel[i].Emp_Location.ToString());
                                emp_Comment = emp_Comment.Replace("[EmpStartDate]", strempstartdate);
                                emp_Comment = emp_Comment.Replace("[EmpEndDate]", strempenddate);
                                if (i == uSDDIndividual.EmployerModel.Count - 1) { }
                                else
                                { emp_Comment = string.Concat(emp_Comment, "\n"); }
                            }
                        }
                    }
                    doc.Replace("EMPLOYEEDESCRIPTION", emp_Comment, true, true);
                    doc.SaveToFile(savePath);
                    string bnres = "";
                    for (int i = 1; i < uSDDIndividual.EmployerModel.Count; i++)
                    {
                        bnres = "";
                        if (uSDDIndividual.EmployerModel[i].Emp_Confirmed.ToString().Equals("Result Pending"))
                        {
                            for (int j = 1; j < 8; j++)
                            {
                                Section section = doc.Sections[j];
                                foreach (Paragraph para in section.Paragraphs)
                                {
                                    DocumentObject obj = null;
                                    for (int k = 0; k < para.ChildObjects.Count; k++)
                                    {
                                        obj = para.ChildObjects[k];
                                        if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                        {
                                            TextRange textRange = obj as TextRange;
                                            string abc = textRange.Text;
                                            // Find the word "Civil" in paragraph1
                                            if (abc.ToString().Contains("APPENDEMPTEXT"))
                                            {
                                                textRange = para.AppendText(" Efforts to independently confirm the subject’s tenure at the same are currently ongoing, the results of which will be provided under separate cover, if and when received.");
                                                textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                textRange.CharacterFormat.FontSize = 11;
                                                textRange.CharacterFormat.Italic = true;
                                                textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                                doc.SaveToFile(savePath);
                                                doc.Replace(string.Concat("APPENDEMPTEXT", i.ToString()), "", true, false);
                                                doc.SaveToFile(savePath);
                                                bnres = "true";
                                                break;
                                            }
                                        }
                                    }
                                    if (bnres.Equals("true")) { break; }
                                }
                                if (bnres.Equals("true")) { break; }
                            }
                        }
                    }
                    doc.SaveToFile(savePath);
                    //Education Details
                    string edu_comment = "";
                    string edu_header = "";
                    string edu_summcomment = "";
                    for (int i = 0; i < uSDDIndividual.educationModels.Count; i++)
                    {
                        CommentModel comment1 = _context.DbComment
                                     .Where(u => u.Comment_type == "Edu1")
                                     .FirstOrDefault();
                        CommentModel edurescomment = _context.DbComment
                                     .Where(u => u.Comment_type == "EDU_Resultpending")
                                     .FirstOrDefault();
                        CommentModel eduattcomment = _context.DbComment
                                     .Where(u => u.Comment_type == "EDU_Attendance")
                                     .FirstOrDefault();
                        CommentModel edusumcomment1 = _context.DbComment
                                     .Where(u => u.Comment_type == "EDU_sumTable1")
                                     .FirstOrDefault();
                        CommentModel edusumrescomment = _context.DbComment
                                     .Where(u => u.Comment_type == "EDU_Sumresultpending")
                                     .FirstOrDefault();
                        CommentModel edusumattcomment = _context.DbComment
                                     .Where(u => u.Comment_type == "EDU_SumAttendance")
                                     .FirstOrDefault();
                        string edustartdate = "<not provided>";
                        string eduenddate = "<not provided>";
                        string edugraddate = "<not provided>";
                        string edustartyr = "<not provided>";
                        string eduendyr = "<not provided>";
                        string edugradyr = "<not provided>";
                        if (uSDDIndividual.educationModels[i].Edu_History.ToString().Equals("Yes"))
                        {
                            try
                            {
                                if (!uSDDIndividual.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && !uSDDIndividual.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                                {
                                    edustartdate = uSDDIndividual.educationModels[i].Edu_StartDateMonth + " " + uSDDIndividual.educationModels[i].Edu_StartDateDay + ", " + uSDDIndividual.educationModels[i].Edu_StartDateYear;
                                    edustartyr = uSDDIndividual.educationModels[i].Edu_StartDateYear;
                                    //educationModel.Edu_StartDate = edustartdate;

                                }
                                if (uSDDIndividual.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && uSDDIndividual.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && uSDDIndividual.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                                {
                                    edustartdate = "<not provided>";
                                    edustartyr = "<not provided>";
                                }
                                if (uSDDIndividual.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && !uSDDIndividual.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                                {
                                    edustartdate = uSDDIndividual.educationModels[i].Edu_StartDateMonth + " " + uSDDIndividual.educationModels[i].Edu_StartDateYear;
                                    edustartyr = uSDDIndividual.educationModels[i].Edu_StartDateYear;
                                    //educationModel.Edu_StartDate = edustartdate;
                                }
                                if (uSDDIndividual.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && uSDDIndividual.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                                {
                                    edustartdate = uSDDIndividual.educationModels[i].Edu_StartDateYear;
                                    edustartyr = uSDDIndividual.educationModels[i].Edu_StartDateYear;
                                    //educationModel.Edu_StartDate = edustartdate;
                                }

                                if (!uSDDIndividual.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && !uSDDIndividual.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                                {
                                    eduenddate = uSDDIndividual.educationModels[i].Edu_EndDateMonth + " " + uSDDIndividual.educationModels[i].Edu_EndDateDay + ", " + uSDDIndividual.educationModels[i].Edu_EndDateYear;
                                    eduendyr = uSDDIndividual.educationModels[i].Edu_EndDateYear;
                                    //educationModel.Edu_EndDate = eduenddate;
                                }
                                if (uSDDIndividual.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && uSDDIndividual.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && uSDDIndividual.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                                {
                                    eduenddate = "<not provided>";
                                    eduendyr = "<not provided>";
                                }
                                if (uSDDIndividual.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && !uSDDIndividual.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                                {
                                    eduenddate = uSDDIndividual.educationModels[i].Edu_EndDateMonth + " " + uSDDIndividual.educationModels[i].Edu_EndDateYear;
                                    eduendyr = uSDDIndividual.educationModels[i].Edu_EndDateYear;
                                    //educationModel.Edu_EndDate = eduenddate;
                                }
                                if (uSDDIndividual.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && uSDDIndividual.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                                {
                                    eduenddate = uSDDIndividual.educationModels[i].Edu_EndDateYear;
                                    eduendyr = uSDDIndividual.educationModels[i].Edu_EndDateYear;
                                    //educationModel.Edu_EndDate = eduenddate;
                                }

                                if (!uSDDIndividual.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && !uSDDIndividual.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                                {
                                    edugraddate = uSDDIndividual.educationModels[i].Edu_GradDateMonth + " " + uSDDIndividual.educationModels[i].Edu_GradDateDay + ", " + uSDDIndividual.educationModels[i].Edu_GradDateYear;
                                    edugradyr = uSDDIndividual.educationModels[i].Edu_GradDateYear;
                                    // educationModel.Edu_Graddate = edugraddate;
                                }
                                if (uSDDIndividual.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && uSDDIndividual.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && uSDDIndividual.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                                {
                                    edugraddate = "<not provided>";
                                    edugradyr = "<not provided>";
                                }
                                if (uSDDIndividual.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && !uSDDIndividual.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                                {
                                    edugraddate = uSDDIndividual.educationModels[i].Edu_GradDateMonth + " " + uSDDIndividual.educationModels[i].Edu_GradDateYear;
                                    edugradyr = uSDDIndividual.educationModels[i].Edu_GradDateYear;
                                    //educationModel.Edu_Graddate = edugraddate;
                                }
                                if (uSDDIndividual.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && uSDDIndividual.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !uSDDIndividual.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                                {
                                    edugraddate = uSDDIndividual.educationModels[i].Edu_GradDateYear;
                                    edugradyr = uSDDIndividual.educationModels[i].Edu_GradDateYear;
                                    //educationModel.Edu_Graddate = edugraddate;
                                }
                                if (i == uSDDIndividual.educationModels.Count - 1)
                                {
                                    try
                                    {
                                        switch (uSDDIndividual.educationModels[i].Edu_Confirmed.ToString())
                                        {
                                            case "Yes":
                                                edu_comment = string.Concat(edu_comment, comment1.confirmed_comment.ToString(), "\n\n");
                                                //edu_header = string.Concat(edu_header, uSDDIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Confirmed");
                                                edu_header = string.Concat(edu_header, "Confirmed");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.confirmed_comment.ToString(), "\n\n");
                                                break;
                                            case "No":
                                                edu_comment = string.Concat(edu_comment, comment1.unconfirmed_comment.ToString(), "\n\n");
                                                //edu_header = string.Concat(edu_header, uSDDIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Unconfirmed");
                                                edu_header = string.Concat(edu_header, "Unconfirmed");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.unconfirmed_comment.ToString(), "\n\n");
                                                break;
                                            case "Result Pending":
                                                edu_comment = string.Concat(edu_comment, edurescomment.confirmed_comment.ToString(), "APPENDEDURESULTPEND", i.ToString(), "\n\n");
                                                //edu_header = string.Concat(edu_header, uSDDIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Results Pending");
                                                edu_header = string.Concat(edu_header, "Results Pending");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.confirmed_comment.ToString(), " APPENDRESPENDING", "", "\n\n");
                                                break;
                                            case "Attendance Confirmed":
                                                edu_comment = string.Concat(edu_comment, eduattcomment.confirmed_comment.ToString(), "\n\n");
                                                //edu_header = string.Concat(edu_header, uSDDIndividual.educationModels[i].Edu_Degree.ToString(), " - ", " Confirmed");
                                                edu_header = string.Concat(edu_header, "Confirmed");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.confirmed_comment.ToString(), "\n\n");
                                                break;
                                            case "Attendance Unconfirmed":
                                                edu_comment = string.Concat(edu_comment, eduattcomment.unconfirmed_comment.ToString(), "\n\n");
                                                //edu_header = string.Concat(edu_header, uSDDIndividual.educationModels[i].Edu_Degree.ToString(), " - ", " Unconfirmed");
                                                edu_header = string.Concat(edu_header, "Unconfirmed");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.unconfirmed_comment.ToString(), "\n\n");
                                                break;
                                            case "Attendance Result pending":
                                                edu_comment = string.Concat(edu_comment, edurescomment.unconfirmed_comment.ToString(), "APPENDEDUATTRESULTPEND", i.ToString(), "\n\n");
                                                //edu_header = string.Concat(edu_header, uSDDIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Results Pending");
                                                edu_header = string.Concat(edu_header, "Results Pending");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.unconfirmed_comment.ToString(), " ATTRESPENDING", "\n\n");
                                                break;
                                        }
                                    }
                                    catch { }
                                }
                                else
                                {
                                    try
                                    {
                                        switch (uSDDIndividual.educationModels[i].Edu_Confirmed.ToString())
                                        {
                                            case "Yes":
                                                edu_comment = string.Concat(edu_comment, comment1.confirmed_comment.ToString(), "\n\n");
                                                edu_header = string.Concat(edu_header, "Confirmed", "\n\n");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.confirmed_comment.ToString(), "\n\n");
                                                break;
                                            case "No":
                                                edu_comment = string.Concat(edu_comment, comment1.unconfirmed_comment.ToString(), "\n\n");
                                                edu_header = string.Concat(edu_header, "Unconfirmed", "\n\n");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.unconfirmed_comment.ToString(), "\n\n");
                                                break;
                                            case "Result Pending":
                                                edu_comment = string.Concat(edu_comment, edurescomment.confirmed_comment.ToString(), "APPENDEDURESULTPEND", i.ToString(), "\n\n");
                                                edu_header = string.Concat(edu_header, "Results Pending", "\n\n");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.confirmed_comment.ToString(), " APPENDRESPENDING", "", "\n\n");
                                                break;
                                            case "Attendance Confirmed":
                                                edu_comment = string.Concat(edu_comment, eduattcomment.confirmed_comment.ToString(), "\n\n");
                                                edu_header = string.Concat(edu_header, "Confirmed", "\n\n");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.confirmed_comment.ToString(), "\n\n");
                                                break;
                                            case "Attendance Unconfirmed":
                                                edu_comment = string.Concat(edu_comment, eduattcomment.unconfirmed_comment.ToString(), "\n\n");
                                                edu_header = string.Concat(edu_header, "Unconfirmed", "\n\n");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.unconfirmed_comment.ToString(), "\n\n");
                                                break;
                                            case "Attendance Result pending":
                                                edu_comment = string.Concat(edu_comment, edurescomment.unconfirmed_comment.ToString(), "APPENDEDUATTRESULTPEND", i.ToString(), "\n\n");
                                                edu_header = string.Concat(edu_header, "Results Pending", "\n\n");
                                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.unconfirmed_comment.ToString(), " ATTRESPENDING", "\n\n");
                                                break;
                                        }
                                    }
                                    catch { }
                                }
                                try
                                {
                                    edu_comment = edu_comment.Replace("[Degree]", uSDDIndividual.educationModels[i].Edu_Degree.ToString());
                                }
                                catch
                                {
                                    edu_comment = edu_comment.Replace("[Degree]", "<not provided>");
                                }
                                try
                                {
                                    edu_comment = edu_comment.Replace("[School]", uSDDIndividual.educationModels[i].Edu_School.ToString());
                                }
                                catch
                                {
                                    edu_comment = edu_comment.Replace("[School]", "<not provided>");
                                }
                                try
                                {
                                    edu_comment = edu_comment.Replace("[EduLocation]", uSDDIndividual.educationModels[i].Edu_Location.ToString());
                                }
                                catch
                                {
                                    edu_comment = edu_comment.Replace("[EduLocation]", "<not provided>");
                                }
                                edu_comment = edu_comment.Replace("[GradDate]", edugraddate);
                                edu_comment = edu_comment.Replace("[EduStartDate]", edustartdate);
                                edu_comment = edu_comment.Replace("[EduEndDate]", eduenddate);
                                try
                                {
                                    edu_summcomment = edu_summcomment.Replace("[Degree]", uSDDIndividual.educationModels[i].Edu_Degree.ToString());
                                }
                                catch
                                {
                                    edu_summcomment = edu_summcomment.Replace("[Degree]", "<not provided>");
                                }
                                try
                                {
                                    edu_summcomment = edu_summcomment.Replace("[School]", uSDDIndividual.educationModels[i].Edu_School.ToString());
                                }
                                catch
                                {
                                    edu_summcomment = edu_summcomment.Replace("[School]", "<not provided>");
                                }
                                try
                                {
                                    edu_summcomment = edu_summcomment.Replace("[EduLocation]", uSDDIndividual.educationModels[i].Edu_Location.ToString());
                                }
                                catch
                                {
                                    edu_summcomment = edu_summcomment.Replace("[EduLocation]", "<not provided>");
                                }
                                edu_summcomment = edu_summcomment.Replace("[GradDate]", edugradyr);
                                edu_summcomment = edu_summcomment.Replace("[EDUFROMDATE]", edustartyr);
                                edu_summcomment = edu_summcomment.Replace("[EDUTODATE]", eduendyr);
                            }
                            catch { }
                        }
                        else
                        {
                            edu_comment = "Research efforts did not identify any educational credentials in connection with [LastName], and additional information -- if any -- would be required from the subject in order to pursue confirmation of the same.\n\n";
                            edu_summcomment = comment1.other_comment.ToString();
                            edu_summcomment = edu_summcomment.Replace("[Last Name]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                            edu_header = "N/A";
                        }
                    }
                    edu_comment = edu_comment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                    doc.Replace("SUMMDEGDESCRIPTION", edu_summcomment.TrimEnd(), true, true);
                    doc.Replace("DEGREEDESCRIPTION", edu_comment, true, true);
                    doc.Replace("DEGREEHEADER", edu_header, true, true);
                    doc.SaveToFile(savePath);
                    table = doc.Sections[1].Tables[0] as Table;
                    cell1 = table.Rows[6].Cells[2];
                    cell2 = table.Rows[6].Cells[1];
                    foreach (Paragraph p1 in cell1.Paragraphs)
                    {
                        DocumentObject obj = null;
                        for (int k = 0; k < p1.ChildObjects.Count; k++)
                        {
                            obj = p1.ChildObjects[k];
                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                TextRange textRange = obj as TextRange;
                                string abc = textRange.Text;
                                // Find the word "Civil" in paragraph1
                                if (abc.ToString().EndsWith("APPENDRESPENDING"))
                                {
                                    textRange = p1.AppendText("DONEhowever efforts to independently verify the same are currently ongoing (the results of which will be provided under separate cover, if and when received)");
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
                                    textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                    doc.SaveToFile(savePath);
                                    //doc.Replace("APPENDRESPENDINGDONE", "", true, false);                            
                                    doc.SaveToFile(savePath);
                                    bnres = "true";
                                    break;
                                }
                                else
                                {
                                    if (abc.ToString().EndsWith("ATTRESPENDING"))
                                    {
                                        textRange = p1.AppendText("DONEefforts to independently verify the same are currently ongoing (the results of which will be provided under separate cover, if and when received)");
                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                        textRange.CharacterFormat.FontSize = 11;
                                        textRange.CharacterFormat.Italic = true;
                                        textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                        doc.SaveToFile(savePath);
                                        //doc.Replace("ATTRESPENDINGDONE", "", true, false);

                                        bnres = "true";
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    doc.Replace("APPENDRESPENDINGDONE", "", true, false);
                    doc.Replace("ATTRESPENDINGDONE", "", true, false);
                    doc.SaveToFile(savePath);
                    if (uSDDIndividual.educationModels[0].Edu_History.ToString().Equals("Yes"))
                    {
                        for (int i = 0; i < uSDDIndividual.educationModels.Count; i++)
                        {
                            bnres = "";
                            if (uSDDIndividual.educationModels[i].Edu_Confirmed.ToString().Equals("Result Pending") || uSDDIndividual.educationModels[i].Edu_Confirmed.ToString().Equals("Attendance Result pending"))
                            {
                                for (int j = 1; j < 8; j++)
                                {
                                    Section section = doc.Sections[j];
                                    foreach (Paragraph para in section.Paragraphs)
                                    {
                                        DocumentObject obj = null;
                                        for (int k = 0; k < para.ChildObjects.Count; k++)
                                        {
                                            obj = para.ChildObjects[k];
                                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                            {
                                                TextRange textRange = obj as TextRange;
                                                string abc = textRange.Text;
                                                // Find the word "Civil" in paragraph1
                                                if (abc.ToString().Contains("APPENDEDURESULTPEND"))
                                                {
                                                    textRange = para.AppendText(" Efforts to independently confirm the subject’s tenure at the same are currently ongoing, the results of which will be provided under separate cover, if and when received.");
                                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                    textRange.CharacterFormat.FontSize = 11;
                                                    textRange.CharacterFormat.Italic = true;
                                                    textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                                    doc.SaveToFile(savePath);
                                                    doc.Replace(string.Concat("APPENDEDURESULTPEND", i.ToString()), "", false, false);
                                                    doc.SaveToFile(savePath);
                                                    bnres = "true";
                                                    break;
                                                }
                                                if (abc.ToString().Contains("APPENDEDUATTRESULTPEND"))
                                                {
                                                    textRange = para.AppendText(" Efforts to independently verify the same are currently ongoing, the results of which will be provided under separate cover, if and when received.");
                                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                    textRange.CharacterFormat.FontSize = 11;
                                                    textRange.CharacterFormat.Italic = true;
                                                    textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                                    doc.SaveToFile(savePath);
                                                    doc.Replace(string.Concat("APPENDEDUATTRESULTPEND", i.ToString()), "", false, false);
                                                    doc.SaveToFile(savePath);
                                                    bnres = "true";
                                                    break;
                                                }
                                            }
                                        }
                                        if (bnres.Equals("true")) { break; }
                                    }
                                    if (bnres.Equals("true")) { break; }
                                }
                            }
                        }
                    }
                    doc.SaveToFile(savePath);
                    //PL License details
                    string pl_comment = "";
                    string plstartdate = "";
                    string plenddate = "";
                    for (int i = 0; i < uSDDIndividual.pllicenseModels.Count; i++)
                    {
                        plstartdate = "<not provided>";
                        plenddate = "<not provided>";
                        try
                        {
                            if (uSDDIndividual.pllicenseModels[i].General_PL_License.ToString().Equals("Yes"))
                            {
                                if (!uSDDIndividual.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && !uSDDIndividual.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                                {
                                    plstartdate = uSDDIndividual.pllicenseModels[i].PL_StartDateMonth + " " + uSDDIndividual.pllicenseModels[i].PL_StartDateDay + ", " + uSDDIndividual.pllicenseModels[i].PL_StartDateYear;
                                    //pllicenseModel.PL_Start_Date = plstartdate;

                                }
                                if (uSDDIndividual.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && uSDDIndividual.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && uSDDIndividual.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                                {
                                    plstartdate = "<not provided>";
                                }
                                if (uSDDIndividual.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && !uSDDIndividual.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                                {
                                    plstartdate = uSDDIndividual.pllicenseModels[i].PL_StartDateMonth + " " + uSDDIndividual.pllicenseModels[i].PL_StartDateYear;
                                    //pllicenseModel.PL_Start_Date = plstartdate;
                                }
                                if (uSDDIndividual.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && uSDDIndividual.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !uSDDIndividual.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                                {
                                    plstartdate = uSDDIndividual.pllicenseModels[i].PL_StartDateYear;
                                    //pllicenseModel.PL_Start_Date = plstartdate;
                                }

                                if (!uSDDIndividual.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && !uSDDIndividual.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                                {
                                    plenddate = uSDDIndividual.pllicenseModels[i].PL_EndDateMonth + " " + uSDDIndividual.pllicenseModels[i].PL_EndDateDay + ", " + uSDDIndividual.pllicenseModels[i].PL_EndDateYear;
                                    //pllicenseModel.PL_End_Date = plenddate;

                                }
                                if (uSDDIndividual.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && uSDDIndividual.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && uSDDIndividual.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                                {
                                    plenddate = "<not provided>";
                                }
                                if (uSDDIndividual.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && !uSDDIndividual.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                                {
                                    plenddate = uSDDIndividual.pllicenseModels[i].PL_EndDateMonth + " " + uSDDIndividual.pllicenseModels[i].PL_EndDateYear;
                                    //pllicenseModel.PL_End_Date = plenddate;
                                }
                                if (uSDDIndividual.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && uSDDIndividual.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !uSDDIndividual.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                                {
                                    plenddate = uSDDIndividual.pllicenseModels[i].PL_EndDateYear;
                                    //pllicenseModel.PL_End_Date = plenddate;
                                }

                                CommentModel comment2 = _context.DbComment
                                                .Where(u => u.Comment_type == "PL1")
                                                .FirstOrDefault();
                                string strplorgfont = string.Concat(uSDDIndividual.pllicenseModels[i].PL_Organization.ToString(), " CHANGEFONTHEADER");
                                if (i == 0) { pl_comment = "\n"; }
                                if (uSDDIndividual.pllicenseModels[i].PL_Confirmed.Equals("Yes"))
                                {
                                    if (uSDDIndividual.pllicenseModels.Count - 1 == i)
                                    {
                                        pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.confirmed_comment.ToString(), "\n\nThere was no disciplinary history identified in connection with the subject’s license. <investigator to modify if disciplinary history exists>");
                                    }
                                    else
                                    {
                                        pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.confirmed_comment.ToString(), "\n\nThere was no disciplinary history identified in connection with the subject’s license. <investigator to modify if disciplinary history exists>", "\n\n");
                                    }
                                }
                                else
                                {
                                    if (uSDDIndividual.pllicenseModels[i].PL_Confirmed.Equals("No"))
                                    {
                                        if (uSDDIndividual.pllicenseModels.Count - 1 == i)
                                        {
                                            pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.unconfirmed_comment.ToString());
                                        }
                                        else
                                        {
                                            pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.unconfirmed_comment.ToString(), "\n\n");
                                        }
                                    }
                                    else
                                    {
                                        if (uSDDIndividual.pllicenseModels.Count - 1 == i)
                                        {
                                            pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.other_comment.ToString(), "APPENDPLTEXT", i.ToString());
                                        }
                                        else
                                        {
                                            pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.other_comment.ToString(), "APPENDPLTEXT", i.ToString(), "\n\n");
                                        }
                                    }
                                }
                                pl_comment = pl_comment.Replace("[Last Name]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                                pl_comment = pl_comment.Replace("[PL Organization]", uSDDIndividual.pllicenseModels[i].PL_Organization);
                                pl_comment = pl_comment.Replace("[Professional License Type]", uSDDIndividual.pllicenseModels[i].PL_License_Type.ToString());
                                if (uSDDIndividual.pllicenseModels[i].PL_Number.Equals(""))
                                { pl_comment = pl_comment.Replace(", with a license number [PL Number]", ""); }
                                else
                                {
                                    pl_comment = pl_comment.Replace("[PL Number]", uSDDIndividual.pllicenseModels[i].PL_Number.ToString());
                                }

                                if (plenddate == "<not provided>" && plstartdate == "<not provided>")
                                {
                                    pl_comment = pl_comment.Replace(", which was issued on [PL Start Date] and is set to expire on [PL End Date], unless renewed", "");
                                    pl_comment = pl_comment.Replace(", which is valid from [PL Start Date] and is set to expire on [PL End Date], unless renewed", "");
                                }
                                else
                                {
                                    pl_comment = pl_comment.Replace("[PL Start Date]", plstartdate);
                                    pl_comment = pl_comment.Replace("[PL End Date]", plenddate);
                                }
                            }
                            else
                            {
                                pl_comment = "";
                            }
                        }
                        catch
                        {
                            pl_comment = "";
                        }
                    }
                    doc.Replace("PLLICENSEDESCRIPTION", pl_comment, true, false);
                    doc.SaveToFile(savePath);
                    if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                    {
                        for (int j = 1; j < 6; j++)
                        {
                            Section section = doc.Sections[j];
                            foreach (Paragraph para in section.Paragraphs)
                            {
                                DocumentObject obj = null;
                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                {
                                    obj = para.ChildObjects[i];
                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                    {
                                        TextRange textRange = obj as TextRange;
                                        string abc = textRange.Text;
                                        // Find the word "Civil" in paragraph1
                                        if (abc.ToString().EndsWith(" CHANGEFONTHEADER"))
                                        {
                                            textRange.CharacterFormat.FontName = "Calibri(Body)";
                                            textRange.CharacterFormat.FontSize = 11;
                                            textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                            textRange.CharacterFormat.Italic = true;
                                            doc.SaveToFile(savePath);
                                            break;
                                        }
                                    }
                                }

                            }

                        }
                    }
                    doc.SaveToFile(savePath);
                    doc.Replace("CHANGEFONTHEADER", "", true, false);
                    doc.SaveToFile(savePath);
                    for (int j = 1; j < 6; j++)
                    {
                        Section section = doc.Sections[j];
                        foreach (Paragraph para in section.Paragraphs)
                        {
                            DocumentObject obj = null;
                            for (int i = 0; i < para.ChildObjects.Count; i++)
                            {
                                obj = para.ChildObjects[i];
                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                {
                                    TextRange textRange = obj as TextRange;
                                    string abc = textRange.Text;
                                    // Find the word "Civil" in paragraph1
                                    if (abc.ToString().EndsWith("[CRIMFONTCHANGE]"))
                                    {
                                        textRange.CharacterFormat.FontName = "Calibri(Body)";
                                        textRange.CharacterFormat.FontSize = 11;
                                        textRange.CharacterFormat.Italic = true;
                                        textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                        doc.SaveToFile(savePath);
                                        break;
                                    }
                                    if (abc.ToString().EndsWith("[CIVILFONTCHANGE]"))
                                    {
                                        textRange.CharacterFormat.FontName = "Calibri(Body)";
                                        textRange.CharacterFormat.FontSize = 11;
                                        textRange.CharacterFormat.Italic = true;
                                        textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                        doc.SaveToFile(savePath);
                                        break;
                                    }
                                    if (abc.ToString().EndsWith("[BANKFONTCHANGE]"))
                                    {
                                        textRange.CharacterFormat.FontName = "Calibri(Body)";
                                        textRange.CharacterFormat.FontSize = 11;
                                        textRange.CharacterFormat.Italic = true;
                                        textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                        doc.SaveToFile(savePath);
                                    }
                                }
                            }
                        }
                    }
                    doc.Replace("[CRIMFONTCHANGE]", "", true, false);
                    doc.Replace("[CIVILFONTCHANGE]", "", true, false);
                    doc.Replace("[BANKFONTCHANGE]", "", true, false);
                    doc.SaveToFile(savePath);
                    if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                    {
                        for (int i = 0; i < uSDDIndividual.pllicenseModels.Count; i++)
                        {
                            bnres = "";
                            if (uSDDIndividual.pllicenseModels[i].PL_Confirmed.ToString().Equals("Result Pending"))
                            {
                                for (int j = 1; j < 6; j++)
                                {
                                    Section section = doc.Sections[j];
                                    foreach (Paragraph para in section.Paragraphs)
                                    {
                                        DocumentObject obj = null;
                                        for (int k = 0; k < para.ChildObjects.Count; k++)
                                        {
                                            obj = para.ChildObjects[k];
                                            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                            {
                                                TextRange textRange = obj as TextRange;
                                                string abc = textRange.Text;
                                                // Find the word "Civil" in paragraph1
                                                if (abc.ToString().Contains("APPENDPLTEXT"))
                                                {
                                                    textRange = para.AppendText(" Efforts to independently verify the same are currently ongoing, the results of which will be provided under separate cover, if and when received.");
                                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                    textRange.CharacterFormat.FontSize = 11;
                                                    textRange.CharacterFormat.Italic = true;
                                                    textRange.CharacterFormat.HighlightColor = Color.Aqua;
                                                    doc.SaveToFile(savePath);
                                                    doc.Replace(string.Concat("APPENDPLTEXT", i.ToString()), "", false, false);
                                                    doc.Replace("[PLLICENSEDESC]", "[PLLICENSEDESC1]", true, false);
                                                    doc.SaveToFile(savePath);
                                                    bnres = "true";
                                                    break;
                                                }
                                            }
                                        }
                                        if (bnres.Equals("true")) { break; }
                                    }
                                    if (bnres.Equals("true")) { break; }
                                }
                                if (i == uSDDIndividual.pllicenseModels.Count - 1)
                                {
                                    bnres = "";
                                    for (int j = 1; j < 6; j++)
                                    {
                                        Section section = doc.Sections[j];
                                        foreach (Paragraph para in section.Paragraphs)
                                        {
                                            DocumentObject obj = null;
                                            for (int k = 0; k < para.ChildObjects.Count; k++)
                                            {
                                                obj = para.ChildObjects[k];
                                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                                {
                                                    TextRange textRange = obj as TextRange;
                                                    string abc = textRange.Text;
                                                    if (abc.ToString().Contains("[PLLICENSEDESC1]"))
                                                    {
                                                        textRange = para.AppendText("[PLLICENSEDESC]");
                                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                        textRange.CharacterFormat.FontSize = 11;
                                                        doc.SaveToFile(savePath);
                                                        doc.Replace(string.Concat("APPENDPLTEXT", i.ToString()), "", false, false);
                                                        doc.Replace("[PLLICENSEDESC1]", "", true, false);
                                                        doc.SaveToFile(savePath);
                                                        bnres = "true";
                                                        break;
                                                    }
                                                }
                                            }
                                            if (bnres.Equals("true")) { break; }
                                        }
                                        if (bnres.Equals("true")) { break; }
                                    }
                                }
                            }
                        }
                    }
                    //EDUCATIONALANDLICENSINGHITS
                    string eduhistory = ""; string plhistory = "";
                    string drivhistory = "";
                    string strEDULICENcomment = "";
                    try
                    {
                        if (uSDDIndividual.educationModels[0].Edu_History.ToString().Equals("Yes"))
                        {
                            eduhistory = "Additionally, efforts included verification of the subject’s educational credentials, where available";
                            strEDULICENcomment = eduhistory;
                        }
                    }
                    catch { }
                    try
                    {
                        if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                        {
                            plhistory = "Additionally, efforts included verification of the subject’s licensing credentials, where available";
                            strEDULICENcomment = plhistory;
                        }
                    }
                    catch { }
                    if (uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_FINRA.ToString().StartsWith("Yes"))
                    {
                        plhistory = "Additionally, efforts included verification of the subject’s licensing credentials, where available";
                        strEDULICENcomment = plhistory;
                    }
                    if (eduhistory != "" && plhistory != "")
                    {
                        strEDULICENcomment = "Additionally, efforts included verification of the subject’s educational and licensing credentials, where available";
                    }
                    if (uSDDIndividual.otherdetails.Has_Driving_Hits.ToUpper().ToString() != "No Records, No Consent".ToUpper())
                    {
                        if (strEDULICENcomment == "")
                        {
                            strEDULICENcomment = "Additionally, efforts included a review of the subject's driving history.";
                            drivhistory = strEDULICENcomment;
                        }
                        else
                        {
                            drivhistory = string.Concat(strEDULICENcomment, ", as well as a review of [LastName]’s driving history.");
                        }

                    }
                    if (uSDDIndividual.otherdetails.Was_credited_obtained.ToUpper().ToString() != "No".ToUpper())
                    {
                        if (strEDULICENcomment == "")
                        {
                            if (drivhistory != "")
                            {
                                strEDULICENcomment = "Additionally, efforts included a review of the subject's credit and driving histories.";
                            }
                            else
                            {
                                strEDULICENcomment = "Additionally, efforts included a review of the subject's credit history.";
                            }
                        }
                        else
                        {
                            if (eduhistory != "" || plhistory != "")
                            {
                                if (drivhistory != "")
                                {
                                    strEDULICENcomment = string.Concat(strEDULICENcomment, ", as well as a review of [LastName]’s credit and driving histories.");
                                }
                                else
                                {
                                    strEDULICENcomment = string.Concat(strEDULICENcomment, ", as well as a review of [LastName]’s credit history.");
                                }
                            }
                        }
                    }
                    else
                    {
                        if (drivhistory != "")
                        {
                            strEDULICENcomment = drivhistory;
                        }
                    }
                    if (strEDULICENcomment.EndsWith("where available")) { strEDULICENcomment = string.Concat(strEDULICENcomment, "."); }
                    doc.Replace("EDUCATIONALANDLICENSINGHITS", strEDULICENcomment, true, false);

                    //Results Pending in Legal Records Section?
                    if (uSDDIndividual.otherdetails.HasBankruptcyRecHits_resultpending == true || uSDDIndividual.otherdetails.Has_CriminalRecHit_resultpending == true || uSDDIndividual.otherdetails.has_civil_resultpending == true)
                    {
                        doc.Replace("[RESULTSPENSECTIONS1]", "\n\n<Search type> searches are currently ongoing in <jurisdiction>, the results of which will be provided under separate cover upon receipt.", true, false);
                    }
                    else
                    {
                        doc.Replace("[RESULTSPENSECTIONS1]", "", true, false);
                    }
                    if (uSDDIndividual.otherdetails.HasBankruptcyRecHits.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_CriminalRecHit.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Bureau_PrisonHit.ToString().Equals("Yes") || uSDDIndividual.otherdetails.Has_Sex_Offender_RegHit.ToString().Equals("Yes") || uSDDIndividual.otherdetails.Has_Civil_Records.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Name_Only == true || uSDDIndividual.otherdetails.Has_US_Tax_Court_Hit.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Property_Records.ToString().StartsWith("Yes"))
                    {
                        uSDDIndividual.otherdetails.Has_Legal_Records_Hits = "Yes";
                    }
                    else
                    {
                        uSDDIndividual.otherdetails.Has_Legal_Records_Hits = "No";
                    }
                    if (uSDDIndividual.otherdetails.Has_Legal_Records_Hits.ToString().Equals("Yes") || uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Yes-With Adverse") || regflag == "Records" || uSDDIndividual.otherdetails.RegulatoryFlag.ToString().Equals("Yes") || uSDDIndividual.otherdetails.Global_Security_Hits.ToString().Equals("Yes") || uSDDIndividual.otherdetails.ICIJ_Hits.ToString().Equals("Yes"))
                    {
                        uSDDIndividual.otherdetails.Has_Regulatory_Hits = "Yes";
                    }
                    else
                    {
                        uSDDIndividual.otherdetails.Has_Regulatory_Hits = "No";
                    }
                    if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Yes-With Adverse") || regflag == "Records" || uSDDIndividual.otherdetails.RegulatoryFlag.ToString().Equals("Yes") || uSDDIndividual.otherdetails.Global_Security_Hits.ToString().Equals("Yes") || uSDDIndividual.otherdetails.ICIJ_Hits.ToString().Equals("Yes"))
                    {
                        uSDDIndividual.otherdetails.Has_Hits_Above = "Yes";
                    }
                    else
                    {
                        uSDDIndividual.otherdetails.Has_Hits_Above = "No";
                    }

                    //Legal_Record_Judgments_Liens_Hits
                    string Legrechitcommentmodel = "";
                    CommentModel Legrechitcommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "Legal_Record_Judgments_Liens_Hits")
                                    .FirstOrDefault();
                    if (uSDDIndividual.otherdetails.Has_Legal_Records_Hits.ToString().Equals("Yes"))
                    {
                        Legrechitcommentmodel = Legrechitcommentmodel1.confirmed_comment.ToString();
                    }
                    else
                    {
                        Legrechitcommentmodel = Legrechitcommentmodel1.unconfirmed_comment.ToString();
                    }
                    doc.Replace("HasLegalRecordsJudgmentsorLiensHits", Legrechitcommentmodel, true, true);
                    //Has Regulatory or Global Security Hits?
                    string hasreghitcommentmodel = "";
                    CommentModel hasreghitcommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "Regulatory_Global_Security_Hits")
                                    .FirstOrDefault();
                    if (regflag == "Records")
                    {
                        hasreghitcommentmodel = hasreghitcommentmodel1.confirmed_comment.ToString();
                    }
                    else
                    {
                        hasreghitcommentmodel = hasreghitcommentmodel1.unconfirmed_comment.ToString();
                    }
                    hasreghitcommentmodel = hasreghitcommentmodel.Replace(" and the United States", "");
                    doc.Replace("HasRegulatoryorGlobalSecurityHits", hasreghitcommentmodel, true, true);
                    string hitcompcommentmodel = "";
                    string searchtext = "";
                    CommentModel hitcompcommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "Hits_with_Companion_Report")
                                    .FirstOrDefault();
                    CommentModel nohitcompcommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "No_Hits_with_Companion_Report")
                                    .FirstOrDefault();
                    string executivesumhitabove = "";
                    string searchtext2 = "";
                    if (regflag == "Records")
                    {
                        executivesumhitabove = "Yes";
                    }
                    else
                    {
                        executivesumhitabove = "No";
                    }
                    if (executivesumhitabove == "Yes")
                    {
                        if (uSDDIndividual.otherdetails.Has_Companion_Report.ToString().Equals("Yes"))
                        {
                            hitcompcommentmodel = hitcompcommentmodel1.confirmed_comment.ToString();

                            searchtext2 = "In sum, with the exception of the above, no other issues of potential relevance were identified in connection with";
                        }
                        else
                        {
                            hitcompcommentmodel = hitcompcommentmodel1.unconfirmed_comment.ToString();
                        }
                    }
                    else
                    {
                        if (uSDDIndividual.otherdetails.Has_Companion_Report.ToString().Equals("Yes"))
                        {
                            hitcompcommentmodel = nohitcompcommentmodel1.confirmed_comment.ToString();
                            searchtext = "In sum, no other issues of potential relevance were identified in connection with";
                        }
                        else
                        {
                            hitcompcommentmodel = nohitcompcommentmodel1.unconfirmed_comment.ToString();
                        }

                    }
                    hitcompcommentmodel = hitcompcommentmodel.Replace("[FirstName]", uSDDIndividual.diligenceInputModel.FirstName.ToString());
                    // hitcompcommentmodel = hitcompcommentmodel.Replace("[MiddleInitial]", uSDDIndividual.diligenceInputModel.MiddleInitial.ToString());
                    doc.Replace("HasHitsAboveAndCompanionReport", hitcompcommentmodel, true, true);
                    doc.SaveToFile(savePath);
                    if (searchtext.ToString().Equals("")) { }
                    else
                    {
                        string blnresult = "";
                        for (int j = 1; j < 6; j++)
                        {
                            Section section = doc.Sections[j];
                            foreach (Paragraph para in section.Paragraphs)
                            {
                                DocumentObject obj = null;
                                for (int i = 0; i < para.ChildObjects.Count; i++)
                                {
                                    obj = para.ChildObjects[i];
                                    if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                    {
                                        TextRange textRange = obj as TextRange;
                                        string abc = textRange.Text;
                                        // Find the word "Civil" in paragraph1
                                        if (abc.ToString().Contains(searchtext))
                                        {
                                            textRange = para.AppendText("  A Companion Report was prepared in connection with <Companion Subject> and was provided under separate cover. <Investigator to modify for multiple Companion Reports if applicable>");
                                            textRange.CharacterFormat.FontName = "Calibri (Body)";
                                            textRange.CharacterFormat.FontSize = 11;
                                            textRange.CharacterFormat.Italic = false;
                                            doc.SaveToFile(savePath);
                                            blnresult = "true";
                                            break;
                                        }
                                        if (abc.ToString().Contains(searchtext2))
                                        {
                                            textRange = para.AppendText("  A Companion Report was prepared in connection with <Companion Subject> and was provided under separate cover. <Investigator to modify for multiple Companion Reports if applicable>");
                                            textRange.CharacterFormat.FontName = "Calibri (Body)";
                                            textRange.CharacterFormat.FontSize = 11;
                                            //textRange.CharacterFormat.Italic = false;
                                            doc.SaveToFile(savePath);
                                            blnresult = "true";
                                            break;
                                        }
                                    }
                                }
                                if (blnresult.Equals("true")) { break; }
                            }
                            if (blnresult.Equals("true")) { break; }
                        }
                    }
                    if (current_full_address.ToString().Equals("")) { doc.Replace("[CurrentFullAddress]", "<not provided>", true, true); }
                    else
                    {
                        doc.Replace("[CurrentFullAddress]", current_full_address, true, true);
                    }
                    doc.Replace("[FirstName]", uSDDIndividual.diligenceInputModel.FirstName, true, true);
                    try
                    {
                        if (uSDDIndividual.diligenceInputModel.MiddleName == "") { doc.Replace(" [MiddleName]", "", true, false); }
                        else
                        {
                            doc.Replace("[MiddleName]", uSDDIndividual.diligenceInputModel.MiddleName, true, true);
                        }
                    }
                    catch
                    {
                        doc.Replace(" [MiddleName]", "", true, false);
                    }
                    doc.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName, true, true);
                    doc.Replace("[Country]", "the United States", true, false);
                    doc.SaveToFile(savePath);
                    string blnresultfound = "";
                    for (int j = 0; j < 8; j++)
                    {
                        Section section = doc.Sections[j];
                        foreach (Paragraph para in section.Paragraphs)
                        {
                            DocumentObject obj = null;
                            for (int i = 0; i < para.ChildObjects.Count; i++)
                            {
                                obj = para.ChildObjects[i];
                                if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                                {
                                    TextRange textRange = obj as TextRange;
                                    string abc = textRange.Text;
                                    // Find the word "Spire.Doc" in paragraph1
                                    if (abc.ToString().Contains("(“OFAC”),") || abc.ToString().EndsWith("(“OFAC”),"))
                                    {
                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                        //Insert footnote1 after the word "Spire.Doc"
                                        para.ChildObjects.Insert(i + 1, footnote1);
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("It should be emphasized that updates are made to these lists on a periodic and irregular basis, and for purposes of preparing this memorandum, these lists were searched on <investigator to insert date of research>.  ");
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        //Append the line
                                        string strglobaltextappended = "";
                                        if (uSDDIndividual.otherdetails.Global_Security_Hits.Equals("Yes"))
                                        {
                                            strglobaltextappended = " and the following information was identified in connection with [LastName]: ";
                                        }
                                        else
                                        {
                                            strglobaltextappended = " and it is noted that [LastName] was not identified on any of these lists.";
                                        }
                                        strglobaltextappended = strglobaltextappended.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName);
                                        TextRange tr = para.AppendText(strglobaltextappended);
                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                        tr.CharacterFormat.FontSize = 11;
                                        doc.SaveToFile(savePath);
                                        blnresultfound = "true";
                                        break;
                                    }
                                }
                            }
                            if (blnresultfound == "true")
                            {
                                break;
                            }
                        }
                        if (blnresultfound == "true")
                        {
                            break;
                        }
                    }
                    doc.SaveToFile(savePath);

                    if (regflag == "Records")
                    {
                        doc.Replace("REGCOMMENT", "<investigator to insert regulatory hits here>", true, true);
                        doc.Replace("REGRESULT", "Records", true, true);
                    }
                    else
                    {
                        doc.Replace("REGCOMMENT", "", true, true);
                        doc.Replace("REGRESULT", "Clear", true, true);
                    }
                    doc.Replace("[First Name]", uSDDIndividual.diligenceInputModel.FirstName, true, false);
                    doc.SaveToFile(savePath);

                    string strPLComment = "";
                    if (uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse") || uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse") || uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                    {
                        if (uSDDIndividual.pllicenseModels[0].General_PL_License.Equals("Yes"))
                        {
                            strPLComment = "[LastName] holds a [ProfessionalLicenseType1] license with the [PLOrganization1]";
                        }
                        if (uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. SEC");
                        }
                        if (uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.K. FCA");
                        }
                        if (uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. FINRA");
                        }
                        if (uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. NFA");
                        }
                        if (uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with HKSFC");
                        }

                        if (uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse"))
                        {
                            strPLComment = string.Concat(strPLComment, "\n\nFurther, <investigator to insert regulatory hits here>");
                        }
                        strPLComment = strPLComment.Replace("[ProfessionalLicenseType1]", uSDDIndividual.pllicenseModels[0].PL_License_Type);
                        strPLComment = strPLComment.Replace("[PLOrganization1]", uSDDIndividual.pllicenseModels[0].PL_Organization);
                        doc.Replace("PROFRESULT", "Records", true, true);
                    }
                    else
                    {
                        strPLComment = "";
                        doc.Replace("PROFRESULT", "Clear", true, true);
                    }
                   doc.Replace("PROFCOMMENT", strPLComment, true, true);
                    doc.SaveToFile(savePath);
                }
                    HttpContext.Session.SetString("first_name", uSDDIndividual.diligenceInputModel.FirstName);
                    HttpContext.Session.SetString("last_name", uSDDIndividual.diligenceInputModel.LastName);
                    HttpContext.Session.SetString("case_number", uSDDIndividual.diligenceInputModel.CaseNumber);
                    HttpContext.Session.SetString("middleinitial", uSDDIndividual.diligenceInputModel.MiddleInitial);
                   
                    HttpContext.Session.SetString("pl_generallicense", uSDDIndividual.pllicenseModels[0].General_PL_License.ToString());                   
                //SummaryResulttableModel ST = new SummaryResulttableModel();
                //ST = uSDDIndividual.summarymodel;

                return Save_Page2(uSDDIndividuals);
            }
            catch (IOException e)
            {
                string recordid = HttpContext.Session.GetString("recordid");
                HttpContext.Session.SetString("recordid", recordid);
                TempData["message"] = e;
                return US_Citizen_Family_Newpage();
            }
        }       
        public IActionResult Save_Page2(MainModel US_CITI)
        {
            string record_Id = HttpContext.Session.GetString("record_Id");
            string first_name = HttpContext.Session.GetString("first_name");
            string last_name = HttpContext.Session.GetString("last_name");
            string country = "the United States";
            string add_states = HttpContext.Session.GetString("additionalstates");
            //string country = HttpContext.Session.GetString("country");
            string case_number = HttpContext.Session.GetString("case_number");
            string city = HttpContext.Session.GetString("city");
            //string employer1City = HttpContext.Session.GetString("employer1City");
            //string employer1State = HttpContext.Session.GetString("employer1State");
            string employer1 = HttpContext.Session.GetString("employer1");
            string emp_position1 = HttpContext.Session.GetString("emp_position1");
            string emp_startdate1 = HttpContext.Session.GetString("emp_startdate1");
            string pl_generallicense = HttpContext.Session.GetString("pl_generallicense");
            string pl_licensetype = "";
            string pl_organization = "";
            string pl_startdate = "";
            string pl_enddate = "";
            if (pl_generallicense.Equals("Yes"))
            {
                pl_licensetype = HttpContext.Session.GetString("pl_licensetype");
                pl_organization = HttpContext.Session.GetString("pl_organization");
                pl_startdate = HttpContext.Session.GetString("pl_startdate");
                // pl_startdate= (Convert.ToDateTime(pl_startdate.ToString()).Year).ToString();
                pl_enddate = HttpContext.Session.GetString("pl_enddate");
                // pl_enddate = (Convert.ToDateTime(pl_enddate.ToString()).Year).ToString();
            }
            string MiddleInitial = HttpContext.Session.GetString("MiddleInitial");
            string savePath = _config.GetValue<string>("ReportPath");
            savePath = string.Concat(savePath, last_name.ToString(), "_SterlingDiligenceReport(", case_number.ToString(), ")_DRAFT.docx");
            Document doc = new Document(savePath);
            string strcomment;
            Table table = doc.Sections[1].Tables[0] as Table;
           
            int adult_count = 0;

            for (int famcount = 0; famcount < US_CITI.familyModels.Count; famcount++)
            {
                if (US_CITI.familyModels[famcount].adult_minor == "Adult")
                {
                    adult_count = famcount + 1;

                    //SummaryResulttableModel ST = new SummaryResulttableModel();
                    SummaryResulttableModel ST = _context.summaryResulttableModels
                        .Where(u => u.record_Id == US_CITI.familyModels[famcount].Family_record_id)
                        .FirstOrDefault();
                    DiligenceInputModel DI = _context.DbPersonalInfo
                        .Where(u => u.record_Id == US_CITI.familyModels[famcount].Family_record_id)
                        .FirstOrDefault();
                    try
                    {
                        if (US_CITI.familyModels[famcount + 1].adult_minor == "Adult")
                        {
                            doc.Replace("NAMECOMMENT", "NAMECOMMENT\n\nNAME2COMMENT", true, true);
                            doc.Replace("NAMERESULT", "NAMERESULT\n\nNAME2RESULT", true, true);
                            doc.Replace("PICOMMENT", "PICOMMENT\n\nPI2COMMENT", true, true);
                            doc.Replace("PLRESULT", "PLRESULT\n\nPL2RESULT", true, true);
                            doc.Replace("REGCOMMENT", "REGCOMMENT\n\nREG2COMMENT", true, true);
                            doc.Replace("REGRESULT", "REGRESULT\n\nREG2RESULT", true, true);
                            doc.Replace("BANKRUPCOMMENT", "BANKRUPCOMMENT\n\nBANKRUP2COMMENT", true, true);
                            doc.Replace("BANKRUPRESULT", "BANKRUPRESULT\n\nBANKRUP2RESULT", true, true);
                            doc.Replace("CIVILCOMMENT", "CIVILCOMMENT\n\nCIVIL2COMMENT", true, true);
                            doc.Replace("CIVILRESULT", "CIVILRESULT\n\nCIVIL2RESULT", true, true);
                            doc.Replace("SOCSECCOMMENT", "SOCSECCOMMENT\n\nSOCSEC2COMMENT", true, true);
                            doc.Replace("SOCSECRESULT", "SOCSECRESULT\n\nSOCSEC2RESULT", true, true);
                            doc.Replace("DEPTFORCOMMENT", "DEPTFORCOMMENT\n\nDEPTFOR2COMMENT", true, true);
                            doc.Replace("DEPTFORRESULT", "DEPTFORRESULT\n\nDEPTFOR2RESULT", true, true);
                            doc.Replace("EURUNCOMMENT", "EURUNCOMMENT\n\nEURUN2COMMENT", true, true);
                            doc.Replace("EURUNRESULT", "EURUNRESULT\n\nEURUN2RESULT", true, true);
                            doc.Replace("HMTRECOMMENT", "HMTRECOMMENT\n\nHMTRE2COMMENT", true, true);
                            doc.Replace("HMTRERESULT", "HMTRERESULT\n\nHMTRE2RESULT", true, true);
                            doc.Replace("USBEUCOMMENT", "USBEUCOMMENT\n\nUSBEU2COMMENT", true, true);
                            doc.Replace("USBEURESULT", "USBEURESULT\n\nUSBEU2RESULT", true, true);
                            doc.Replace("USBDEPCOMMENT", "USBDEPCOMMENT\n\nUSBDEP2COMMENT", true, true);
                            doc.Replace("USBDEPRESULT", "USBDEPRESULT\n\nUSBDEP2RESULT", true, true);
                            doc.Replace("USBDIRCOMMENT", "USBDIRCOMMENT\n\nUSBDIR2COMMENT", true, true);
                            doc.Replace("USBDIRRESULT", "USBDIRRESULT\n\nUSBDIR2RESULT", true, true);
                            doc.Replace("USGENCOMMENT", "USGENCOMMENT\n\nUSGEN2COMMENT", true, true);
                            doc.Replace("USGENRESULT", "USGENRESULT\n\nUSGEN2RESULT", true, true);
                            doc.Replace("USOFFCOMMENT", "USOFFCOMMENT\n\nUSOFF2COMMENT", true, true);
                            doc.Replace("USOFFRESULT", "USOFFRESULT\n\nUSOFF2RESULT", true, true);
                            doc.Replace("UNCONSOCOMMENT", "UNCONSOCOMMENT\n\nUNCONSO2COMMENT", true, true);
                            doc.Replace("UNCONSORESULT", "UNCONSORESULT\n\nUNCONSO2RESULT", true, true);
                            doc.Replace("WORLDBANKCOMMENT", "WORLDBANKCOMMENT\n\nWORLDBANK2COMMENT", true, true);
                            doc.Replace("WORLDBANKRESULT", "WORLDBANKRESULT\n\nWORLDBANK2RESULT", true, true);
                            doc.Replace("INTERPOLCOMMENT", "INTERPOLCOMMENT\n\nINTERPOL2COMMENT", true, true);
                            doc.Replace("INTERPOLRESULT", "INTERPOLRESULT\n\nINTERPOL2RESULT", true, true);
                            doc.Replace("USFEDERALCOMMENT", "USFEDERALCOMMENT\n\nUSFEDERAL2COMMENT", true, true);
                            doc.Replace("USFEDERALRESULT", "USFEDERALRESULT\n\nUSFEDERAL2RESULT", true, true);
                            doc.Replace("USSECRETCOMMENT", "USSECRETCOMMENT\n\nUSSECRET2COMMENT", true, true);
                            doc.Replace("USSECRETRESULT", "USSECRETRESULT\n\nUSSECRET2RESULT", true, true);
                            doc.Replace("COMMODITYCOMMENT", "COMMODITYCOMMENT\n\nCOMMODITY2COMMENT", true, true);
                            doc.Replace("COMMODITYRESULT", "COMMODITYRESULT\n\nCOMMODITY2RESULT", true, true);
                            doc.Replace("FDEPOCOMMENT", "FDEPOCOMMENT\n\nFDEPO2COMMENT", true, true);
                            doc.Replace("FDEPORESULT", "FDEPORESULT\n\nFDEPO2RESULT", true, true);
                            doc.Replace("FRESCOMMENT", "FRESCOMMENT\n\nFRES2COMMENT", true, true);
                            doc.Replace("FRESRESULT", "FRESRESULT\n\nFRES2RESULT", true, true);
                            doc.Replace("FCRIMECOMMENT", "FCRIMECOMMENT\n\nFCRIME2COMMENT", true, true);
                            doc.Replace("FCRIMERESULT", "FCRIMERESULT\n\nFCRIME2RESULT", true, true);
                            doc.Replace("NATCRECOMMENT", "NATCRECOMMENT\n\nNATCRE2COMMENT", true, true);
                            doc.Replace("NATCRERESULT", "NATCRERESULT\n\nNATCRE2RESULT", true, true);
                            doc.Replace("NYCOMMENT", "NYCOMMENT\n\nNY2COMMENT", true, true);
                            doc.Replace("NYRESULT", "NYRESULT\n\nNY2RESULT", true, true);
                            doc.Replace("OFFCPTCOMMENT", "OFFCPTCOMMENT\n\nOFFCPT2COMMENT", true, true);
                            doc.Replace("OFFCPTRESULT", "OFFCPTRESULT\n\nOFFCPT2RESULT", true, true);
                            doc.Replace("OFFSUPCOMMENT", "OFFSUPCOMMENT\n\nOFFSUP2COMMENT", true, true);
                            doc.Replace("OFFSUPRESULT", "OFFSUPRESULT\n\nOFFSUP2RESULT", true, true);
                            doc.Replace("RESTCOMMENT", "RESTCOMMENT\n\nREST2COMMENT", true, true);
                            doc.Replace("RESTRESULT", "RESTRESULT\n\nREST2RESULT", true, true);
                            doc.Replace("USCOURCOMMENT", "USCOURCOMMENT\n\nUSCOUR2COMMENT", true, true);
                            doc.Replace("USCOURESULT", "USCOURESULT\n\nUSCOU2RESULT", true, true);
                            doc.Replace("USDPJSCOMMENT", "USDPJSCOMMENT\n\nUSDPJS2COMMENT", true, true);
                            doc.Replace("USDPJSRESULT", "USDPJSRESULT\n\nUSDPJS2RESULT", true, true);
                            doc.Replace("USFEDCOMMENT", "USFEDCOMMENT\n\nUSFED2COMMENT", true, true);
                            doc.Replace("USFEDRESULT", "USFEDRESULT\n\nUSFED2RESULT", true, true);
                            doc.Replace("USOTCOMMENT", "USOTCOMMENT\n\nUSOT2COMMENT", true, true);
                            doc.Replace("USOTRESULT", "USOTRESULT\n\nUSOT2RESULT", true, true);
                            doc.Replace("CICOMMENT", "CICOMMENT\n\nCI2COMMENT", true, true);
                            doc.Replace("CIRESULT", "CIRESULT\n\nCI2RESULT", true, true);
                            doc.Replace("CITYCOMMENT", "CITYCOMMENT\n\nCITY2COMMENT", true, true);
                            doc.Replace("CITYRESULT", "CITYRESULT\n\nCITY2RESULT", true, true);
                            doc.Replace("COSTABCOMMENT", "COSTABCOMMENT\n\nCOSTAB2COMMENT", true, true);
                            doc.Replace("COSTABRESULT", "COSTABRESULT\n\nCOSTAB2RESULT", true, true);
                            doc.Replace("HAMPSPCOMMENT", "HAMPSPCOMMENT\n\nHAMPSP2COMMENT", true, true);
                            doc.Replace("HAMPSPRESULT", "HAMPSPRESULT\n\nHAMPSP2RESULT", true, true);
                            doc.Replace("HONGKONGCOMMENT", "HONGKONGCOMMENT\n\nHONGKONG2COMMENT", true, true);
                            doc.Replace("HONGKONGRESULT", "HONGKONGRESULT\n\nHONGKONG2RESULT", true, true);
                            doc.Replace("METROPOLICOMMENT", "METROPOLICOMMENT\n\nMETROPOLI2COMMENT", true, true);
                            doc.Replace("METROLIRESULT", "METROLIRESULT\n\nMETROLI2RESULT", true, true);
                            doc.Replace("NATLSQADCOMMENT", "NATLSQADCOMMENT\n\nNATLSQAD2COMMENT", true, true);
                            doc.Replace("NATLSQADRESULT", "NATLSQADRESULT\n\nNATLSQAD2RESULT", true, true);
                            doc.Replace("NOYOKCOMMENT", "NOYOKCOMMENT\n\nNOYOK2COMMENT", true, true);
                            doc.Replace("NOYOKRESULT", "NOYOKRESULT\n\nNOYOK2RESULT", true, true);
                            doc.Replace("NOTTINGCOMMENT", "NOTTINGCOMMENT\n\nNOTTING2COMMENT", true, true);
                            doc.Replace("NOTTINGRESULT", "NOTTINGRESULT\n\nNOTTING2RESULT", true, true);
                            doc.Replace("SURREYCOMMENT", "SURREYCOMMENT\n\nSURREY2COMMENT", true, true);
                            doc.Replace("SURREYRESULT", "SURREYRESULT\n\nSURREY2RESULT", true, true);
                            doc.Replace("THAMCOMMENT", "THAMCOMMENT\n\nTHAM2COMMENT", true, true);
                            doc.Replace("THAMRESULT", "THAMRESULT\n\nTHAM2RESULT", true, true);
                            doc.Replace("WARWCOMMENT", "WARWCOMMENT\n\nWARW2COMMENT", true, true);
                            doc.Replace("WARWRESULT", "WARWRESULT\n\nWARW2RESULT", true, true);
                            doc.Replace("ALBCOMMENT", "ALBCOMMENT\n\nALB2COMMENT", true, true);
                            doc.Replace("ALBRESULT", "ALBRESULT\n\nALB2RESULT", true, true);
                            doc.Replace("ASSECOMMENT", "ASSECOMMENT\n\nASSE2COMMENT", true, true);
                            doc.Replace("ASSERESULT", "ASSERESULT\n\nASSE2RESULT", true, true);
                            doc.Replace("AUSPRCOMMENT", "AUSPRCOMMENT\n\nAUSPR2COMMENT", true, true);
                            doc.Replace("AUSPRRESULT", "AUSPRRESULT\n\nAUSPR2RESULT", true, true);
                            doc.Replace("AUSSECCOMMENT", "AUSSECCOMMENT\n\nAUSSEC2COMMENT", true, true);
                            doc.Replace("AUSSECRESULT", "AUSSECRESULT\n\nAUSSEC2RESULT", true, true);
                            doc.Replace("BAQUECOMMENT", "BAQUECOMMENT\n\nBAQUE2COMMENT", true, true);
                            doc.Replace("BAQUERESULT", "BAQUERESULT\n\nBAQUE2RESULT", true, true);
                            doc.Replace("BACOMCOMMENT", "BACOMCOMMENT\n\nBACOM2COMMENT", true, true);
                            doc.Replace("BACOMRESULT", "BACOMRESULT\n\nBACOM2RESULT", true, true);
                            doc.Replace("BIRVIRCOMMENT", "BIRVIRCOMMENT\n\nBIRVIR2COMMENT", true, true);
                            doc.Replace("BRIVIRRESULT", "BRIVIRRESULT\n\nBRIVIR2RESULT", true, true);
                            doc.Replace("CAYCOMMENT", "CAYCOMMENT\n\nCAY2COMMENT", true, true);
                            doc.Replace("CAYRESULT", "CAYRESULT\n\nCAY2RESULT", true, true);
                            doc.Replace("COMDECOMMENT", "COMDECOMMENT\n\nCOMDE2COMMENT", true, true);
                            doc.Replace("COMDERESULT", "COMDERESULT\n\nCOMDE2RESULT", true, true);
                            doc.Replace("COUNFINCOMMENT", "COUNFINCOMMENT\n\nCOUNFIN2COMMENT", true, true);
                            doc.Replace("COUNFINRESULT", "COUNFINRESULT\n\nCOUNFIN2RESULT", true, true);
                            doc.Replace("MENTODECOMMENT", "MENTODECOMMENT\n\nMENTODE2COMMENT", true, true);
                            doc.Replace("MENTODERESULT", "MENTODERESULT\n\nMENTODE2RESULT", true, true);
                            doc.Replace("DEPLABCOMMENT", "DEPLABCOMMENT\n\nDEPLAB2COMMENT", true, true);
                            doc.Replace("DEPLABRESULT", "DEPLABRESULT\n\nDEPLAB2RESULT", true, true);
                            doc.Replace("FINACCOMMENT", "FINACCOMMENT\n\nFINAC2COMMENT", true, true);
                            doc.Replace("FINACRESULT", "FINACRESULT\n\nFINAC2RESULT", true, true);
                            doc.Replace("REGIRECOMMENT", "REGIRECOMMENT\n\nREGIRE2COMMENT", true, true);
                            doc.Replace("REGIRERESULT", "REGIRERESULT\n\nREGIRE2RESULT", true, true);
                            doc.Replace("KONMONCOMMENT", "KONMONCOMMENT\n\nKONMON2COMMENT", true, true);
                            doc.Replace("KONMONRESULT", "KONMONRESULT\n\nKONMON2RESULT", true, true);
                            doc.Replace("INDEASCOMMENT", "INDEASCOMMENT\n\nINDEAS2COMMENT", true, true);
                            doc.Replace("INDEASRESULT", "INDEASRESULT\n\nINDEAS2RESULT", true, true);
                            doc.Replace("INMARECOMMENT", "INMARECOMMENT\n\nINMARE2COMMENT", true, true);
                            doc.Replace("INMARERESULT", "INMARERESULT\n\nINMARE2RESULT", true, true);
                            doc.Replace("ISMASUCOMMENT", "ISMASUCOMMENT\n\nISMASU2COMMENT", true, true);
                            doc.Replace("ISMASURESULT", "ISMASURESULT\n\nISMASU2RESULT", true, true);
                            doc.Replace("JESECOCOMMENT", "JESECOCOMMENT\n\nJESECO2COMMENT", true, true);
                            doc.Replace("JESECORESULT", "JESECORESULT\n\nJESECO2RESULT", true, true);
                            doc.Replace("LIARCOMMENT", "LIARCOMMENT\n\nLIAR2COMMENT", true, true);
                            doc.Replace("LIARRESULT", "LIARRESULT\n\nLIAR2RESULT", true, true);
                            doc.Replace("MOSICOMMENT", "MOSICOMMENT\n\nMOSI2COMMENT", true, true);
                            doc.Replace("MOSIRESULT", "MOSIRESULT\n\nMOSI2RESULT", true, true);
                            doc.Replace("SECOBRCOMMENT", "SECOBRCOMMENT\n\nSECOBR2COMMENT", true, true);
                            doc.Replace("SECOBRRESULT", "SECOBRRESULT\n\nSECOBR2RESULT", true, true);
                            doc.Replace("SEFUAUCOMMENT", "SEFUAUCOMMENT\n\nSEFUAU2COMMENT", true, true);
                            doc.Replace("SEFUAURESULT", "SEFUAURESULT\n\nSEFUAU2RESULT", true, true);
                            doc.Replace("SWFICOMMENT", "SWFICOMMENT\n\nSWFI2COMMENT", true, true);
                            doc.Replace("SWFIRESULT", "SWFIRESULT\n\nSWFI2RESULT", true, true);
                            doc.Replace("SWBACOMMENT", "SWBACOMMENT\n\nSWBA2COMMENT", true, true);
                            doc.Replace("SWBARESULT", "SWBARESULT\n\nSWBA2RESULT", true, true);
                            doc.Replace("UKCOHOCOMMENT", "UKCOHOCOMMENT\n\nUKCOHO2COMMENT", true, true);
                            doc.Replace("UKCOHORESULT", "UKCOHORESULT\n\nUKCOHO2RESULT", true, true);

                        }
                    }
                    catch { }
                    switch (ST.personal_Identification.ToString())
                    {
                        case "Clear":
                            doc.Replace("PICOMMENT", "Confirmed", true, true);
                            doc.Replace("PLRESULT", string.Concat(DI.FirstName, " ", DI.LastName, " - Clear"), true, true);
                            break;
                        case "Discrepancy Identified":
                            doc.Replace("PICOMMENT", "<Investigator to insert discrepancy here>", true, true);
                            doc.Replace("PLRESULT", string.Concat(DI.FirstName, " ", DI.LastName, " - Discrepancy Identified"), true, true);
                            break;
                    }
                    doc.SaveToFile(savePath);
                    switch (ST.name_add_Ver_History.ToString())
                    {
                        case "Clear":
                            strcomment = "LastName has jurisdictional ties to United States";
                            strcomment = strcomment.Replace("LastName", string.Concat(DI.FirstName, " ", DI.LastName));
                            doc.Replace("NAMECOMMENT", strcomment, true, true);
                            doc.Replace("NAMERESULT", string.Concat(DI.FirstName, " ", DI.LastName, " - Clear"), true, true);
                            break;
                        case "Discrepancy Identified":
                            strcomment = "LastName has jurisdictional ties to United States\n\nHowever, while not reported by the subject, LastName was also identified as having current ties to < Investigator to insert additional jurisdictions here>";
                            strcomment = strcomment.Replace("LastName", string.Concat(DI.FirstName, " ", DI.LastName));
                            doc.Replace("NAMECOMMENT", strcomment, true, true);
                            doc.Replace("NAMERESULT", string.Concat(DI.FirstName, " ", DI.LastName, " - Discrepancy Identified"), true, true);
                            break;
                    }                  
                    //string bankrupcomment = "";
                    //string bankrupresult = "";
                    //switch (ST.bankruptcy_filings.ToString())
                    //{
                    //    case "Clear":
                    //        bankrupcomment = "";
                    //        bankrupresult = string.Concat(DI.FirstName, " ", DI.LastName, " - Clear");
                    //        break;
                    //    case "Records":
                    //        bankrupcomment = "LastName was identified as a <party type> in connection with at least <number> bankruptcy filings in [Country], filed between <date> and <date>, which are currently <status> ";
                    //        bankrupcomment = bankrupcomment.Replace("LastName", string.Concat(DI.FirstName, " ", DI.LastName));
                    //        bankrupresult = string.Concat(DI.FirstName, " ", DI.LastName, " - Records");
                    //        break;
                    //    case "Record":
                    //        bankrupcomment = "[LastName] was identified as a <party type> in connection with a bankruptcy filing in [Country], which were recorded in <date> and <date>, and is currently <status>";
                    //        bankrupresult = string.Concat(DI.FirstName, " ", DI.LastName, " - Record");
                    //        break;
                    //    case "Results Pending":
                    //        bankrupcomment = "Bankruptcy court records searches are currently pending in [Country] the results of which will be provided under separate cover upon receipt";
                    //        bankrupresult = string.Concat(DI.FirstName, " ", DI.LastName, " - Results Pending");
                    //        break;
                    //}
                    //if (ST.bankruptcy_filings1 == true)
                    //{
                    //    bankrupresult = string.Concat(bankrupresult, "\n\nPossible Records");
                    //    bankrupcomment = string.Concat(bankrupcomment, "\n\nOne or more individuals known only as “[FirstName] [LastName]” were identified as Petitioners to at least <number> bankruptcy filings in [Country], which were recorded between <date> and <date>, and are currently <status>");
                    //}
                    //doc.Replace("BANKRUPCOMMENT", bankrupcomment, true, true);
                    //doc.Replace("BANKRUPRESULT", bankrupresult, true, true);                   
                  
                    switch (ST.social_securitytrace.ToString())
                    {
                        case "Clear":
                            doc.Replace("SOCSECCOMMENT", "Confirmed", true, true);
                            doc.Replace("SOCSECRESULT", "Clear", true, true);
                            break;
                        case "Discrepancy Identified":
                            doc.Replace("SOCSECCOMMENT", "<Investigator to insert discrepancy here>", true, true);
                            doc.Replace("SOCSECRESULT", "Discrepancy Identified", true, true);
                            break;
                    }
                    switch (ST.news_media_searches.ToString())
                    {
                        case "Clear":
                            doc.Replace("[NewsRES]", "", true, false);
                            doc.Replace("NEWSCOMMENT", "No adverse or materially-significant information was identified ", true, true);
                            doc.Replace("NEWSRES", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, false);
                            break;
                        case "Records":
                            doc.Replace("NEWSCOMMENT", "<investigator to insert press summaries here>", true, true);
                            doc.Replace("NEWSRES", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, false);
                            doc.Replace("[NewsRES]", "", true, false);
                            break;
                        case "Potentially-Relevant Information":
                            doc.Replace("NEWSCOMMENT", "<investigator to insert press summaries here>", true, true);
                            doc.Replace("NEWSRES", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - "), true, false);
                            doc.Replace("[NewsRES]", "Potentially-Relevant Information", true, false);
                            break;
                    }
                    switch (ST.department_foreign.ToString())
                    {
                        case "Clear":
                            doc.Replace("DEPTFORCOMMENT", "", true, true);
                            doc.Replace("DEPTFORRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("DEPTFORCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("DEPTFORRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.european_union.ToString())
                    {
                        case "Clear":
                            doc.Replace("EURUNCOMMENT", "", true, true);
                            doc.Replace("EURUNRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("EURUNCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("EURUNRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.HM_treasury.ToString())
                    {
                        case "Clear":
                            doc.Replace("HMTRECOMMENT", "", true, true);
                            doc.Replace("HMTRERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("HMTRECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("HMTRERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_bureau.ToString())
                    {
                        case "Clear":
                            doc.Replace("USBEUCOMMENT", "", true, true);
                            doc.Replace("USBEURESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USBEUCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USBEURESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_department.ToString())
                    {
                        case "Clear":
                            doc.Replace("USBDEPCOMMENT", "", true, true);
                            doc.Replace("USBDEPRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USBDEPCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USBDEPRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_Directorate.ToString())
                    {
                        case "Clear":
                            doc.Replace("USBDIRCOMMENT", "", true, true);
                            doc.Replace("USBDIRRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USBDIRCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USBDIRRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_general.ToString())
                    {
                        case "Clear":
                            doc.Replace("USGENCOMMENT", "", true, true);
                            doc.Replace("USGENRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USGENCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USGENRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    try
                    {
                        switch (ST.US_office.ToString())
                        {
                            case "Clear":
                                doc.Replace("USOFFCOMMENT", "", true, true);
                                doc.Replace("USOFFRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                                break;
                            case "Records":
                                doc.Replace("USOFFCOMMENT", "<investigator to insert summary here>", true, true);
                                doc.Replace("USOFFRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                                break;
                        }
                    }
                    catch
                    {
                        doc.Replace("USOFFCOMMENT", "", true, true);
                        doc.Replace("USOFFRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                    }
                    switch (ST.UN_consolidated.ToString())
                    {
                        case "Clear":
                            doc.Replace("UNCONSOCOMMENT", "", true, true);
                            doc.Replace("UNCONSORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("UNCONSOCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("UNCONSORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.world_bank_list.ToString())
                    {
                        case "Clear":
                            doc.Replace("WORLDBANKCOMMENT", "", true, true);
                            doc.Replace("WORLDBANKRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("WORLDBANKCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("WORLDBANKRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.interpol.ToString())
                    {
                        case "Clear":
                            doc.Replace("INTERPOLCOMMENT", "", true, true);
                            doc.Replace("INTERPOLRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("INTERPOLCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("INTERPOLRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_Federal.ToString())
                    {
                        case "Clear":
                            doc.Replace("USFEDERALCOMMENT", "", true, true);
                            doc.Replace("USFEDERALRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USFEDERALCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USFEDERALRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_secret_service.ToString())
                    {
                        case "Clear":
                            doc.Replace("USSECRETCOMMENT", "", true, true);
                            doc.Replace("USSECRETRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USSECRETCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USSECRETRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.commodity_futures.ToString())
                    {
                        case "Clear":
                            doc.Replace("COMMODITYCOMMENT", "", true, true);
                            doc.Replace("COMMODITYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("COMMODITYCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("COMMODITYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.federal_deposit.ToString())
                    {
                        case "Clear":
                            doc.Replace("FDEPOCOMMENT", "", true, true);
                            doc.Replace("FDEPORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - ", "Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("FDEPOCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("FDEPORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - ", "Records"), true, true);
                            break;
                    }
                    switch (ST.federal_reserve.ToString())
                    {
                        case "Clear":
                            doc.Replace("FRESCOMMENT", "", true, true);
                            doc.Replace("FRESRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("FRESCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("FRESRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.financial_crimes.ToString())
                    {
                        case "Clear":
                            doc.Replace("FCRIMECOMMENT", "", true, true);
                            doc.Replace("FCRIMERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("FCRIMECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("FCRIMERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.national_credit.ToString())
                    {
                        case "Clear":
                            doc.Replace("NATCRECOMMENT", "", true, true);
                            doc.Replace("NATCRERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("NATCRECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("NATCRERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.new_york_stock.ToString())
                    {
                        case "Clear":
                            doc.Replace("NYCOMMENT", "", true, true);
                            doc.Replace("NYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("NYCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("NYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.Office_of_comptroller.ToString())
                    {
                        case "Clear":
                            doc.Replace("OFFCPTCOMMENT", "", true, true);
                            doc.Replace("OFFCPTRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("OFFCPTCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("OFFCPTRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.Office_of_superintendent.ToString())
                    {
                        case "Clear":
                            doc.Replace("OFFSUPCOMMENT", "", true, true);
                            doc.Replace("OFFSUPRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("OFFSUPCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("OFFSUPRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.resolution_trust.ToString())
                    {
                        case "Clear":
                            doc.Replace("RESTCOMMENT", "", true, true);
                            doc.Replace("RESTRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("RESTCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("RESTRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_court.ToString())
                    {
                        case "Clear":
                            doc.Replace("USCOURCOMMENT", "", true, true);
                            doc.Replace("USCOURESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USCOURCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USCOURESULT", "Records", true, true);
                            break;
                    }
                    switch (ST.US_department_justice.ToString())
                    {
                        case "Clear":
                            doc.Replace("USDPJSCOMMENT", "", true, true);
                            doc.Replace("USDPJSRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USDPJSCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USDPJSRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_federal_trade.ToString())
                    {
                        case "Clear":
                            doc.Replace("USFEDCOMMENT", "", true, true);
                            doc.Replace("USFEDRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USFEDCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USFEDRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.US_office_thrifts.ToString())
                    {
                        case "Clear":
                            doc.Replace("USOTCOMMENT", "", true, true);
                            doc.Replace("USOTRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("USOTCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USOTRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.central_intelligence.ToString())
                    {
                        case "Clear":
                            doc.Replace("CICOMMENT", "", true, true);
                            doc.Replace("CIRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("CICOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("CIRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.city_london_police.ToString())
                    {
                        case "Clear":
                            doc.Replace("CITYCOMMENT", "", true, true);
                            doc.Replace("CITYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("CITYCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("CITYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.constabularies_cheshire.ToString())
                    {
                        case "Clear":
                            doc.Replace("COSTABCOMMENT", "", true, true);
                            doc.Replace("COSTABRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("COSTABCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("COSTABRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.hampshire_police.ToString())
                    {
                        case "Clear":
                            doc.Replace("HAMPSPCOMMENT", "", true, true);
                            doc.Replace("HAMPSPRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("HAMPSPCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("HAMPSPRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.hong_kong_police.ToString())
                    {
                        case "Clear":
                            doc.Replace("HONGKONGCOMMENT", "", true, true);
                            doc.Replace("HONGKONGRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("HONGKONGCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("HONGKONGRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.metropolitan_police.ToString())
                    {
                        case "Clear":
                            doc.Replace("METROPOLICOMMENT", "", true, true);
                            doc.Replace("METROLIRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("METROPOLICOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("METROLIRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.national_crime.ToString())
                    {
                        case "Clear":
                            doc.Replace("NATLSQADCOMMENT", "", true, true);
                            doc.Replace("NATLSQADRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("NATLSQADCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("NATLSQADRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.north_yorkshire_polic.ToString())
                    {
                        case "Clear":
                            doc.Replace("NOYOKCOMMENT", "", true, true);
                            doc.Replace("NOYOKRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("NOYOKCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("NOYOKRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.nottinghamshire_police.ToString())
                    {
                        case "Clear":
                            doc.Replace("NOTTINGCOMMENT", "", true, true);
                            doc.Replace("NOTTINGRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("NOTTINGCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("NOTTINGRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.surrey_police.ToString())
                    {
                        case "Clear":
                            doc.Replace("SURREYCOMMENT", "", true, true);
                            doc.Replace("SURREYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("SURREYCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("SURREYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.thames_valley_police.ToString())
                    {
                        case "Clear":
                            doc.Replace("THAMCOMMENT", "", true, true);
                            doc.Replace("THAMRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("THAMCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("THAMRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.warwickshire_police.ToString())
                    {
                        case "Clear":
                            doc.Replace("WARWCOMMENT", "", true, true);
                            doc.Replace("WARWRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("WARWCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("WARWRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.alberta_securities_commission.ToString())
                    {
                        case "Clear":
                            doc.Replace("ALBCOMMENT", "", true, true);
                            doc.Replace("ALBRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("ALBCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("ALBRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.asset_recovery_agency.ToString())
                    {
                        case "Clear":
                            doc.Replace("ASSECOMMENT", "", true, true);
                            doc.Replace("ASSERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("ASSECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("ASSERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.australian_prudential.ToString())
                    {
                        case "Clear":
                            doc.Replace("AUSPRCOMMENT", "", true, true);
                            doc.Replace("AUSPRRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("AUSPRCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("AUSPRRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.australian_securities.ToString())
                    {
                        case "Clear":
                            doc.Replace("AUSSECCOMMENT", "", true, true);
                            doc.Replace("AUSSECRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("AUSSECCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("AUSSECRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.banque_de_CECEI.ToString())
                    {
                        case "Clear":
                            doc.Replace("BAQUECOMMENT", "", true, true);
                            doc.Replace("BAQUERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("BAQUECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("BAQUERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.banque_de_commission.ToString())
                    {
                        case "Clear":
                            doc.Replace("BACOMCOMMENT", "", true, true);
                            doc.Replace("BACOMRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("BACOMCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("BACOMRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.british_virgin_islands.ToString())
                    {
                        case "Clear":
                            doc.Replace("BIRVIRCOMMENT", "", true, true);
                            doc.Replace("BRIVIRRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("BIRVIRCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("BRIVIRRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.cayman_islands_monetary.ToString())
                    {
                        case "Clear":
                            doc.Replace("CAYCOMMENT", "", true, true);
                            doc.Replace("CAYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("CAYCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("CAYRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.commission_de_surveillance.ToString())
                    {
                        case "Clear":
                            doc.Replace("COMDECOMMENT", "", true, true);
                            doc.Replace("COMDERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("COMDECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("COMDERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.council_financial_activities.ToString())
                    {
                        case "Clear":
                            doc.Replace("COUNFINCOMMENT", "", true, true);
                            doc.Replace("COUNFINRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("COUNFINCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("HONGKONGRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.departamento_de_investigacoes.ToString())
                    {
                        case "Clear":
                            doc.Replace("MENTODECOMMENT", "", true, true);
                            doc.Replace("MENTODERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("MENTODECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("MENTODERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.department_labour_inspection.ToString())
                    {
                        case "Clear":
                            doc.Replace("DEPLABCOMMENT", "", true, true);
                            doc.Replace("DEPLABRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("DEPLABCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("DEPLABRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.financial_action_task.ToString())
                    {
                        case "Clear":
                            doc.Replace("FINACCOMMENT", "", true, true);
                            doc.Replace("FINACRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("FINACCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("FINACRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.financial_regulator_ireland.ToString())
                    {
                        case "Clear":
                            doc.Replace("REGIRECOMMENT", "", true, true);
                            doc.Replace("REGIRERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("REGIRECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("REGIRERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.hongkong_monetary_authority.ToString())
                    {
                        case "Clear":
                            doc.Replace("KONMONCOMMENT", "", true, true);
                            doc.Replace("KONMONRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("KONMONCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("KONMONRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.investment_association_Canada.ToString())
                    {
                        case "Clear":
                            doc.Replace("INDEASCOMMENT", "", true, true);
                            doc.Replace("INDEASRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("INDEASCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("INDEASRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.investment_management_regulatory.ToString())
                    {
                        case "Clear":
                            doc.Replace("INMARECOMMENT", "", true, true);
                            doc.Replace("INMARERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("INMARECOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("INMARERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.isle_financial_supervision.ToString())
                    {
                        case "Clear":
                            doc.Replace("ISMASUCOMMENT", "", true, true);
                            doc.Replace("ISMASURESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("ISMASUCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("ISMASURESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.jersey_financial_commission.ToString())
                    {
                        case "Clear":
                            doc.Replace("JESECOCOMMENT", "", true, true);
                            doc.Replace("JESECORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("JESECOCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("JESECORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.lloyd_insurance_arimbolaet.ToString())
                    {
                        case "Clear":
                            doc.Replace("LIARCOMMENT", "", true, true);
                            doc.Replace("LIARRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("LIARCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("LIARRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.monetary_authority_singapore.ToString())
                    {
                        case "Clear":
                            doc.Replace("MOSICOMMENT", "", true, true);
                            doc.Replace("MOSIRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("MOSICOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("MOSIRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.securities_exchange_commission.ToString())
                    {
                        case "Clear":
                            doc.Replace("SECOBRCOMMENT", "", true, true);
                            doc.Replace("SECOBRRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("SECOBRCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("SECOBRRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.securities_futuresauthority.ToString())
                    {
                        case "Clear":
                            doc.Replace("SEFUAUCOMMENT", "", true, true);
                            doc.Replace("SEFUAURESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("SEFUAUCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("SEFUAURESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.swedish_financial_supervisory.ToString())
                    {
                        case "Clear":
                            doc.Replace("SWFICOMMENT", "", true, true);
                            doc.Replace("SWFIRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("SWFICOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("SWFIRESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.swiss_federal_banking.ToString())
                    {
                        case "Clear":
                            doc.Replace("SWBACOMMENT", "", true, true);
                            doc.Replace("SWBARESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("SWBACOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("SWBARESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    switch (ST.U_K_companies_disqualified.ToString())
                    {
                        case "Clear":
                            doc.Replace("UKCOHOCOMMENT", "", true, true);
                            doc.Replace("UKCOHORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("UKCOHOCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("UKCOHORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records"), true, true);
                            break;
                    }
                    doc.Replace("[Last Name]", string.Concat(DI.FirstName, " ", DI.LastName), true, true);
                    doc.SaveToFile(savePath);
                }
                doc.Replace("NAME2COMMENT", "NAMECOMMENT", true, true);
                doc.Replace("NAME2RESULT", "NAMERESULT", true, true);
                doc.Replace("PI2COMMENT", "PICOMMENT", true, true);
                doc.Replace("PL2RESULT", "PLRESULT", true, true);
                doc.Replace("REG2COMMENT", "REGCOMMENT", true, true);
                doc.Replace("REG2RESULT", "REGRESULT", true, true);
                doc.Replace("BANKRUP2COMMENT", "BANKRUPCOMMENT", true, true);
                doc.Replace("BANKRUP2RESULT", "BANKRUPRESULT", true, true);
                doc.Replace("CIVIL2COMMENT", "CIVILCOMMENT", true, true);
                doc.Replace("CIVIL2RESULT", "CIVILRESULT", true, true);
                doc.Replace("SOCSEC2COMMENT", "SOCSECCOMMENT", true, true);
                doc.Replace("SOCSEC2RESULT", "SOCSECRESULT", true, true);
                doc.Replace("DEPTFOR2COMMENT", "DEPTFORCOMMENT", true, true);
                doc.Replace("DEPTFOR2RESULT", "DEPTFORRESULT", true, true);
                doc.Replace("EURUN2COMMENT", "EURUNCOMMENT", true, true);
                doc.Replace("EURUN2RESULT", "EURUNRESULT", true, true);
                doc.Replace("HMTRE2COMMENT", "HMTRECOMMENT", true, true);
                doc.Replace("HMTRE2RESULT", "HMTRERESULT", true, true);
                doc.Replace("USBEU2COMMENT", "USBEUCOMMENT", true, true);
                doc.Replace("USBEU2RESULT", "USBEURESULT", true, true);
                doc.Replace("USBDEP2COMMENT", "USBDEPCOMMENT", true, true);
                doc.Replace("USBDEP2RESULT", "USBDEPRESULT", true, true);
                doc.Replace("USBDIR2COMMENT", "USBDIRCOMMENT", true, true);
                doc.Replace("USBDIR2RESULT", "USBDIRRESULT", true, true);
                doc.Replace("USGEN2COMMENT", "USGENCOMMENT", true, true);
                doc.Replace("USGEN2RESULT", "USGENRESULT", true, true);
                doc.Replace("USOFF2COMMENT", "USOFFCOMMENT", true, true);
                doc.Replace("USOFF2RESULT", "USOFFRESULT", true, true);
                doc.Replace("UNCONSO2COMMENT", "UNCONSOCOMMENT", true, true);
                doc.Replace("UNCONSO2RESULT", "UNCONSORESULT", true, true);
                doc.Replace("WORLDBANK2COMMENT", "WORLDBANKCOMMENT", true, true);
                doc.Replace("WORLDBANK2RESULT", "WORLDBANKRESULT", true, true);
                doc.Replace("INTERPOL2COMMENT", "INTERPOLCOMMENT", true, true);
                doc.Replace("INTERPOL2RESULT", "INTERPOLRESULT", true, true);
                doc.Replace("USFEDERAL2COMMENT", "USFEDERALCOMMENT", true, true);
                doc.Replace("USFEDERAL2RESULT", "USFEDERALRESULT", true, true);
                doc.Replace("USSECRET2COMMENT", "USSECRETCOMMENT", true, true);
                doc.Replace("USSECRET2RESULT", "USSECRETRESULT", true, true);
                doc.Replace("COMMODITY2COMMENT", "COMMODITYCOMMENT", true, true);
                doc.Replace("COMMODITY2RESULT", "COMMODITYRESULT", true, true);
                doc.Replace("FDEPO2COMMENT", "FDEPOCOMMENT", true, true);
                doc.Replace("FDEPO2RESULT", "FDEPORESULT", true, true);
                doc.Replace("FRES2COMMENT", "FRESCOMMENT", true, true);
                doc.Replace("FRES2RESULT", "FRESRESULT", true, true);
                doc.Replace("FCRIME2COMMENT", "FCRIMECOMMENT", true, true);
                doc.Replace("FCRIME2RESULT", "FCRIMERESULT", true, true);
                doc.Replace("NATCRE2COMMENT", "NATCRECOMMENT", true, true);
                doc.Replace("NATCRE2RESULT", "NATCRERESULT", true, true);
                doc.Replace("NY2COMMENT", "NYCOMMENT", true, true);
                doc.Replace("NY2RESULT", "NYRESULT", true, true);
                doc.Replace("OFFCPT2COMMENT", "OFFCPTCOMMENT", true, true);
                doc.Replace("OFFCPT2RESULT", "OFFCPTRESULT", true, true);
                doc.Replace("OFFSUP2COMMENT", "OFFSUPCOMMENT", true, true);
                doc.Replace("OFFSUP2RESULT", "OFFSUPRESULT", true, true);
                doc.Replace("REST2COMMENT", "RESTCOMMENT", true, true);
                doc.Replace("REST2RESULT", "RESTRESULT", true, true);
                doc.Replace("USCOUR2COMMENT", "USCOURCOMMENT", true, true);
                doc.Replace("USCOU2RESULT", "USCOURESULT", true, true);
                doc.Replace("USDPJS2COMMENT", "USDPJSCOMMENT", true, true);
                doc.Replace("USDPJS2RESULT", "USDPJSRESULT", true, true);
                doc.Replace("USFED2COMMENT", "USFEDCOMMENT", true, true);
                doc.Replace("USFED2RESULT", "USFEDRESULT", true, true);
                doc.Replace("USOT2COMMENT", "USOTCOMMENT", true, true);
                doc.Replace("USOT2RESULT", "USOTRESULT", true, true);
                doc.Replace("CI2COMMENT", "CICOMMENT", true, true);
                doc.Replace("CI2RESULT", "CIRESULT", true, true);
                doc.Replace("CITY2COMMENT", "CITYCOMMENT", true, true);
                doc.Replace("CITY2RESULT", "CITYRESULT", true, true);
                doc.Replace("COSTAB2COMMENT", "COSTABCOMMENT", true, true);
                doc.Replace("COSTAB2RESULT", "COSTABRESULT", true, true);
                doc.Replace("HAMPSP2COMMENT", "HAMPSPCOMMENT", true, true);
                doc.Replace("HAMPSP2RESULT", "HAMPSPRESULT", true, true);
                doc.Replace("HONGKONG2COMMENT", "HONGKONGCOMMENT", true, true);
                doc.Replace("HONGKONG2RESULT", "HONGKONGRESULT", true, true);
                doc.Replace("METROPOLI2COMMENT", "METROPOLICOMMENT", true, true);
                doc.Replace("METROLI2RESULT", "METROLIRESULT", true, true);
                doc.Replace("NATLSQAD2COMMENT", "NATLSQADCOMMENT", true, true);
                doc.Replace("NATLSQAD2RESULT", "NATLSQADRESULT", true, true);
                doc.Replace("NOYOK2COMMENT", "NOYOKCOMMENT", true, true);
                doc.Replace("NOYOK2RESULT", "NOYOKRESULT", true, true);
                doc.Replace("NOTTING2COMMENT", "NOTTINGCOMMENT", true, true);
                doc.Replace("NOTTING2RESULT", "NOTTINGRESULT", true, true);
                doc.Replace("SURREY2COMMENT", "SURREYCOMMENT", true, true);
                doc.Replace("SURREY2RESULT", "SURREYRESULT", true, true);
                doc.Replace("THAM2COMMENT", "THAMCOMMENT", true, true);
                doc.Replace("THAM2RESULT", "THAMRESULT", true, true);
                doc.Replace("WARW2COMMENT", "WARWCOMMENT", true, true);
                doc.Replace("WARW2RESULT", "WARWRESULT", true, true);
                doc.Replace("ALB2COMMENT", "ALBCOMMENT", true, true);
                doc.Replace("ALB2RESULT", "ALBRESULT", true, true);
                doc.Replace("ASSE2COMMENT", "ASSECOMMENT", true, true);
                doc.Replace("ASSE2RESULT", "ASSERESULT", true, true);
                doc.Replace("AUSPR2COMMENT", "AUSPRCOMMENT", true, true);
                doc.Replace("AUSPR2RESULT", "AUSPRRESULT", true, true);
                doc.Replace("AUSSEC2COMMENT", "AUSSECCOMMENT", true, true);
                doc.Replace("AUSSEC2RESULT", "AUSSECRESULT", true, true);
                doc.Replace("BAQUE2COMMENT", "BAQUECOMMENT", true, true);
                doc.Replace("BAQUE2RESULT", "BAQUERESULT", true, true);
                doc.Replace("BACOM2COMMENT", "BACOMCOMMENT", true, true);
                doc.Replace("BACOM2RESULT", "BACOMRESULT", true, true);
                doc.Replace("BIRVIR2COMMENT", "BIRVIRCOMMENT", true, true);
                doc.Replace("BRIVIR2RESULT", "BRIVIRRESULT", true, true);
                doc.Replace("CAY2COMMENT", "CAYCOMMENT", true, true);
                doc.Replace("CAY2RESULT", "CAYRESULT", true, true);
                doc.Replace("COMDE2COMMENT", "COMDECOMMENT", true, true);
                doc.Replace("COMDE2RESULT", "COMDERESULT", true, true);
                doc.Replace("COUNFIN2COMMENT", "COUNFINCOMMENT", true, true);
                doc.Replace("COUNFIN2RESULT", "COUNFINRESULT", true, true);
                doc.Replace("MENTODE2COMMENT", "MENTODECOMMENT", true, true);
                doc.Replace("MENTODE2RESULT", "MENTODERESULT", true, true);
                doc.Replace("DEPLAB2COMMENT", "DEPLABCOMMENT", true, true);
                doc.Replace("DEPLAB2RESULT", "DEPLABRESULT", true, true);
                doc.Replace("FINAC2COMMENT", "FINACCOMMENT", true, true);
                doc.Replace("FINAC2RESULT", "FINACRESULT", true, true);
                doc.Replace("REGIRE2COMMENT", "REGIRECOMMENT", true, true);
                doc.Replace("REGIRE2RESULT", "REGIRERESULT", true, true);
                doc.Replace("KONMON2COMMENT", "KONMONCOMMENT", true, true);
                doc.Replace("KONMON2RESULT", "KONMONRESULT", true, true);
                doc.Replace("INDEAS2COMMENT", "INDEASCOMMENT", true, true);
                doc.Replace("INDEAS2RESULT", "INDEASRESULT", true, true);
                doc.Replace("INMARE2COMMENT", "INMARECOMMENT", true, true);
                doc.Replace("INMARE2RESULT", "INMARERESULT", true, true);
                doc.Replace("ISMASU2COMMENT", "ISMASUCOMMENT", true, true);
                doc.Replace("ISMASU2RESULT", "ISMASURESULT", true, true);
                doc.Replace("JESECO2COMMENT", "JESECOCOMMENT", true, true);
                doc.Replace("JESECO2RESULT", "JESECORESULT", true, true);
                doc.Replace("LIAR2COMMENT", "LIARCOMMENT", true, true);
                doc.Replace("LIAR2RESULT", "LIARRESULT", true, true);
                doc.Replace("MOSI2COMMENT", "MOSICOMMENT", true, true);
                doc.Replace("MOSI2RESULT", "MOSIRESULT", true, true);
                doc.Replace("SECOBR2COMMENT", "SECOBRCOMMENT", true, true);
                doc.Replace("SECOBR2RESULT", "SECOBRRESULT", true, true);
                doc.Replace("SEFUAU2COMMENT", "SEFUAUCOMMENT", true, true);
                doc.Replace("SEFUAU2RESULT", "SEFUAURESULT", true, true);
                doc.Replace("SWFI2COMMENT", "SWFICOMMENT", true, true);
                doc.Replace("SWFI2RESULT", "SWFIRESULT", true, true);
                doc.Replace("SWBA2COMMENT", "SWBACOMMENT", true, true);
                doc.Replace("SWBA2RESULT", "SWBARESULT", true, true);
                doc.Replace("UKCOHO2COMMENT", "UKCOHOCOMMENT", true, true);
                doc.Replace("UKCOHO2RESULT", "UKCOHORESULT", true, true);

                doc.SaveToFile(savePath);
            }

                   
            
            doc.Replace("the United Kingdom’s Financial Conduct Authority", "United Kingdom’s Financial Conduct Authority", true, true);
            doc.Replace("< investigator to remove if inaccurate >", "<investigator to remove if inaccurate>", false, true);
            doc.Replace("<investigator to insert regulatory hits here >", "<investigator to insert regulatory hits here>", true, true);
            doc.SaveToFile(savePath);
            Regex regex = new Regex(@"\<([^>]*)\>");
            TextSelection[] textSelections1 = doc.FindAllPattern(regex);
            if (textSelections1 != null)
            {
                foreach (TextSelection seletion in textSelections1)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            doc.SaveToFile(savePath);

            //Regex regex1 = new Regex(@"/\p{No}/ug");
            //TextSelection[] footnotesupSelections1 = doc.FindAllPattern(regex1);
            //if (footnotesupSelections1 != null)
            //{
            //    foreach (TextSelection seletion in footnotesupSelections1)
            //    {
            //        seletion.GetAsOneRange().CharacterFormat.FontSize = 9;
            //        seletion.GetAsOneRange().CharacterFormat.FontName = "Calibri (Body)";
            //        seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.DarkGreen;
            //    }
            //}
            //doc.SaveToFile(savePath);
            TextSelection[] text24 = doc.FindAllString(" 	 • 	", false, true);
            if (text24 != null)
            {
                foreach (TextSelection seletion in text24)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.White;
                }
            }
            TextSelection[] text21 = doc.FindAllString("• 	 <Investigator to insert bulleted list of results here>", false, true);
            if (text21 != null)
            {
                foreach (TextSelection seletion in text21)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] text31 = doc.FindAllString(" 	•", false, true);
            if (text31 != null)
            {
                foreach (TextSelection seletion in text31)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.White;
                }
            }
            TextSelection[] text33 = doc.FindAllString("As confirmed", false, true);
            if (text33 != null)
            {
                foreach (TextSelection seletion in text33)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] text34 = doc.FindAllString("as confirmed", false, true);
            if (text34 != null)
            {
                foreach (TextSelection seletion in text34)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] text32 = doc.FindAllString("	•", false, true);
            if (text32 != null)
            {
                foreach (TextSelection seletion in text32)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.White;
                }
            }
            TextSelection[] text30 = doc.FindAllString("•	<Investigator to insert bulleted list of results here>", false, false);
            if (text30 != null)
            {
                foreach (TextSelection seletion in text30)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] text03 = doc.FindAllString("<Search type> searches are currently ongoing in <jurisdiction>, the results of which will be provided under separate cover upon receipt.", false, false);
            if (text03 != null)
            {
                foreach (TextSelection seletion in text03)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                }
            }
            //TextSelection[] text21 = doc.FindAllString("<Investigator to insert bulleted list of results here>", false, true);
            //if (text21 != null)
            //{
            //    foreach (TextSelection seletion in text21)
            //    {
            //        seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
            //    }
            //}           
            doc.Replace("[Last Name]", last_name.ToString(), true, true);
            doc.Replace("[LastName]", last_name.ToString(), false, false);
            //doc.Replace("The United Kingdom", "United Kingdom", true, true);            
            //if (employer1City.ToString().ToUpper().Equals("NA") || employer1City.ToString().Equals(""))
            //{
            //    doc.Replace("[Employer1City], ", "", true, false);
            //}
            //else
            //{
            //    doc.Replace("[Employer1City]", employer1City, true, false);
            //}
            //doc.SaveToFile(savePath);
            //if (employer1State.ToString().ToUpper().Equals("NA") || employer1State.ToString().Equals(""))
            //{
            //    doc.Replace(", [Employer1State]", "", true, false);
            //    doc.Replace("[Employer1State]", "", true, false);
            //}
            //else
            //{
            //    doc.Replace("[Employer1State]", employer1State, true, false);
            //}
            doc.Replace("countries", "counties", true, false);
            doc.SaveToFile(savePath);
            doc.Replace("investigator", "Investigator", true, true);
            doc.Replace("  ", " ", true, false);
            doc.SaveToFile(savePath);
            doc.Replace(" ,", ",", true, false);
            doc.SaveToFile(savePath);
            doc.Replace("  (", " (", true, false);
            doc.Replace("Result Pending", "Results Pending", true, false);
            doc.SaveToFile(savePath);
            TextSelection[] textresult1 = doc.FindAllString("Results Pending", false, false);
            if (textresult1 != null)
            {
                foreach (TextSelection seletion in textresult1)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                }
            }
            doc.SaveToFile(savePath);
            doc.Replace(",,", ",", true, false);
            doc.SaveToFile(savePath);
            doc.Replace(".  ", ". ", true, false);
            doc.Replace(". ", ".  ", true, false);
            doc.SaveToFile(savePath);
            try
            {
                if (MiddleInitial == "")
                {
                    doc.Replace(" [MiddleInitial]", "", true, false);
                    doc.Replace(" Middle_Initial", "", true, false);
                }
                else
                {
                    doc.Replace("[MiddleInitial]", MiddleInitial, true, true);
                    doc.Replace("Middle_Initial", MiddleInitial.TrimEnd().ToUpper().ToString(), true, true);
                }
            }
            catch
            {
                doc.Replace(" [MiddleInitial]", "", true, false);
                doc.Replace(" Middle_Initial", "", true, false);
            }
            doc.Replace("[ADDOOFOOTNOTE]", "", true, false);
            doc.Replace("[ADDNAMEONLYFN]", ".", true, false);
            doc.SaveToFile(savePath);
            doc.Replace("the District of Columbia", "District of Columbia", false, false);
            doc.SaveToFile(savePath);
            doc.Replace("District of Columbia", "the District of Columbia", false, false);
            doc.SaveToFile(savePath);
            doc.Replace("U.S.  ", "U.S. ", true, false);
            doc.Replace("U.K.  ", "U.K. ", true, false);
            if (US_CITI.familyModels.Count == 1)
            {
                doc.Replace("subjects", "subject", true, true);
                doc.Replace("Subjects", "Subject", true, true);
            }
            doc.SaveToFile(savePath);
            return RedirectToAction("GenerateFile", "Diligence");
        }
        public IActionResult GenerateFile()
        {
            return View();
        }
        public ActionResult DownloadFile()
        {
            string first_name = HttpContext.Session.GetString("first_name");
            string last_name = HttpContext.Session.GetString("last_name");
            string case_number = HttpContext.Session.GetString("case_number");
            string savePath = _config.GetValue<string>("ReportPath");
            savePath = string.Concat(savePath, last_name, "_SterlingDiligenceReport(", case_number, ")_DRAFT.docx");
            // string path = AppDomain.CurrentDomain.BaseDirectory + "FolderName/";
            byte[] fileBytes = System.IO.File.ReadAllBytes(savePath);
            string fileName = string.Concat(last_name, "_SterlingDiligenceReport(", case_number, ")_DRAFT.docx");
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }
    }
}