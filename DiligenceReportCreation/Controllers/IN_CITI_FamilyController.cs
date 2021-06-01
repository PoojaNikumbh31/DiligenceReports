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
  public class IN_CITI_FamilyController : Controller
    {
        private readonly DataBaseContext _context;
        private readonly IConfiguration _config;
        public IN_CITI_FamilyController(IConfiguration config, DataBaseContext context)
        {
            _config = config;
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult INTL_Citizen_Family_Newpage()
        {
            string recordid = HttpContext.Session.GetString("recordid");
            MainModel mainModel = new MainModel();
            mainModel.familyModels = _context.familyModels
                        .Where(u => u.record_Id == recordid)
                        .ToList();
            return View(mainModel);
        }
        [HttpPost]
        public IActionResult INTL_Citizen_Family_Newpage(MainModel main, string SaveData, string Submit)
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
                CountrySpecificModel countrySpecificModel = new CountrySpecificModel();
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
                    if (mainmodel.familyModels[i].adult_minor == "Adult" )
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
                        //try
                        //{
                            countrySpecificModel = _context.CSComment
                                         .Where(u => u.record_Id == mainmodel.familyModels[i].Family_record_id)
                                         .FirstOrDefault();
                            if (countrySpecificModel == null)
                            {
                                TempData["message"] = "Please enter the details in country specific section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                            }
                        //}
                        //catch
                        //{
                        //    TempData["message"] = "Please enter the details in country specific section for " + mainmodel.familyModels[i].first_name + " " + mainmodel.familyModels[i].last_name;
                        //}
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
            try {
                mainModel.pllicenseModels = _context.DbPLLicense
                    .Where(u => u.record_Id == id.ToString())
                    .ToList();
            } catch { }
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
            try {
                mainModel.additional_Countries = _context.DbadditionalCountries
                                      .Where(u => u.record_Id == id.ToString())
                                      .ToList();
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

                return RedirectToAction("INTL_Citizen_Family_Newpage", "IN_CITI_Family");
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
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Nationality = mainModel.diligenceInputModel.Nationality;
                                inputModel.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                                inputModel.Natlidissuedate = mainModel.diligenceInputModel.Natlidissuedate;
                                inputModel.Natlidexpirationdate = mainModel.diligenceInputModel.Natlidexpirationdate;
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
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
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Nationality = mainModel.diligenceInputModel.Nationality;
                                inputModel.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                                inputModel.Natlidissuedate = mainModel.diligenceInputModel.Natlidissuedate;
                                inputModel.Natlidexpirationdate = mainModel.diligenceInputModel.Natlidexpirationdate;
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
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
                                strupdate.USLegal_Record_Hits = mainModel.otherdetails.USLegal_Record_Hits;
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
                                _context.Entry(strupdate).State = EntityState.Modified;
                                _context.SaveChanges();
                            }
                            catch
                            {
                                Otherdatails strupdate = new Otherdatails();
                                strupdate.record_Id = family_record_id;
                                strupdate.Has_Property_Records = mainModel.otherdetails.Has_Property_Records;
                                strupdate.Press_Media = mainModel.otherdetails.Press_Media;
                                strupdate.USLegal_Record_Hits = mainModel.otherdetails.USLegal_Record_Hits;
                                strupdate.RegulatoryFlag = mainModel.otherdetails.RegulatoryFlag;
                                strupdate.Has_Reg_FINRA = mainModel.otherdetails.Has_Reg_FINRA;
                                strupdate.Has_Reg_UK_FCA = mainModel.otherdetails.Has_Reg_UK_FCA;
                                strupdate.Has_Reg_US_NFA = mainModel.otherdetails.Has_Reg_US_NFA;
                                strupdate.Has_Reg_US_SEC = mainModel.otherdetails.Has_Reg_US_SEC;
                                strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;
                                strupdate.HoldsAnyLicense = mainModel.otherdetails.HoldsAnyLicense;
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
                            DiligenceInputModel diligence = _context.DbPersonalInfo
                            .Where(a => a.record_Id == family_record_id)
                            .FirstOrDefault();
                            try
                            {
                                diligence.Country = mainModel.diligenceInputModel.Country;
                                _context.DbPersonalInfo.Update(diligence);
                                _context.SaveChanges();
                            }
                            catch
                            {
                                diligence = new DiligenceInputModel();
                                diligence.record_Id = family_record_id;
                                diligence.Country = mainModel.diligenceInputModel.Country;
                                _context.DbPersonalInfo.Add(diligence);
                                _context.SaveChanges();
                            }
                            CountrySpecificModel countrySpecific = new CountrySpecificModel();
                            try
                            {
                                countrySpecific = _context.CSComment
                                   .Where(u => u.record_Id == family_record_id)
                                                  .FirstOrDefault();
                           
                                countrySpecific.record_Id = family_record_id;
                            }
                            catch
                            {
                                countrySpecific = new CountrySpecificModel();
                                countrySpecific.record_Id = family_record_id;
                            }
                            countrySpecific.country = mainModel.diligenceInputModel.Country;
                            switch (mainModel.diligenceInputModel.Country)
                            {
                                case "Australia":
                                    countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                    countrySpecific.criminal_record_hits = mainModel.csModel.criminal_record_hits;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                    countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.pprrecordhits = mainModel.csModel.pprrecordhits;                                    
                                    break;
                                case "Canada":
                                    countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                    countrySpecific.criminal_record_hits = mainModel.csModel.criminal_record_hits;
                                    countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.pprrecordhits = mainModel.csModel.pprrecordhits;
                                    break;
                                case "United Kingdom":
                                    countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                    countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                    countrySpecific.criminal_record_hits = mainModel.csModel.criminal_record_hits;
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.reg_trust_hits = mainModel.csModel.reg_trust_hits;
                                    countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                    break;
                                case "France":
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;

                                    break;
                                case "India":
                                    countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                    countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                    countrySpecific.india_corpregistry = mainModel.csModel.india_corpregistry;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "Germany":
                                    countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                    countrySpecific.germany_regsearches = mainModel.csModel.germany_regsearches;
                                    countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "Switzerland":
                                    countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                    countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "United Arab Emirates":
                                    countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                    countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                default:
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                    countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                    break;
                            }
                            try
                            {
                                _context.CSComment.Update(countrySpecific);
                                _context.SaveChanges();
                            }
                            catch
                            {
                                _context.CSComment.Add(countrySpecific);
                                _context.SaveChanges();
                            }

                          
                            for (int i = 0; i < mainModel.additional_Countries.Count; i++)
                            {
                                countrySpecific = new CountrySpecificModel();
                                try
                                {
                                    countrySpecific = _context.CSComment
                                       .Where(u => u.record_Id == family_record_id && u.country == mainModel.additional_Countries[i].additionalCountries)
                                                      .FirstOrDefault();

                                    countrySpecific.record_Id = family_record_id;
                                }
                                catch
                                {
                                    countrySpecific = new CountrySpecificModel();
                                    countrySpecific.record_Id = family_record_id;
                                }
                                switch (mainModel.additional_Countries[i].additionalCountries)
                                {
                                    case "Australia":
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                        countrySpecific.criminal_record_hits = mainModel.csModel.criminal_record_hits;
                                        countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                        countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                        countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                        countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                        countrySpecific.pprrecordhits = mainModel.csModel.pprrecordhits;
                                        break;
                                    case "Canada":
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                        countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                        countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                        countrySpecific.criminal_record_hits = mainModel.csModel.criminal_record_hits;
                                        countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                        countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                        countrySpecific.pprrecordhits = mainModel.csModel.pprrecordhits;
                                        break;
                                    case "United Kingdom":
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                        countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                        countrySpecific.criminal_record_hits = mainModel.csModel.criminal_record_hits;
                                        countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                        countrySpecific.reg_trust_hits = mainModel.csModel.reg_trust_hits;
                                        countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                        break;
                                    case "France":
                                        countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        break;
                                    case "India":
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                        countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                        countrySpecific.india_corpregistry = mainModel.csModel.india_corpregistry;
                                        countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                        break;
                                    case "Germany":
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                        countrySpecific.germany_regsearches = mainModel.csModel.germany_regsearches;
                                        countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                        countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                        break;
                                    case "Switzerland":
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                        countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                        countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                        break;
                                    case "United Arab Emirates":
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                        countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                        countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                        break;
                                    default:
                                        countrySpecific.country = mainModel.additional_Countries[i].additionalCountries;
                                        countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                        countrySpecific.driving_hits = mainModel.csModel.driving_hits;
                                        countrySpecific.credit_hits = mainModel.csModel.credit_hits;
                                        break;
                                }
                                try
                                {
                                    _context.CSComment.Update(countrySpecific);
                                    _context.SaveChanges();
                                }
                                catch
                                {
                                    _context.CSComment.Add(countrySpecific);
                                    _context.SaveChanges();
                                }
                            }

                            TempData["CS"] = "Done";
                        }
                    }
                    catch (IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                case "SaveACData":                    
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
                                    List<Additional_CountriesModel> additional_Countries = _context.DbadditionalCountries
                                       .Where(a => a.record_Id == family_record_id)
                                       .ToList();
                                    if (additional_Countries == null || additional_Countries.Count == 0)
                                    {
                                        for (int i = 0; i < mainModel.diligence.additional_Countries.Count; i++)
                                        {
                                            Additional_CountriesModel additional_Countries1 = new Additional_CountriesModel();
                                            additional_Countries1.record_Id = family_record_id;
                                            additional_Countries1.additionalCountries = mainModel.diligence.additional_Countries[i].additionalCountries;
                                            _context.DbadditionalCountries.Add(additional_Countries1);
                                            _context.SaveChanges();
                                        }
                                    }
                                    else
                                    {
                                        for (int i = 0; i < additional_Countries.Count; i++)
                                        {
                                            additional_Countries[i].record_Id = family_record_id;
                                            additional_Countries[i].additionalCountries = additional_Countries[i].additionalCountries;
                                            _context.DbadditionalCountries.Update(additional_Countries[i]);
                                            _context.SaveChanges();
                                        }
                                        try
                                        {
                                            for (int i = 0; i < mainModel.diligence.additional_Countries.Count; i++)
                                            {
                                                if (mainModel.diligence.additional_Countries[i].additionalCountries.ToString().Equals("Afghanistan")) { }
                                                else
                                                {
                                                    Additional_CountriesModel additional_Countries1 = new Additional_CountriesModel();
                                                    additional_Countries1.record_Id = family_record_id;
                                                    additional_Countries1.additionalCountries = mainModel.diligence.additional_Countries[i].additionalCountries;
                                                    _context.DbadditionalCountries.Add(additional_Countries1);
                                                    _context.SaveChanges();
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                catch
                                {
                                    for (int i = 0; i < mainModel.diligence.additional_Countries.Count; i++)
                                    {
                                        Additional_CountriesModel additional_Countries1 = new Additional_CountriesModel();
                                        additional_Countries1.record_Id = family_record_id;                                       
                                        _context.DbadditionalCountries.Add(additional_Countries1);
                                        _context.SaveChanges();
                                    }
                                }
                                mainModel.additional_Countries = _context.DbadditionalCountries
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
                                            if (mainModel.diligence.currentResidences[i].CurrentStreet == "" && mainModel.diligence.currentResidences[i].CurrentCity == "" && mainModel.diligence.currentResidences[i].CurrentCountry == "" && mainModel.diligence.currentResidences[i].CurrentZipcode == "") { }
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
                           .Where(u => u.record_Id == family_record_id)
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
                            resulttableModel.bankruptcy_filings = mainModel.summarymodel.bankruptcy_filings;
                            var checkBox = Request.Form["bankruptcy_filings1"];
                            if (checkBox == "on")
                            {
                                resulttableModel.bankruptcy_filings1 = true;
                            }
                            else
                            {
                                resulttableModel.bankruptcy_filings1 = mainModel.summarymodel.bankruptcy_filings1;
                            }
                            resulttableModel.civil_court_Litigation = mainModel.summarymodel.civil_court_Litigation;
                            var checkBox1 = Request.Form["civil_court_Litigation1"];
                            if (checkBox1 == "on")
                            {
                                resulttableModel.civil_court_Litigation1 = true;
                            }
                            else
                            {
                                resulttableModel.civil_court_Litigation1 = mainModel.summarymodel.civil_court_Litigation1;
                            }
                            resulttableModel.civil_judge_Liens = mainModel.summarymodel.civil_judge_Liens;
                            var chekbox5 = Request.Form["civil_judge_Liens1"];
                            if (chekbox5 == "on")
                            {
                                resulttableModel.civil_judge_Liens1 = true;
                            }
                            else
                            {
                                resulttableModel.civil_judge_Liens1 = mainModel.summarymodel.civil_judge_Liens1;
                            }
                            resulttableModel.criminal_records = mainModel.summarymodel.criminal_records;
                            var checkBox2 = Request.Form["criminal_records1"];
                            if (checkBox2 == "on")
                            {
                                resulttableModel.criminal_records1 = true;
                            }
                            else
                            {
                                resulttableModel.criminal_records1 = mainModel.summarymodel.criminal_records1;
                            }
                            resulttableModel.social_securitytrace = mainModel.summarymodel.social_securitytrace;
                            resulttableModel.real_estate_prop = mainModel.summarymodel.real_estate_prop;
                            resulttableModel.secretary_state_director = mainModel.summarymodel.secretary_state_director;
                            resulttableModel.driving_history = mainModel.summarymodel.driving_history;
                            resulttableModel.credit_history = mainModel.summarymodel.credit_history;
                            resulttableModel.uniform_commercial = mainModel.summarymodel.uniform_commercial;
                            resulttableModel.uniform_commercial1 = mainModel.summarymodel.uniform_commercial1;
                            resulttableModel.news_media_searches = mainModel.summarymodel.news_media_searches;
                            resulttableModel.department_foreign = mainModel.summarymodel.department_foreign;
                            resulttableModel.european_union = mainModel.summarymodel.european_union;
                            resulttableModel.HM_treasury = mainModel.summarymodel.HM_treasury;
                            resulttableModel.US_bureau = mainModel.summarymodel.US_bureau;
                            resulttableModel.US_department = mainModel.summarymodel.US_department;
                            resulttableModel.US_Directorate = mainModel.summarymodel.US_Directorate;
                            resulttableModel.US_general = mainModel.summarymodel.US_general;
                            if (mainModel.summarymodel.US_office == null)
                            {
                                resulttableModel.US_office = "Clear";
                            }
                            else
                            {
                                resulttableModel.US_office = mainModel.summarymodel.US_office;
                            }
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
                            resulttableModel.financial_regulator_ireland = mainModel.summarymodel.financial_regulator_ireland;
                            resulttableModel.hongkong_monetary_authority = mainModel.summarymodel.hongkong_monetary_authority;
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
                            resulttableModel.securities_exchange_commission = mainModel.summarymodel.securities_exchange_commission;
                            resulttableModel.securities_futuresauthority = mainModel.summarymodel.securities_futuresauthority;
                            resulttableModel.swedish_financial_supervisory = mainModel.summarymodel.swedish_financial_supervisory;
                            resulttableModel.swiss_federal_banking = mainModel.summarymodel.swiss_federal_banking;
                            resulttableModel.U_K_companies_disqualified = mainModel.summarymodel.U_K_companies_disqualified;
                            resulttableModel.US_court = mainModel.summarymodel.US_court;
                            resulttableModel.US_department_justice = mainModel.summarymodel.US_department_justice;
                            resulttableModel.US_federal_trade = mainModel.summarymodel.US_federal_trade;
                            resulttableModel.US_office_thrifts = mainModel.summarymodel.US_office_thrifts;
                            resulttableModel.central_intelligence = mainModel.summarymodel.central_intelligence;
                            try
                            {
                                _context.summaryResulttableModels.Update(resulttableModel);
                                _context.SaveChanges();
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
                        if (SaveData.Contains("Emp_"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            EmployerModel employerModel1 = _context.DbEmployer
                                      .Where(a => a.id == id_c).FirstOrDefault();
                            _context.DbEmployer.Remove(employerModel1);
                            _context.SaveChanges();
                            mainModel.EmployerModel = _context.DbEmployer
                            .Where(u => u.record_Id == family_record_id)
                            .ToList();
                        }
                        if (SaveData.Contains("Edu_"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            EducationModel educationModel = _context.DbEducation
                                      .Where(a => a.id == id_c).FirstOrDefault();

                            _context.DbEducation.Remove(educationModel);
                            _context.SaveChanges();
                            mainModel.EmployerModel = _context.DbEmployer
                            .Where(u => u.record_Id == family_record_id)
                            .ToList();
                        }
                        if (SaveData.Contains("PL_"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            PLLicenseModel pLLicenseModel = _context.DbPLLicense
                                      .Where(a => a.id == id_c).FirstOrDefault();

                            _context.DbPLLicense.Remove(pLLicenseModel);
                            _context.SaveChanges();
                            mainModel.pllicenseModels = _context.DbPLLicense
                            .Where(u => u.record_Id == family_record_id)
                            .ToList();
                        }
                        if (SaveData.Contains("Ref_"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            ReferencetableModel referencetable = _context.DbReferencetableModel
                                      .Where(a => a.id == id_c).FirstOrDefault();

                            _context.DbReferencetableModel.Remove(referencetable);
                            _context.SaveChanges();

                            mainModel.referencetableModels = _context.DbReferencetableModel
                            .Where(u => u.record_Id == family_record_id)
                            .ToList();
                        }
                        if (SaveData.Contains("CurrentRes_"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            CurrentResidenceModel current = _context.DbcurrentResidenceModels
                                      .Where(a => a.id == id_c).FirstOrDefault();

                            _context.DbcurrentResidenceModels.Remove(current);
                            _context.SaveChanges();

                            mainModel.currentResidenceModels = _context.DbcurrentResidenceModels
                            .Where(u => u.record_Id == family_record_id)
                            .ToList();
                        }
                        if (SaveData.Contains("PreviousRes_"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            PreviousResidenceModel previous = _context.DbpreviousResidenceModels
                                      .Where(a => a.id == id_c).FirstOrDefault();

                            _context.DbpreviousResidenceModels.Remove(previous);
                            _context.SaveChanges();

                            mainModel.PreviousResidenceModels = _context.DbpreviousResidenceModels
                            .Where(u => u.record_Id == family_record_id)
                            .ToList();
                        }
                        if (SaveData.Contains("Delete_Country"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            Additional_CountriesModel additional = _context.DbadditionalCountries
                                      .Where(a => a.id == id_c).FirstOrDefault();

                            _context.DbadditionalCountries.Remove(additional);
                            _context.SaveChanges();

                            mainModel.additional_Countries = _context.DbadditionalCountries
                               .Where(u => u.record_Id == family_record_id)
                               .ToList();
                        }
                    }
                    catch { }
                    break;
            }
            mainModel.additional_Countries = _context.DbadditionalCountries
                                 .Where(u => u.record_Id == family_record_id)
                                 .ToList();
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
        public IActionResult Save_Page1(MainModel diligenceInputs)
        {
            string regflag = "";
            string holdresult = "";
            string templatePath;
            string savePath = _config.GetValue<string>("ReportPath");
            MainModel diligenceInput = new MainModel();
            string recordid = HttpContext.Session.GetString("recordid");
            diligenceInput.familyModels = _context.familyModels
                 .Where(u => u.Family_record_id == diligenceInputs.familyModels[0].Family_record_id)
                 .ToList();
            diligenceInput.clientspecific = _context.DbclientspecificModels
                 .Where(u => u.record_Id == diligenceInputs.familyModels[0].Family_record_id)
                 .FirstOrDefault();                     
            diligenceInput.diligenceInputModel = _context.DbPersonalInfo
                .Where(u => u.record_Id == diligenceInputs.familyModels[0].Family_record_id)
                .FirstOrDefault();            
            if (diligenceInputs.familyModels.Count == 1)
            {
                templatePath = _config.GetValue<string>("INCITIInditemplatePath");
            }
            else
            {
                templatePath = _config.GetValue<string>("INCITIFamilytemplatePath");
            }
            switch (diligenceInput.clientspecific.clientname)
            {
                case "Malta":
                    templatePath = string.Concat(templatePath, "INTL_CITI_Individual_SterlingDiligenceReport(XXXXX)_DRAFT(MALTA).docx");
                    break;
                case "Dominica":
                    templatePath = string.Concat(templatePath, "INTL_CITI_Individual_SterlingDiligenceReport(XXXXX)_DRAFT(Dominica).docx");
                    break;
                case "St. Kitts":
                    templatePath = string.Concat(templatePath, "INTL_CITI_Individual_SterlingDiligenceReport(XXXXX)_DRAFT(St_Kitts).docx");
                    break;
                default:
                    templatePath = string.Concat(templatePath, "INTL_CITI_Individual_SterlingDiligenceReport(XXXXX)_DRAFT(Cyprus).docx");
                    break;
            }
            string pathTo = _config.GetValue<string>("OlderReport"); // the destination file name would be appended later            
            savePath = string.Concat(savePath, diligenceInput.diligenceInputModel.LastName.ToString(), "_SterlingDiligenceReport(", diligenceInput.diligenceInputModel.CaseNumber.ToString(), ")_DRAFT.docx");
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
                fileInfo.MoveTo(Path.Combine(pathTo, string.Format("{0}{1}", filename, newExtension)));
            }
            string strblnres = "";
            Document doc = new Document(templatePath);
            string Client_Name = diligenceInput.clientspecific.clientname;
            Table table1 = doc.Sections[3].Tables[0] as Table;
            
            for (int i = 1; i < diligenceInputs.familyModels.Count(); i++)
            {
                Table table2 = table1.Clone();
                doc.Sections[3].Tables.Add(table2);
                doc.SaveToFile(savePath);
            }
            doc.SaveToFile(savePath);
           // int childcount = 0;
            for (int j = 0; j < diligenceInputs.familyModels.Count(); j++)
            {
                string strpreviousresidence = "";
                string strcurrentresidence = "";
                int currowcount = 5;
                diligenceInput.diligenceInputModel = _context.DbPersonalInfo
              .Where(u => u.record_Id == diligenceInputs.familyModels[j].Family_record_id)
              .FirstOrDefault();
                diligenceInput.currentResidenceModels = _context.DbcurrentResidenceModels
                .Where(u => u.record_Id == diligenceInputs.familyModels[j].Family_record_id)
                .ToList();
                diligenceInput.PreviousResidenceModels = _context.DbpreviousResidenceModels
                .Where(u => u.record_Id == diligenceInputs.familyModels[j].Family_record_id)
                .ToList();
                diligenceInput.additional_Countries = _context.DbadditionalCountries
               .Where(u => u.record_Id == diligenceInputs.familyModels[j].Family_record_id)
               .ToList();
                Table table3 = doc.Sections[3].Tables[j] as Table;
                TableCell cell11 = table3.Rows[0].Cells[0];
                Paragraph p1 = cell11.Paragraphs[0];
                p1.Text = string.Concat("\n", diligenceInput.diligenceInputModel.FirstName.ToUpper(), " ", diligenceInput.diligenceInputModel.MiddleName.ToUpper(), " ", diligenceInput.diligenceInputModel.LastName.ToUpper(), " (", diligenceInput.diligenceInputModel.foreignlanguage, ") (", diligenceInputs.familyModels[j].applicant_type, ")");
                TableCell cel2 = table3.Rows[1].Cells[1];
                Paragraph p2 = cel2.Paragraphs[0];
                try
                {
                    p2.Text = diligenceInput.diligenceInputModel.FullaliasName;
                }
                catch
                {
                    p2.Text = "<not provided>";
                }
                TableCell cel3 = table3.Rows[2].Cells[1];
                Paragraph p3 = cel3.Paragraphs[0];
                try
                {
                    p3.Text = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Month) + " " + Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Day + ", " + Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Year;
                }
                catch
                {
                    p3.Text = "<not provided>";
                }
                TableCell cel4 = table3.Rows[3].Cells[1];
                Paragraph p4 = cel4.Paragraphs[0];
                try
                {
                    if (diligenceInput.diligenceInputModel.Pob.ToString().Equals("") || diligenceInput.diligenceInputModel.Pob.ToString().Equals("NA") || diligenceInput.diligenceInputModel.Pob.ToString().Equals("N/A"))
                    {
                        p4.Text = "<not provided>";
                    }
                    else
                    {
                        p4.Text = diligenceInput.diligenceInputModel.Pob;
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
                    p5.Text = diligenceInput.diligenceInputModel.Nationality;
                }
                catch
                {
                    p5.Text = "<not provided>";
                }
                TableCell cel6 = table3.Rows[5].Cells[1];
                Paragraph p6 = cel6.Paragraphs[0];

                for (int i = 0; i < diligenceInput.currentResidenceModels.Count(); i++)
                {
                    if (diligenceInput.currentResidenceModels[i].CurrentStreet.ToString().Equals("")) { }
                    else
                    {
                        strcurrentresidence = diligenceInput.currentResidenceModels[i].CurrentStreet;
                    }
                    if (diligenceInput.currentResidenceModels[i].CurrentCity.ToString().Equals("")) { }
                    else
                    {
                        strcurrentresidence = string.Concat(strcurrentresidence, ", ", diligenceInput.currentResidenceModels[i].CurrentCity);
                    }
                    if (diligenceInput.currentResidenceModels[i].CurrentCountry.ToString().Equals("")) { }
                    else
                    {
                        strcurrentresidence = string.Concat(strcurrentresidence, ", ", diligenceInput.currentResidenceModels[i].CurrentCountry);
                    }
                    if (diligenceInput.currentResidenceModels[i].CurrentZipcode.ToString().Equals("")) { }
                    else
                    {
                        strcurrentresidence = string.Concat(strcurrentresidence, ", ", diligenceInput.currentResidenceModels[i].CurrentZipcode);
                    }
                    if (i == 0)
                    {
                        if (strcurrentresidence.ToString().Equals(""))
                        {
                            
                        }
                        else
                        {
                            p6.Text = string.Concat(strcurrentresidence, ", reportedly since <date>");
                        }
                        if (diligenceInput.currentResidenceModels.Count() == 1)
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
                            if (diligenceInput.currentResidenceModels.Count() == 2)
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
                    if (diligenceInput.currentResidenceModels.Count() == 1)
                    {
                        doc.Replace("Current Residences", "Current Residence", true, true);
                    }
                }
                catch { }
                //Previous Residence                
                for (int i = 0; i < diligenceInput.PreviousResidenceModels.Count(); i++)
                {
                    TableCell cel10 = table3.Rows[currowcount + 1].Cells[1];
                    Paragraph p10 = cel10.Paragraphs[0];
                    if (diligenceInput.PreviousResidenceModels[i].PreviousStreet.ToString().Equals("")) { }
                    else
                    {
                        strpreviousresidence = diligenceInput.PreviousResidenceModels[i].PreviousStreet;
                    }
                    if (diligenceInput.PreviousResidenceModels[i].PreviousCity.ToString().Equals("")) { }
                    else
                    {
                        strpreviousresidence = string.Concat(strpreviousresidence, ", ", diligenceInput.PreviousResidenceModels[i].PreviousCity);
                    }
                    if (diligenceInput.PreviousResidenceModels[i].PreviousCountry.ToString().Equals("")) { }
                    else
                    {
                        strpreviousresidence = string.Concat(strpreviousresidence, ", ", diligenceInput.PreviousResidenceModels[i].PreviousCountry);
                    }
                    if (diligenceInput.PreviousResidenceModels[i].PreviousZipcode.ToString().Equals("")) { }
                    else
                    {
                        strpreviousresidence = string.Concat(strpreviousresidence, ", ", diligenceInput.PreviousResidenceModels[i].PreviousZipcode);
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
                        if (diligenceInput.PreviousResidenceModels.Count() == 1)
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
                            if (diligenceInput.PreviousResidenceModels.Count() == 2)
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
                    if (diligenceInput.PreviousResidenceModels.Count() == 1)
                    {
                        doc.Replace("Previous Residences", "Previous Residence", true, true);
                    }
                    else
                    {
                        if (diligenceInput.PreviousResidenceModels.Count() == 0)
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
                if (diligenceInputs.familyModels[j].adult_minor == "Adult")
                {                  
                    //Marital_Status
                    switch (diligenceInput.diligenceInputModel.Marital_Status)
                    {
                        case "Married":
                            p9.Text = string.Concat("According to their Application Form  (a copy of which was provided), ", diligenceInput.diligenceInputModel.LastName, " married <spouse> on <date> in <location> (a copy of their Marriage Certificate, issued by <authority> number <number>, was provided and was authenticated).[MARRITALSTATUSDESC]");
                            break;
                        case "Never Married":
                            p9.Text = string.Concat("According to their Application Form  (a copy of which was provided), ", diligenceInput.diligenceInputModel.LastName, " has never been married.[MARRITALSTATUSDESC]");
                            break;
                        case "Divorced":
                            p9.Text = string.Concat("According to their Application Form  (a copy of which was provided), ", diligenceInput.diligenceInputModel.LastName, " previously married <spouse> on <date> in <location> (a copy of their Marriage Certificate, issued by <authority> number <number>, was provided and was authenticated).  They divorced on <date> in <location>  (a copy of their Divorce Certificate, issued by <authority> number <number>, was provided and was authenticated).[MARRITALSTATUSDESC]");
                            break;

                    }
                    doc.SaveToFile(savePath);
                    switch (diligenceInput.diligenceInputModel.Children)
                    {
                        case "Children":
                            doc.Replace("[MARRITALSTATUSDESC]", string.Concat("\n\nThey have <number> children together, who reside with their parents at the aforenoted residential address in [Country]: <Name>, born on <date>; and <Name> born on <date>.  As noted above, ", diligenceInput.diligenceInputModel.LastName, "’s spouse and children are not included in this application."), true, false);
                            break;
                        case "No Children":
                            doc.Replace("[MARRITALSTATUSDESC]", string.Concat("\n\nAccording to their Application Form, ", diligenceInput.diligenceInputModel.LastName, " has no children."), true, false);
                            break;
                    }
                    //if (j == 0)
                    //{
                        //Military_Service
                        if (diligenceInput.clientspecific.clientname.ToString().Equals("Malta") || diligenceInput.clientspecific.clientname.ToString().Equals("Dominica") || diligenceInput.clientspecific.clientname.ToString().Equals("St. Kitts"))
                        {
                            currowcount = currowcount + 1;
                            TableCell cel11 = table3.Rows[currowcount].Cells[1];
                            Paragraph p11 = cel11.Paragraphs[0];
                            switch (diligenceInput.clientspecific.Military_Service)
                            {
                                case "Service Confirmed":
                                    p11.Text = string.Concat("As confirmed, ", diligenceInput.diligenceInputModel.LastName, " served as a <rank/position> in the <Country> military from <year> to <year>.  <investigator to specify whether military certificate was provided and/or authenticated>");
                                    doc.Replace("MILITARYRES", "Confirmed", true, true);
                                    doc.Replace("MILITARYCOMMENT", "As confirmed, [Last Name] served as a <rank/position> in the <Country> military from <year> to <year>", true, true);
                                    break;
                                case "Service Unconfirmed":
                                    p11.Text = string.Concat(diligenceInput.diligenceInputModel.LastName, " reportedly served as a <rank/position> in the <Country> military from <year> to <year>, however, <investigator to specify why not confirmed and whether certificate was provided>.");
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
                TextSelection[] text30 = doc.FindAllString(string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName), false, false);
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
                    //doc.Replace(string.Concat(diligenceInput.diligenceInputModel.MiddleName, "_", j), diligenceInput.diligenceInputModel.MiddleName, true, true);
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
            for (int i = 0; i < diligenceInputs.familyModels.Count; i++)
            {
                diligenceInput.diligenceInputModel = _context.DbPersonalInfo
                .Where(u => u.record_Id == diligenceInputs.familyModels[i].Family_record_id)
                .FirstOrDefault();
                if (diligenceInputs.familyModels.Count == 1)
                {
                    Family_namesub = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                    try { if (diligenceInput.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.foreignlanguage); } } catch { }
                    Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.Country);
                    familyname = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                    family_middle = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName);
                    family_applicant = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName);
                }
                else
                {
                    if (diligenceInputs.familyModels.Count == 2)
                    {
                        if (diligenceInputs.familyModels[i].adult_minor == "Adult")
                        {
                            if (i == 0)
                            {
                                Family_namesub = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                                try { if (diligenceInput.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.foreignlanguage); } } catch { }
                                Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.Country);
                                familyname = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName, " and ");
                                family_middle = string.Concat(family_middle, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, " and ");
                                family_applicant = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type, " and ");
                            }
                            else
                            {
                                Family_namesub = string.Concat(Family_namesub,"\n\n",diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                                try { if (diligenceInput.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.foreignlanguage); } } catch { }
                                Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.Country);
                                familyname = string.Concat(familyname, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                                family_middle = string.Concat(family_middle, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName);
                                family_applicant = string.Concat(family_applicant, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type);
                            }
                        }
                        else
                        {
                            familyname = string.Concat(familyname, "child");
                            family_middle = string.Concat(family_middle, "child");
                            age = DateTime.Now.Year - Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Year;
                            family_applicant = string.Concat(family_applicant, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type, ", who is ", age, " years old, ");
                        }
                    }
                    else
                    {
                        if (diligenceInputs.familyModels[i].adult_minor == "Adult")
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
                for (int i = 0; i < diligenceInputs.familyModels.Count; i++)
                {
                    diligenceInput.diligenceInputModel = _context.DbPersonalInfo
                                   .Where(u => u.record_Id == diligenceInputs.familyModels[i].Family_record_id)
                                   .FirstOrDefault();
                    if (minor_count == 0)
                    {
                        if (i == 0)
                        {
                            Family_namesub = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                            try { if (diligenceInput.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.foreignlanguage); } } catch { }
                            Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.Country);
                        }
                        else
                        {
                            Family_namesub = string.Concat(Family_namesub, "\n\n", diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                            try { if (diligenceInput.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.foreignlanguage); } } catch { }
                            Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.Country);
                        }
                        if (i == diligenceInputs.familyModels.Count - 1)
                        {
                            familyname = string.Concat(familyname, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                            family_middle = string.Concat(family_middle, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName);
                            family_applicant = string.Concat(family_applicant, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type, ", ");
                        }
                        else
                        {
                            if (i == diligenceInputs.familyModels.Count - 2)
                            {
                                familyname = string.Concat(familyname, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName, " and ");
                                family_middle = string.Concat(family_middle, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, " and ");
                                family_applicant = string.Concat(family_applicant, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type, " and ");
                            }
                            else
                            {
                                familyname = string.Concat(familyname, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName, ", ");
                                family_middle = string.Concat(family_middle, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ");
                                family_applicant = string.Concat(family_applicant, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type, ", ");
                            }
                        }
                    }
                    else
                    {
                        if (diligenceInputs.familyModels[i].adult_minor == "Adult")
                        {
                            if (i == 0)
                            {
                                Family_namesub = string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                                try { if (diligenceInput.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.foreignlanguage); } } catch { }
                                Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.Country);
                            }
                            else
                            {
                                Family_namesub = string.Concat(Family_namesub,"\n\n",diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName);
                                try { if (diligenceInput.diligenceInputModel.foreignlanguage == "") { } else { Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.foreignlanguage); } } catch { }
                                Family_namesub = string.Concat(Family_namesub, "\n", diligenceInput.diligenceInputModel.Country);
                            }
                            if (i == diligenceInputs.familyModels.Count - 2)
                            {
                                familyname = string.Concat(familyname, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName, " and ");
                                family_middle = string.Concat(family_middle, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, " and ");
                                family_applicant = string.Concat(family_applicant, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type, " and ");
                            }
                            else
                            {
                                familyname = string.Concat(familyname, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleInitial, " ", diligenceInput.diligenceInputModel.LastName, ", ");
                                family_middle = string.Concat(family_middle, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ");
                                family_applicant = string.Concat(family_applicant, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type, ", ");
                            }
                        }
                        else
                        {
                            if (i == diligenceInputs.familyModels.Count - 2)
                            {                                
                                familyname = string.Concat(familyname, " and ");
                                family_middle = string.Concat(family_middle, " and ");                                
                            }
                            if (i == diligenceInputs.familyModels.Count - 1)
                            {
                                family_applicant = string.Concat(family_applicant, " and ");
                            }
                            age = DateTime.Now.Year - Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Year;
                            family_applicant = string.Concat(family_applicant, diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.MiddleName, " ", diligenceInput.diligenceInputModel.LastName, ", ", diligenceInputs.familyModels[i].applicant_type, ", who is ", age, " years old, ");
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
            //Additional Country
            string country_comment = "";
            for (int i = 0; i < diligenceInputs.familyModels.Count; i++)
            {
                diligenceInput.diligenceInputModel = _context.DbPersonalInfo
                .Where(u => u.record_Id == diligenceInputs.familyModels[i].Family_record_id)
                .FirstOrDefault();
                if (diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Equals(""))
                {
                }
                else
                {
                    CommentModel comment2 = _context.DbComment
                                  .Where(u => u.Comment_type == "NonScopeCountry")
                                  .FirstOrDefault();
                    if (diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Contains(",") || diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Contains(" and "))
                    {
                        country_comment = string.Concat(country_comment, "\n", diligenceInput.diligenceInputModel.FirstName.ToUpper(), " ", diligenceInput.diligenceInputModel.MiddleName.ToUpper(), " ", diligenceInput.diligenceInputModel.LastName.ToUpper(), "\n\n", comment2.confirmed_comment.ToString(), "\n");
                        country_comment = country_comment.Replace("[Non-ScopeCountries]", diligenceInput.diligenceInputModel.Nonscopecountry1.ToString());
                    }
                    else
                    {
                        country_comment = string.Concat(country_comment, "\n", diligenceInput.diligenceInputModel.FirstName.ToUpper(), " ", diligenceInput.diligenceInputModel.MiddleName.ToUpper(), " ", diligenceInput.diligenceInputModel.LastName.ToUpper(), "\n\n", comment2.unconfirmed_comment.ToString(), "\n");
                        country_comment = country_comment.Replace("[Non-ScopeCountry]", diligenceInput.diligenceInputModel.Nonscopecountry1.ToString());
                    }
                }
            }
            country_comment = country_comment.Replace("the subjects", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName));
            country_comment = country_comment.Replace("the subject", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName));
            doc.Replace("ADDITIONALJURIDICTIONS", country_comment, true, true);
            doc.SaveToFile(savePath);
           
            for (int famcount = 0; famcount < diligenceInputs.familyModels.Count; famcount++)
            {
                diligenceInput.familyModels = _context.familyModels
               .Where(u => u.Family_record_id == diligenceInputs.familyModels[famcount].Family_record_id)
               .ToList();
                diligenceInput.clientspecific = _context.DbclientspecificModels
                .Where(u => u.record_Id == diligenceInputs.familyModels[0].Family_record_id)
                .FirstOrDefault();
                diligenceInput.pllicenseModels = _context.DbPLLicense
                    .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                    .ToList();
                diligenceInput.EmployerModel = _context.DbEmployer
                    .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                    .ToList();
                diligenceInput.educationModels = _context.DbEducation
                    .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                    .ToList();
                diligenceInput.referencetableModels = _context.DbReferencetableModel
                   .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                   .ToList();
                diligenceInput.summarymodel = _context.summaryResulttableModels
                     .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                     .FirstOrDefault();
                diligenceInput.otherdetails = _context.othersModel
                    .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                    .FirstOrDefault();
                diligenceInput.diligenceInputModel = _context.DbPersonalInfo
                    .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                    .FirstOrDefault();
                diligenceInput.csModel = _context.CSComment
                   .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                   .FirstOrDefault();
                if (diligenceInput.diligenceInputModel.Country.ToString().Equals("Hong Kong SAR"))
                {
                    diligenceInput.diligenceInputModel.Country = "Hong Kong";
                }
                string record_Id = HttpContext.Session.GetString("record_Id");
                string strpreviousresidence = "";
                string strcurrentresidence = "";
                int currowcount = 5;
                if (famcount == 0)
                {
                    if (Client_Name.ToString().Equals("Malta"))
                    {
                        switch (diligenceInput.clientspecific.source_of_wealth)
                        {
                            case "Wealth Reported":
                                doc.Replace("SOURCERESULT", "Records & Unconfirmed", true, true);
                                doc.Replace("SOURCECOMMENT", "[Last Name] has an estimated total net worth of <amount>, which is comprised of <list of assets> in <location>\n\n<Investigator to specify whether confirmation is available and results if undertaken>", true, true);
                                doc.Replace("SOURCEOFWEALTHDESC", "According to the subject’s Form SSFW and Asset Declaration Template (a copy of which was provided), [Last Name] has an estimated total net worth of <amount>, which is comprised of <sources of wealth>.\n\nIt should be noted that confirmation of the subject’s salary cannot be undertaken without directly contacting their current employer.\n\nAdditionally, per the subject’s Form SSFW and Asset Declaration Template, [Last Name] holds the following assets in <locations>:\n\n\t•	<investigator to bullet list of assets with description, amounts and locations, and details\t\ton confirmation efforts>", true, true);
                                break;
                            case "No Wealth Reported":
                                doc.Replace("SOURCERESULT", "N/A", true, true);
                                doc.Replace("SOURCECOMMENT", "[Last Name] did not provide any details regarding their source of wealth, and information in this regard -- if any -- would be required in order to pursue confirmation of the same", true, true);
                                doc.Replace("SOURCEOFWEALTHDESC", "[Last Name] did not provide any details regarding their source of wealth, and information in this regard -- if any -- would be required in order to pursue confirmation of the same.\n\nIt should be noted that confirmation of the subject’s salary cannot be undertaken without directly contacting their current employer.", true, true);
                                break;
                        }

                        if (diligenceInput.clientspecific.discreet_reputational_inquiries.ToString().Equals("Clear") || diligenceInput.clientspecific.discreet_reputational_inquiries.ToString().Equals("Potentially-Relevant Information"))
                        {

                            doc.Replace("DISCREETREPUTATIONALINQUIRIESDESC", "Discreet reputational inquiries were undertaken in connection with sources familiar with [Last Name] in [Country]. We contacted these individuals and asked them about their knowledge of the applicant’s background, reputation, character, and integrity.\n\nThe following are the summaries of the inquiries:\n\nSOURCE1HEADER\n\nAccording to the source, <investigator to insert summary>.\n\nSOURCE2HEADER\n\nAccording to the source, <investigator to insert summary>.\n\nSOURCE3HEADER\n\nAccording to the source, <investigator to insert summary>.", true, true);
                            doc.SaveToFile(savePath);
                            for (int l = 0; l < 3; l++)
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
                                                    if (abc.ToString().Contains("SOURCE1HEADER"))
                                                    {
                                                        textRange = para.AppendText("Source 1:  <Description of Source>");
                                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                        textRange.CharacterFormat.FontSize = 11;
                                                        textRange.CharacterFormat.Italic = true;
                                                        textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                                        doc.SaveToFile(savePath);
                                                        doc.Replace("SOURCE1HEADER", "", true, false);
                                                        doc.SaveToFile(savePath);
                                                        disbnres = "true";
                                                        break;
                                                    }
                                                    if (abc.ToString().Contains("SOURCE2HEADER"))
                                                    {
                                                        textRange = para.AppendText("Source 2:  <Description of Source>");
                                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                        textRange.CharacterFormat.FontSize = 11;
                                                        textRange.CharacterFormat.Italic = true;
                                                        textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                                        doc.SaveToFile(savePath);
                                                        doc.Replace("SOURCE2HEADER", "", true, false);
                                                        doc.SaveToFile(savePath);
                                                        disbnres = "true";
                                                        break;
                                                    }
                                                    if (abc.ToString().Contains("SOURCE3HEADER"))
                                                    {
                                                        textRange = para.AppendText("Source 3:  <Description of Source>");
                                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                        textRange.CharacterFormat.FontSize = 11;
                                                        textRange.CharacterFormat.Italic = true;
                                                        textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                                        doc.SaveToFile(savePath);
                                                        doc.Replace("SOURCE3HEADER", "", true, false);
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
                        }
                        else
                        {
                            doc.Replace("DISCREETREPUTATIONALINQUIRIESDESC", "Discreet reputational inquiries were undertaken in connection with sources familiar with [Last Name] in [Country].  However, numerous attempts to identify individuals familiar with the subject were unsuccessful.", true, true);
                        }

                        switch (diligenceInput.clientspecific.discreet_reputational_inquiries)
                        {
                            case "Clear":
                                doc.Replace("DISCRETRES", "Clear", true, true);
                                doc.Replace("DISCRETCOMMENT", "Discreet reputational inquiries did not reveal any adverse or materially-significant information in connection with [Last Name] in [Country]", true, true);
                                break;
                            case "Potentially-Relevant Information":
                                doc.Replace("DISCRETRES", "Potentially-Relevant Information", true, true);
                                doc.Replace("DISCRETCOMMENT", "Discreet reputational inquiries conducted in connection with [Last Name] in [Country] revealed <investigator to insert summary of findings>", true, true);
                                break;
                            case "Unsuccessful":
                                doc.Replace("DISCRETRES", "Unsuccessful", true, true);
                                doc.Replace("DISCRETCOMMENT", "Discreet reputational inquiries were undertaken in connection with sources familiar with [Last Name] in [Country].  However, numerous attempts to identify individuals familiar with the subject were unsuccessful", true, true);
                                break;
                        }

                        //Military_Service
                        if (Client_Name.ToString().Equals("Malta") || Client_Name.ToString().Equals("Dominica") || Client_Name.ToString().Equals("St. Kitts"))
                        {
                            switch (diligenceInput.clientspecific.Military_Service)
                            {
                                case "Service Confirmed":
                                    doc.Replace("[MILITARYSERVICEDESC]", "As confirmed, [Last Name] served as a <rank/position> in the <Country> military from <year> to <year>.  <investigator to specify whether military certificate was provided and/or authenticated>", true, true);
                                    doc.Replace("MILITARYRES", "Confirmed", true, true);
                                    doc.Replace("MILITARYCOMMENT", "As confirmed, [Last Name] served as a <rank/position> in the <Country> military from <year> to <year>", true, true);
                                    break;
                                case "Service Unconfirmed":
                                    doc.Replace("[MILITARYSERVICEDESC]", "[Last Name] reportedly served as a <rank/position> in the <Country> military from <year> to <year>, however, <investigator to specify why not confirmed and whether certificate was provided>.", true, true);
                                    doc.Replace("MILITARYRES", "Unconfirmed", true, true);
                                    doc.Replace("MILITARYCOMMENT", "[Last Name] reportedly served as a <rank/position> in the <Country> military from <year> to <year>, however, <investigator to specify why not confirmed>", true, true);
                                    break;
                                case "No Military Service":
                                    doc.Replace("[MILITARYSERVICEDESC]", "The subject reportedly has not served in the military <investigator to specify whether Exemption certificate was provided, adjust confirmation language if confirmed>.", true, true);
                                    doc.Replace("MILITARYRES", "N/A", true, true);
                                    doc.Replace("MILITARYCOMMENT", "The subject reportedly has not served in the military <investigator to specify whether Exemption certificate was provided, adjust confirmation language if confirmed>", true, true);
                                    break;
                            }
                        }
                    }
                    if (Client_Name.ToString().Equals("Dominica"))
                    {
                        string strcontactref = "";
                        switch (diligenceInput.clientspecific.Character_Reference)
                        {
                            case "Not Provided":
                                doc.Replace("CHARACTERREFERENCESDESC", "It is noted that [Last Name] did not provide any character references, and details in this regard would be required in order to pursue reference interviews in connection with the subject.", true, true);
                                break;
                            case "Provided and Contacted":
                                doc.Replace("CHARACTERREFERENCESDESC", "[Last Name] provided the following character references in [Country].  We contacted these individuals and asked them about their knowledge of the subject’s background, reputation, character and integrity.  It is noted that a reference letter was provided by an individual for [Last Name].\n\nThe following are summaries of the interviews and reference letter:[ADDCONTACTEDREFERENCE]", true, true);
                                strcontactref = "[Reference Full Name] stated that <investigator to insert summary>.\n\nFurther, [Reference Full Name] provided a reference letter, wherein <he/she> stated that <investigator to insert summary or remove if no letter provided>.";
                                break;
                            case "Provided but Unsuccessful":
                                doc.Replace("CHARACTERREFERENCESDESC", "[Last Name] provided the following character references in [Country].  We attempted to contact these individuals to ask them about their knowledge of the subject’s background, reputation, character and integrity.  It is noted that a reference letter was provided by an individual for [Last Name].\n\nThe following are summaries of the interviews and reference letter:[ADDCONTACTEDREFERENCE]", true, true);
                                strcontactref = "Numerous attempts to contact [Reference Full Name] utilizing the contact details provided have remained unsuccessful to date.\n\nHowever, [Reference Full Name] provided a reference letter, wherein <he/she> stated that <investigator to insert summary or remove if no letter provided>.";
                                break;
                        }
                        doc.SaveToFile(savePath);
                        if (diligenceInput.clientspecific.Character_Reference.ToString().Equals("Provided and Contacted") || diligenceInput.clientspecific.Character_Reference.ToString().Equals("Provided but Unsuccessful"))
                        {
                            string strrefadded = "";
                            int j = 1;
                            for (int i = 0; i < diligenceInput.referencetableModels.Count; i++)
                            {

                                strrefadded = string.Concat(strrefadded, "\n\nReference ", j, ":	[Reference_Full_Name_",i,"][Reference Position][Reference Employer][Reference Location]\n\n", strcontactref);
                                if (diligenceInput.referencetableModels[i].ref_position.Equals("")) { strrefadded = strrefadded.Replace("[Reference Position]", ""); }
                                else
                                {
                                    strrefadded = strrefadded.Replace("[Reference Position]", string.Concat("\n\t\t\t\t", diligenceInput.referencetableModels[i].ref_position));
                                }
                                if (diligenceInput.referencetableModels[i].ref_employer.Equals(""))
                                {
                                    strrefadded = strrefadded.Replace("[Reference Employer]", "");
                                }
                                else { strrefadded = strrefadded.Replace("[Reference Employer]", string.Concat("\n\t\t\t\t", diligenceInput.referencetableModels[i].ref_employer)); }
                                if (diligenceInput.referencetableModels[i].ref_location.Equals(""))
                                {
                                    strrefadded = strrefadded.Replace("[Reference Location]", "");
                                }
                                else
                                {
                                    strrefadded = strrefadded.Replace("[Reference Location]", string.Concat("\n\t\t\t\t", diligenceInput.referencetableModels[i].ref_location));
                                }

                               
                                strrefadded = strrefadded.Replace("[Reference Full Name]", diligenceInput.referencetableModels[i].ref_full_name);

                                j = j + 1;
                            }
                            doc.Replace("[ADDCONTACTEDREFERENCE]", strrefadded, true, false);
                            doc.SaveToFile(savePath);
                          
                            for (int i = 0; i < diligenceInput.referencetableModels.Count; i++)
                            {
                                TextSelection[] textref = doc.FindAllString(string.Concat("[Reference_Full_Name_", i, "]"), false, false);
                                if (textref != null)
                                {
                                    foreach (TextSelection seletion in textref)
                                    {
                                        seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                    }
                                }
                                doc.Replace(string.Concat("[Reference_Full_Name_", i, "]"), diligenceInput.referencetableModels[i].ref_full_name, false, false);
                                doc.SaveToFile(savePath);
                            }
                        }
                    }
                    doc.Replace("ClientName", diligenceInput.clientspecific.clientname.TrimEnd(), true, true);
                    doc.Replace("ClientName", diligenceInput.clientspecific.clientname.TrimEnd(), true, true);
                    doc.Replace("[Last Name]", diligenceInput.diligenceInputModel.LastName.ToString().TrimEnd(), false, false);
                    if (diligenceInput.diligenceInputModel.Country.ToString().Equals("Hong Kong SAR"))
                    {
                        diligenceInput.diligenceInputModel.Country = "Hong Kong";
                    }                    
                    doc.Replace("[Country]", diligenceInput.diligenceInputModel.Country.ToString().TrimEnd(), false, false);
                    doc.SaveToFile(savePath);                   
                    if (diligenceInputs.familyModels.Count == 1)
                    {
                        //Current Residence
                       
                        for (int i = 0; i < diligenceInput.currentResidenceModels.Count(); i++)
                        {
                            if (diligenceInput.currentResidenceModels[i].CurrentStreet.ToString().Equals("")) { }
                            else
                            {
                                strcurrentresidence = diligenceInput.currentResidenceModels[i].CurrentStreet;
                            }
                            if (diligenceInput.currentResidenceModels[i].CurrentCity.ToString().Equals("")) { }
                            else
                            {
                                strcurrentresidence = string.Concat(strcurrentresidence, ", ", diligenceInput.currentResidenceModels[i].CurrentCity);
                            }
                            if (diligenceInput.currentResidenceModels[i].CurrentCountry.ToString().Equals("")) { }
                            else
                            {
                                strcurrentresidence = string.Concat(strcurrentresidence, ", ", diligenceInput.currentResidenceModels[i].CurrentCountry);
                            }
                            if (diligenceInput.currentResidenceModels[i].CurrentZipcode.ToString().Equals("")) { }
                            else
                            {
                                strcurrentresidence = string.Concat(strcurrentresidence, ", ", diligenceInput.currentResidenceModels[i].CurrentZipcode);
                            }
                            if (i == 0)
                            {
                                if (strcurrentresidence.ToString().Equals(""))
                                {
                                    doc.Replace("[CURRENTFULLADDRESSDESC]1, reportedly since <date>", "", true, true);
                                }
                                else
                                {
                                    doc.Replace("[CURRENTFULLADDRESSDESC]1", string.Concat(strcurrentresidence), true, true);
                                }
                                if (diligenceInput.currentResidenceModels.Count() == 1)
                                {
                                    currowcount = 5;
                                    table1.Rows.RemoveAt(6);
                                    table1.Rows.RemoveAt(6);
                                }
                            }
                            else
                            {
                                if (i == 1)
                                {
                                    doc.Replace("[CURRENTFULLADDRESSDESC]2", string.Concat(strcurrentresidence), true, true);
                                    if (diligenceInput.currentResidenceModels.Count() == 2)
                                    {
                                        currowcount = 6;
                                        table1.Rows.RemoveAt(7);
                                    }
                                }
                                else
                                {
                                    currowcount = 7;
                                    doc.Replace("[CURRENTFULLADDRESSDESC]3", string.Concat(strcurrentresidence), true, true);
                                }
                                //rowcount1 = rowcount1 + 1;
                                //Table table1 = doc.Sections[3].Tables[0] as Table;
                                //TableRow row = table1.AddRow();                    
                                //table1.Rows.Insert(rowcount1, row);
                                //table1.AddRow(true, 1);                    
                            }

                        }
                        doc.SaveToFile(savePath);
                        try
                        {
                            if (diligenceInput.currentResidenceModels.Count() == 1)
                            {
                                doc.Replace("Current Residences", "Current Residence", true, true);
                            }
                        }
                        catch { }
                        //Previous Residence                        
                        for (int i = 0; i < diligenceInput.PreviousResidenceModels.Count(); i++)
                        {
                            if (diligenceInput.PreviousResidenceModels[i].PreviousStreet.ToString().Equals("")) { }
                            else
                            {
                                strpreviousresidence = diligenceInput.PreviousResidenceModels[i].PreviousStreet;
                            }
                            if (diligenceInput.PreviousResidenceModels[i].PreviousCity.ToString().Equals("")) { }
                            else
                            {
                                strpreviousresidence = string.Concat(strpreviousresidence, ", ", diligenceInput.PreviousResidenceModels[i].PreviousCity);
                            }
                            if (diligenceInput.PreviousResidenceModels[i].PreviousCountry.ToString().Equals("")) { }
                            else
                            {
                                strpreviousresidence = string.Concat(strpreviousresidence, ", ", diligenceInput.PreviousResidenceModels[i].PreviousCountry);
                            }
                            if (diligenceInput.PreviousResidenceModels[i].PreviousZipcode.ToString().Equals("")) { }
                            else
                            {
                                strpreviousresidence = string.Concat(strpreviousresidence, ", ", diligenceInput.PreviousResidenceModels[i].PreviousZipcode);
                            }
                            if (i == 0)
                            {
                                if (strpreviousresidence.ToString().Equals(""))
                                {
                                    doc.Replace("[PREVIOUSFULLADDRESSDESC]1, reportedly from <date> to <date>", "", true, true);

                                }
                                else
                                {
                                    doc.Replace("[PREVIOUSFULLADDRESSDESC]1", string.Concat(strpreviousresidence), true, true);
                                }
                                if (diligenceInput.PreviousResidenceModels.Count() == 1)
                                {
                                    table1.Rows.RemoveAt(currowcount + 2);
                                    table1.Rows.RemoveAt(currowcount + 2);
                                    if (strpreviousresidence.ToString().Equals(""))
                                    {
                                        table1.Rows.RemoveAt(currowcount + 1);
                                    }
                                }
                            }
                            else
                            {
                                if (i == 1)
                                {
                                    doc.Replace("[PREVIOUSFULLADDRESSDESC]2", string.Concat(strpreviousresidence), true, true);
                                    if (diligenceInput.PreviousResidenceModels.Count() == 2)
                                    {
                                        table1.Rows.RemoveAt(currowcount + 3);
                                    }
                                }
                                else
                                {
                                    doc.Replace("[PREVIOUSFULLADDRESSDESC]3", string.Concat(strpreviousresidence), true, true);
                                }                            
                            }
                        }
                        try
                        {
                            if (diligenceInput.PreviousResidenceModels.Count() == 1)
                            {
                                doc.Replace("Previous Residences", "Previous Residence", true, true);
                            }
                        }
                        catch { }
                        //Marital_Status
                        switch (diligenceInput.diligenceInputModel.Marital_Status)
                        {
                            case "Married":
                                doc.Replace("[MARRITALSTATUSDESC]", "According to their Application Form  (a copy of which was provided), [Last Name] married <spouse> on <date> in <location> (a copy of their Marriage Certificate, issued by <authority> number <number>, was provided and was authenticated).[MARRITALSTATUSDESC]", true, true);
                                break;
                            case "Never Married":
                                doc.Replace("[MARRITALSTATUSDESC]", "According to their Application Form  (a copy of which was provided), [Last Name] has never been married.[MARRITALSTATUSDESC]", true, true);
                                break;
                            case "Divorced":
                                doc.Replace("[MARRITALSTATUSDESC]", "According to their Application Form  (a copy of which was provided), [Last Name] previously married <spouse> on <date> in <location> (a copy of their Marriage Certificate, issued by <authority> number <number>, was provided and was authenticated).  They divorced on <date> in <location>  (a copy of their Divorce Certificate, issued by <authority> number <number>, was provided and was authenticated).[MARRITALSTATUSDESC]", true, true);
                                break;

                        }
                        doc.SaveToFile(savePath);
                        switch (diligenceInput.diligenceInputModel.Children)
                        {
                            case "Children":
                                doc.Replace("[MARRITALSTATUSDESC]", "\n\nThey have <number> children together, who reside with their parents at the aforenoted residential address in [Country]: <Name>, born on <date>; and <Name> born on <date>.  As noted above, [Last Name]’s spouse and children are not included in this application.", true, false);
                                break;
                            case "No Children":
                                doc.Replace("[MARRITALSTATUSDESC]", "\n\nAccording to their Application Form, [Last Name] has no children.", true, false);
                                break;
                        }
                        try
                        {
                            if (diligenceInput.diligenceInputModel.FullaliasName.ToString().Equals("") || diligenceInput.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("NA") || diligenceInput.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("N/A"))
                            {
                                doc.Replace("FullAliasName", "", true, true);
                            }
                            else
                            {
                                doc.Replace("FullAliasName", diligenceInput.diligenceInputModel.FullaliasName, true, true);
                            }
                        }
                        catch
                        {

                        }
                        if (diligenceInput.diligenceInputModel.foreignlanguage.ToString().Equals(""))
                        {
                            doc.Replace("[FORSUB]", "", true, false);
                            doc.Replace("([FORSUBJECT])", "", true, false);
                        }
                        else
                        {
                            string strforsub = string.Concat("\n", diligenceInput.diligenceInputModel.foreignlanguage);
                            doc.Replace("[FORSUB]", strforsub, true, false);
                            strforsub = string.Concat(diligenceInput.diligenceInputModel.foreignlanguage, " ");
                            doc.Replace("[FORSUBJECT]", strforsub, true, false);
                        }
                        doc.Replace("First_Name", diligenceInput.diligenceInputModel.FirstName.ToUpper().TrimEnd().ToString(), true, true);
                        doc.Replace("Last_Name", diligenceInput.diligenceInputModel.LastName.ToUpper().ToString().TrimEnd(), true, true);
                        try
                        {
                            if (diligenceInput.diligenceInputModel.MiddleName == "")
                            {
                                doc.Replace(" Middle_Name", "", false, false);
                            }
                            else
                            {
                                doc.Replace("Middle_Name", diligenceInput.diligenceInputModel.MiddleName.ToUpper().ToString().TrimEnd(), true, true);
                            }
                        }
                        catch
                        {
                            doc.Replace(" Middle_Name", "", false, false);
                        }
                        doc.Replace("FirstName", diligenceInput.diligenceInputModel.FirstName.TrimEnd(), true, true);
                        doc.Replace("LastName", diligenceInput.diligenceInputModel.LastName.TrimEnd(), true, true);
                        try
                        {
                            doc.Replace("Position1", diligenceInput.EmployerModel[0].Emp_Position.TrimEnd(), true, true);
                        }
                        catch
                        {
                            doc.Replace("Position1", "<not provided>", true, true);
                        }
                        try
                        {
                            doc.Replace("Employer1", diligenceInput.EmployerModel[0].Emp_Employer.TrimEnd(), true, true);
                        }
                        catch
                        {
                            doc.Replace("Employer1", "<not provided>", true, true);
                        }

                        try
                        {
                            if (diligenceInput.diligenceInputModel.Dob.ToString().Equals(""))
                            {
                                doc.Replace("DateofBirth", "<not provided>", true, true);
                            }
                            else
                            {
                                doc.Replace("DateofBirth", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Month) + " " + Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Day + ", " + Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Year, true, true);
                            }
                        }
                        catch
                        {
                            doc.Replace("DateofBirth", "<not provided>", true, true);
                        }
                        try
                        {
                            if (diligenceInput.diligenceInputModel.Pob.ToString().Equals("") || diligenceInput.diligenceInputModel.Pob.ToString().Equals("NA") || diligenceInput.diligenceInputModel.Pob.ToString().Equals("N/A"))
                            {
                                doc.Replace("PlaceofBirth", "<not provided>", true, true);
                            }
                            else
                            {
                                doc.Replace("PlaceofBirth", diligenceInput.diligenceInputModel.Pob, true, true);
                            }
                        }
                        catch
                        {
                            doc.Replace("PlaceofBirth", "<not provided>", true, true);
                        }
                        doc.Replace("[CaseNumber]", diligenceInput.diligenceInputModel.CaseNumber.TrimEnd(), true, true);
                        try
                        {
                            if (diligenceInput.diligenceInputModel.MiddleName.ToString() == "")
                            {
                                doc.Replace(" MiddleName", "", true, false);
                            }
                            else
                            {
                                doc.Replace("MiddleName", diligenceInput.diligenceInputModel.MiddleName.ToString().TrimEnd(), true, true);
                            }
                        }
                        catch
                        {
                            doc.Replace(" MiddleName", "", true, false);
                        }
                        doc.Replace("Nationality", diligenceInput.diligenceInputModel.Nationality.ToString().TrimEnd(), true, true);
                        doc.SaveToFile(savePath);
                    }
                }
                string bnres = "";

                if (diligenceInputs.familyModels[famcount].adult_minor == "Adult")
                {
                    doc.Replace("First_Name", diligenceInput.diligenceInputModel.FirstName.ToUpper(), true, true);
                    try
                    {
                        doc.Replace("Middle_Name", diligenceInput.diligenceInputModel.MiddleName.ToUpper(), true, true);
                    }
                    catch
                    {
                        doc.Replace(" Middle_Name", "", false, false);
                    }
                    doc.Replace("Last_Name", diligenceInput.diligenceInputModel.LastName.ToUpper(), true, true);
                    doc.SaveToFile(savePath);

                    if (famcount == 0 || famcount < diligenceInputs.familyModels.Count - minor_count )
                    {
                        if (famcount == diligenceInputs.familyModels.Count - minor_count - 1)
                        {

                        }
                        else
                        {
                            doc.Replace("COMPRESULT", "COMPRESULT\n\nCOMP2RESULT", true, true);
                            doc.Replace("COMPCOMMENT", "COMPCOMMENT\n\nCOMP2COMMENT", true, true);
                            doc.Replace("PROFRESULT", "PROFRESULT\n\nPROF2RESULT", true, true);
                            doc.Replace("PROFCOMMENT", "PROFCOMMENT\n\nPROF2COMMENT", true, true);                            
                            doc.Replace("FININDRESULT", "FININDRESULT\n\nFIN2INDRESULT", true, true);
                            doc.Replace("FININDCOMMENT", "FININDCOMMENT\n\nFIN2INDCOMMENT", true, true);
                            doc.Replace("KONSECRESULT", "KONSECRESULT\n\nKON2SECRESULT", true, true);
                            doc.Replace("KONSECCOMMENT", "KONSECCOMMENT\n\nKONSEC2COMMENT", true, true);
                            doc.Replace("SECRESULT", "SECRESULT\n\nSEC2RESULT", true, true);
                            doc.Replace("SECCOMMENT", "SECCOMMENT\n\nSEC2COMMENT", true, true);
                            doc.Replace("UKFICORESULT", "UKFICORESULT\n\nUKFI2CORESULT", true, true);
                            doc.Replace("UKFICOCOMMENT", "UKFICOCOMMENT\n\nUKFI2COCOMMENT", true, true);
                            doc.Replace("CRIMINALCOMMENT", "CRIMINALCOMMENT\n\nCRIMINAL2COMMENT", true, true);
                            doc.Replace("CRIMINALRESULT", "CRIMINALRESULT\n\nCRIMINAL2RESULT", true, true);
                            doc.Replace("USNFRESULT", "USNFRESULT\n\nUSNF2RESULT", true, true);
                            doc.Replace("USNFCOMMENT", "USNFCOMMENT\n\nUSNF2COMMENT", true, true);
                            doc.Replace("INCOJORESULT", "INCOJORESULT\n\nINCOJORESULT", true, true);
                            doc.Replace("INCOJOCOMMENT", "INCOJOCOMMENT\n\nINCOJO2COMMENT", true, true);
                            doc.Replace("POLITICRESULT", "POLITICRESULT\n\nPOLITIC2RESULT", true, true);
                            doc.Replace("POLITICCOMMENT", "POLITICCOMMENT\n\nPOLITIC2COMMENT", true, true);
                            if (famcount == 0) {
                                doc.Replace("[CountryHeader][COURTDESC]", "First_Name0 Middle_Name0 Last_Name0\n\n[CountryHeader][COURTDESC]\nFirst_Name Middle_Name Last_Name\n\n[Country2Header][COURT2DESC]", true, false);
                                doc.Replace("GLOBALSECURITYHITSDESCRIPTION", "First_Name0 Middle_Name0 Last_Name0\n\nGLOBALSECURITYHITSDESCRIPTION", true, false);
                            } else
                            {
                                doc.Replace("[CountryHeader][COURTDESC]", "[CountryHeader][COURTDESC]\nFirst_Name Middle_Name Last_Name\n\n[Country2Header][COURT2DESC]", true, false);
                            }                           
                           
                            doc.Replace("ICIJHITSDESCRIPTION", "ICIJHITSDESCRIPTION\nFirst_Name Middle_Name Last_Name\n\nGLOBAL2SECURITYHITSDESCRIPTION\nGLOBAL2SECFAMILYHITSDESCRIPTION\nPEP2HITSDESCRIPTION\nICIJ2HITSDESCRIPTION", true, false);                            
                            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "PRESSANDMEDIASEARCHDESCRIPTION\n\nFirst_Name Middle_Name Last_Name\n\nPRESS2ANDMEDIASEARCHDESCRIPTION", true, false);
                            doc.Replace("CHARACTERREFERENCESDESC", "First_Name0 Middle_Name0 Last_Name0\n\nCHARACTERREFERENCESDESC\n\nFirst_Name Middle_Name Last_Name\n\nCHARACTER2REFERENCESDESC", true, true);
                            doc.Replace("SOURCEOFWEALTHDESC", "First_Name0 Middle_Name0 Last_Name0\n\nSOURCEOFWEALTHDESC\n\nFirst_Name Middle_Name Last_Name\n\nSOURCE2OFWEALTHDESC", true, true);
                            doc.Replace("DISCREETREPUTATIONALINQUIRIESDESC", "First_Name0 Middle_Name0 Last_Name0\n\nDISCREETREPUTATIONALINQUIRIESDESC\n\nFirst_Name Middle_Name Last_Name\n\nDISCREET2REPUTATIONALINQUIRIESDESC", true, true);
                        }
                    }
                    else
                    {
                        
                    }                        
                    doc.SaveToFile(savePath);
                    TextSelection[] text1 = doc.FindAllString("[CountryHeader]", false, false);
                    if (text1 != null)
                    {
                        foreach (TextSelection seletion in text1)
                        {
                            seletion.GetAsOneRange().CharacterFormat.Italic = true;
                            seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                        }
                    }
                    doc.SaveToFile(savePath);
                    //Employee details 
                    string emp_Comment = "";
                    string strempstartdate = "<not provided>";
                    string strempenddate = "<not provided>";
                    for (int i = 0; i < diligenceInput.EmployerModel.Count; i++)
                    {
                        strempstartdate = "<not provided>";
                        strempenddate = "<not provided>";
                        if (!diligenceInput.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && !diligenceInput.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strempstartdate = diligenceInput.EmployerModel[i].Emp_StartDateMonth + " " + diligenceInput.EmployerModel[i].Emp_StartDateDay + ", " + diligenceInput.EmployerModel[i].Emp_StartDateYear;
                            // employerModel.Emp_StartDate = strempstartdate;
                        }
                        if (diligenceInput.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && diligenceInput.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strempstartdate = "<not provided>";
                        }
                        if (diligenceInput.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && !diligenceInput.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strempstartdate = diligenceInput.EmployerModel[i].Emp_StartDateMonth + " " + diligenceInput.EmployerModel[i].Emp_StartDateYear;
                            //employerModel.Emp_StartDate = strempstartdate;
                        }
                        if (diligenceInput.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && diligenceInput.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strempstartdate = diligenceInput.EmployerModel[i].Emp_StartDateYear;
                            //employerModel.Emp_StartDate = strempstartdate;
                        }
                        if (i == 0 && diligenceInput.EmployerModel[0].Emp_Status.ToString().Equals("Current")) { }
                        else
                        {
                            if (!diligenceInput.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && !diligenceInput.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                            {
                                strempenddate = diligenceInput.EmployerModel[i].Emp_EndDateMonth + " " + diligenceInput.EmployerModel[i].Emp_EndDateDay + ", " + diligenceInput.EmployerModel[i].Emp_EndDateYear;
                                //  employerModel.Emp_EndDate = strempenddate;
                            }
                            if (diligenceInput.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && diligenceInput.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && diligenceInput.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                            {
                                strempenddate = "<not provided>";
                            }
                            if (diligenceInput.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && !diligenceInput.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                            {
                                strempenddate = diligenceInput.EmployerModel[i].Emp_EndDateMonth + " " + diligenceInput.EmployerModel[i].Emp_EndDateYear;
                                //employerModel.Emp_EndDate = strempenddate;
                            }
                            if (diligenceInput.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && diligenceInput.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                            {
                                strempenddate = diligenceInput.EmployerModel[i].Emp_EndDateYear;
                                //employerModel.Emp_EndDate = strempenddate;
                            }
                        }
                        string stremp1textappended = "";
                        if (i == 0)
                        {
                            if (diligenceInput.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                            {
                                emp_Comment = "[LastName] is a [Position1] at [Employer1Footnote]\n";
                                emp_Comment = emp_Comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString().TrimEnd());
                                emp_Comment = emp_Comment.Replace("[Position1]", diligenceInput.EmployerModel[0].Emp_Position.ToString().TrimEnd());
                                try
                                {
                                    emp_Comment = emp_Comment.Replace("[Employer1]", diligenceInput.EmployerModel[0].Emp_Employer.ToString().TrimEnd());
                                }
                                catch
                                {
                                    emp_Comment = emp_Comment.Replace("[Employer1]", "<not provided>");
                                }
                                stremp1textappended = " in [EmpLocation1], since [EmpStartDate1].  <Investigator to insert company website, brief description, registration details, and any other applicant-provided data here>. EMPLOYEEDESCRIPTION";
                            }
                            else
                            {
                                if (diligenceInput.EmployerModel[i].Emp_Confirmed.Equals("No"))
                                {
                                    emp_Comment = "According to self-reported biographical information, [LastName] is a [Position1] at [Employer1Footnote]\n";
                                    emp_Comment = emp_Comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString().TrimEnd());
                                    emp_Comment = emp_Comment.Replace("[Position1]", diligenceInput.EmployerModel[0].Emp_Position.ToString().TrimEnd());
                                    try
                                    {
                                        emp_Comment = emp_Comment.Replace("[Employer1]", diligenceInput.EmployerModel[0].Emp_Employer.ToString().TrimEnd());
                                    }
                                    catch
                                    {
                                        emp_Comment = emp_Comment.Replace("[Employer1]", "<not provided>");
                                    }
                                    stremp1textappended = " in [EmpLocation1], since [EmpStartDate1].  <Investigator to insert company website, brief description, registration details, and any other applicant-provided data here>. EMPLOYEEDESCRIPTION";
                                }
                                else
                                {
                                    emp_Comment = "According to self-reported biographical information, [LastName] is a [Position1] at [Employer1Footnote]\n";
                                    emp_Comment = emp_Comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString().TrimEnd());
                                    emp_Comment = emp_Comment.Replace("[Position1]", diligenceInput.EmployerModel[0].Emp_Position.ToString().TrimEnd());
                                    try
                                    {
                                        emp_Comment = emp_Comment.Replace("[Employer1]", diligenceInput.EmployerModel[0].Emp_Employer.ToString().TrimEnd());
                                    }
                                    catch
                                    {
                                        emp_Comment = emp_Comment.Replace("[Employer1]", "<not provided>");
                                    }
                                    stremp1textappended = " in [EmpLocation1], since [EmpStartDate1].  <Investigator to insert company website, brief description, registration details, and any other applicant-provided data here>.  Efforts to independently confirm the subject’s tenure at the same have remained unsuccessful to date. EMPLOYEEDESCRIPTION";
                                }
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
                                                    try
                                                    {
                                                        if (diligenceInput.EmployerModel[0].Emp_Employer.ToString() == "")
                                                        {
                                                            doc.Replace("[Employer1Footnote]", "<not provided>", true, false);
                                                        }
                                                        else
                                                        {
                                                            doc.Replace("[Employer1Footnote]", diligenceInput.EmployerModel[0].Emp_Employer.ToString(), true, true);
                                                        }
                                                    }
                                                    catch {
                                                        doc.Replace("[Employer1Footnote]", "<not provided>", true, false);
                                                    }
                                                    stremp1textappended = stremp1textappended.Replace("[EmpLocation1]", diligenceInput.EmployerModel[0].Emp_Location.ToString());
                                                    stremp1textappended = stremp1textappended.Replace("[EmpStartDate1]", strempstartdate);
                                                    TextRange tr = para.AppendText(stremp1textappended);
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
                            catch { }
                            doc.SaveToFile(savePath);
                            emp_Comment = "";
                        }
                        else
                        {
                            if (i == 1) { emp_Comment = "\n"; }
                            CommentModel comment = _context.DbComment
                                  .Where(u => u.Comment_type == "Emp2")
                                  .FirstOrDefault();
                            if (diligenceInput.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                            {
                                emp_Comment = string.Concat(emp_Comment, "\n", comment.confirmed_comment.ToString());
                            }
                            else
                            {
                                if (diligenceInput.EmployerModel[i].Emp_Confirmed.Equals("No"))
                                {
                                    emp_Comment = string.Concat(emp_Comment, "\n", comment.unconfirmed_comment.ToString());
                                }
                                else
                                {
                                    emp_Comment = string.Concat(emp_Comment, "\n", comment.other_comment.ToString(), " APPENDEMPTEXT", i.ToString());
                                }
                            }
                            emp_Comment = emp_Comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
                            emp_Comment = emp_Comment.Replace("[Position]", diligenceInput.EmployerModel[i].Emp_Position.ToString());
                            try
                            {
                                emp_Comment = emp_Comment.Replace("[Employer]", diligenceInput.EmployerModel[i].Emp_Employer.ToString());
                            }
                            catch
                            {
                                emp_Comment = emp_Comment.Replace("[Employer]", "<not provided>");
                            }
                            emp_Comment = emp_Comment.Replace("[EmpLocation]", diligenceInput.EmployerModel[i].Emp_Location.ToString());
                            emp_Comment = emp_Comment.Replace("[EmpStartDate]", strempstartdate);
                            emp_Comment = emp_Comment.Replace("[EmpEndDate]", strempenddate);
                            if (i != diligenceInput.EmployerModel.Count - 1)
                            {
                                emp_Comment = string.Concat(emp_Comment, "\n");
                            }
                        }
                    }
                    doc.Replace("EMPLOYEEDESCRIPTION", emp_Comment, true, true);
                    doc.Replace("LastName", emp_Comment, true, false);
                    doc.SaveToFile(savePath);
                    string bnres24 = "";
                    for (int i = 1; i < diligenceInput.EmployerModel.Count; i++)
                    {
                        bnres24 = "";
                        if (diligenceInput.EmployerModel[i].Emp_Confirmed.ToString().Equals("Unreturned"))
                        {
                            try
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
                                                    textRange = para.AppendText("Efforts to independently confirm the subject’s tenure at the same are currently ongoing, the results of which will be provided under separate cover, if and when received.");
                                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                    textRange.CharacterFormat.FontSize = 11;
                                                    textRange.CharacterFormat.Italic = true;
                                                    doc.SaveToFile(savePath);
                                                    doc.Replace(string.Concat("APPENDEMPTEXT", i.ToString()), "", true, false);
                                                    doc.SaveToFile(savePath);
                                                    bnres24 = "true";
                                                    break;
                                                }
                                            }
                                        }
                                        if (bnres24.Equals("true")) { break; }
                                    }
                                    if (bnres24.Equals("true")) { break; }
                                }
                            }
                            catch { }
                        }
                    }

                    //Business Affiliation
                    string strbusumcomm = "";
                    string strbusinesscomm = "";
                    if (diligenceInput.otherdetails.undisclosedBA.ToString().Equals("Yes"))
                    {
                        doc.Replace("COMPRESULT", string.Concat(diligenceInput.diligenceInputModel.FirstName," ",diligenceInput.diligenceInputModel.LastName, " - Discrepancy Identified"), true, true);
                        strbusinesscomm = "[LastName] is reportedly a [Position1] of [Employer1] in [Country], since [EmpStartDate1]\n\nIn addition, [LastName] is identified in connection with additional entities in <investigator to add countries>";
                        strbusinesscomm = strbusinesscomm.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName);
                    }
                    string business_comment;
                    CommentModel busi_comment2 = _context.DbComment
                                      .Where(u => u.Comment_type == "Business_Affiliation")
                                      .FirstOrDefault();
                    if (diligenceInput.otherdetails.Has_Business_Affiliations.ToString().Equals("Yes"))
                    {
                        doc.Replace("COMPRESULT", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName, " - Records"), true, true);
                        strbusumcomm = "[LastName] is a [Position1] of [Employer1] in [Employer1Location], since [EmpStartDate1]\n\nIn addition, [LastName] is identified in connection with additional entities in <investigator to add countries>";
                        business_comment = busi_comment2.confirmed_comment.ToString();
                        business_comment = string.Concat(business_comment, "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>");
                    }
                    else
                    {
                        business_comment = busi_comment2.unconfirmed_comment.ToString();
                        doc.Replace("COMPRESULT", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName, " - Clear"), true, true);
                        strbusumcomm = "[LastName] is a [Position1] of [Employer1] in [Employer1Location], since [EmpStartDate1]";
                    }
                    doc.Replace("[BUSINESSAFFILIATIONSIDENTIFIED]", business_comment, true, true);
                    try
                    {
                        if (diligenceInput.EmployerModel[0].Emp_Position.Equals(""))
                        {
                            strbusumcomm = strbusumcomm.Replace("[Position1]", "<not provided>");
                            strbusinesscomm = strbusinesscomm.Replace("[Position1]", "<not provided>");
                        }
                        else
                        {
                            strbusumcomm = strbusumcomm.Replace("[Position1]", diligenceInput.EmployerModel[0].Emp_Position);
                            strbusinesscomm = strbusinesscomm.Replace("[Position1]", diligenceInput.EmployerModel[0].Emp_Position);
                        }
                    }
                    catch
                    {
                        strbusumcomm = strbusumcomm.Replace("[Position1]", "<not provided>");
                        strbusinesscomm = strbusinesscomm.Replace("[Position1]", "<not provided>");
                    }
                    try
                    {
                        if (diligenceInput.EmployerModel[0].Emp_Employer.Equals(""))
                        {
                            strbusumcomm = strbusumcomm.Replace("[Employer1]", "<not provided>");
                            strbusinesscomm = strbusinesscomm.Replace("[Employer1]", "<not provided>");
                        }
                        else
                        {
                            strbusumcomm = strbusumcomm.Replace("[Employer1]", diligenceInput.EmployerModel[0].Emp_Employer);
                            strbusinesscomm = strbusinesscomm.Replace("[Employer1]", diligenceInput.EmployerModel[0].Emp_Employer);
                        }
                    }
                    catch
                    {
                        strbusumcomm = strbusumcomm.Replace("[Employer1]", "<not provided>");
                        strbusinesscomm = strbusinesscomm.Replace("[Employer1]", "<not provided>");
                    }
                    try
                    {
                        if (diligenceInput.EmployerModel[0].Emp_Location.Equals(""))
                        {
                            strbusumcomm = strbusumcomm.Replace("[Employer1Location]", "<not provided>");
                            strbusinesscomm = strbusinesscomm.Replace("[Employer1Location]", "<not provided>");
                        }
                        else
                        {
                            strbusumcomm = strbusumcomm.Replace("[Employer1Location]", diligenceInput.EmployerModel[0].Emp_Location);
                            strbusinesscomm = strbusinesscomm.Replace("[Employer1Location]", diligenceInput.EmployerModel[0].Emp_Location);
                        }
                    }
                    catch
                    {
                        strbusumcomm = strbusumcomm.Replace("[Employer1Location]", "<not provided>");
                        strbusinesscomm = strbusinesscomm.Replace("[Employer1Location]", "<not provided>");
                    }
                    try
                    {
                        if (diligenceInput.EmployerModel[0].Emp_StartDateYear.ToString().Equals("Year"))
                        {
                            strbusumcomm = strbusumcomm.Replace("[EmpStartDate1]", "<not provided>");
                            strbusinesscomm = strbusinesscomm.Replace("[EmpStartDate1]", "<not provided>");
                        }
                        else
                        {
                            strbusumcomm = strbusumcomm.Replace("[EmpStartDate1]", diligenceInput.EmployerModel[0].Emp_StartDateYear);
                            strbusinesscomm = strbusinesscomm.Replace("[EmpStartDate1]", diligenceInput.EmployerModel[0].Emp_StartDateYear);
                        }
                    }
                    catch
                    {
                        strbusumcomm = strbusumcomm.Replace("[EmpStartDate1]", "<not provided>");
                        strbusinesscomm = strbusinesscomm.Replace("[EmpStartDate1]", "<not provided>");
                    }
                    if (strbusinesscomm.ToString().Equals(""))
                    {
                        doc.Replace("COMPCOMMENT", strbusumcomm, true, true);
                    }
                    else
                    {
                        doc.Replace("COMPCOMMENT", strbusinesscomm, true, true);
                    }
                    doc.SaveToFile(savePath);
                    //Intellectual Hits
                    string intellectual_comment;
                    CommentModel intellec_comment2 = _context.DbComment
                                       .Where(u => u.Comment_type == "Intellectual_hits")
                                       .FirstOrDefault();
                    if (diligenceInput.otherdetails.Has_Intellectual_Hits.ToString().Equals("Yes"))
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
                    switch (diligenceInput.otherdetails.worldcheck_discloseBA)
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
                    if (diligenceInput.otherdetails.worldcheck_discloseBA.ToString().Equals("Yes") || diligenceInput.otherdetails.worldcheck_discloseBA.ToString().Equals("No"))
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
                    switch (diligenceInput.otherdetails.undisclosedBA)
                    {
                        case "Yes":
                            doc.Replace("[UNDISCLOSEDBADESC]", "While not reported by the subject, research of records maintained by the [Corp Registry], as well as other sources, revealed the subject in connection with the following additional business entities:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>", true, false);
                            break;
                        case "No":
                            doc.Replace("[UNDISCLOSEDBADESC]", "In addition to the above, research of records maintained by the [Corp Registry], as well as other sources, did not identify the subject in connection with any unreported business entities.", true, false);
                            break;
                    }
                    //worldcheck_undiscloseBA

                    if (famcount == diligenceInputs.familyModels.Count - minor_count-1 || famcount == diligenceInputs.familyModels.Count - 1)
                    {
                        switch (diligenceInput.otherdetails.worldcheck_undiscloseBA)
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
                        switch (diligenceInput.otherdetails.worldcheck_undiscloseBA)
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
                    if (diligenceInput.otherdetails.worldcheck_undiscloseBA.ToString().Equals("Yes") || diligenceInput.otherdetails.worldcheck_undiscloseBA.ToString().Equals("No"))
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
                                            if (abc.ToString().Contains("[UNDISCLOSEFONTCHANGE]"))
                                            {
                                                textRange.CharacterFormat.Italic = true;
                                                doc.SaveToFile(savePath);
                                                doc.Replace("[UNDISCLOSEFONTCHANGE]", "", true, false);
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
                    if (famcount == 0)
                    {
                        if (diligenceInput.educationModels[0].Edu_History == "No") { }
                        else
                        {
                            doc.Replace("DEGREEDESCRIPTION", "First_Name0 Middle_Name0 Last_Name0\n\nDEGREEDESCRIPTION", true, true);
                        }
                    }

                    string edu_comment = "";
                    string edu_header = "";
                    string edu_summcomment = "";

                    //Education details
                    for (int i = 0; i < diligenceInput.educationModels.Count; i++)
                    {
                        EducationModel educationModel = new EducationModel();
                        //educationModel.record_Id = record_Id;
                        educationModel.Edu_History = diligenceInput.educationModels[i].Edu_History;
                        educationModel.Edu_Location = diligenceInput.educationModels[i].Edu_Location;
                        educationModel.Edu_Degree = diligenceInput.educationModels[i].Edu_Degree;
                        educationModel.Edu_School = diligenceInput.educationModels[i].Edu_School;
                        educationModel.CreatedBy = Environment.UserName;
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
                        if (diligenceInput.educationModels[i].Edu_History.ToString().Equals("Yes"))
                        {
                            if (!diligenceInput.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && !diligenceInput.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                            {
                                edustartdate = diligenceInput.educationModels[i].Edu_StartDateMonth + " " + diligenceInput.educationModels[i].Edu_StartDateDay + ", " + diligenceInput.educationModels[i].Edu_StartDateYear;
                                edustartyr = diligenceInput.educationModels[i].Edu_StartDateYear;
                                //educationModel.Edu_StartDate = edustartdate;

                            }
                            if (diligenceInput.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && diligenceInput.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && diligenceInput.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                            {
                                edustartdate = "<not provided>";
                                edustartyr = "<not provided>";
                            }
                            if (diligenceInput.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && !diligenceInput.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                            {
                                edustartdate = diligenceInput.educationModels[i].Edu_StartDateMonth + " " + diligenceInput.educationModels[i].Edu_StartDateYear;
                                edustartyr = diligenceInput.educationModels[i].Edu_StartDateYear;
                                //educationModel.Edu_StartDate = edustartdate;
                            }
                            if (diligenceInput.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && diligenceInput.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                            {
                                edustartdate = diligenceInput.educationModels[i].Edu_StartDateYear;
                                edustartyr = diligenceInput.educationModels[i].Edu_StartDateYear;
                                //educationModel.Edu_StartDate = edustartdate;
                            }

                            if (!diligenceInput.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && !diligenceInput.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                            {
                                eduenddate = diligenceInput.educationModels[i].Edu_EndDateMonth + " " + diligenceInput.educationModels[i].Edu_EndDateDay + ", " + diligenceInput.educationModels[i].Edu_EndDateYear;
                                eduendyr = diligenceInput.educationModels[i].Edu_EndDateYear;
                                //educationModel.Edu_EndDate = eduenddate;
                            }
                            if (diligenceInput.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && diligenceInput.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && diligenceInput.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                            {
                                eduenddate = "<not provided>";
                                eduendyr = "<not provided>";
                            }
                            if (diligenceInput.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && !diligenceInput.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                            {
                                eduenddate = diligenceInput.educationModels[i].Edu_EndDateMonth + " " + diligenceInput.educationModels[i].Edu_EndDateYear;
                                eduendyr = diligenceInput.educationModels[i].Edu_EndDateYear;
                                //educationModel.Edu_EndDate = eduenddate;
                            }
                            if (diligenceInput.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && diligenceInput.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                            {
                                eduenddate = diligenceInput.educationModels[i].Edu_EndDateYear;
                                eduendyr = diligenceInput.educationModels[i].Edu_EndDateYear;
                                //educationModel.Edu_EndDate = eduenddate;
                            }

                            if (!diligenceInput.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && !diligenceInput.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                            {
                                edugraddate = diligenceInput.educationModels[i].Edu_GradDateMonth + " " + diligenceInput.educationModels[i].Edu_GradDateDay + ", " + diligenceInput.educationModels[i].Edu_GradDateYear;
                                edugradyr = diligenceInput.educationModels[i].Edu_GradDateYear;
                                // educationModel.Edu_Graddate = edugraddate;
                            }
                            if (diligenceInput.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && diligenceInput.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && diligenceInput.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                            {
                                edugraddate = "<not provided>";
                                edugradyr = "<not provided>";
                            }
                            if (diligenceInput.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && !diligenceInput.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                            {
                                edugraddate = diligenceInput.educationModels[i].Edu_GradDateMonth + " " + diligenceInput.educationModels[i].Edu_GradDateYear;
                                edugradyr = diligenceInput.educationModels[i].Edu_GradDateYear;
                                //educationModel.Edu_Graddate = edugraddate;
                            }
                            if (diligenceInput.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && diligenceInput.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                            {
                                edugraddate = diligenceInput.educationModels[i].Edu_GradDateYear;
                                edugradyr = diligenceInput.educationModels[i].Edu_GradDateYear;
                                //educationModel.Edu_Graddate = edugraddate;
                            }

                            //_context.DbEducation.Add(educationModel);
                            //_context.SaveChanges();
                            if (i == diligenceInput.educationModels.Count - 1)
                            {
                                switch (diligenceInput.educationModels[i].Edu_Confirmed.ToString())
                                {
                                    case "Yes":
                                        edu_comment = string.Concat(edu_comment, comment1.confirmed_comment.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName," - " ,diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Confirmed".TrimEnd());
                                        edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.confirmed_comment.ToString(), "\n\n");
                                        break;
                                    case "No":
                                        edu_comment = string.Concat(edu_comment, comment1.unconfirmed_comment.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, " - ",  diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Unconfirmed".TrimEnd());
                                        edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.unconfirmed_comment.ToString(), "\n\n");
                                        break;
                                    case "Result Pending":
                                        edu_comment = string.Concat(edu_comment, edurescomment.confirmed_comment.ToString(), "APPENDEDURESULTPEND", i.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, " - " , diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Results Pending".TrimEnd());
                                        edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.confirmed_comment.ToString(), " APPENDRESPENDING", "", "\n\n");
                                        break;
                                    case "Attendance Confirmed":
                                        edu_comment = string.Concat(edu_comment, eduattcomment.confirmed_comment.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, " - " , diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Confirmed".TrimEnd());
                                        edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.confirmed_comment.ToString(), "\n\n");
                                        break;
                                    case "Attendance Unconfirmed":
                                        edu_comment = string.Concat(edu_comment, eduattcomment.unconfirmed_comment.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, " - ", diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Unconfirmed".TrimEnd());
                                        edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.unconfirmed_comment.ToString(), "\n\n");
                                        break;
                                    case "Attendance Result pending":
                                        edu_comment = string.Concat(edu_comment, edurescomment.unconfirmed_comment.ToString(), "APPENDEDUATTRESULTPEND", i.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, " - ", diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Results Pending".TrimEnd());
                                        edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.unconfirmed_comment.ToString(), " ATTRESPENDING", "\n\n");
                                        break;
                                }
                            }
                            else
                            {
                                switch (diligenceInput.educationModels[i].Edu_Confirmed.ToString())
                                {
                                    case "Yes":
                                        edu_comment = string.Concat(edu_comment, comment1.confirmed_comment.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName,"-",diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Confirmed", "\n\n");
                                        edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.confirmed_comment.ToString(), "\n\n");
                                        break;
                                    case "No":
                                        edu_comment = string.Concat(edu_comment, comment1.unconfirmed_comment.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, "-", diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Unconfirmed", "\n\n");
                                        edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.unconfirmed_comment.ToString(), "\n\n");
                                        break;
                                    case "Result Pending":
                                        edu_comment = string.Concat(edu_comment, edurescomment.confirmed_comment.ToString(), "APPENDEDURESULTPEND", i.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, "-", diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Results Pending", "\n\n");
                                        edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.confirmed_comment.ToString(), " APPENDRESPENDING", "", "\n\n");
                                        break;
                                    case "Attendance Confirmed":
                                        edu_comment = string.Concat(edu_comment, eduattcomment.confirmed_comment.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, "-", diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Confirmed", "\n\n");
                                        edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.confirmed_comment.ToString(), "\n\n");
                                        break;
                                    case "Attendance Unconfirmed":
                                        edu_comment = string.Concat(edu_comment, eduattcomment.unconfirmed_comment.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, "-", diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Unconfirmed", "\n\n");
                                        edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.unconfirmed_comment.ToString(), "\n\n");
                                        break;
                                    case "Attendance Result pending":
                                        edu_comment = string.Concat(edu_comment, edurescomment.unconfirmed_comment.ToString(), "APPENDEDUATTRESULTPEND", i.ToString(), "\n\n");
                                        edu_header = string.Concat(edu_header, diligenceInput.diligenceInputModel.FirstName, "-", diligenceInput.educationModels[i].Edu_Degree.ToString(), " - ", "Results Pending", "\n\n");
                                        edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.unconfirmed_comment.ToString(), " ATTRESPENDING", "\n\n");
                                        break;
                                }
                            }
                            try
                            {
                                edu_comment = edu_comment.Replace("[Degree]", diligenceInput.educationModels[i].Edu_Degree.ToString());
                            }
                            catch
                            {
                                edu_comment = edu_comment.Replace("[Degree]", "not provided");
                            }
                            try
                            {
                                edu_comment = edu_comment.Replace("[School]", diligenceInput.educationModels[i].Edu_School.ToString());
                            }
                            catch
                            {
                                edu_comment = edu_comment.Replace("[School]", "not provided");
                            }
                            try
                            {
                                edu_comment = edu_comment.Replace("[EduLocation]", diligenceInput.educationModels[i].Edu_Location.ToString());
                            }
                            catch
                            {
                                edu_comment = edu_comment.Replace("[EduLocation]", "not provided");
                            }
                            edu_comment = edu_comment.Replace("[GradDate]", edugraddate);
                            edu_comment = edu_comment.Replace("[EduStartDate]", edustartdate);
                            edu_comment = edu_comment.Replace("[EduEndDate]", eduenddate);
                            try
                            {
                                edu_summcomment = edu_summcomment.Replace("[Degree]", diligenceInput.educationModels[i].Edu_Degree.ToString());
                            }
                            catch
                            {
                                edu_summcomment = edu_summcomment.Replace("[Degree]", "not provided");
                            }
                            try
                            {
                                edu_summcomment = edu_summcomment.Replace("[School]", diligenceInput.educationModels[i].Edu_School.ToString());
                            }
                            catch
                            {
                                edu_summcomment = edu_summcomment.Replace("[School]", "not provided");
                            }
                            try
                            {
                                edu_summcomment = edu_summcomment.Replace("[EduLocation]", diligenceInput.educationModels[i].Edu_Location.ToString());
                            }
                            catch
                            {
                                edu_summcomment = edu_summcomment.Replace("[EduLocation]", "not provided");
                            }
                            edu_summcomment = edu_summcomment.Replace("[GradDate]", edugradyr);
                            edu_summcomment = edu_summcomment.Replace("[EDUFROMDATE]", edustartyr);
                            edu_summcomment = edu_summcomment.Replace("[EDUTODATE]", eduendyr);
                        }
                        else
                        {
                            edu_comment = "";
                            edu_summcomment = comment1.other_comment.ToString();
                            edu_summcomment = edu_summcomment.Replace("[Last Name]", diligenceInput.diligenceInputModel.LastName.ToString());
                            edu_header = "N/A";
                        }
                    }
                    edu_comment = edu_comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
                    edu_summcomment = edu_summcomment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
                    if (famcount == (diligenceInputs.familyModels.Count - minor_count - 1)|| famcount == diligenceInputs.familyModels.Count - 1)
                    {
                        doc.Replace("SUMMDEGDESCRIPTION", edu_summcomment.TrimEnd(), true, true);
                        doc.Replace("DEGREEDESCRIPTION", edu_comment, true, true);
                        doc.Replace("DEGREEHEADER", edu_header, true, true);
                    }
                    else
                    {
                     
                        doc.Replace("SUMMDEGDESCRIPTION", string.Concat(edu_summcomment.TrimEnd(), "\n\nSUMMDEGDESCRIPTION"), true, true);
                        doc.Replace("DEGREEDESCRIPTION", string.Concat(edu_comment, "First_Name Middle_Name Last_Name\n\nDEGREEDESCRIPTION"), true, true);
                        doc.Replace("DEGREEHEADER", string.Concat(edu_header, "\n\nDEGREEHEADER"), true, true);
                    }
                    doc.SaveToFile(savePath);
                    doc.Replace(" 	", " ", false, false);
                    doc.SaveToFile(savePath);
                    int Edurowcount = 0;
                    Table table = doc.Sections[1].Tables[0] as Table;
                    if (Client_Name.ToString().ToUpper().Equals("MALTA"))
                    {
                        Edurowcount = 6;
                    }
                    else
                    {
                        Edurowcount = 4;
                    }
                    TableCell cell1 = table.Rows[Edurowcount].Cells[2];
                    TableCell cell2 = table.Rows[Edurowcount].Cells[1];
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
                                    textRange = p1.AppendText("DONEhowever, efforts to independently verify the same are currently ongoing (the results of which will be provided under separate cover -- if and when received)");
                                    textRange.CharacterFormat.FontName = "Calibri (Body)";
                                    textRange.CharacterFormat.FontSize = 11;
                                    textRange.CharacterFormat.Italic = true;
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
                                        textRange = p1.AppendText("DONEefforts to independently verify the same are currently ongoing (the results of which will be provided under separate cover -- if and when received)");
                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                        textRange.CharacterFormat.FontSize = 11;
                                        textRange.CharacterFormat.Italic = true;
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
                    if (diligenceInput.educationModels[0].Edu_History.ToString().Equals("Yes"))
                    {
                        for (int i = 0; i < diligenceInput.educationModels.Count; i++)
                        {
                            bnres = "";
                            if (diligenceInput.educationModels[i].Edu_Confirmed.ToString().Equals("Result Pending") || diligenceInput.educationModels[i].Edu_Confirmed.ToString().Equals("Attendance Result pending"))
                            {
                                try
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
                                catch { }
                            }
                        }
                    }
                    doc.SaveToFile(savePath);
                    if (famcount == 0)
                    {
                        doc.Replace("PLLICENSEDESCRIPTION[PLLICENSEDESC]", "\nFirst_Name0 Middle_Name0 Last_Name0\nPLLICENSEDESCRIPTION[PLLICENSEDESC]", true, true);
                    }
                    //PL License details
                    string pl_comment = "";
                    for (int i = 0; i < diligenceInput.pllicenseModels.Count; i++)
                    {
                        PLLicenseModel pllicenseModel = new PLLicenseModel();
                        //pllicenseModel.record_Id = record_Id;
                        pllicenseModel.General_PL_License = diligenceInput.pllicenseModels[i].General_PL_License;
                        pllicenseModel.PL_License_Type = diligenceInput.pllicenseModels[i].PL_License_Type;
                        pllicenseModel.PL_Location = diligenceInput.pllicenseModels[i].PL_Location;
                        pllicenseModel.PL_Number = diligenceInput.pllicenseModels[i].PL_Number;
                        pllicenseModel.PL_Organization = diligenceInput.pllicenseModels[i].PL_Organization;
                        pllicenseModel.PL_Confirmed = diligenceInput.pllicenseModels[i].PL_Confirmed;
                        pllicenseModel.CreatedBy = Environment.UserName;

                        string plstartdate = "<not provided>";
                        string plenddate = "<not provided>";
                        if (diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                        {
                            if (!diligenceInput.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && !diligenceInput.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !diligenceInput.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                            {
                                plstartdate = diligenceInput.pllicenseModels[i].PL_StartDateMonth + " " + diligenceInput.pllicenseModels[i].PL_StartDateDay + ", " + diligenceInput.pllicenseModels[i].PL_StartDateYear;
                                //pllicenseModel.PL_Start_Date = plstartdate;

                            }
                            if (diligenceInput.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && diligenceInput.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && diligenceInput.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                            {
                                plstartdate = "<not provided>";
                            }
                            if (diligenceInput.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && !diligenceInput.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !diligenceInput.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                            {
                                plstartdate = diligenceInput.pllicenseModels[i].PL_StartDateMonth + " " + diligenceInput.pllicenseModels[i].PL_StartDateYear;
                                //pllicenseModel.PL_Start_Date = plstartdate;
                            }
                            if (diligenceInput.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && diligenceInput.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !diligenceInput.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                            {
                                plstartdate = diligenceInput.pllicenseModels[i].PL_StartDateYear;
                                //pllicenseModel.PL_Start_Date = plstartdate;
                            }

                            if (!diligenceInput.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && !diligenceInput.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !diligenceInput.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                            {
                                plenddate = diligenceInput.pllicenseModels[i].PL_EndDateMonth + " " + diligenceInput.pllicenseModels[i].PL_EndDateDay + ", " + diligenceInput.pllicenseModels[i].PL_EndDateYear;
                                //pllicenseModel.PL_End_Date = plenddate;

                            }
                            if (diligenceInput.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && diligenceInput.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && diligenceInput.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                            {
                                plenddate = "<not provided>";
                            }
                            if (diligenceInput.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && !diligenceInput.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !diligenceInput.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                            {
                                plenddate = diligenceInput.pllicenseModels[i].PL_EndDateMonth + " " + diligenceInput.pllicenseModels[i].PL_EndDateYear;
                                //pllicenseModel.PL_End_Date = plenddate;
                            }
                            if (diligenceInput.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && diligenceInput.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !diligenceInput.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                            {
                                plenddate = diligenceInput.pllicenseModels[i].PL_EndDateYear;
                                //pllicenseModel.PL_End_Date = plenddate;
                            }


                            CommentModel comment2 = _context.DbComment
                                        .Where(u => u.Comment_type == "PL1")
                                        .FirstOrDefault();
                            if (i == 0) { pl_comment = "\n"; }
                            string strplorgfont = string.Concat(diligenceInput.pllicenseModels[i].PL_Organization.ToString(), " CHANGEFONTHEADER");
                            if (diligenceInput.pllicenseModels[i].PL_Confirmed.Equals("Yes"))
                            {
                                if (diligenceInput.pllicenseModels.Count - 1 == i)
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
                                if (diligenceInput.pllicenseModels[i].PL_Confirmed.Equals("No"))
                                {
                                    if (diligenceInput.pllicenseModels.Count - 1 == i)
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
                                    if (diligenceInput.pllicenseModels.Count - 1 == i)
                                    {
                                        pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.other_comment.ToString(), "APPENDPLTEXT", i.ToString());
                                    }
                                    else
                                    {
                                        pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.other_comment.ToString(), "APPENDPLTEXT", i.ToString(), "\n\n");
                                    }
                                }
                            }
                            pl_comment = pl_comment.Replace("[Last Name]", diligenceInput.diligenceInputModel.LastName.ToString());
                            pl_comment = pl_comment.Replace("[PL Organization]", diligenceInput.pllicenseModels[i].PL_Organization);
                            pl_comment = pl_comment.Replace("[Professional License Type]", diligenceInput.pllicenseModels[i].PL_License_Type.ToString());
                            if (diligenceInput.pllicenseModels[i].PL_Number.Equals(""))
                            { pl_comment = pl_comment.Replace(", with a license number [PL Number]", ""); }
                            else
                            {
                                pl_comment = pl_comment.Replace("[PL Number]", diligenceInput.pllicenseModels[i].PL_Number.ToString());
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
                    doc.Replace("PLLICENSEDESCRIPTION", pl_comment, true, false);
                    doc.SaveToFile(savePath);
                    if (diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                    {
                        try
                        {
                            for (int j = 1; j < 8; j++)
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
                        catch { }
                    }
                    doc.SaveToFile(savePath);
                    doc.Replace("CHANGEFONTHEADER", "", true, false);
                    doc.SaveToFile(savePath);
                    bnres = "";
                    if (diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                    {
                        for (int i = 0; i < diligenceInput.pllicenseModels.Count; i++)
                        {
                            bnres = "";
                            if (diligenceInput.pllicenseModels[i].PL_Confirmed.ToString().Equals("Result Pending"))
                            {
                                try
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
                                                    if (abc.ToString().Contains("APPENDPLTEXT"))
                                                    {
                                                        textRange = para.AppendText(" Efforts to independently verify the same are currently ongoing, the results of which will be provided under separate cover, if and when received.");
                                                        textRange.CharacterFormat.FontName = "Calibri (Body)";
                                                        textRange.CharacterFormat.FontSize = 11;
                                                        textRange.CharacterFormat.Italic = true;
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
                                }
                                catch { }
                                if (i == diligenceInput.pllicenseModels.Count - 1)
                                {
                                    bnres = "";
                                    try
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
                                    catch { }
                                }
                            }
                        }
                    }
                    //US_SEC
                    string usseccommentmodel = "";
                    CommentModel usseccommentModel1 = _context.DbComment
                                     .Where(u => u.Comment_type == "Reg_US_SEC")
                                     .FirstOrDefault();
                    if (diligenceInput.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered"))
                    {
                        usseccommentmodel = "";
                        doc.Replace("SECCOMMENT", "", true, true);
                        doc.Replace("SECRESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Has_Reg_US_SEC.ToString().Equals("Yes-With Adverse"))
                        {
                            usseccommentmodel = usseccommentModel1.confirmed_comment.ToString();
                            doc.Replace("SECCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("SECRESULT", "Records", true, true);
                        }
                        else
                        {
                            usseccommentmodel = usseccommentModel1.unconfirmed_comment.ToString();
                            doc.Replace("SECCOMMENT", "", true, true);
                            doc.Replace("SECRESULT", "Clear", true, true);
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
                    if (diligenceInput.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered"))
                    {
                        ukfcacommentmodel = "";
                        doc.Replace("UK_FCAHEADER", "", true, true);
                        doc.Replace("UKFICOCOMMENT", "", true, true);
                        doc.Replace("UKFICORESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Yes-With Adverse"))
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
                        // doc.Replace("UK_FCAHEADER", "\nUnited Kingdom’s Financial Conduct Authority \n", true, true);
                        ukfcacommentmodel = string.Concat("\nUnited Kingdom’s Financial Conduct Authority\n\n", ukfcacommentmodel, "\n");
                    }
                    //FINRA
                    string finracommentmodel = "";
                    CommentModel finracommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "FINRA")
                                    .FirstOrDefault();
                    if (diligenceInput.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered"))
                    {
                        finracommentmodel = "";
                        doc.Replace("FINRAHEADER", "", true, true);
                        doc.Replace("FININDCOMMENT", "", true, true);
                        doc.Replace("FININDRESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Has_Reg_FINRA.ToString().Equals("Yes-With Adverse"))
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
                        //doc.Replace("FINRAHEADER", "\nUnited States Financial Industry Regulatory Authority \n", true, true);
                        finracommentmodel = string.Concat("\nUnited States Financial Industry Regulatory Authority\n\n", finracommentmodel, "\n");
                    }
                    //NFA
                    string nfacommentmodel = "";
                    CommentModel nfacommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "NFA")
                                    .FirstOrDefault();
                    if (diligenceInput.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered"))
                    {
                        nfacommentmodel = "";
                        doc.Replace("US_NFAHEADER", "", true, true);
                        doc.Replace("USNFCOMMENT", "", true, true);
                        doc.Replace("USNFRESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Has_Reg_US_NFA.ToString().Equals("Yes-With Adverse"))
                        {
                            nfacommentmodel = nfacommentmodel1.confirmed_comment.ToString();
                            doc.Replace("USNFCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("USNFRESULT", "Records", true, true);
                        }
                        else
                        {
                            nfacommentmodel = nfacommentmodel1.unconfirmed_comment.ToString();
                            doc.Replace("USNFCOMMENT", "", true, true);
                            doc.Replace("USNFRESULT", "Clear", true, true);
                        }
                        nfacommentmodel = nfacommentmodel.Replace("*n ", "\n");
                        nfacommentmodel = nfacommentmodel.Replace("*n", "\n");
                        nfacommentmodel = nfacommentmodel.Replace("*t", "\t");
                        //doc.Replace("US_NFAHEADER", "\nUnited States National Futures Association  \n", true, true);
                        nfacommentmodel = string.Concat("\nUnited States National Futures Association\n\n", nfacommentmodel, "\n");
                    }
                    string hksfccommentmodel = "";
                    string holdslicensecommentmodel = "";

                    //HKFSC
                    if (diligenceInput.otherdetails.Registered_with_HKSFC.ToString().Equals("Not Registered"))
                    {
                        hksfccommentmodel = "";
                        doc.Replace("KONSECCOMMENT", "", true, true);
                        doc.Replace("KONSECRESULT", "Clear", true, true);
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes - Without Adverse"))
                        {
                            hksfccommentmodel = "[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>)), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nThere are no public disciplinary actions on file against the subject within the past five years.";
                            doc.Replace("KONSECCOMMENT", "", true, true);
                            doc.Replace("KONSECRESULT", "Clear", true, true);
                        }
                        else
                        {
                            hksfccommentmodel = "[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>)), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nAdditionally, the following public disciplinary actions were on file in connection with the subject:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                            doc.Replace("KONSECCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("KONSECRESULT", "Records", true, true);
                        }
                        hksfccommentmodel = string.Concat("\nHong Kong Securities and Futures Commission\n\n", hksfccommentmodel, "\n");
                    }
                    //Holds Any License 
                    string strotherheader = "";
                    CommentModel holdslicensecommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "Other_License_Language")
                                    .FirstOrDefault();
                    if (diligenceInput.otherdetails.worldcheck_discloseBA.ToString().Equals("Yes") || diligenceInput.otherdetails.worldcheck_undiscloseBA.ToString().Equals("Yes"))
                    {
                        regflag = "Records";
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                        {
                            regflag = "Records";
                        }
                        else
                        {
                            regflag = "Clear";

                        }
                    }
                    if (regflag == "Records")
                    {
                        holdslicensecommentmodel = "Investigative efforts did not reveal any additional professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction and license-type basis to confirm the same. ";
                        strotherheader = "\nOther Professional Licensures and/or Designations\n\n";
                    }
                    else
                    {
                        holdslicensecommentmodel = "Investigative efforts did not reveal any professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction and license-type basis to confirm the same. ";
                        strotherheader = "\nOther Professional Licensures and/or Designations\n\n";
                    }
                    if (diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Not Registered") || diligenceInput.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered") || diligenceInput.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered") || diligenceInput.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered") || diligenceInput.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered"))
                    {
                        strotherheader = "\n";
                    }
                    holdslicensecommentmodel = string.Concat(strotherheader, holdslicensecommentmodel, "\n");
                    if (famcount == diligenceInputs.familyModels.Count - minor_count-1 || famcount == diligenceInputs.familyModels.Count - 1)
                    {
                        if (diligenceInput.pllicenseModels[0].General_PL_License.ToString().ToUpper().Equals("YES")) { doc.Replace("[PLLICENSEDESC]", string.Concat("\n", usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false); }
                        else
                        {
                            doc.Replace("[PLLICENSEDESC]", string.Concat(usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false);
                        }
                    }
                    else
                    {
                        if (diligenceInput.pllicenseModels[0].General_PL_License.ToString().ToUpper().Equals("YES")) { doc.Replace("[PLLICENSEDESC]", string.Concat("\n", usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel, "\nFirst_Name Middle_Name Last_Name\nPLLICENSEDESCRIPTION[PLLICENSEDESC]"), true, false); }
                        else
                        {
                            doc.Replace("[PLLICENSEDESC]", string.Concat(usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel, "\nFirst_Name Middle_Name Last_Name\nPLLICENSEDESCRIPTION[PLLICENSEDESC]"), true, false);
                        }
                    }
                    doc.SaveToFile(savePath);
                    //CountrySpecific
                    try
                    {
                        if (diligenceInput.additional_Countries.Count() > 0)
                        {
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "DRIVINGHISTORYDESCRIPTION\nDRIVINGHISTORY2DESCRIPTION", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "CREDITHISTORYDESCRIPTION\nCREDITHISTORY2DESCRIPTION", true, true);
                            doc.Replace("DRIVINGCOMMENT", "DRIVINGCOMMENT\nDRIVING2COMMENT", true, true);
                            doc.Replace("DRIVINGRESULT", "DRIVINGRESULT\nDRIVING2RESULT", true, true);
                            doc.Replace("PERCREDITCOMMENT", "PERCREDITCOMMENT\nPERCREDIT2COMMENT", true, true);
                            doc.Replace("PERCREDITRESULT", "PERCREDITRESULT\nPERCREDIT2RESULT", true, true);                          
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "COUNTRYSPECIFICREGHITDISCRIPTION\nCOUNTRYSPECIFICREGHIT2DISCRIPTION", true, true);                            
                        }
                    }
                    catch { }
                    
                    if (famcount == 0)
                    {
                        doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "\n\nFirst_Name0 Middle_Name0 Last_Name0 COUNTRYSPECIFICREGHITDISCRIPTION", true, false);
                    }
                    CountrySpecificModel CS1 = new CountrySpecificModel();
                    string strothers_SpeRegHits = "";

                    switch (diligenceInput.diligenceInputModel.Country.ToString())
                    {
                        case "Iraq":
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Iraq.", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Iraq.", true, true);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Iraq.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Iraq.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "<Local> Chamber of Commerce", true, true);
                            doc.Replace("[Property Registry]", "Iraqi Ministry of Justice, Directorate General of Real Estate Registration", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[Courts]", "Iraqi Ministry of Interior, Directorate of Criminal Identification, media and other information from sources in judicial circles in Iraq, and the Iraqi Central Bank", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Russia":
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Russia.", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Russia.", true, true);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Russia.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Russia.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "the Russian Business Register and the Ministry of Labor and Social Protection of Russia", true, true);
                            doc.Replace("[Property Registry]", "<Local> Property Registry", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[Courts]", "Russia Ministry of Justice, Russian Business Register, and the Moscow Tax Office", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Jordan":
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Jordan.", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Jordan.", true, true);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Jordan.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Jordan.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "the Amman Chamber of Commerce", true, true);
                            doc.Replace("[Property Registry]", "<Local> Property Registry", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[Courts]", "Justice Palace of Jordan,  media and other information from sources in judicial circles in Jordan, and the Jordan Central Bank", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Qatar":
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Qatar.", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Qatar.", true, true);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Qatar.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Qatar.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "Doha Chamber of Commerce and Industry", true, true);
                            doc.Replace("[Property Registry]", "<Local> Property Registry", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[Courts]", "the Qatari Ministry of Interior Criminal Evidence and Information Department,  media and other information from sources in judicial circles in Qatar, and the Qatar Central Bank", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Lebanon":
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Lebanon.", true, true);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Lebanon.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Lebanon.", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Lebanon.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "Beirut Chamber of Commerce", true, true);
                            doc.Replace("[Property Registry]", "<Local> Property Registry", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[Courts]", "the Lebanon Directorate General of Interior, Security Forces, media and other information from sources in judicial circles in Lebanon, and the Lebanon Central Bank", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Singapore":
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Singapore.", true, true);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Singapore.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Singapore.", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Singapore.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "the Singapore Accounting and Corporate Regulatory Authority", true, true);
                            doc.Replace("[Property Registry]", "property records in Singapore", true, true);
                            strothers_SpeRegHits = "";
                            if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                            {
                                strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                            }
                            else
                            {
                                strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].\n";
                            }
                            strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "Monetary Authority of Singapore. It should be noted that the Monetary Authority of Singapore (“MAS”) regulates companies, and not individuals, and while individuals representing investment and fund management firms should have a Capital Market License or Financial Advice License, for which they have to pass an examination to get a certificate, the onus is on the companies that they represent to ensure that they have the necessary qualifications and credentials.  Further, individual fund managers and investment advisors are not regulated by the MAS. In addition, the subject’s MAS Representative Number -- if any -- would be required in order to confirm if the subject is registered with the MAS, and if any disciplinary actions were filed in connection with this registration, as well as to confirm if the subject has met MAS’s “fit and proper” criteria");
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                            doc.Replace("[Courts]", "the Insolvency & Public Trustee’s Office, the High Court, District Court, Magistrate’s Court and the Subordinate Courts of Singapore", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Germany":
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Germany.", true, true);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Germany.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "German Corporate Registry (“Handelsregister”)", true, true);
                            doc.Replace("[Property Registry]", "It is noted that third-party access to property ownership records is restricted in Germany.", true, true);
                            string strregulatortsearch = "";
                            if (diligenceInput.csModel.germany_regsearches.ToString().Equals("BaFin with Hits"))
                            {
                                strregulatortsearch = "Searches were conducted through the German Federal Financial Supervisory Authority (“Bundesanstalt für Finanzdienstleistungsaufsicht,” or “BaFin”), which revealed <Investigator to insert summary of results here>";
                            }
                            else
                            {
                                strregulatortsearch = "Searches were conducted through the German Federal Financial Supervisory Authority (“Bundesanstalt für Finanzdienstleistungsaufsicht,” or “BaFin”), which did not identify any records in connection with the subject";
                            }
                            strothers_SpeRegHits = "";
                            if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                            {
                                strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                            }
                            else
                            {
                                strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].";
                            }
                            strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", strregulatortsearch);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                            doc.Replace("[Courts]", "the Federal Court of Justice (Bundesgerichtshof), the Federal Labor Court (Bundesarbeitsgericht), the Federal Administrative Court (Bundesverwaltungsgericht), and the Federal Finance Court (Bundessozialgericht), <investigator to add other relevant courts as applicable>", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Switzerland":
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Switzerland.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Switzerland.", true, true);
                            doc.Replace("[Corp Registry]", "Corporate records in Switzerland", true, true);
                            doc.Replace("[Property Registry]", "property records in Switzerland", true, true);
                            strothers_SpeRegHits = "";
                            if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                            {
                                strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                            }
                            else
                            {
                                strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].\n";
                            }
                            strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "It is noted that the Swiss Financial Market Supervisory Authority (“FINMA”) regulates certain companies, and not individuals, and in this regard, the subject is not registered with FINMA");
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                            doc.Replace("[Courts]", "the Swiss Official Gazette of Commercial (“SOGC”), the Federal Supreme Court of Switzerland, the Federal Criminal, Patent and Administrative Courts of Switzerland, <investigator to add other relevant courts as applicable>", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "India":
                            string strcorpreg = "";
                            if (diligenceInput.csModel.india_corpregistry.ToString().Equals("Has Business Affiliations"))
                            {
                                strcorpreg = "In addition to the above, while records maintained on individuals by the Indian Registrar of Companies and the Ministry of Corporate Affairs are not available to third-party companies, research efforts identified [LastName] as an Officer, Director and/or Shareholder of the following business entities in [Country]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                            }
                            else
                            {
                                strcorpreg = "It should be noted that records maintained on individuals by the Indian Registrar of Companies and the Ministry of Corporate Affairs are not available to third-party companies.";
                            }
                            doc.Replace("[Corp Registry]", strcorpreg, true, true);
                            doc.Replace("[Property Registry]", "Third-party access to property records is restricted in India.", true, true);
                            strothers_SpeRegHits = "";
                            if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                            {
                                strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                            }
                            else
                            {
                                strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].\n";
                            }
                            strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "Securities Exchange Board of India (“SEBI”)");
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                            doc.Replace("[Courts]", "the Local Police, Taluka/Small Causes Courts, the District/Session Courts, as well as the High Court and Supreme Court of India", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Saudi Arabia":
                            doc.SaveToFile(savePath);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Saudi Arabia.", true, true);
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Saudi Arabia.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Saudi Arabia.", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Saudi Arabia.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "Third-party access to property ownership records is restricted in Saudi Arabia.", true, true);
                            doc.Replace("[Property Registry]", "Riyadh Chamber of Commerce", true, true);
                            doc.Replace("[Courts]", "Saudi Arabia Ministry of Interior, searches of media and other information from sources in judicial circles in Saudi Arabia, and the Saudi Arabia Central Bank ", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "United Arab Emirates":
                            doc.Replace("[Corp Registry]", "<Local Emirate> Chamber of Commerce ", true, true);
                            doc.Replace("[Property Registry]", "<Local Emirate> Land Department ", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[Courts]", "<Local Emirate> Ministry of Interior, the <Local Emirate> Court of First Instance,  searches of media and other information from sources in judicial circles in the U.A.E., and the U.A.E. Central Bank", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "China":
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in China.", true, true);
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in China.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in China.", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in China.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "available corporate registry records in China ", true, true);
                            doc.Replace("[Property Registry]", "It is noted that there are no publicly available real estate databases in China that provide information on historical ownership or private leases. ", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[Courts]", "<investigator to insert source from PI report>", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            break;
                        case "Hong Kong":
                            doc.SaveToFile(savePath);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Hong Kong.", true, true);
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Hong Kong.", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Hong Kong.", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Hong Kong.", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("[Corp Registry]", "Hong Kong companies registry ", true, true);
                            doc.Replace("[Property Registry]", "Hong Kong property records ", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.Replace("[Courts]", "High Court, District Court, Magistrate Courts or the Small Claims Tribunal, as well as the Official Receiver’s Office, and proprietary litigation databases", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "France":

                            CommentModel fr_dhcommentmodel1 = _context.DbComment
                                   .Where(u => u.Comment_type == "fr_drivinghistory")
                                   .FirstOrDefault();
                            CommentModel fr_chcommentmodel = _context.DbComment
                                   .Where(u => u.Comment_type == "fr_credithistory")
                                   .FirstOrDefault();
                            CommentModel fr_crcommentmodel = _context.DbComment
                                   .Where(u => u.Comment_type == "fr_corpreg")
                                   .FirstOrDefault();
                            CommentModel fr_prcommentmodel = _context.DbComment
                                   .Where(u => u.Comment_type == "fr_propertyreg")
                                   .FirstOrDefault();
                            doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                            doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in [Country]", true, true);
                            doc.Replace("DRIVINGRESULT", "N/A", true, true);
                            doc.Replace("PERCREDITCOMMENT", "Third-party access to credit history records is restricted in [Country] ", true, true);
                            doc.Replace("PERCREDITRESULT", "N/A", true, true);
                            doc.Replace("DRIVINGHISTORYDESCRIPTION", fr_dhcommentmodel1.confirmed_comment.ToString(), true, true);
                            doc.Replace("CREDITHISTORYDESCRIPTION", fr_chcommentmodel.confirmed_comment.ToString(), true, true);
                            doc.Replace("[Corp Registry]", fr_crcommentmodel.confirmed_comment.ToString(), true, true);
                            doc.Replace("[Property Registry]", fr_prcommentmodel.confirmed_comment.ToString(), true, true);
                            string fr_srhcommentmodel = "";
                            CommentModel fr_srhcommentmodel1 = _context.DbComment
                                  .Where(u => u.Comment_type == "fr_specreghit")
                                  .FirstOrDefault();
                            CommentModel fr_lrhcommentmodel1 = _context.DbComment
                                   .Where(u => u.Comment_type == "fr_legalrechit")
                                   .FirstOrDefault();
                            if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                            {
                                fr_srhcommentmodel = fr_srhcommentmodel1.confirmed_comment.ToString();
                            }
                            else
                            {
                                fr_srhcommentmodel = fr_srhcommentmodel1.unconfirmed_comment.ToString();
                            }
                            fr_srhcommentmodel = fr_srhcommentmodel.Replace("*n ", "\n");
                            fr_srhcommentmodel = fr_srhcommentmodel.Replace("*n", "\n");
                            fr_srhcommentmodel = fr_srhcommentmodel.Replace("*t", "\t");
                            fr_srhcommentmodel = string.Concat("\n\n", fr_srhcommentmodel, "\n");
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", fr_srhcommentmodel, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "United Kingdom":
                            CommentModel uk_crcommentmodel = _context.DbComment
                                   .Where(u => u.Comment_type == "uk_corpreg")
                                   .FirstOrDefault();
                            CommentModel uk_prcommentmodel = _context.DbComment
                                   .Where(u => u.Comment_type == "uk_propertyreg")
                                   .FirstOrDefault();
                            doc.Replace("COUNTRYLEGDESC/", "", true, true);
                            doc.SaveToFile(savePath);
                            doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                            doc.Replace("[Corp Registry]", uk_crcommentmodel.confirmed_comment.ToString(), true, true);
                            doc.Replace("[Property Registry]", uk_prcommentmodel.confirmed_comment.ToString(), true, true);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Canada":
                            doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                            string can_csregcommentmodel = "";
                            CommentModel can_albercommentmodel = _context.DbComment
                                   .Where(u => u.Comment_type == "can_alberta")
                                   .FirstOrDefault();
                            CommentModel can_csregcommentmodel1 = _context.DbComment
                                   .Where(u => u.Comment_type == "can_specreghit")
                                   .FirstOrDefault();
                            doc.Replace("[Corp Registry]", "<Investigator to add source for corp records>", true, true);
                            doc.Replace("[Property Registry]", "<Investigator to add source for property records>", true, true);
                            switch (diligenceInput.csModel.country_specific_reg_hits.ToString())
                            {
                                case "Yes":
                                    can_csregcommentmodel = can_csregcommentmodel1.confirmed_comment.ToString();
                                    break;
                                case "No":
                                    can_csregcommentmodel = can_csregcommentmodel1.unconfirmed_comment.ToString();
                                    break;
                            }
                            can_csregcommentmodel = can_csregcommentmodel.Replace("*n ", "\n");
                            can_csregcommentmodel = can_csregcommentmodel.Replace("*n", "\n");
                            can_csregcommentmodel = can_csregcommentmodel.Replace("*t", "\t");
                            can_csregcommentmodel = string.Concat("\n\n", can_csregcommentmodel);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", can_csregcommentmodel, true, true);
                            doc.SaveToFile(savePath);
                            break;
                        case "Australia":
                            string aus_asiccommentmodel = "";
                            CommentModel aus_asiccommentmodel1 = _context.DbComment
                                   .Where(u => u.Comment_type == "aus_asicreg")
                                   .FirstOrDefault();
                            CommentModel aus_cpregcommentmodel1 = _context.DbComment
                                   .Where(u => u.Comment_type == "aus_corpreg")
                                   .FirstOrDefault();
                            CommentModel aus_pregcommentmodel1 = _context.DbComment
                                   .Where(u => u.Comment_type == "aus_propreg")
                                   .FirstOrDefault();
                            doc.Replace("[Corp Registry]", aus_cpregcommentmodel1.confirmed_comment.ToString(), true, true);
                            doc.Replace("[Property Registry]", aus_pregcommentmodel1.confirmed_comment.ToString(), true, true);
                            switch (diligenceInput.csModel.country_specific_reg_hits.ToString())
                            {
                                case "Yes":
                                    aus_asiccommentmodel = aus_asiccommentmodel1.confirmed_comment.ToString();
                                    break;
                                case "No":
                                    aus_asiccommentmodel = aus_asiccommentmodel1.unconfirmed_comment.ToString();
                                    break;
                            }
                            aus_asiccommentmodel = aus_asiccommentmodel.Replace("*n ", "\n");
                            aus_asiccommentmodel = aus_asiccommentmodel.Replace("*n", "\n");
                            aus_asiccommentmodel = aus_asiccommentmodel.Replace("*t", "\t");
                            aus_asiccommentmodel = string.Concat("\n\n", aus_asiccommentmodel);
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", aus_asiccommentmodel, true, true);
                            doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                            doc.SaveToFile(savePath);
                            break;
                        default:
                            doc.Replace("[Corp Registry]", "<Investigator to add source for corp records>", true, true);
                            doc.Replace("[Property Registry]", "<Investigator to add source for property records>", true, true);
                            strothers_SpeRegHits = "";
                            if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                            {
                                strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                            }
                            else
                            {
                                strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].\n";
                            }
                            strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "<insert Country Regulatory Body or delete section if not applicable>");
                            doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                            doc.SaveToFile(savePath);
                            doc.Replace("[Courts]", "<insert relevant courts>", true, true);
                            doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                            doc.SaveToFile(savePath);
                            break;
                    }
                    doc.SaveToFile(savePath);
                    for (int k = 0; k < diligenceInput.additional_Countries.Count(); k++) 
                    {
                        if (k == diligenceInput.additional_Countries.Count() - 1)
                        {
                            doc.Replace("DRIVINGHISTORY2DESCRIPTION", "DRIVINGHISTORYDESCRIPTION", true, true);
                            doc.Replace("CREDITHISTORY2DESCRIPTION", "CREDITHISTORYDESCRIPTION", true, true);
                            doc.Replace("DRIVING2COMMENT", "DRIVINGCOMMENT", true, true);
                            doc.Replace("DRIVING2RESULT", "DRIVINGRESULT", true, true);
                            doc.Replace("PERCREDIT2COMMENT", "PERCREDITCOMMENT", true, true);
                            doc.Replace("PERCREDIT2RESULT", "PERCREDITRESULT", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHIT2DISCRIPTION", "COUNTRYSPECIFICREGHITDISCRIPTION", true, true);
                        }
                        else
                        {
                            doc.Replace("DRIVINGHISTORY2DESCRIPTION", "DRIVINGHISTORYDESCRIPTION\nDRIVINGHISTORY2DESCRIPTION", true, true);
                            doc.Replace("CREDITHISTORY2DESCRIPTION", "CREDITHISTORYDESCRIPTION\nCREDITHISTORY2DESCRIPTION", true, true);
                            doc.Replace("DRIVING2COMMENT", "DRIVINGCOMMENT\nDRIVING2COMMENT", true, true);
                            doc.Replace("DRIVING2RESULT", "DRIVINGRESULT\nDRIVING2RESULT", true, true);
                            doc.Replace("PERCREDIT2COMMENT", "PERCREDITCOMMENT\nPERCREDIT2COMMENT", true, true);
                            doc.Replace("PERCREDIT2RESULT", "PERCREDITRESULT\nPERCREDIT2RESULT", true, true);
                            doc.Replace("COUNTRYSPECIFICREGHIT2DISCRIPTION", "COUNTRYSPECIFICREGHITDISCRIPTION\nCOUNTRYSPECIFICREGHIT2DISCRIPTION", true, true);
                        }
                        switch (diligenceInput.additional_Countries[k].additionalCountries.ToString())
                        {
                            case "Iraq":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Iraq.", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Iraq.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Iraq.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Iraq.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "<Local> Chamber of Commerce", true, true);
                                doc.Replace("[Property Registry]", "Iraqi Ministry of Justice, Directorate General of Real Estate Registration", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[Courts]", "Iraqi Ministry of Interior, Directorate of Criminal Identification, media and other information from sources in judicial circles in Iraq, and the Iraqi Central Bank", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Russia":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Russia.", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Russia.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Russia.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Russia.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "the Russian Business Register and the Ministry of Labor and Social Protection of Russia", true, true);
                                doc.Replace("[Property Registry]", "<Local> Property Registry", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[Courts]", "Russia Ministry of Justice, Russian Business Register, and the Moscow Tax Office", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Jordan":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Jordan.", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Jordan.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Jordan.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Jordan.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "the Amman Chamber of Commerce", true, true);
                                doc.Replace("[Property Registry]", "<Local> Property Registry", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[Courts]", "Justice Palace of Jordan,  media and other information from sources in judicial circles in Jordan, and the Jordan Central Bank", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Qatar":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Qatar.", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Qatar.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Qatar.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Qatar.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "Doha Chamber of Commerce and Industry", true, true);
                                doc.Replace("[Property Registry]", "<Local> Property Registry", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[Courts]", "the Qatari Ministry of Interior Criminal Evidence and Information Department,  media and other information from sources in judicial circles in Qatar, and the Qatar Central Bank", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Lebanon":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Lebanon.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Lebanon.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Lebanon.", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Lebanon.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "Beirut Chamber of Commerce", true, true);
                                doc.Replace("[Property Registry]", "<Local> Property Registry", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[Courts]", "the Lebanon Directorate General of Interior, Security Forces, media and other information from sources in judicial circles in Lebanon, and the Lebanon Central Bank", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Singapore":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Singapore.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Singapore.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Singapore.", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Singapore.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "the Singapore Accounting and Corporate Regulatory Authority", true, true);
                                doc.Replace("[Property Registry]", "property records in Singapore", true, true);
                                strothers_SpeRegHits = "";
                                if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                                {
                                    strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                                }
                                else
                                {
                                    strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].\n";
                                }
                                strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "Monetary Authority of Singapore. It should be noted that the Monetary Authority of Singapore (“MAS”) regulates companies, and not individuals, and while individuals representing investment and fund management firms should have a Capital Market License or Financial Advice License, for which they have to pass an examination to get a certificate, the onus is on the companies that they represent to ensure that they have the necessary qualifications and credentials.  Further, individual fund managers and investment advisors are not regulated by the MAS. In addition, the subject’s MAS Representative Number -- if any -- would be required in order to confirm if the subject is registered with the MAS, and if any disciplinary actions were filed in connection with this registration, as well as to confirm if the subject has met MAS’s “fit and proper” criteria");
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                                doc.Replace("[Courts]", "the Insolvency & Public Trustee’s Office, the High Court, District Court, Magistrate’s Court and the Subordinate Courts of Singapore", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Germany":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Germany.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Germany.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "German Corporate Registry (“Handelsregister”)", true, true);
                                doc.Replace("[Property Registry]", "It is noted that third-party access to property ownership records is restricted in Germany.", true, true);
                                string strregulatortsearch = "";
                                if (diligenceInput.csModel.germany_regsearches.ToString().Equals("BaFin with Hits"))
                                {
                                    strregulatortsearch = "Searches were conducted through the German Federal Financial Supervisory Authority (“Bundesanstalt für Finanzdienstleistungsaufsicht,” or “BaFin”), which revealed <Investigator to insert summary of results here>";
                                }
                                else
                                {
                                    strregulatortsearch = "Searches were conducted through the German Federal Financial Supervisory Authority (“Bundesanstalt für Finanzdienstleistungsaufsicht,” or “BaFin”), which did not identify any records in connection with the subject";
                                }
                                strothers_SpeRegHits = "";
                                if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                                {
                                    strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                                }
                                else
                                {
                                    strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].";
                                }
                                strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", strregulatortsearch);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                                doc.Replace("[Courts]", "the Federal Court of Justice (Bundesgerichtshof), the Federal Labor Court (Bundesarbeitsgericht), the Federal Administrative Court (Bundesverwaltungsgericht), and the Federal Finance Court (Bundessozialgericht), <investigator to add other relevant courts as applicable>", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Switzerland":
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Switzerland.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Switzerland.", true, true);
                                doc.Replace("[Corp Registry]", "Corporate records in Switzerland", true, true);
                                doc.Replace("[Property Registry]", "property records in Switzerland", true, true);
                                strothers_SpeRegHits = "";
                                if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                                {
                                    strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                                }
                                else
                                {
                                    strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].\n";
                                }
                                strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "It is noted that the Swiss Financial Market Supervisory Authority (“FINMA”) regulates certain companies, and not individuals, and in this regard, the subject is not registered with FINMA");
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                                doc.Replace("[Courts]", "the Swiss Official Gazette of Commercial (“SOGC”), the Federal Supreme Court of Switzerland, the Federal Criminal, Patent and Administrative Courts of Switzerland, <investigator to add other relevant courts as applicable>", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "India":
                                string strcorpreg = "";
                                if (diligenceInput.csModel.india_corpregistry.ToString().Equals("Has Business Affiliations"))
                                {
                                    strcorpreg = "In addition to the above, while records maintained on individuals by the Indian Registrar of Companies and the Ministry of Corporate Affairs are not available to third-party companies, research efforts identified [LastName] as an Officer, Director and/or Shareholder of the following business entities in [Country]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                                }
                                else
                                {
                                    strcorpreg = "It should be noted that records maintained on individuals by the Indian Registrar of Companies and the Ministry of Corporate Affairs are not available to third-party companies.";
                                }
                                doc.Replace("[Corp Registry]", strcorpreg, true, true);
                                doc.Replace("[Property Registry]", "Third-party access to property records is restricted in India.", true, true);
                                strothers_SpeRegHits = "";
                                if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                                {
                                    strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                                }
                                else
                                {
                                    strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].\n";
                                }
                                strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "Securities Exchange Board of India (“SEBI”)");
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                                doc.Replace("[Courts]", "the Local Police, Taluka/Small Causes Courts, the District/Session Courts, as well as the High Court and Supreme Court of India", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Saudi Arabia":
                                doc.SaveToFile(savePath);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Saudi Arabia.", true, true);
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Saudi Arabia.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Saudi Arabia.", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Saudi Arabia.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "Third-party access to property ownership records is restricted in Saudi Arabia.", true, true);
                                doc.Replace("[Property Registry]", "Riyadh Chamber of Commerce", true, true);
                                doc.Replace("[Courts]", "Saudi Arabia Ministry of Interior, searches of media and other information from sources in judicial circles in Saudi Arabia, and the Saudi Arabia Central Bank ", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "United Arab Emirates":
                                doc.Replace("[Corp Registry]", "<Local Emirate> Chamber of Commerce ", true, true);
                                doc.Replace("[Property Registry]", "<Local Emirate> Land Department ", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[Courts]", "<Local Emirate> Ministry of Interior, the <Local Emirate> Court of First Instance,  searches of media and other information from sources in judicial circles in the U.A.E., and the U.A.E. Central Bank", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "China":
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in China.", true, true);
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in China.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in China.", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in China.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "available corporate registry records in China ", true, true);
                                doc.Replace("[Property Registry]", "It is noted that there are no publicly available real estate databases in China that provide information on historical ownership or private leases. ", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[Courts]", "<investigator to insert source from PI report>", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                break;
                            case "Hong Kong":
                                doc.SaveToFile(savePath);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in Hong Kong.", true, true);
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in Hong Kong.", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to personal credit history records is restricted in Hong Kong.", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to personal credit history records is restricted in Hong Kong.", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("[Corp Registry]", "Hong Kong companies registry ", true, true);
                                doc.Replace("[Property Registry]", "Hong Kong property records ", true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.Replace("[Courts]", "High Court, District Court, Magistrate Courts or the Small Claims Tribunal, as well as the Official Receiver’s Office, and proprietary litigation databases", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "France":

                                CommentModel fr_dhcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "fr_drivinghistory")
                                       .FirstOrDefault();
                                CommentModel fr_chcommentmodel = _context.DbComment
                                       .Where(u => u.Comment_type == "fr_credithistory")
                                       .FirstOrDefault();
                                CommentModel fr_crcommentmodel = _context.DbComment
                                       .Where(u => u.Comment_type == "fr_corpreg")
                                       .FirstOrDefault();
                                CommentModel fr_prcommentmodel = _context.DbComment
                                       .Where(u => u.Comment_type == "fr_propertyreg")
                                       .FirstOrDefault();
                                doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in [Country]", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to credit history records is restricted in [Country] ", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", fr_dhcommentmodel1.confirmed_comment.ToString(), true, true);
                                doc.Replace("CREDITHISTORYDESCRIPTION", fr_chcommentmodel.confirmed_comment.ToString(), true, true);
                                doc.Replace("[Corp Registry]", fr_crcommentmodel.confirmed_comment.ToString(), true, true);
                                doc.Replace("[Property Registry]", fr_prcommentmodel.confirmed_comment.ToString(), true, true);
                                string fr_srhcommentmodel = "";
                                CommentModel fr_srhcommentmodel1 = _context.DbComment
                                      .Where(u => u.Comment_type == "fr_specreghit")
                                      .FirstOrDefault();
                                CommentModel fr_lrhcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "fr_legalrechit")
                                       .FirstOrDefault();
                                if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                                {
                                    fr_srhcommentmodel = fr_srhcommentmodel1.confirmed_comment.ToString();
                                }
                                else
                                {
                                    fr_srhcommentmodel = fr_srhcommentmodel1.unconfirmed_comment.ToString();
                                }
                                fr_srhcommentmodel = fr_srhcommentmodel.Replace("*n ", "\n");
                                fr_srhcommentmodel = fr_srhcommentmodel.Replace("*n", "\n");
                                fr_srhcommentmodel = fr_srhcommentmodel.Replace("*t", "\t");
                                fr_srhcommentmodel = string.Concat("\n\n", fr_srhcommentmodel, "\n");
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", fr_srhcommentmodel, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "United Kingdom":
                                CommentModel uk_crcommentmodel = _context.DbComment
                                       .Where(u => u.Comment_type == "uk_corpreg")
                                       .FirstOrDefault();
                                CommentModel uk_prcommentmodel = _context.DbComment
                                       .Where(u => u.Comment_type == "uk_propertyreg")
                                       .FirstOrDefault();
                                doc.Replace("COUNTRYLEGDESC/", "", true, true);
                                doc.SaveToFile(savePath);
                                doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                                doc.Replace("[Corp Registry]", uk_crcommentmodel.confirmed_comment.ToString(), true, true);
                                doc.Replace("[Property Registry]", uk_prcommentmodel.confirmed_comment.ToString(), true, true);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Canada":
                                doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                                string can_csregcommentmodel = "";
                                CommentModel can_albercommentmodel = _context.DbComment
                                       .Where(u => u.Comment_type == "can_alberta")
                                       .FirstOrDefault();
                                CommentModel can_csregcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "can_specreghit")
                                       .FirstOrDefault();
                                doc.Replace("[Corp Registry]", "<Investigator to add source for corp records>", true, true);
                                doc.Replace("[Property Registry]", "<Investigator to add source for property records>", true, true);
                                switch (diligenceInput.csModel.country_specific_reg_hits.ToString())
                                {
                                    case "Yes":
                                        can_csregcommentmodel = can_csregcommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No":
                                        can_csregcommentmodel = can_csregcommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                }
                                can_csregcommentmodel = can_csregcommentmodel.Replace("*n ", "\n");
                                can_csregcommentmodel = can_csregcommentmodel.Replace("*n", "\n");
                                can_csregcommentmodel = can_csregcommentmodel.Replace("*t", "\t");
                                can_csregcommentmodel = string.Concat("\n\n", can_csregcommentmodel);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", can_csregcommentmodel, true, true);
                                doc.SaveToFile(savePath);
                                break;
                            case "Australia":
                                string aus_asiccommentmodel = "";
                                CommentModel aus_asiccommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "aus_asicreg")
                                       .FirstOrDefault();
                                CommentModel aus_cpregcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "aus_corpreg")
                                       .FirstOrDefault();
                                CommentModel aus_pregcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "aus_propreg")
                                       .FirstOrDefault();
                                doc.Replace("[Corp Registry]", aus_cpregcommentmodel1.confirmed_comment.ToString(), true, true);
                                doc.Replace("[Property Registry]", aus_pregcommentmodel1.confirmed_comment.ToString(), true, true);
                                switch (diligenceInput.csModel.country_specific_reg_hits.ToString())
                                {
                                    case "Yes":
                                        aus_asiccommentmodel = aus_asiccommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No":
                                        aus_asiccommentmodel = aus_asiccommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                }
                                aus_asiccommentmodel = aus_asiccommentmodel.Replace("*n ", "\n");
                                aus_asiccommentmodel = aus_asiccommentmodel.Replace("*n", "\n");
                                aus_asiccommentmodel = aus_asiccommentmodel.Replace("*t", "\t");
                                aus_asiccommentmodel = string.Concat("\n\n", aus_asiccommentmodel);
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", aus_asiccommentmodel, true, true);
                                doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                                doc.SaveToFile(savePath);
                                break;
                            default:
                                doc.Replace("[Corp Registry]", "<Investigator to add source for corp records>", true, true);
                                doc.Replace("[Property Registry]", "<Investigator to add source for property records>", true, true);
                                strothers_SpeRegHits = "";
                                if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                                {
                                    strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                                }
                                else
                                {
                                    strothers_SpeRegHits = "\n\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].\n";
                                }
                                strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "<insert Country Regulatory Body or delete section if not applicable>");
                                doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                                doc.SaveToFile(savePath);
                                doc.Replace("[Courts]", "<insert relevant courts>", true, true);
                                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, true);
                                doc.SaveToFile(savePath);
                                break;
                        }
                    }
                    //Regulatory_Red_Flag
                    string regredflagommentmodel = "";
                    CommentModel regredflagcommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "Regulatory_Red_Flag")
                                    .FirstOrDefault();
                    if (diligenceInput.otherdetails.RegulatoryFlag.ToString().Equals("Yes"))
                    {
                        regredflagommentmodel = regredflagcommentmodel1.confirmed_comment.ToString();
                    }
                    else
                    {
                        regredflagommentmodel = regredflagcommentmodel1.unconfirmed_comment.ToString();
                    }
                    regredflagommentmodel = regredflagommentmodel.Replace("*n ", "\n");
                    regredflagommentmodel = regredflagommentmodel.Replace("*n", "\n");
                    regredflagommentmodel = regredflagommentmodel.Replace("*t", "\t");
                    regredflagommentmodel = string.Concat("\n", regredflagommentmodel, "\n");
                    if (famcount == diligenceInputs.familyModels.Count - minor_count-1 || famcount == diligenceInputs.familyModels.Count - 1)
                    {
                        doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", regredflagommentmodel, true, true);
                    }
                    else
                    {
                        doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", string.Concat(regredflagommentmodel, "\nFirst_Name Middle_Name Last_Name COUNTRYSPECIFICREGHITDISCRIPTION OTHERREGULATORYREDFLAGSDESCRIPTION"), true, true);
                       // doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", string.Concat(regredflagommentmodel, "\nCOUNTRYSPECIFICREGHITDISCRIPTION OTHERREGULATORYREDFLAGSDESCRIPTION"), true, true);
                    }
                    doc.SaveToFile(savePath);
                    string blnredresultfound = "";
                    doc.Replace("mortgage and securities industries", "mortgage[addingFootnote] and securities industries", true, true);
                    doc.SaveToFile(savePath);
                    try
                    {
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
                                        if (abc.ToString().EndsWith("mortgage[addingFootnote] and securities industries,"))
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
                                            if (diligenceInput.otherdetails.RegulatoryFlag.Equals("Yes"))
                                            {
                                                strredflagtextappended = " and the following information was identified in connection with [LastName]:   <Investigator to insert results here>";
                                            }
                                            else
                                            {
                                                strredflagtextappended = " and it is noted that [LastName] was not identified in any of these records.";
                                            }
                                            strredflagtextappended = strredflagtextappended.Replace("[LastName]",string.Concat(diligenceInput.diligenceInputModel.FirstName," ", diligenceInput.diligenceInputModel.LastName));
                                            TextRange tr = para.AppendText(strredflagtextappended);
                                            tr.CharacterFormat.FontName = "Calibri (Body)";
                                            tr.CharacterFormat.FontSize = 11;
                                            doc.SaveToFile(savePath);
                                            doc.Replace("mortgage[addingFootnote] and securities industries", "mortgage[addedFootnote] and securities industries", true, true);
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
                    }
                    catch { }

                    //EDUCATIONALANDLICENSINGHITS
                    if (diligenceInput.educationModels[0].Edu_History.ToString().Equals("No"))
                    {
                        if (diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse") || diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                        {
                            doc.Replace("EDUCATIONALANDLICENSINGHITS", "Additionally, efforts included verification of the subject’s licensing credentials, where available.[DESCRESULTPENDINGCRIMM]", true, true);
                        }
                        else
                        {
                            doc.Replace("EDUCATIONALANDLICENSINGHITS", "", true, true);
                        }
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse") || diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                        {
                            doc.Replace("EDUCATIONALANDLICENSINGHITS", "Additionally, efforts included verification of the subject’s educational and licensing credentials, where available.[DESCRESULTPENDINGCRIMM]", true, true);
                        }
                        else
                        {
                            doc.Replace("EDUCATIONALANDLICENSINGHITS", "Additionally, efforts included verification of the subject’s educational credentials, where available.[DESCRESULTPENDINGCRIMM]", true, true);
                        }
                    }
                    doc.SaveToFile(savePath);
                    if (diligenceInput.summarymodel.bankruptcy_filings.ToString().Equals("Results Pending") || diligenceInput.summarymodel.civil_court_Litigation.ToString().Equals("Results Pending") || diligenceInput.summarymodel.criminal_records.ToString().Equals("Results Pending"))
                    {
                        doc.Replace("[DESCRESULTPENDINGCRIMM]", "\n\n<Search type> searches are currently ongoing in [Country], the results of which will be provided under separate cover upon receipt.", true, false);
                    }
                    else
                    {
                        doc.Replace("[DESCRESULTPENDINGCRIMM]", "", true, false);
                    }
                    doc.SaveToFile(savePath);
                    
                    //Global security hits
                    string globalhit_comment;
                    CommentModel globalhit_comment2 = _context.DbComment
                                      .Where(u => u.Comment_type == "Global_Security")
                                      .FirstOrDefault();
                    if (diligenceInput.otherdetails.Global_Security_Hits.ToString().Equals("Yes"))
                    {
                        globalhit_comment = globalhit_comment2.confirmed_comment.ToString();
                        globalhit_comment = string.Concat(globalhit_comment, "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n");
                    }
                    else
                    {
                        globalhit_comment = globalhit_comment2.unconfirmed_comment.ToString();
                        globalhit_comment = string.Concat(globalhit_comment, "\n");
                    }
                    doc.Replace("GLOBALSECURITYHITSDESCRIPTION", globalhit_comment, true, true);
                    doc.SaveToFile(savePath);
                    //Global_Sec_Family_Hits
                    if (diligenceInput.otherdetails.Global_Sec_Family_Hits.ToString().Equals("Yes"))
                    {
                        doc.Replace("GLOBALSECFAMILYHITSDESCRIPTION", "Searches of the same sources were also conducted in connection with the subject’s <family members> (utilizing their names and provided in their Application Forms) and the following records were identified:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n", true, true);
                    }
                    else
                    {
                        doc.Replace("GLOBALSECFAMILYHITSDESCRIPTION", "Searches of the same sources were also conducted in connection with the subject’s <family members> (utilizing their names and provided in their Application Forms) and no such records were identified.\n", true, true);
                    }
                    doc.SaveToFile(savePath);
                    //PEP_Hits
                    if (diligenceInput.otherdetails.PEP_Hits.ToString().Equals("Yes"))
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
                    if (diligenceInput.otherdetails.ICIJ_Hits.ToString().Equals("Yes"))
                    {
                        doc.Replace("INCOJOCOMMENT", "<investigator to insert summary here>", true, true);
                        doc.Replace("INCOJORESULT", "Records", true, true);
                        icij_comment = icij_comment2.confirmed_comment.ToString();
                        icij_comment = string.Concat(icij_comment, "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n");
                    }
                    else
                    {
                        doc.Replace("INCOJOCOMMENT", "", true, true);
                        doc.Replace("INCOJORESULT", "Clear", true, true);
                        icij_comment = icij_comment2.unconfirmed_comment.ToString();
                        icij_comment = string.Concat(icij_comment, "\n");
                    }
                    doc.Replace("ICIJHITSDESCRIPTION", icij_comment, true, true);
                    doc.SaveToFile(savePath);
                    //Press and media
                    string PMCommentModel = "";
                    CommentModel PMcommentModel1 = _context.DbComment
                                      .Where(u => u.Comment_type == "Press_Media_CommonName")
                                      .FirstOrDefault();
                    CommentModel PMcommentModel2 = _context.DbComment
                                      .Where(u => u.Comment_type == "Press_Media_HighVolume")
                                      .FirstOrDefault();
                    CommentModel PMcommentModel3 = _context.DbComment
                                      .Where(u => u.Comment_type == "Press_Media_StandardSearch")
                                      .FirstOrDefault();
                    switch (diligenceInput.otherdetails.Press_Media.ToString())
                    {
                        case "Common name with adverse Hits":
                            PMCommentModel = PMcommentModel1.confirmed_comment.ToString();
                            PMCommentModel = string.Concat(PMCommentModel, "   <Investigator to insert article summaries here>");
                            break;
                        case "Common name without adverse Hits":
                            PMCommentModel = PMcommentModel1.unconfirmed_comment.ToString();
                            break;
                        case "High volume with adverse Hits":
                            PMCommentModel = PMcommentModel2.confirmed_comment.ToString();
                            PMCommentModel = string.Concat(PMCommentModel, "   <Investigator to insert article summaries here>");
                            break;
                        case "High volume without adverse Hits":
                            PMCommentModel = PMcommentModel2.unconfirmed_comment.ToString();
                            break;
                        case "Standard search with adverse Hits":
                            PMCommentModel = PMcommentModel3.confirmed_comment.ToString();
                            PMCommentModel = string.Concat(PMCommentModel, "   <Investigator to insert article summaries here>");
                            break;
                        case "Standard search without adverse Hits":
                            PMCommentModel = PMcommentModel3.unconfirmed_comment.ToString();
                            break;
                        case "No Hits":
                            PMCommentModel = PMcommentModel1.other_comment.ToString();
                            break;
                    }
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", PMCommentModel.ToString(), true, true);

                    //Country_legal_desc
                    string mediacommentmodel = "";
                    CommentModel mediacommentmodel1 = _context.DbComment
                                        .Where(u => u.Comment_type == "Media_Based_Legal_Record")
                                        .FirstOrDefault();
                    string uslegcommentmodel = "";
                    string strcountry = "";
                    string strlegalrechitdesc = "";
                    if (diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom"))
                    {
                        strcountry = "The United Kingdom";
                    }
                    else
                    {
                        strcountry = diligenceInput.diligenceInputModel.Country.ToString();
                    }
                    if (diligenceInput.summarymodel.bankruptcy_filings.ToString().Equals("Results Pending") || diligenceInput.summarymodel.civil_court_Litigation.ToString().Equals("Results Pending") || diligenceInput.summarymodel.criminal_records.ToString().Equals("Results Pending"))
                    {
                        doc.Replace("[CountryHEADER]", string.Concat(strcountry, "\n\n<Search type> searches are currently ongoing through <source> in [Country], the results of which will be provided under separate cover upon receipt."), true, true);
                    }
                    else
                    {
                        if (diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom")) { }
                        else
                        {
                            doc.Replace("[CountryHEADER]", strcountry, true, true);
                        }
                    }
                    doc.SaveToFile(savePath);
                    if (diligenceInput.summarymodel.criminal_records.ToString().Equals("No Consent"))
                    {
                        if (diligenceInput.summarymodel.bankruptcy_filings.ToString().StartsWith("Record") || diligenceInput.summarymodel.civil_court_Litigation.ToString().StartsWith("Record"))
                        {
                            strlegalrechitdesc = "Yes + Consent Needed for Crim";
                        }
                        else
                        {
                            strlegalrechitdesc = "No + Consent Needed for Crim";
                        }
                    }
                    else
                    {
                        if (diligenceInput.summarymodel.criminal_records.ToString().Equals("N/A"))
                        {

                            if (diligenceInput.summarymodel.bankruptcy_filings.ToString().StartsWith("Record") || diligenceInput.summarymodel.civil_court_Litigation.ToString().StartsWith("Record"))
                            {
                                strlegalrechitdesc = "Yes + Crim Not Available in Country";
                            }
                            else
                            {
                                strlegalrechitdesc = "No + Crim Not Available in Country";
                            }
                        }
                        else
                        {
                            if (diligenceInput.summarymodel.bankruptcy_filings.ToString().StartsWith("Record") || diligenceInput.summarymodel.civil_court_Litigation.ToString().StartsWith("Record") || diligenceInput.summarymodel.criminal_records.ToString().StartsWith("Record"))
                            {
                                strlegalrechitdesc = "Yes";
                            }
                            else
                            {
                                strlegalrechitdesc = "No";
                            }
                        }
                    }
                    if (diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom"))
                    {
                        doc.Replace("[CountryHEADER]", "", true, true);
                    }
                    else
                    {
                        switch (strlegalrechitdesc)
                        {
                            case "Yes":
                                doc.Replace("[COURTDESC]", "\n\nSearches of all available legal records in [Country], which include the [Courts], identified the subject, personally, as a party to any bankruptcy filings, civil litigation matters, judgments and/or criminal convictions:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n[COURTDESC]", true, false);
                                break;
                            case "No":
                                doc.Replace("[COURTDESC]", "\n\nSearches of all available legal records in [Country], which include the [Courts], did not identify the subject, personally, as a party to any bankruptcy filings, civil litigation matters, judgments or criminal convictions.\n[COURTDESC]", true, false);
                                break;
                            case "Yes + Consent Needed for Crim":
                                doc.Replace("[COURTDESC]", "\n\nWhile the subject’s express written authorization would be required in order to conduct a criminal records search in [Country], searches of all available legal records in [Country], which include the [Courts], identified the subject, personally, as a party to any bankruptcy filings, civil litigation matters and/or judgments:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n[COURTDESC]", true, false);
                                break;
                            case "No + Consent Needed for Crim":
                                doc.Replace("[COURTDESC]", "\n\nWhile the subject’s express written authorization would be required in order to conduct a criminal records search in [Country], searches of all available legal records in [Country], which include the [Courts], did not identify the subject, personally, as a party to any bankruptcy filings, civil litigation matters or judgments.\n[COURTDESC]", true, false);
                                break;
                            case "Yes + Crim Not Available in Country":
                                doc.Replace("[COURTDESC]", "\n\nWhile third-party access to criminal records information is restricted in [Country], searches of all available legal records in [Country], which include the [Courts], identified the subject, personally, as a party to any bankruptcy filings, civil litigation matters and/or judgments:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n[COURTDESC]", true, false);
                                break;
                            case "No + Crim Not Available in Country":
                                doc.Replace("[COURTDESC]", "\n\nWhile third-party access to criminal records information is restricted in [Country], searches of all available legal records in [Country], which include the [Courts], did not identify the subject, personally, as a party to any bankruptcy filings, civil litigation matters or judgments.\n[COURTDESC]", true, false);
                                break;
                        }
                    }
                    doc.SaveToFile(savePath);
                    string legrechitcommon = "";
                    if (diligenceInput.summarymodel.bankruptcy_filings1 == true)
                    {
                        legrechitcommon = "\nIt should be noted that one or more individuals known only as [FirstName] [LastName] were identified as <party type> in at least <number> bankruptcy filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n";
                    }
                    if (diligenceInput.summarymodel.civil_court_Litigation1 == true)
                    {
                        legrechitcommon = string.Concat(legrechitcommon, "\nIt should be noted that one or more individuals known only as [FirstName] [LastName] were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
                    }
                    if (diligenceInput.summarymodel.criminal_records1 == true)
                    {
                        legrechitcommon = string.Concat(legrechitcommon, "\nIt should be noted that one or more individuals known only as [FirstName] [LastName] were identified as <party type> in at least <number> criminal records, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
                    }
                    //Criminal Clearance Certificate Provided?
                    string crim_certificate = "";
                    string strconcatcrimsumm = "";
                    if (diligenceInput.otherdetails.Crim_Clearance_Certifi.ToString().Equals("Yes – Authenticated"))
                    {
                        crim_certificate = string.Concat("\n", diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName, " provided a copy of their <jurisdiction> Criminal Clearance Certificate, number <certificate #> issued on <date> by <issuing authority>, which was authenticated by <entity>, and reported the following criminal records in connection with the subject: \n\n\t•	<Investigator to insert results here>\n");
                        strconcatcrimsumm = string.Concat("CRIMINALCOMMENT\n\n", diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName, " provided a copy of their [Country] Criminal Clearance Certificate, which was authenticated");
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Crim_Clearance_Certifi.ToString().Equals("Yes - Not Authenticated"))
                        {
                            crim_certificate = string.Concat("\n",diligenceInput.diligenceInputModel.FirstName," ", diligenceInput.diligenceInputModel.LastName, " provided a copy of their <jurisdiction> Criminal Clearance Certificate, number <certificate #> issued on <date> by <issuing authority>, which was authenticated by <entity>, and did not report any criminal records in connection with the subject.\n");
                            strconcatcrimsumm = string.Concat("CRIMINALCOMMENT\n\n", diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName, " provided a copy of their [Country] Criminal Clearance Certificate, however, the same could not be authenticated <investigator to provide clarification>");
                        }
                        else
                        {
                            crim_certificate = "\nIt is noted that the subject did not provide a Criminal Clearance Certificate for [Country], and a copy of the same would be required in order to confirm the authenticity of this document.\n";
                            strconcatcrimsumm = "CRIMINALCOMMENT\n\nIt is noted that the subject did not provide a Criminal Clearance Certificate for [Country], and a copy of the same would be required in order to confirm the authenticity of this document";
                        }
                    }
                    doc.Replace("CRIMINALCOMMENT", strconcatcrimsumm, true, true);
                    string criminalresult = "";
                    string criminalcomment = "";
                    switch (diligenceInput.summarymodel.criminal_records.ToString())
                    {
                        case "Clear":
                            criminalcomment = "No criminal records were identified in connection with LastName";
                            criminalcomment = criminalcomment.Replace("LastName", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName));
                            criminalresult = "Clear";
                            break;
                        case "Records":
                            criminalcomment = "LastName was identified as a Defendant in connection with at least <number> criminal records in [Country] , filed between <date> and <date>, which pertain to <type of charge> and are currently <status>";
                            criminalcomment = criminalcomment.Replace("LastName", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName));
                            criminalresult = "Records";
                            break;
                        case "N/A":
                            criminalcomment = "It is noted that the subject’s express written authorization would be required to conduct criminal searches in [Country]";
                            criminalresult = "N/A";
                            break;
                        case "Record":
                            criminalcomment = "[LastName] was identified as a Defendant in connection with a criminal record in [Country], which was filed in <date>, and pertain to <type of charge> and is <status>";
                            criminalresult = "Record";
                            break;
                        case "Results Pending":
                            criminalcomment = "Criminal records searches are currently pending through the <investigator to insert source/court/law enforcement agency>, the results of which will be provided under separate cover upon receipt";
                            criminalresult = "Results Pending";
                            break;
                        case "No Consent":
                            criminalcomment = "The subject’s express written authorization would be required to conduct criminal records research in [Country]";
                            criminalresult = "No Consent";
                            break;
                    }
                    doc.Replace("CRIMINALCOMMENT", criminalcomment, true, true);
                    doc.Replace("CRIMINALRESULT", criminalresult, true, true);
                    //Court description
                    CommentModel uslegcommentmodel1 = _context.DbComment
                                        .Where(u => u.Comment_type == "U.S_Legal_Record")
                                        .FirstOrDefault();
                    if (diligenceInput.diligenceInputModel.Country.ToString().Equals("Australia") || diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom") || diligenceInput.diligenceInputModel.Country.ToString().Equals("Canada") || diligenceInput.diligenceInputModel.Country.ToString().Equals("France"))
                    {
                        string strlegrechit = "";
                        string strinsolvency = "";
                        string strcivil = "";
                        string strcriminal = "";
                        string strregtrust = "";
                        string strppr = "";
                        string strcourtdesc = "";
                        switch (diligenceInput.diligenceInputModel.Country.ToString())
                        {
                            case "France":
                                string fr_lrhcommentmodel = "";
                                fr_lrhcommentmodel = legrechitcommon;
                                strlegrechit = string.Concat(fr_lrhcommentmodel, "\n[MEDIABASEDLEGALDESCRIPTION]");
                                doc.Replace("[COURTDESC]", strlegrechit, true, true);
                                break;
                            case "United Kingdom":
                                string uk_ihcommentmodel = "";
                                string uk_crhcommentmodel = "";
                                string uk_rthcommentmodel = "";
                                CommentModel uk_ihcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "uk_insolvencyhit")
                                       .FirstOrDefault();
                                CommentModel uk_crhcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "uk_civilrecordhit")
                                       .FirstOrDefault();
                                CommentModel uk_rthcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "uk_registrytrusthit")
                                       .FirstOrDefault();
                                uk_ihcommentmodel = "Civil Litigation, Insolvency and Other Filings in the Country[HEADER]\n\nRecords maintained by the Insolvency Service in the United Kingdom[INSERTFOOTNOTE]\n";
                                strinsolvency = uk_ihcommentmodel;
                                switch (diligenceInput.csModel.civil_record_hits.ToString())
                                {
                                    case "Yes":
                                        uk_crhcommentmodel = uk_crhcommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No":
                                        uk_crhcommentmodel = uk_crhcommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                }
                                uk_crhcommentmodel = uk_crhcommentmodel.Replace("*n ", "\n");
                                uk_crhcommentmodel = uk_crhcommentmodel.Replace("*n", "\n");
                                uk_crhcommentmodel = uk_crhcommentmodel.Replace("*t", "\t");
                                strcivil = string.Concat("\n", uk_crhcommentmodel, "\n");
                                switch (diligenceInput.csModel.reg_trust_hits.ToString())
                                {
                                    case "N/A":
                                        uk_rthcommentmodel = "Further, the subject’s residential address history information within the past six years would be required in order to conduct a search through the Registry of Judgments, Orders and Fines.";
                                        break;
                                    case "Yes":
                                        uk_rthcommentmodel = uk_rthcommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No":
                                        uk_rthcommentmodel = uk_rthcommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                }
                                uk_rthcommentmodel = uk_rthcommentmodel.Replace("*n ", "\n");
                                uk_rthcommentmodel = uk_rthcommentmodel.Replace("*n", "\n");
                                uk_rthcommentmodel = uk_rthcommentmodel.Replace("*t", "\t");
                                strregtrust = string.Concat("\n", uk_rthcommentmodel, "\n");
                                string uk_ecrcommentmodel = "";
                                string uk_appendtext = "";
                                string uk_footnotecomment = "Information available in the Basic Disclosure is limited to convictions considered “unspent” under The Rehabilitation of Offenders Act 1974, meaning that the applicable rehabilitation period has not expired.  A custodial sentence of more than two and a half years will never become spent.";
                                CommentModel uk_ecrcommentmodel1 = _context.DbComment
                                      .Where(u => u.Comment_type == "uk_engcriminalrecord")
                                      .FirstOrDefault();
                                CommentModel uk_scrcommentmodel1 = _context.DbComment
                                      .Where(u => u.Comment_type == "uk_scotlandcriminalrecord")
                                      .FirstOrDefault();
                                CommentModel uk_ncrcommentmodel1 = _context.DbComment
                                      .Where(u => u.Comment_type == "uk_Northencriminalrecord")
                                      .FirstOrDefault();
                                CommentModel uk_manualcommentmodel1 = _context.DbComment
                                      .Where(u => u.Comment_type == "uk_manualcriminalrecord")
                                      .FirstOrDefault();

                                switch (diligenceInput.csModel.criminal_record_hits.ToString())
                                {
                                    case "England, Wales, the Channel Islands and the Isle of Man: Basic Disclosure without Hits":
                                        uk_ecrcommentmodel = uk_ecrcommentmodel1.unconfirmed_comment.ToString();
                                        uk_appendtext = "Research efforts conducted by the Disclosure and Barring Service, as well as information held by local police forces, did not identify any such records in connection with [LastName].";
                                        break;
                                    case "England, Wales, the Channel Islands and the Isle of Man: Basic Disclosure with Hits":
                                        uk_ecrcommentmodel = uk_ecrcommentmodel1.confirmed_comment.ToString();
                                        uk_appendtext = "Research efforts conducted by the Disclosure and Barring Service, as well as information held by local police forces, identified the following records in connection with [LastName]:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                                        break;
                                    case "Scotland: Basic Disclosure without Hits":
                                        uk_ecrcommentmodel = uk_scrcommentmodel1.unconfirmed_comment.ToString();
                                        uk_appendtext = "Research efforts conducted by Disclosure Scotland through the Police National Computer, as well as information held by local police forces, did not identify any such records in connection with [LastName].";
                                        break;
                                    case "Scotland: Basic Disclosure with Hits":
                                        uk_ecrcommentmodel = uk_scrcommentmodel1.confirmed_comment.ToString();
                                        uk_appendtext = "Research efforts conducted by Disclosure Scotland through the Police National Computer, as well as information held by local police forces, identified the following records in connection with [LastName]:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                                        break;
                                    case "Northern Ireland: Basic Disclosure without Hits":
                                        uk_ecrcommentmodel = uk_ncrcommentmodel1.unconfirmed_comment.ToString();
                                        uk_appendtext = "Research efforts conducted by NIDirect, as well as information held by local police forces, did not identify any such records in connection with [LastName].";
                                        break;
                                    case "Northern Ireland: Basic Disclosure with Hits":
                                        uk_ecrcommentmodel = uk_ncrcommentmodel1.confirmed_comment.ToString();
                                        uk_appendtext = "Research efforts conducted by NIDirect, as well as information held by local police forces, identified the following records in connection with [LastName]:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                                        break;
                                    case "Manual Local without Hits":
                                        uk_ecrcommentmodel = "Pursuant to [LastName]’s express written authorization, criminal court records searches were conducted through the Crown, High and Magistrate Courts associated with the subject’s address.  In addition, research efforts were also undertaken in all levels of the roughly 300 or more United Kingdom courts to ensure that no other records were identified in connection with the subject's linked address history, which did not locate any references to criminal records involving [LastName].";
                                        break;
                                    case "Manual Local with Hits":
                                        uk_ecrcommentmodel = "Pursuant to [LastName]’s express written authorization, criminal court records searches were conducted through the Crown, High and Magistrate Courts associated with the subject’s address.  In addition, research efforts were also undertaken in all levels of the roughly 300 or more United Kingdom courts to ensure that no other records were identified in connection with the subject's linked address history, which located the following references to criminal records involving [LastName]:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                                        break;
                                    case "No Consent":
                                        uk_ecrcommentmodel = uk_ecrcommentmodel1.other_comment.ToString();
                                        break;
                                }
                                uk_ecrcommentmodel = uk_ecrcommentmodel.Replace("*n ", "\n");
                                uk_ecrcommentmodel = uk_ecrcommentmodel.Replace("*n", "\n");
                                uk_ecrcommentmodel = uk_ecrcommentmodel.Replace("*t", "\t");
                                strcriminal = string.Concat("\n", "Criminal Records[CRIMHEADER]\n\n", uk_ecrcommentmodel);
                                strcourtdesc = string.Concat(legrechitcommon, strinsolvency, strcivil, strregtrust, strcriminal, "\n");
                                doc.Replace("[COURTDESC]", strcourtdesc, true, false);
                                doc.SaveToFile(savePath);
                                //Criminal Footnote
                                string blninvresultfound = "";
                                if (diligenceInput.csModel.criminal_record_hits.ToString().Equals("England, Wales, the Channel Islands and the Isle of Man: Basic Disclosure without Hits") || diligenceInput.csModel.criminal_record_hits.ToString().Equals("England, Wales, the Channel Islands and the Isle of Man: Basic Disclosure with Hits") || diligenceInput.csModel.criminal_record_hits.ToString().Equals("Scotland: Basic Disclosure without Hits") || diligenceInput.csModel.criminal_record_hits.ToString().Equals("Scotland: Basic Disclosure with Hits") || diligenceInput.csModel.criminal_record_hits.ToString().Equals("Northern Ireland: Basic Disclosure without Hits") || diligenceInput.csModel.criminal_record_hits.ToString().Equals("Northern Ireland: Basic Disclosure with Hits"))
                                {
                                    try
                                    {
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
                                                        // Find the word "Civil" in paragraph1
                                                        if (abc.ToString().EndsWith("Basic Disclosures for individuals throughout England, Wales, the Channel Islands and the Isle of Man.") || abc.ToString().EndsWith("Basic Disclosures for individuals throughout Scotland.") || abc.ToString().EndsWith("Basic Disclosures for individuals throughout Northern Ireland."))
                                                        {
                                                            Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                            footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                            footnote1.MarkerCharacterFormat.FontSize = 11;
                                                            //Insert footnote1 after the word "Spire.Doc"
                                                            para.ChildObjects.Insert(i + 1, footnote1);
                                                            TextRange _text = footnote1.TextBody.AddParagraph().AppendText(uk_footnotecomment);
                                                            footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                            _text.CharacterFormat.FontName = "Calibri (Body)";
                                                            _text.CharacterFormat.FontSize = 9;
                                                            //Append the line                                             
                                                            uk_appendtext = string.Concat(uk_appendtext, "\n[MEDIABASEDLEGALDESCRIPTION]");
                                                            TextRange tr = para.AppendText(uk_appendtext);
                                                            tr.CharacterFormat.FontName = "Calibri (Body)";
                                                            tr.CharacterFormat.FontSize = 11;
                                                            doc.SaveToFile(savePath);
                                                            blninvresultfound = "true";
                                                            break;
                                                        }
                                                    }
                                                }
                                                if (blninvresultfound == "true")
                                                {
                                                    break;
                                                }
                                            }
                                            if (blninvresultfound == "true")
                                            {
                                                break;
                                            }
                                        }
                                    }
                                    catch { }
                                }
                                //Insolvency footnote
                                blninvresultfound = "";
                                try
                                {
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
                                                    if (abc.ToString().EndsWith("[INSERTFOOTNOTE]"))
                                                    {
                                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                                        //Insert footnote1 after the word "Spire.Doc"
                                                        para.ChildObjects.Insert(i + 1, footnote1);
                                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("The Insolvency Service maintains details that are either current or have terminated within the past three months or so, as well as current individual voluntary arrangements, fast track voluntary arrangements and current bankruptcy restrictions orders and undertakings.  Further, a search was also conducted in connection with bankruptcy orders and applications made within the past five years.");
                                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                                        _text.CharacterFormat.FontSize = 9;
                                                        //Append the line 
                                                        uk_ihcommentmodel = "";
                                                        if (diligenceInput.csModel.insolvency_hits.ToString().Equals("Yes"))
                                                        {
                                                            uk_ihcommentmodel = "identified the following bankruptcy filings in connection with [LastName]:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                                                        }
                                                        else
                                                        {
                                                            uk_ihcommentmodel = "did not identify any bankruptcy filings in connection with [LastName], personally.";
                                                        }
                                                        TextRange tr = para.AppendText(uk_ihcommentmodel);
                                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                                        tr.CharacterFormat.FontSize = 11;
                                                        doc.SaveToFile(savePath);
                                                        doc.Replace("[INSERTFOOTNOTE]", "", true, false);
                                                        doc.SaveToFile(savePath);
                                                        blninvresultfound = "true";
                                                        break;
                                                    }
                                                }
                                            }
                                            if (blninvresultfound == "true")
                                            {
                                                break;
                                            }
                                        }
                                        if (blninvresultfound == "true")
                                        {
                                            break;
                                        }
                                    }
                                }
                                catch { }
                                //Civil footnote
                                blninvresultfound = "";
                                try
                                {
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
                                                    if (abc.ToString().EndsWith("litigation and/or judgments registered against [LastName]") || abc.ToString().Contains("litigation and/or judgments registered against [LastName]"))
                                                    {
                                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                                        //Insert footnote1 after the word "Spire.Doc"
                                                        para.ChildObjects.Insert(i + 1, footnote1);
                                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("Research was conducted through court decisions of the following England and Wales Courts: Court of Appeal (Civil and Criminal Divisions); High Court (Administrative, Commercial, Exchequer, Mercantile, Patents, and Technology and Construction Courts, as well as Admiralty, Chancery, Family, King’s Bench and Queen’s Bench Divisions); Patents County Court; and Care Standards and Lands Tribunals.  Research of court decisions was also conducted through the following the United Kingdom Courts: House of Lords; Financial Services and Markets Tribunals; Privy Council; Special Commissioners of Income Tax; Social Security and Child Support Commissioners’ Tribunal; VAT & Duties Tribunals (Customs, Excise, Insurance Premium Tax and Landfill Tax); Competition Appeals Tribunal; Special Immigrations Appeals Commission; Employment Appeal Tribunal; Asylum and Immigration Tribunal; and Information Tribunal including the National Security Appeals Panel.");
                                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                                        _text.CharacterFormat.FontSize = 9;
                                                        blninvresultfound = "true";
                                                        break;
                                                    }

                                                }
                                            }
                                            if (blninvresultfound == "true")
                                            {
                                                break;
                                            }
                                        }
                                        if (blninvresultfound == "true")
                                        {
                                            break;
                                        }
                                    }
                                }
                                catch { }
                                doc.Replace("[INSERTFOOTNOTE]", "", true, false);
                                break;
                            case "Canada":
                                string can_Inhitcommentmodel = "";
                                string can_civilcommentmodel = "";
                                string can_pprcommentmodel = "";
                                string can_cricommentmodel = "";
                                CommentModel can_Inhitcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "can_insolvencyhit")
                                       .FirstOrDefault();
                                CommentModel can_civilcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "can_civilrechit")
                                       .FirstOrDefault();
                                CommentModel can_pprcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "can_pprrechit")
                                       .FirstOrDefault();
                                CommentModel can_cricommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "can_criminalrecord")
                                       .FirstOrDefault();
                                switch (diligenceInput.csModel.insolvency_hits.ToString())
                                {
                                    case "Yes":
                                        can_Inhitcommentmodel = can_Inhitcommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No":
                                        can_Inhitcommentmodel = can_Inhitcommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                }
                                can_Inhitcommentmodel = can_Inhitcommentmodel.Replace("*n ", "\n");
                                can_Inhitcommentmodel = can_Inhitcommentmodel.Replace("*n", "\n");
                                can_Inhitcommentmodel = can_Inhitcommentmodel.Replace("*t", "\t");
                                strinsolvency = string.Concat(can_Inhitcommentmodel, "\n");
                                switch (diligenceInput.csModel.civil_record_hits.ToString())
                                {
                                    case "Yes":
                                        can_civilcommentmodel = can_civilcommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No":
                                        can_civilcommentmodel = can_civilcommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                }
                                can_civilcommentmodel = can_civilcommentmodel.Replace("*n ", "\n");
                                can_civilcommentmodel = can_civilcommentmodel.Replace("*n", "\n");
                                can_civilcommentmodel = can_civilcommentmodel.Replace("*t", "\t");
                                strcivil = string.Concat("\n", can_civilcommentmodel, "\n");
                                switch (diligenceInput.csModel.criminal_record_hits.ToString())
                                {
                                    case "Criminal Search without Hits":
                                        can_cricommentmodel = can_cricommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                    case "Criminal Search with Hits":
                                        can_cricommentmodel = can_cricommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No Consent":
                                        can_cricommentmodel = can_cricommentmodel1.other_comment.ToString();
                                        break;
                                }
                                can_cricommentmodel = can_cricommentmodel.Replace("*n ", "\n");
                                can_cricommentmodel = can_cricommentmodel.Replace("*n", "\n");
                                can_cricommentmodel = can_cricommentmodel.Replace("*t", "\t");
                                strcriminal = string.Concat("\n", can_cricommentmodel, "\n");
                                switch (diligenceInput.csModel.pprrecordhits.ToString())
                                {
                                    case "Yes":
                                        can_pprcommentmodel = can_pprcommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No":
                                        can_pprcommentmodel = can_pprcommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                }
                                can_pprcommentmodel = can_pprcommentmodel.Replace("*n ", "\n");
                                can_pprcommentmodel = can_pprcommentmodel.Replace("*n", "\n");
                                can_pprcommentmodel = can_pprcommentmodel.Replace("*t", "\t");
                                strppr = string.Concat("\n", can_pprcommentmodel, "\n");
                                strcourtdesc = string.Concat(legrechitcommon, strinsolvency, strcivil, strppr, "\nCriminal Records[CRIMHEADER]\n", strcriminal, "[MEDIABASEDLEGALDESCRIPTION]");
                                doc.Replace("[COURTDESC]", strcourtdesc, true, false);
                                doc.SaveToFile(savePath);
                                break;
                            case "Australia":
                                string aus_inhitcommentmodel = "";
                                string aus_civilcommentmodel = "";
                                string aus_pprcommentmodel = "";
                                string aus_cricommentmodel = "";
                                CommentModel aus_inhitcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "aus_insolvencyhit")
                                       .FirstOrDefault();
                                CommentModel aus_civilcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "aus_civilrechit")
                                       .FirstOrDefault();
                                CommentModel aus_pprcommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "aus_ppsrrechit")
                                       .FirstOrDefault();
                                CommentModel aus_cricommentmodel1 = _context.DbComment
                                       .Where(u => u.Comment_type == "aus_criminalrecord")
                                       .FirstOrDefault();
                                strinsolvency = "\nA search of the Australian National Bankruptcy Registrar\n";
                                strcivil = "Additionally, searches conducted on all available legal records in Australia, including the Federal Court, District Courts, County Courts, Magistrate Courts and Tribunals,\n";
                                doc.SaveToFile(savePath);
                                string blnausinvresultfound = "";
                                switch (diligenceInput.csModel.criminal_record_hits.ToString())
                                {
                                    case "Criminal Search without Hits":
                                        aus_cricommentmodel = aus_cricommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                    case "Criminal Search with Hits":
                                        aus_cricommentmodel = aus_cricommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No Consent":
                                        aus_cricommentmodel = aus_cricommentmodel1.other_comment.ToString();
                                        break;
                                }
                                aus_cricommentmodel = aus_cricommentmodel.Replace("*n ", "\n");
                                aus_cricommentmodel = aus_cricommentmodel.Replace("*n", "\n");
                                aus_cricommentmodel = aus_cricommentmodel.Replace("*t", "\t");
                                strcriminal = string.Concat("\n", aus_cricommentmodel, "\n");
                                switch (diligenceInput.csModel.pprrecordhits.ToString())
                                {
                                    case "Yes":
                                        aus_pprcommentmodel = aus_pprcommentmodel1.confirmed_comment.ToString();
                                        break;
                                    case "No":
                                        aus_pprcommentmodel = aus_pprcommentmodel1.unconfirmed_comment.ToString();
                                        break;
                                }
                                aus_pprcommentmodel = aus_pprcommentmodel.Replace("*n ", "\n");
                                aus_pprcommentmodel = aus_pprcommentmodel.Replace("*n", "\n");
                                aus_pprcommentmodel = aus_pprcommentmodel.Replace("*t", "\t");
                                strppr = string.Concat(aus_pprcommentmodel, "\n");
                                strcourtdesc = string.Concat(legrechitcommon, strinsolvency, strcivil, strppr, "\nCriminal Records[CRIMHEADER]\n", strcriminal, "[MEDIABASEDLEGALDESCRIPTION]");
                                doc.Replace("[COURTDESC]", strcourtdesc, true, false);
                                doc.SaveToFile(savePath);
                                try
                                {
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
                                                    // Find the word "Civil" in paragraph1
                                                    if (abc.ToString().EndsWith("A search of the Australian National Bankruptcy Registrar"))
                                                    {
                                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                                        //Insert footnote1 after the word "Spire.Doc"
                                                        para.ChildObjects.Insert(i + 1, footnote1);
                                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("Searches were conducted of the National Personal Insolvency Index (“NPII”) through the Australian Financial Security Authority (“AFSA”).");
                                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                                        _text.CharacterFormat.FontSize = 9;
                                                        //Append the line 
                                                        switch (diligenceInput.csModel.insolvency_hits.ToString())
                                                        {
                                                            case "Yes":
                                                                aus_inhitcommentmodel = aus_inhitcommentmodel1.confirmed_comment.ToString();
                                                                break;
                                                            case "No":
                                                                aus_inhitcommentmodel = aus_inhitcommentmodel1.unconfirmed_comment.ToString();
                                                                break;
                                                        }
                                                        aus_inhitcommentmodel = aus_inhitcommentmodel.Replace("*n ", "\n");
                                                        aus_inhitcommentmodel = aus_inhitcommentmodel.Replace("*n", "\n");
                                                        aus_inhitcommentmodel = aus_inhitcommentmodel.Replace("*t", "\t");
                                                        aus_inhitcommentmodel = string.Concat(aus_inhitcommentmodel, "\n");
                                                        TextRange tr = para.AppendText(aus_inhitcommentmodel);
                                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                                        tr.CharacterFormat.FontSize = 11;
                                                        doc.SaveToFile(savePath);
                                                        blnausinvresultfound = "true";
                                                        break;
                                                    }
                                                }
                                            }
                                            if (blnausinvresultfound == "true")
                                            {
                                                break;
                                            }
                                        }
                                        if (blnausinvresultfound == "true")
                                        {
                                            break;
                                        }
                                    }
                                }
                                catch
                                {

                                }
                                blnausinvresultfound = "";
                                try
                                {
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
                                                    // Find the word "Civil" in paragraph1
                                                    if (abc.ToString().EndsWith("Additionally, searches conducted on all available legal records in Australia, including the Federal Court, District Courts, County Courts, Magistrate Courts and Tribunals,"))
                                                    {
                                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                                        //Insert footnote1 after the word "Spire.Doc"
                                                        para.ChildObjects.Insert(i + 1, footnote1);
                                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("Research efforts were pursued through the Queensland Online Supreme and District Court Registries, Victorian County Court Registry, Federal Courts of Australia, Federal Magistrates Courts of Australia, as well as other Australian Courts and Tribunals, including Commonwealth Case Law, Australian Capital Territory (“ACT”) Case Law, New South Wales (“NSW”) Case Law, Northern Territory Case Law, Queensland  Case Law, South Australia Case Law, Tasmania Case Law, Victoria Case Law, and Western Australia Case Law.");
                                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                                        _text.CharacterFormat.FontSize = 9;
                                                        //Append the line 
                                                        switch (diligenceInput.csModel.civil_record_hits.ToString())
                                                        {
                                                            case "Yes":
                                                                aus_civilcommentmodel = aus_civilcommentmodel1.confirmed_comment.ToString();
                                                                break;
                                                            case "No":
                                                                aus_civilcommentmodel = aus_civilcommentmodel1.unconfirmed_comment.ToString();
                                                                break;
                                                        }
                                                        aus_civilcommentmodel = aus_civilcommentmodel.Replace("*n ", "\n");
                                                        aus_civilcommentmodel = aus_civilcommentmodel.Replace("*n", "\n");
                                                        aus_civilcommentmodel = aus_civilcommentmodel.Replace("*t", "\t");
                                                        aus_civilcommentmodel = string.Concat(aus_civilcommentmodel, "\n");
                                                        TextRange tr = para.AppendText(aus_civilcommentmodel);
                                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                                        tr.CharacterFormat.FontSize = 11;
                                                        doc.SaveToFile(savePath);
                                                        blnausinvresultfound = "true";
                                                        break;
                                                    }
                                                }
                                            }
                                            if (blnausinvresultfound == "true")
                                            {
                                                break;
                                            }
                                        }
                                        if (blnausinvresultfound == "true")
                                        {
                                            break;
                                        }
                                    }
                                }
                                catch { }
                                break;
                        }
                        //Media_Based_Legal_Record              
                        if (diligenceInput.otherdetails.Media_Based_Hits.ToString().Equals("Yes"))
                        {
                            mediacommentmodel = mediacommentmodel1.confirmed_comment.ToString();
                        }
                        else
                        {
                            mediacommentmodel = mediacommentmodel1.unconfirmed_comment.ToString();
                        }
                        mediacommentmodel = mediacommentmodel.Replace("*n ", "\n");
                        mediacommentmodel = mediacommentmodel.Replace("*n", "\n");
                        mediacommentmodel = mediacommentmodel.Replace("*t", "\t");
                        mediacommentmodel = string.Concat("\n", mediacommentmodel, "\n");
                        //U.S_Legal_Record            
                        if (diligenceInput.otherdetails.USLegal_Record_Hits.ToString().Equals("Yes"))
                        {
                            uslegcommentmodel = uslegcommentmodel1.confirmed_comment.ToString();
                            uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                        }
                        else
                        {
                            if (diligenceInput.otherdetails.USLegal_Record_Hits.ToString().Equals("No"))
                            {
                                uslegcommentmodel = uslegcommentmodel1.unconfirmed_comment.ToString();
                                uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                            }
                            else
                            {
                                uslegcommentmodel = uslegcommentmodel1.other_comment.ToString();
                                uslegcommentmodel = string.Concat("\n", uslegcommentmodel, "\n");
                            }
                        }
                        uslegcommentmodel = uslegcommentmodel.Replace("*n ", "\n");
                        uslegcommentmodel = uslegcommentmodel.Replace("*n", "\n");
                        uslegcommentmodel = uslegcommentmodel.Replace("*t", "\t");
                        strcourtdesc = string.Concat(crim_certificate, mediacommentmodel, uslegcommentmodel);
                        doc.Replace("[MEDIABASEDLEGALDESCRIPTION]", strcourtdesc, true, true);
                        doc.SaveToFile(savePath);

                        if (strcivil.ToString().Equals("")) { }
                        else
                        {
                            holdresult = "";
                            try
                            {
                                for (int j = 1; j < 8; j++)
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
                                                if (abc.ToString().Equals("Civil Litigation, Insolvency and Other Filings in the Country[HEADER]"))
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
                            }
                            catch { }
                        }

                        if (strcriminal.ToString().Equals("")) { }
                        else
                        {
                            holdresult = "";
                            try
                            {
                                for (int j = 1; j < 8; j++)
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
                                                if (abc.ToString().Equals("Criminal Records[CRIMHEADER]"))
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
                            }
                            catch { }
                        }

                        doc.Replace("[HEADER]", "", true, false);
                        doc.Replace("[CRIMHEADER]", "", true, false);
                        doc.SaveToFile(savePath);
                    }
                    else
                    {
                        string strlegrechit = "";
                        string strcourtsection = "";
                        if (legrechitcommon == "") { }
                        else
                        {
                            strlegrechit = string.Concat(legrechitcommon);
                        }
                        //Media_Based_Legal_Record                               
                        if (diligenceInput.otherdetails.Media_Based_Hits.ToString().Equals("Yes"))
                        {
                            mediacommentmodel = mediacommentmodel1.confirmed_comment.ToString();
                        }
                        else
                        {
                            mediacommentmodel = mediacommentmodel1.unconfirmed_comment.ToString();
                        }
                        mediacommentmodel = mediacommentmodel.Replace("*n ", "\n");
                        mediacommentmodel = mediacommentmodel.Replace("*n", "\n");
                        mediacommentmodel = mediacommentmodel.Replace("*t", "\t");
                        mediacommentmodel = string.Concat("\n", mediacommentmodel, "\n");
                        //U.S_Legal_Record                
                        if (diligenceInput.otherdetails.USLegal_Record_Hits.ToString().Equals("Yes"))
                        {
                            uslegcommentmodel = uslegcommentmodel1.confirmed_comment.ToString();
                            uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                        }
                        else
                        {
                            if (diligenceInput.otherdetails.USLegal_Record_Hits.ToString().Equals("No"))
                            {
                                uslegcommentmodel = uslegcommentmodel1.unconfirmed_comment.ToString();
                                uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                            }
                            else
                            {
                                uslegcommentmodel = uslegcommentmodel1.other_comment.ToString();
                                uslegcommentmodel = string.Concat("\n", uslegcommentmodel, "\n");
                            }
                        }
                        uslegcommentmodel = uslegcommentmodel.Replace("*n ", "\n");
                        uslegcommentmodel = uslegcommentmodel.Replace("*n", "\n");
                        uslegcommentmodel = uslegcommentmodel.Replace("*t", "\t");
                        strcourtsection = string.Concat(strlegrechit, crim_certificate, mediacommentmodel, uslegcommentmodel);
                        doc.Replace("[COURTDESC]", strcourtsection, true, false);
                        doc.SaveToFile(savePath);
                    }
                    holdresult = "";
                    try
                    {
                        for (int j = 1; j < 8; j++)
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
                                        if (abc.ToString().Equals("Other Legal Records Searches"))
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
                    }
                    catch { }

                    CommentModel uk_dhcommentmodel1 = _context.DbComment
                                 .Where(u => u.Comment_type == "uk_driving_history")
                                 .FirstOrDefault();
                    CommentModel uk_chcommentmodel1 = _context.DbComment
                           .Where(u => u.Comment_type == "uk_credithistory")
                           .FirstOrDefault();
                    //Driving history
                    if (diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom"))
                    {
                        string strdrivingComment = "";
                        string uk_dhcommentmodel = "";
                        switch (diligenceInput.csModel.driving_hits.ToString())
                        {
                            case "Driving History without Hits":
                                strdrivingComment = "";
                                uk_dhcommentmodel = uk_dhcommentmodel1.unconfirmed_comment.ToString();
                                doc.Replace("DRIVINGRESULT", "Clear", true, true);
                                break;
                            case "Driving check with Hits":
                                strdrivingComment = "With the subject’s express written authorization, [LastName]’s driving record was retrieved in [Country], which revealed the following records: \n\n<investigator to insert summary here>";
                                uk_dhcommentmodel = uk_dhcommentmodel1.confirmed_comment.ToString();
                                doc.Replace("DRIVINGRESULT", "Records", true, true);
                                break;
                            case "No Consent":
                                strdrivingComment = "The subject’s express written authorization, driver’s license details, as well as their unique search code from the United Kingdom Driver and Vehicle Licensing Agency, would be required in order to pursue driving history records searches in the United Kingdom";
                                uk_dhcommentmodel = uk_dhcommentmodel1.other_comment.ToString();
                                doc.Replace("DRIVINGRESULT", "No Consent", true, true);
                                break;
                            case "Not Applicable":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in [Country].", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in [Country]", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                break;
                            case "Results Pending":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "With the subject’s express written authorization, efforts to retrieve [LastName]’s driving record are currently underway, the results of which will be provided under separate cover upon receipt.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "With the subject’s express written authorization, efforts to retrieve [LastName]’s driving record are currently underway, the results of which will be provided under separate cover upon receipt", true, true);
                                doc.Replace("DRIVINGRESULT", "Results Pending", true, true);
                                break;
                        }
                        strdrivingComment = strdrivingComment.Replace("[LastName]", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName));
                        doc.Replace("DRIVINGCOMMENT", strdrivingComment, true, true);
                        uk_dhcommentmodel = uk_dhcommentmodel.Replace("*n ", "\n");
                        uk_dhcommentmodel = uk_dhcommentmodel.Replace("*n", "\n");
                        uk_dhcommentmodel = uk_dhcommentmodel.Replace("*t", "\t");
                        doc.Replace("DRIVINGHISTORYDESCRIPTION", uk_dhcommentmodel, true, true);
                    }
                    else
                    {
                        switch (diligenceInput.csModel.driving_hits)
                        {
                            case "Driving History without Hits":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "With the subject’s express written authorization, a driving history check was conducted in [Country], which did not reveal any traffic infractions or license suspensions in connection with the subject.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "", true, true);
                                doc.Replace("DRIVINGRESULT", "Clear", true, true);
                                break;
                            case "Driving check with Hits":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "With the subject’s express written authorization, a driving history check was conducted in [Country], which identified the following records in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>", true, true);
                                doc.Replace("DRIVINGCOMMENT", "With the subject’s express written authorization, [LastName]’s driving record was retrieved in [Country], which revealed the following records:\n\n<investigator to insert summary here>", true, true);
                                doc.Replace("DRIVINGRESULT", "Records", true, true);
                                break;
                            case "No Consent":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "The subject’s express written authorization, as well as other information (such as driver’s license details), would be required in order to pursue driving history records searches in [Country].", true, true);
                                doc.Replace("DRIVINGCOMMENT", "The subject’s express written authorization would be required in order to conduct a driving history search in [Country]", true, true);
                                doc.Replace("DRIVINGRESULT", "No Consent", true, true);
                                break;
                            case "Not Applicable":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "Third-party access to driving history records is restricted in [Country].", true, true);
                                doc.Replace("DRIVINGCOMMENT", "Third-party access to driving history records is restricted in [Country]", true, true);
                                doc.Replace("DRIVINGRESULT", "N/A", true, true);
                                break;
                            case "Results Pending":
                                doc.Replace("DRIVINGHISTORYDESCRIPTION", "With the subject’s express written authorization, efforts to retrieve [LastName]’s driving record are currently underway, the results of which will be provided under separate cover upon receipt.", true, true);
                                doc.Replace("DRIVINGCOMMENT", "With the subject’s express written authorization, efforts to retrieve [LastName]’s driving record are currently underway, the results of which will be provided under separate cover upon receipt", true, true);
                                doc.Replace("DRIVINGRESULT", "Results Pending", true, true);
                                break;
                        }
                    }
                    doc.SaveToFile(savePath);
                    TextSelection[] text31 = doc.FindAllString("efforts to retrieve [LastName]’s driving record are currently underway, the results of which will be provided under separate cover upon receipt", false, true);
                    if (text31 != null)
                    {
                        foreach (TextSelection seletion in text31)
                        {
                            seletion.GetAsOneRange().CharacterFormat.Italic = true;
                            seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                        }
                    }
                    doc.SaveToFile(savePath);
                    //Credit history
                    if (diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom"))
                    {
                        string strcreditComment = "";
                        string uk_chcommentmodel = "";
                        switch (diligenceInput.csModel.credit_hits.ToString())
                        {
                            case "Credit with No Adverse Hits":
                                uk_chcommentmodel = uk_chcommentmodel1.unconfirmed_comment.ToString();
                                strcreditComment = "With the subject’s express written authorization, [LastName]’s personal credit history was retrieved in [Country], which revealed the following negative credit information: \n\n<investigator to insert summary here>";
                                doc.Replace("PERCREDITCOMMENT", "", true, true);
                                doc.Replace("PERCREDITRESULT", "Clear", true, true);
                                break;
                            case "Credit with Adverse Hits":
                                uk_chcommentmodel = uk_chcommentmodel1.confirmed_comment.ToString();
                                strcreditComment = "With the subject’s express written authorization, [LastName]’s personal credit history was retrieved in [Country], which revealed the following negative credit information: \n\n<investigator to insert summary here>";
                                doc.Replace("PERCREDITCOMMENT", strcreditComment, true, true);
                                doc.Replace("PERCREDITRESULT", "Records", true, true);
                                break;
                            case "No Consent":
                                uk_chcommentmodel = uk_chcommentmodel1.other_comment.ToString();
                                doc.Replace("PERCREDITCOMMENT", "The subject’s express written authorization would be required to retrieve their personal credit history records in [Country]", true, true);
                                doc.Replace("PERCREDITRESULT", "No Consent", true, true);
                                break;
                            case "Results Pending":
                                doc.Replace("CREDITHISTORYDESCRIPTION", "With the subject’s express written authorization, efforts to retrieve [LastName]’s personal credit history in [Country] are currently underway, the results of which will be provided under separate cover upon receipt.", true, true);
                                doc.Replace("PERCREDITCOMMENT", "With the subject’s express written authorization, efforts to retrieve [LastName]’s personal credit history in [Country] are currently underway, the results of which will be provided under separate cover upon receipt", true, true);
                                doc.Replace("PERCREDITRESULT", "Results Pending", true, true);
                                break;
                            case "Not Available":
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to credit history records is restricted in [Country].", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to credit history records is restricted in [Country]", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                break;
                        }
                        uk_chcommentmodel = uk_chcommentmodel.Replace("*n ", "\n");
                        uk_chcommentmodel = uk_chcommentmodel.Replace("*n", "\n");
                        uk_chcommentmodel = uk_chcommentmodel.Replace("*t", "\t");
                        doc.Replace("CREDITHISTORYDESCRIPTION", uk_chcommentmodel, true, true);
                        doc.SaveToFile(savePath);
                    }
                    else
                    {
                        switch (diligenceInput.csModel.credit_hits)
                        {
                            case "Not Available":
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Third-party access to credit history records is restricted in [Country].", true, true);
                                doc.Replace("PERCREDITCOMMENT", "Third-party access to credit history records is restricted in [Country]", true, true);
                                doc.Replace("PERCREDITRESULT", "N/A", true, true);
                                break;
                            case "No Consent":
                                doc.Replace("CREDITHISTORYDESCRIPTION", "The subject’s express written authorization would be required in order to pursue credit history records searches in [Country].", true, true);
                                doc.Replace("PERCREDITCOMMENT", "The subject’s express written authorization would be required in order to conduct a credit history search in [Country]", true, true);
                                doc.Replace("PERCREDITRESULT", "No Consent", true, true);
                                break;
                            case "Credit with Adverse Hits":
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Pursuant to [LastName]’s express authorization, a credit report was obtained, the same of which identified the following:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>", true, true);
                                doc.Replace("PERCREDITCOMMENT", "With the subject’s express written authorization, [LastName]’s personal credit history was retrieved in [Country], which revealed the following negative credit information:\n\n<investigator to insert summary here>", true, true);
                                doc.Replace("PERCREDITRESULT", "Records", true, true);
                                break;
                            case "Credit with No Adverse Hits":
                                doc.Replace("CREDITHISTORYDESCRIPTION", "Pursuant to [LastName]’s express authorization, a credit report was obtained, which did not reveal any negative payment history in connection with subject.", true, true);
                                doc.Replace("PERCREDITCOMMENT", "", true, true);
                                doc.Replace("PERCREDITRESULT", "Clear", true, true);
                                break;
                            case "Results Pending":
                                doc.Replace("CREDITHISTORYDESCRIPTION", "With the subject’s express written authorization, efforts to retrieve [LastName]’s personal credit history in [Country] are currently underway, the results of which will be provided under separate cover upon receipt.", true, true);
                                doc.Replace("PERCREDITCOMMENT", "With the subject’s express written authorization, efforts to retrieve [LastName]’s personal credit history in [Country] are currently underway, the results of which will be provided under separate cover upon receipt", true, true);
                                doc.Replace("PERCREDITRESULT", "Results Pending", true, true);
                                break;

                        }
                    }
                    doc.SaveToFile(savePath);
                    TextSelection[] text32 = doc.FindAllString("efforts to retrieve [LastName]’s personal credit history in [Country] are currently underway, the results of which will be provided under separate cover upon receipt", false, true);
                    if (text32 != null)
                    {
                        foreach (TextSelection seletion in text32)
                        {
                            seletion.GetAsOneRange().CharacterFormat.Italic = true;
                            seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                        }
                    }
                    doc.SaveToFile(savePath);
                    //Summarytable PL License
                    string strPLComment = "";
                    if (diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse") || diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                    {
                        if (diligenceInput.pllicenseModels[0].General_PL_License.Equals("Yes"))
                        {
                            strPLComment = "[LastName] holds a [ProfessionalLicenseType1] license with the [PLOrganization1]";
                        }
                        if (diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. SEC");
                        }
                        if (diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.K. FCA");
                        }
                        if (diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. FINRA");
                        }
                        if (diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. NFA");
                        }
                        if (diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse"))
                        {
                            if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                            strPLComment = string.Concat(strPLComment, "[LastName] has been registered with HKSFC");
                        }

                        if (diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse"))
                        {
                            strPLComment = string.Concat(strPLComment, "\n\nFurther, <investigator to insert regulatory hits here>");
                        }
                        strPLComment = strPLComment.Replace("[ProfessionalLicenseType1]", diligenceInput.pllicenseModels[0].PL_License_Type);
                        strPLComment = strPLComment.Replace("[PLOrganization1]", diligenceInput.pllicenseModels[0].PL_Organization);
                        doc.Replace("PROFRESULT", "Records", true, true);
                    }
                    else
                    {
                        strPLComment = "";
                        doc.Replace("PROFRESULT", "Clear", true, true);
                    }
                    doc.Replace("PROFCOMMENT", strPLComment, true, true);

                    //Legal_Record_Judgments_Liens_Hits
                    string Legrechitcommentmodel = "";
                    string strexecutive_sumlegrec = "";

                    CommentModel Legrechitcommentmodel1 = _context.DbComment
                                    .Where(u => u.Comment_type == "Legal_Record_Judgments_Liens_Hits")
                                    .FirstOrDefault();
                    if (diligenceInput.summarymodel.criminal_records.StartsWith("Record") || diligenceInput.summarymodel.bankruptcy_filings.StartsWith("Record") || diligenceInput.summarymodel.civil_court_Litigation.StartsWith("Record"))
                    {
                        Legrechitcommentmodel = Legrechitcommentmodel1.confirmed_comment.ToString();
                        strexecutive_sumlegrec = "Yes";
                    }
                    else
                    {
                        Legrechitcommentmodel = Legrechitcommentmodel1.unconfirmed_comment.ToString();
                        strexecutive_sumlegrec = "No";
                    }
                    doc.Replace("HasLegalRecordsJudgmentsorLiensHits", Legrechitcommentmodel, true, true);
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
                    if (regflag == "Records" || strexecutive_sumlegrec == "Yes")
                    {
                        executivesumhitabove = "Yes";
                    }
                    else
                    {
                        executivesumhitabove = "No";
                    }
                    if (executivesumhitabove == "Yes")
                    {
                        if (diligenceInput.otherdetails.Has_Companion_Report.ToString().Equals("Yes"))
                        {
                            hitcompcommentmodel = hitcompcommentmodel1.confirmed_comment.ToString();
                            searchtext = "In sum, with the exception of the above, no other issues of potential relevance were identified in connection with";
                        }
                        else
                        {
                            hitcompcommentmodel = hitcompcommentmodel1.unconfirmed_comment.ToString();
                        }
                    }
                    else
                    {
                        if (diligenceInput.otherdetails.Has_Companion_Report.ToString().Equals("Yes"))
                        {
                            hitcompcommentmodel = nohitcompcommentmodel1.confirmed_comment.ToString();
                            searchtext = "In sum, no other issues of potential relevance were identified in connection with";
                        }
                        else
                        {
                            // hitcompcommentmodel = nohitcompcommentmodel1.unconfirmed_comment.ToString();
                            hitcompcommentmodel = "In sum, no issues of potential relevance were identified in connection with [FirstName] [MiddleInitial] [LastName] in [Country].";
                            searchtext = "In sum, no issues of potential relevance were identified in connection with";
                        }
                    }
                    hitcompcommentmodel = hitcompcommentmodel.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName.ToString());
                    // hitcompcommentmodel = hitcompcommentmodel.Replace("[MiddleInitial]", diligenceInput.diligenceInputModel.MiddleInitial.ToString());
                    doc.Replace("HasHitsAboveAndCompanionReport", hitcompcommentmodel, true, true);
                    doc.SaveToFile(savePath);
                    if (searchtext.ToString().Equals("")) { }
                    else
                    {
                        string blnresult = "";
                        try
                        {
                            for (int j = 1; j < 8; j++)
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
                                        }
                                    }
                                    if (blnresult.Equals("true")) { break; }
                                }
                                if (blnresult.Equals("true")) { break; }
                            }
                        }
                        catch { }
                    }
                    doc.Replace("[LastName]", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName), true, true);
                    doc.Replace("(“OFAC”),", "(“OF[AddingOFSTATE]AC”),", true, true);
                    doc.SaveToFile(savePath);
                    string blnresultfound = "";
                    try
                    {
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
                                        if (abc.ToString().Contains("(“OF[AddingOFSTATE]AC”),") || abc.ToString().EndsWith("(“OF[AddingOFSTATE]AC”),"))
                                        {
                                            Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                            footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                            footnote1.MarkerCharacterFormat.FontSize = 11;
                                            //Insert footnote1 after the word "Spire.Doc"
                                            para.ChildObjects.Insert(i + 1, footnote1);
                                            TextRange _text = footnote1.TextBody.AddParagraph().AppendText("It should be emphasized that updates are made to these lists on a periodic and irregular basis, and for purposes of preparing this memorandum, these lists were searched on <investigator to insert date of research>.");
                                            footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                            _text.CharacterFormat.FontName = "Calibri (Body)";
                                            _text.CharacterFormat.FontSize = 9;

                                            //Append the line
                                            string strglobaltextappended = "";
                                            if (diligenceInput.otherdetails.Global_Security_Hits.Equals("Yes"))
                                            {
                                                strglobaltextappended = " and the following information was identified in connection with [LastName]: ";
                                            }
                                            else
                                            {
                                                strglobaltextappended = " and it is noted that [LastName] was not identified on any of these lists.";
                                            }
                                            strglobaltextappended = strglobaltextappended.Replace("[LastName]", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName));
                                            TextRange tr = para.AppendText(strglobaltextappended);
                                            tr.CharacterFormat.FontName = "Calibri (Body)";
                                            tr.CharacterFormat.FontSize = 11;
                                            doc.Replace("(“OF[AddingOFSTATE]AC”),", "(“OF[AddedOFSTATE]AC”),", true, true);
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
                    }
                    catch { }
                    doc.SaveToFile(savePath);
                }
                if (diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom"))
                {
                    doc.Replace("the [Country]", "[country]", true, true);
                    doc.SaveToFile(savePath);
                    doc.Replace("[Country]", string.Concat("the ", diligenceInput.diligenceInputModel.Country.ToString()), true, true);
                    doc.Replace("[CountryHeader]", string.Concat("The ", diligenceInput.diligenceInputModel.Country), true, true);
                    doc.SaveToFile(savePath);
                }
                else
                {
                    doc.Replace("[CountryHeader]", diligenceInput.diligenceInputModel.Country.ToString(), true, true);
                    doc.Replace("[Country]", diligenceInput.diligenceInputModel.Country.ToString(), true, true);
                }
                doc.Replace("[CaseNumber]", diligenceInput.diligenceInputModel.CaseNumber.TrimEnd(), true, false);
                doc.Replace("[LAST NAME]", diligenceInput.diligenceInputModel.LastName, false, true);
                doc.Replace("[LASTNAME]", diligenceInput.diligenceInputModel.LastName, false, true);
                doc.Replace("[FIRST NAME]", diligenceInput.diligenceInputModel.FirstName, false, true);
                try
                {
                    if (diligenceInput.diligenceInputModel.MiddleName == "") { doc.Replace(" [MIDDLE NAME]", "", false, false); }
                    else
                    {
                        doc.Replace("[MIDDLE NAME]", diligenceInput.diligenceInputModel.MiddleName, false, true);
                    }
                }
                catch
                {
                    doc.Replace(" [MIDDLE NAME]", "", false, false);
                }
                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, false, true);
                doc.Replace("[CLIENT NAME]", Client_Name, true, true);
                doc.Replace("[FirstName] [LastName]", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName),true,true);
                doc.SaveToFile(savePath);
                doc.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName, true, false);
                doc.Replace("[LastName]", string.Concat(diligenceInput.diligenceInputModel.FirstName, " ", diligenceInput.diligenceInputModel.LastName), true, false);
                doc.SaveToFile(savePath);
                doc.Replace("COMP2RESULT", "COMPRESULT", true, true);
                doc.Replace("COMP2COMMENT", "COMPCOMMENT", true, true);
                doc.Replace("PROF2RESULT", "PROFRESULT", true, true);
                doc.Replace("PROF2COMMENT", "PROFCOMMENT", true, true);
                doc.Replace("FIN2INDRESULT", "FININDRESULT", true, true);
                doc.Replace("FIN2INDCOMMENT", "FININDCOMMENT", true, true);
                doc.Replace("KON2SECRESULT", "KONSECRESULT", true, true);
                doc.Replace("KONSEC2COMMENT", "KONSECCOMMENT", true, true);                
                doc.Replace("CRIMINAL2COMMENT", "CRIMINALCOMMENT", true, true);
                doc.Replace("CRIMINAL2RESULT", "CRIMINALRESULT", true, true);
                doc.Replace("SEC2RESULT", "SECRESULT", true, true);
                doc.Replace("SEC2COMMENT", "SECCOMMENT", true, true);
                doc.Replace("UKFI2CORESULT", "UKFICORESULT", true, true);
                doc.Replace("UKFI2COCOMMENT", "UKFICOCOMMENT", true, true);
                doc.Replace("USNF2RESULT", "USNFRESULT", true, true);
                doc.Replace("USNF2COMMENT", "USNFCOMMENT", true, true);
                doc.Replace("INCOJORESULT", "INCOJORESULT", true, true);
                doc.Replace("INCOJO2COMMENT", "INCOJOCOMMENT", true, true);
                doc.Replace("POLITIC2RESULT", "POLITICRESULT", true, true);
                doc.Replace("POLITIC2COMMENT", "POLITICCOMMENT", true, true);
                doc.Replace("[Country2Header][COURT2DESC]", "[CountryHeader][COURTDESC]", true, false);
                doc.Replace("GLOBAL2SECURITYHITSDESCRIPTION", "GLOBALSECURITYHITSDESCRIPTION", true, false);
                doc.Replace("GLOBAL2SECFAMILYHITSDESCRIPTION", "GLOBALSECFAMILYHITSDESCRIPTION", true, false);
                doc.Replace("PEP2HITSDESCRIPTION", "PEPHITSDESCRIPTION", true, false);
                doc.Replace("ICIJ2HITSDESCRIPTION", "ICIJHITSDESCRIPTION", true, false);
                doc.Replace("PRESS2ANDMEDIASEARCHDESCRIPTION", "PRESSANDMEDIASEARCHDESCRIPTION", true, false);
                doc.Replace("CHARACTER2REFERENCESDESC", "CHARACTERREFERENCESDESC", true, false);
                doc.Replace("SOURCE2OFWEALTHDESC", "SOURCEOFWEALTHDESC", true, false);
                doc.Replace("DISCREET2REPUTATIONALINQUIRIESDESC", "DISCREETREPUTATIONALINQUIRIESDESC", true, false);
                TextSelection[] text00 = doc.FindAllString("First_Name0 Middle_Name0 Last_Name0", false, false);
                if (text00 != null)
                {
                    foreach (TextSelection seletion in text00)
                    {
                        seletion.GetAsOneRange().CharacterFormat.Bold = true;
                    }
                }
                TextSelection[] text5 = doc.FindAllString("First_Name Middle_Name Last_Name", false, false);
                if (text5 != null)
                {
                    foreach (TextSelection seletion in text5)
                    {
                        seletion.GetAsOneRange().CharacterFormat.Bold = true;
                    }
                }
            doc.SaveToFile(savePath);
                //try
                //{
                //    doc.Replace("First_Name Middle_Name Last_Name", string.Concat(diligenceInput.diligenceInputModel.FirstName.ToUpper(), " ", diligenceInput.diligenceInputModel.MiddleName.ToUpper(), " ", diligenceInput.diligenceInputModel.LastName.ToUpper()), true, true);
                //}
                //catch
                //{
                //    doc.Replace("First_Name Middle_Name Last_Name", string.Concat(diligenceInput.diligenceInputModel.FirstName.ToUpper(), " ", diligenceInput.diligenceInputModel.LastName.ToUpper()), true, true);
                //}
                try
                {
                    doc.Replace("First_Name0 Middle_Name0 Last_Name0", string.Concat(diligenceInput.diligenceInputModel.FirstName.ToUpper(), " ", diligenceInput.diligenceInputModel.MiddleName.ToUpper(), " ", diligenceInput.diligenceInputModel.LastName.ToUpper()), true, true);
                }
                catch
                {
                    doc.Replace("First_Name0 Middle_Name0 Last_Name0", string.Concat(diligenceInput.diligenceInputModel.FirstName.ToUpper(), " ", diligenceInput.diligenceInputModel.LastName.ToUpper()), true, true);
                }           
                //doc.Replace("First_Name0 Middle_Name0 Last_Name0", "First_Name Middle_Name Last_Name", true, true);
                doc.SaveToFile(savePath);
                TextSelection[] text4 = doc.FindAllString(string.Concat(diligenceInput.diligenceInputModel.FirstName.ToUpper(), " ", diligenceInput.diligenceInputModel.MiddleName.ToUpper(), " ", diligenceInput.diligenceInputModel.LastName.ToUpper()), true, true);
                if (text4 != null)
                {
                    foreach (TextSelection seletion in text4)
                    {
                        seletion.GetAsOneRange().CharacterFormat.Bold = true;
                    }
                }
                doc.SaveToFile(savePath);
                HttpContext.Session.SetString("first_name", diligenceInput.diligenceInputModel.FirstName);
                HttpContext.Session.SetString("last_name", diligenceInput.diligenceInputModel.LastName);
                HttpContext.Session.SetString("country", diligenceInput.diligenceInputModel.Country);
                HttpContext.Session.SetString("case_number", diligenceInput.diligenceInputModel.CaseNumber);
                HttpContext.Session.SetString("regflag", regflag);
                HttpContext.Session.SetString("middleinitial", diligenceInput.diligenceInputModel.MiddleInitial);
               
            }
            if (diligenceInputs.familyModels.Count > 1)
            {
                doc.Replace("Passport", "Passports", true, true);
                doc.Replace("Certificate", "Certificates", true, true);
                doc.Replace("National Identity Card", "National Identity Cards", true, true);
                doc.Replace("License", "Licenses", true, true);
                doc.Replace("Disclosure Form", "Disclosure Forms", true, true);
                doc.Replace("Letter", "Letters", true, true);
                doc.Replace("Bank Statement", "Bank Statements", true, true);
                doc.Replace("Investment Agreement", "Investment Agreements", true, true);
                doc.Replace("Photograph","Photographs", true, true);
                doc.Replace("Residency", "Residencies", true, true);
                doc.Replace("Utility Bill", "Utility Bills", true, true);
            }
            doc.Replace(",.", ".", true, true);
            doc.SaveToFile(savePath);
            HttpContext.Session.SetString("recordid", recordid);
            //return Save_Page2(diligenceInputs.summarymodel);
            return Save_Page2(diligenceInputs);
        }
        public IActionResult INTL_Citizen_Individualpage2()
        {
                return View();
        }
        public IActionResult Save_Page2(MainModel diligenceInputs)
            {
            string recordid = HttpContext.Session.GetString("recordid");
            string last_name = HttpContext.Session.GetString("last_name");
            string country = HttpContext.Session.GetString("country");
            string case_number = HttpContext.Session.GetString("case_number");
            string regflag = HttpContext.Session.GetString("regflag");
            string middleinitial = HttpContext.Session.GetString("middleinitial");
            string savePath = _config.GetValue<string>("ReportPath");

            DiligenceInputModel DI2 = _context.DbPersonalInfo
                        .Where(u => u.record_Id == diligenceInputs.familyModels[0].Family_record_id)
                        .FirstOrDefault();
            savePath = string.Concat(savePath, DI2.LastName.ToString(), "_SterlingDiligenceReport(", case_number.ToString(), ")_DRAFT.docx");
            Document doc = new Document(savePath);
            string strcomment;
            Table table = doc.Sections[1].Tables[0] as Table;
            int adult_count = 0;
            
            for (int famcount = 0; famcount < diligenceInputs.familyModels.Count; famcount++)
            {
                if (diligenceInputs.familyModels[famcount].adult_minor == "Adult")
                {
                    adult_count = famcount + 1;
                   
                    //SummaryResulttableModel ST = new SummaryResulttableModel();
                    SummaryResulttableModel ST = _context.summaryResulttableModels
                        .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                        .FirstOrDefault();
                    DiligenceInputModel DI = _context.DbPersonalInfo
                        .Where(u => u.record_Id == diligenceInputs.familyModels[famcount].Family_record_id)
                        .FirstOrDefault();                   
                    try
                    {
                        if (diligenceInputs.familyModels[famcount + 1].adult_minor == "Adult")
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
                            doc.Replace("CIVILJUDGECOMMENT", "CIVILJUDGECOMMENT\n\nCIVILJUDGE2COMMENT", true, true);
                            doc.Replace("CIVILJUDGERESULT", "CIVILJUDGERESULT\n\nCIVILJUDGE2RESULT", true, true);
                            
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
                            doc.Replace("PLRESULT", string.Concat(DI.FirstName, " ", DI.LastName," - Clear"), true, true);
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
                            strcomment = "LastName has jurisdictional ties to [Country]";
                            strcomment = strcomment.Replace("LastName", string.Concat(DI.FirstName, " ", DI.LastName));
                            doc.Replace("NAMECOMMENT", strcomment, true, true);
                            doc.Replace("NAMERESULT", string.Concat(DI.FirstName, " ", DI.LastName, " - Clear"), true, true);
                            break;
                        case "Discrepancy Identified":
                            strcomment = "LastName has jurisdictional ties to [Country]\n\nHowever, while not reported by the subject, LastName was also identified as having current ties to < Investigator to insert additional jurisdictions here>";
                            strcomment = strcomment.Replace("LastName", string.Concat(DI.FirstName, " ", DI.LastName));
                            doc.Replace("NAMECOMMENT", strcomment, true, true);
                            doc.Replace("NAMERESULT", string.Concat(DI.FirstName, " ", DI.LastName, " - Discrepancy Identified"), true, true);
                            break;
                    }
                    switch (regflag.ToString())
                    {
                        case "Clear":
                            doc.Replace("REGCOMMENT", "", true, true);
                            doc.Replace("REGRESULT", string.Concat(DI.FirstName, " ", DI.LastName, " - Clear"), true, true);
                            break;
                        case "Records":
                            strcomment = "<investigator to insert regulatory hits here>";
                            doc.Replace("REGCOMMENT", strcomment, true, true);
                            doc.Replace("REGRESULT", string.Concat(DI.FirstName, " ", DI.LastName, " - Records"), true, true);
                            break;
                    }
                    string bankrupcomment = "";
                    string bankrupresult = "";
                    switch (ST.bankruptcy_filings.ToString())
                    {
                        case "Clear":
                            bankrupcomment = "";
                            bankrupresult = string.Concat(DI.FirstName, " ", DI.LastName, " - Clear");
                            break;
                        case "Records":
                            bankrupcomment = "LastName was identified as a <party type> in connection with at least <number> bankruptcy filings in [Country], filed between <date> and <date>, which are currently <status> ";
                            bankrupcomment = bankrupcomment.Replace("LastName", string.Concat(DI.FirstName, " ", DI.LastName));
                            bankrupresult = string.Concat(DI.FirstName, " ", DI.LastName, " - Records");
                            break;
                        case "Record":
                            bankrupcomment = "[LastName] was identified as a <party type> in connection with a bankruptcy filing in [Country], which were recorded in <date> and <date>, and is currently <status>";
                            bankrupresult = string.Concat(DI.FirstName, " ", DI.LastName, " - Record");
                            break;
                        case "Results Pending":
                            bankrupcomment = "Bankruptcy court records searches are currently pending in [Country] the results of which will be provided under separate cover upon receipt";
                            bankrupresult = string.Concat(DI.FirstName, " ", DI.LastName, " - Results Pending");
                            break;
                    }
                    if (ST.bankruptcy_filings1 == true)
                    {
                        bankrupresult = string.Concat(bankrupresult, "\n\nPossible Records");
                        bankrupcomment = string.Concat(bankrupcomment, "\n\nOne or more individuals known only as “[FirstName] [LastName]” were identified as Petitioners to at least <number> bankruptcy filings in [Country], which were recorded between <date> and <date>, and are currently <status>");
                    }
                    doc.Replace("BANKRUPCOMMENT", bankrupcomment, true, true);
                    doc.Replace("BANKRUPRESULT", bankrupresult, true, true);
                    string civilcourtcomment = "";
                    string civilcourtResult = "";
                    switch (ST.civil_court_Litigation.ToString())
                    {
                        case "Clear":
                            civilcourtcomment = "";
                            civilcourtResult = string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Clear");
                            break;
                        case "Records":
                            civilcourtcomment = "LastName was identified as a <party type> in connection with at least <number> civil litigation filings in [Country], filed between <date> and <date>, which are currently <status>";
                            civilcourtcomment = civilcourtcomment.Replace("LastName", string.Concat(DI.FirstName, " ", DI.LastName));
                            civilcourtResult = string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records");
                            break;
                        case "Record":
                            civilcourtcomment = "[Last Name] was identified as a <party type> in connection with a civil litigation filing in [Country], which was recorded in <date>, which is <status>";
                            civilcourtResult = string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Record");
                            break;
                        case "Results Pending":
                            civilcourtcomment = "A civil court records search is currently pending in [Country] the results of which will be provided under separate cover upon receipt";
                            civilcourtResult = string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Results Pending");
                            break;
                    }
                    if (ST.civil_court_Litigation1 == true)
                    {
                        civilcourtResult = string.Concat(civilcourtResult, "\n\nPossible Records");
                        civilcourtcomment = string.Concat(civilcourtcomment, "\n\nOne or more individuals known only as “[FirstName] [LastName]” were identified as <Party Type> in at least <number> civil litigation filings in [Country], which were recorded between <date> and <date>, which are <status>");
                    }
                    doc.Replace("CIVILCOMMENT", civilcourtcomment, true, true);
                    doc.Replace("CIVILRESULT", civilcourtResult, true, true);
                    switch (ST.civil_judge_Liens.ToString())
                    {
                        case "N/A":
                            doc.Replace("CIVILJUDGECOMMENT", "It is noted that the subject’s address history in the United Kingdom would be required to conduct a search through the Registry of Judgments, Orders and Fines in the United Kingdom CIVILJUDGECOMMENT", true, true);
                            doc.Replace("CIVILJUDGERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - N/A CIVILJUDGERESULT"), true, true);
                            break;
                        case "Clear":
                            doc.Replace("CIVILJUDGECOMMENT", "CIVILJUDGECOMMENT", true, true);
                            doc.Replace("CIVILJUDGERESULT", string.Concat(DI.FirstName.ToString()," ",DI.LastName.ToString()," - Clear CIVILJUDGERESULT"), true, true);
                            break;
                        case "Records":
                            strcomment = "[LastName] was identified as a <party type> in connection with at least <number> judgments filed in [Country] between <date> and <date>, which are currently <status> CIVILJUDGECOMMENT";
                            strcomment = strcomment.Replace("[LastName]", string.Concat(DI.FirstName, " ", DI.LastName));
                            doc.Replace("CIVILJUDGECOMMENT", strcomment, true, true);
                            doc.Replace("CIVILJUDGERESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - Records CIVILJUDGERESULT"), true, true);
                            break;
                    }
                    doc.SaveToFile(savePath);
                    if (ST.civil_judge_Liens1 == true)
                    {
                        doc.Replace("CIVILJUDGECOMMENT", "\n\nOne or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>", true, false);
                        doc.Replace("CIVILJUDGERESULT", "\n\nPossible Records", true, false);
                    }
                    else
                    {
                        doc.Replace(" CIVILJUDGECOMMENT", "", true, true);
                        doc.Replace(" CIVILJUDGERESULT", "", true, true);
                        doc.Replace("CIVILJUDGECOMMENT", "", true, true);
                        doc.Replace("CIVILJUDGERESULT", "", true, true);
                    }
                    //string criminalcomment = "";
                    //string criminalresult = "";
                   
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
                            doc.Replace("FDEPORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - ","Clear"), true, true);
                            break;
                        case "Records":
                            doc.Replace("FDEPOCOMMENT", "<investigator to insert summary here>", true, true);
                            doc.Replace("FDEPORESULT", string.Concat(DI.FirstName.ToString(), " ", DI.LastName.ToString(), " - ","Records"), true, true);
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
                doc.Replace("CIVILJUDGE2COMMENT", "CIVILJUDGECOMMENT", true, true);
                doc.Replace("CIVILJUDGE2RESULT", "CIVILJUDGERESULT", true, true);
                
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
            doc.SaveToFile(savePath);
            for(int i=0;i< table.Rows.Count; i++)
            {
                try
                {
                    string C1text = "";
                    TableCell cell11 = table.Rows[i].Cells[2];
                    int paracount = 0;
                    paracount = cell11.Paragraphs.Count;
                    for (int k=0; k < paracount; k++){ 
                    Paragraph p1 = cell11.Paragraphs[k];
                        if (p1.Text == "")
                        {
                            TableCell cel2 = table.Rows[i].Cells[1];
                            Paragraph p2 = cel2.Paragraphs[k];
                            p2.Text = "".TrimEnd();
                            C1text = "Clear";
                        }
                        else
                        {
                            C1text = "";
                        }
                    }
                    if( C1text== "Clear")
                    {
                        TableCell cel2 = table.Rows[i].Cells[1];
                        TableCell cel3 = table.Rows[i].Cells[2];
                        
                        Paragraph p2 = cel2.Paragraphs[0];
                        p2.Text = "Clear".Trim();
                        
                        //cel2.CellFormat.FitText = true;
                        //cel3.CellFormat.FitText = true;

                    }
                }
                catch { }
            }
            doc.SaveToFile(savePath);
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

                TextSelection[] textresult = doc.FindAllString("Possible Records", true, true);
                if (textresult != null)
                {
                    foreach (TextSelection seletion in textresult)
                    {
                        seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    }
                }
                TextSelection[] textresult1 = doc.FindAllString("Results Pending", false, false);
                if (textresult1 != null)
                {
                    foreach (TextSelection seletion in textresult1)
                    {
                        seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    }
                }
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
                TextSelection[] text33 = doc.FindAllString("As confirmed", false, false);
                if (text33 != null)
                {
                    foreach (TextSelection seletion in text33)
                    {
                        seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;                  
                }
                }
            doc.SaveToFile(savePath);
            //TextSelection[] text34 = doc.FindAllString("as confirmed", false, true);
            //if (text34 != null)
            //{
            //    foreach (TextSelection seletion in text34)
            //    {
            //        seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
            //    }
            //}
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


            TextSelection[] text3 = doc.FindAllString("REPORTED BUSINESS AFFILIATIONS AND EMPLOYMENT HISTORY", false, false);
            if (text3 != null)
            {
                foreach (TextSelection seletion in text3)
                {
                    seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                }
            }
            TextSelection[] text4 = doc.FindAllString("UNDISCLOSED AFFILIATIONS", false, false);
            if (text4 != null)
            {
                foreach (TextSelection seletion in text4)
                {
                    seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                }
            }

            TextSelection[] text5 = doc.FindAllString("Other Legal Records Searches", false, false);
            if (text5 != null)
            {
                foreach (TextSelection seletion in text5)
                {
                    seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                }
            }
            doc.Replace("[AddedOFSTATE]", "", true, false);
            doc.Replace("[addedFootnote]", "", true, false);
            
                if (country.ToString().Equals("United Kingdom"))
                {
                    doc.Replace("the [Country]", "[Country]", true, true);
                    doc.SaveToFile(savePath);
                    doc.Replace("[Country]", string.Concat("the ", country), true, false);
                    doc.SaveToFile(savePath);
                    doc.Replace("the The United Kingdom", "the United Kingdom", true, true);
                }
                else
                {
                    doc.Replace("[Country]", country, true, false);
                }
                doc.SaveToFile(savePath);
                doc.Replace("[HeaderCountry]", country, true, true);
                doc.Replace("  ", " ", true, false);
                doc.Replace("  (", " (", true, false);
                doc.Replace("united states", "United States", true, false);
                doc.Replace("investigator", "Investigator", true, false);
                doc.Replace(" .", ".", true, false);
                doc.Replace(" ,", ",", true, false);
                doc.Replace("[Month Generated] [Year Generated]", string.Concat(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month).ToString(), " ", DateTime.Now.Year.ToString()), true, true);
                doc.SaveToFile(savePath);
                doc.Replace(".  ", ". ", true, false);
                doc.Replace(". ", ".  ", true, false);
                doc.SaveToFile(savePath);
                try
                {
                    if (middleinitial == "")
                    {
                        doc.Replace(" Middle_Initial", "", true, false);
                        doc.Replace(" MiddleInitial", "", true, false);
                        doc.Replace(" [MiddleInitial]", "", true, false);
                    }
                    else
                    {
                        doc.Replace("Middle_Initial", middleinitial.ToUpper().ToString().TrimEnd(), true, true);
                        doc.Replace("MiddleInitial", middleinitial.TrimEnd(), true, true);
                        doc.Replace("[MiddleInitial]", middleinitial, true, true);
                    }
                }
                catch
                {
                    doc.Replace(" Middle_Initial", "", true, false);
                    doc.Replace(" MiddleInitial", "", true, false);
                    doc.Replace(" [MiddleInitial]", "", true, false);
                }
                doc.Replace("U.S.  ", "U.S. ", true, false);
                doc.Replace("U.K.  ", "U.K. ", true, false);
            doc.Replace("the the","the", false, true);            
            try
            {
                if (adult_count > 1)
                {
                    doc.Replace("subject", "subjects", true, true);
                }
            }
            catch { }
            doc.Replace("It is noted that the subjects did not provide", "It is noted that the subject did not provide", true, true);
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