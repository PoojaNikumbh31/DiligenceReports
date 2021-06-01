using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using DiligenceReportCreation.Models;
using DiligenceReportCreation.Data;
using Microsoft.AspNetCore.Http;
using System.DirectoryServices;
using Microsoft.AspNetCore.Mvc.Rendering;
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
    public class DiligenceController : Controller
    {
        private readonly DataBaseContext _context;
        private readonly IConfiguration _config;        

        public DiligenceController(IConfiguration config,DataBaseContext context)
        {
            _config = config;
            _context = context;                                 
        }
        //DataBaseContext db = new DataBaseContext();
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult Formselection()
        {           
            return View();
        }
        [HttpPost]
        public IActionResult Formselection(ReportModel report)
        {
            //Country List               
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
            ReportModel comment12 = _context.reportModel
                             .Where(u => u.casenumber == report.casenumber && u.lastname == report.lastname)
                             .FirstOrDefault();
            try
            {
                if (comment12.ToString().Equals(""))
                {
                    TempData["message"] = "Report not found";
                    return RedirectToAction("Formselection", "Diligence");
                }
                else
                {
                    HttpContext.Session.SetString("recordid", comment12.record_Id);
                    string straction = "";
                    string strcontroller = "";
                    switch (comment12.TemplateType)
                    {
                        case "International DD Individual":
                            straction = "Edit";
                            strcontroller = "Diligence";
                            break;
                        case "US DD Individual":
                            straction = "US_DD_Individual_Edit";
                            strcontroller = "USDDIndividual";
                            break;
                        case "US PE Individual":
                            straction = "US_PE_Individual_Edit";
                            strcontroller = "USPEIndividual";
                            break;
                        case "International Citizenship Individual":
                            straction = "INTL_Citizen_Individual_Edit";
                            strcontroller = "IN_CITI_Individual";
                            break;
                        case "US Company":
                            straction = "US_COMPANY_Edit";
                            strcontroller = "USCOMPANY";
                            break;
                        case "International Citizenship Family":
                            straction = "INTL_Citizen_Family_Newpage";
                            strcontroller = "IN_CITI_Family";
                            break;
                        case "International PE Individual":
                            straction = "IN_PE_Individual_Edit";
                            strcontroller = "IN_PE_Individual";
                            break;
                        case "International Company":
                            straction = "IN_COMPANY_Edit";
                            strcontroller = "INCOMPANY";
                            break;
                        case "US Citizenship Family":
                            straction = "US_Citizen_Family_Newpage";
                            strcontroller = "US_CITI_Family";
                            break;

                    }
                    return RedirectToAction(straction, strcontroller);
                }
        }
            catch
            {
                TempData["message"] = "Report not found";
                return RedirectToAction("Formselection", "Diligence");
    }
        }
        [HttpGet]
        public IActionResult Edit()
        {
            //Country List
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
            string recordid = HttpContext.Session.GetString("recordid");
            MainModel mainModel = new MainModel();            
            DiligenceInputModel diligenceInput = _context.DbPersonalInfo
                       .Where(u => u.record_Id == recordid)
                       .FirstOrDefault();
            Otherdatails otherdatails = _context.othersModel
                       .Where(u => u.record_Id == recordid)
                       .FirstOrDefault();
            CountrySpecificModel countrySpecificModel = _context.CSComment
                       .Where(u => u.record_Id == recordid)
                       .FirstOrDefault();
            mainModel.EmployerModel = _context.DbEmployer
                        .Where(u => u.record_Id == recordid)
                        .ToList();

            mainModel.educationModels = _context.DbEducation
                       .Where(u => u.record_Id == recordid)
                       .ToList();
            mainModel.pllicenseModels = _context.DbPLLicense
                       .Where(u => u.record_Id == recordid)
                        .ToList();
            mainModel.summarymodel = _context.summaryResulttableModels
                       .Where(u => u.record_Id == recordid)
                        .FirstOrDefault();
            mainModel.diligenceInputModel = diligenceInput;
            mainModel.otherdetails = otherdatails;
            mainModel.csModel = countrySpecificModel;
            HttpContext.Session.SetString("recordid", recordid);
            ReportModel report = _context.reportModel
                .Where(u => u.record_Id == recordid)
                                  .FirstOrDefault();
            ViewBag.CaseNumber = report.casenumber;
            ViewBag.LastName = report.lastname;
            HttpContext.Session.SetString("lastname", report.lastname);
            HttpContext.Session.SetString("casenumber", report.casenumber);
            return View(mainModel);
        }      
        [HttpPost]
        public IActionResult Edit(MainModel mainModel, string SaveData,string Submit)
        {
            string recordid = HttpContext.Session.GetString("recordid");
            if (Submit == "SubmitData")
            {
                MainModel main = new MainModel();
                DiligenceInputModel dinput = new DiligenceInputModel();
                Otherdatails otherdatails = new Otherdatails();
                CountrySpecificModel countrySpecificModel = new CountrySpecificModel();
                EmployerModel employer = new EmployerModel();
                EducationModel education = new EducationModel();
                PLLicenseModel pL = new PLLicenseModel();
                SummaryResulttableModel summary = new SummaryResulttableModel();
                try
                {
                    dinput = _context.DbPersonalInfo
                                      .Where(u => u.record_Id == recordid)
                                      .FirstOrDefault();
                    if (dinput == null)
                    {
                        TempData["message"] = "Please enter the details in personal details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in personal details section";
                }
                try
                {
                    otherdatails = _context.othersModel
                               .Where(u => u.record_Id == recordid)
                               .FirstOrDefault();
                    if (otherdatails == null)
                    {
                        TempData["message"] = "Please enter the details in other details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in other details section";
                }
                try
                {
                    countrySpecificModel = _context.CSComment
                                 .Where(u => u.record_Id == recordid)
                                 .FirstOrDefault();
                    if (countrySpecificModel == null)
                    {
                        TempData["message"] = "Please enter the details in country specific section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in country specific section";
                }
                try
                {
                    employer = _context.DbEmployer
                                .Where(u => u.record_Id == recordid)
                                .FirstOrDefault();
                    if (employer == null)
                    {
                        TempData["message"] = "Please enter the details in Employee details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in Employee details section";
                }
                try
                {
                    education = _context.DbEducation
                               .Where(u => u.record_Id == recordid)
                               .FirstOrDefault();
                    if (education == null)
                    {
                        TempData["message"] = "Please enter the details in Education details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in Education details section";
                }
                try
                {
                    pL = _context.DbPLLicense
                               .Where(u => u.record_Id == recordid)
                                .FirstOrDefault();
                    if (pL == null)
                    {
                        TempData["message"] = "Please enter the details in Pllicence details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in Pllicence details section";
                }
                try
                {
                    summary = _context.summaryResulttableModels
                                      .Where(u => u.record_Id == recordid)
                                      .FirstOrDefault();
                    if (summary == null)
                    {
                        TempData["message"] = "Please enter the details in summary result table section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in summary result table section";
                }
                if (TempData["message"] == null)
                {
                    main.diligenceInputModel = dinput;
                    main.educationModels = _context.DbEducation
                             .Where(u => u.record_Id == recordid)
                             .ToList();
                    main.otherdetails = otherdatails;
                    main.EmployerModel = _context.DbEmployer
                                .Where(u => u.record_Id == recordid)
                                .ToList();
                    main.pllicenseModels = _context.DbPLLicense
                               .Where(u => u.record_Id == recordid)
                                .ToList();
                    main.csModel = countrySpecificModel;
                    main.summarymodel = summary;
                    return Save_Page1(main);
                }
            }
            string lastname = HttpContext.Session.GetString("lastname");
            //string firstname = HttpContext.Session.GetString("firstname");
            string casenumber = HttpContext.Session.GetString("casenumber");
            //string middleinitial = HttpContext.Session.GetString("middleinitial");
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
                           .Where(u => u.record_Id == recordid)
                                          .FirstOrDefault();
                                inputModel.record_Id = recordid;
                                inputModel.ClientName = mainModel.diligenceInputModel.ClientName;
                                inputModel.Dob = mainModel.diligenceInputModel.Dob;
                                inputModel.FirstName = mainModel.diligenceInputModel.FirstName;
                                inputModel.LastName = lastname;
                                inputModel.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                                inputModel.MiddleName = mainModel.diligenceInputModel.MiddleName;
                                inputModel.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                                inputModel.CaseNumber = casenumber;
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
                                inputModel.CurrentCity = mainModel.diligenceInputModel.CurrentCity;
                                inputModel.CurrentCountry = mainModel.diligenceInputModel.CurrentCountry;
                                inputModel.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                                inputModel.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
                                _context.Entry(inputModel).State = EntityState.Modified;
                                _context.SaveChanges();
                            }
                            catch
                            {
                                DiligenceInputModel inputModel = new DiligenceInputModel();
                                inputModel.record_Id = recordid;
                                inputModel.ClientName = mainModel.diligenceInputModel.ClientName;
                                inputModel.Dob = mainModel.diligenceInputModel.Dob;
                                inputModel.FirstName = mainModel.diligenceInputModel.FirstName;
                                inputModel.LastName = lastname;
                                inputModel.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                                inputModel.MiddleName = mainModel.diligenceInputModel.MiddleName;
                                inputModel.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                                inputModel.CaseNumber = casenumber;
                                inputModel.MaidenName = mainModel.diligenceInputModel.MaidenName;
                                inputModel.FullaliasName = mainModel.diligenceInputModel.FullaliasName;
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Dob = mainModel.diligenceInputModel.Dob;
                                inputModel.Pob = mainModel.diligenceInputModel.Pob;
                                inputModel.City = mainModel.diligenceInputModel.City;
                                inputModel.Nationality = mainModel.diligenceInputModel.Nationality;
                                inputModel.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                                inputModel.Natlidissuedate = mainModel.diligenceInputModel.Natlidissuedate;
                                inputModel.Natlidexpirationdate = mainModel.diligenceInputModel.Natlidexpirationdate;
                                inputModel.CurrentCity = mainModel.diligenceInputModel.CurrentCity;
                                inputModel.CurrentCountry = mainModel.diligenceInputModel.CurrentCountry;
                                inputModel.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                                inputModel.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
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
                               .Where(u => u.record_Id == recordid)
                                              .FirstOrDefault();
                                strupdate.record_Id = recordid;
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
                                strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                                strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                                strupdate.Has_Companion_Report = mainModel.otherdetails.Has_Companion_Report;
                                strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
                                strupdate.Media_Based_Hits = mainModel.otherdetails.Media_Based_Hits;
                                _context.Entry(strupdate).State = EntityState.Modified;
                                _context.SaveChanges();
                            }
                            catch
                            {
                                Otherdatails strupdate = new Otherdatails();
                                strupdate.record_Id = recordid;
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
                                strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                                strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                                strupdate.Has_Companion_Report = mainModel.otherdetails.Has_Companion_Report;
                                strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
                                strupdate.Media_Based_Hits = mainModel.otherdetails.Media_Based_Hits;
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
                                          .Where(a => a.record_Id == recordid)
                                          .ToList();
                                if (employerModel1.Count==0 || employerModel1 == null) {
                                    for (int i = 0; i < mainModel.diligence.employerList.Count; i++)
                                    {
                                        EmployerModel employerModel = new EmployerModel();
                                        employerModel.record_Id = recordid;
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
                                        employerModel1[i].record_Id = recordid;
                                        employerModel1[i].Emp_Employer = mainModel.EmployerModel[i].Emp_Employer;
                                        employerModel1[i].Emp_Location = mainModel.EmployerModel[i].Emp_Location;
                                        employerModel1[i].Emp_Position = mainModel.EmployerModel[i].Emp_Position;
                                        employerModel1[i].Emp_Confirmed = mainModel.EmployerModel[i].Emp_Confirmed;
                                        employerModel1[i].Emp_Status = mainModel.EmployerModel[i].Emp_Status;
                                        try 
                                        {
                                            if (mainModel.EmployerModel[i].Emp_StartDateDay == "")
                                            {
                                                employerModel1[i].Emp_StartDateDay = "Day";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_StartDateDay = mainModel.EmployerModel[i].Emp_StartDateDay;
                                            }
                                        }
                                        catch {
                                            employerModel1[i].Emp_StartDateDay = "Day";
                                        }                                        
                                        employerModel1[i].Emp_StartDateMonth = mainModel.EmployerModel[i].Emp_StartDateMonth;
                                        try { 
                                        if (mainModel.EmployerModel[i].Emp_StartDateYear == "")
                                            {
                                                employerModel1[i].Emp_StartDateYear = "Year";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_StartDateYear = mainModel.EmployerModel[i].Emp_StartDateYear;
                                            }
                                        }
                                        catch {
                                            employerModel1[i].Emp_StartDateYear = "Year";
                                        }
                                        try {
                                            if (mainModel.EmployerModel[i].Emp_EndDateDay == "")
                                            {
                                                employerModel1[i].Emp_EndDateDay = "Day";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_EndDateDay = mainModel.EmployerModel[i].Emp_EndDateDay;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel1[i].Emp_EndDateDay = "Day";
                                        }
                                        employerModel1[i].Emp_EndDateMonth = mainModel.EmployerModel[i].Emp_EndDateMonth;
                                        try { if (mainModel.EmployerModel[i].Emp_EndDateYear == "")
                                            {
                                                employerModel1[i].Emp_EndDateYear = "Year";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_EndDateYear = mainModel.EmployerModel[i].Emp_EndDateYear;
                                            }
                                            } catch {
                                            employerModel1[i].Emp_EndDateYear = "Year";
                                        }                                        
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
                                                    employerModel.record_Id = recordid;
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
                                    employerModel.record_Id = recordid;
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
                    .Where(u => u.record_Id == recordid)
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
                                   .Where(a => a.record_Id == recordid)
                                   .ToList();
                                if (educationModel1.Count==0 || educationModel1 == null)
                                {
                                    for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                    {
                                        EducationModel educationModel = new EducationModel();
                                        educationModel.record_Id = recordid;
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
                                        educationModel1[i].record_Id = recordid;
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
                                                if (educationModel1[0].Edu_History == "No")
                                                {
                                                    TempData["message"] = "Cannot add additional details in Educational section as the Education history is selected as No in first row.";
                                                }
                                                else
                                                {
                                                    EducationModel educationModel = new EducationModel();
                                                    educationModel.record_Id = recordid;
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
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                {
                                    EducationModel educationModel = new EducationModel();
                                    educationModel.record_Id = recordid;
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
                              .Where(u => u.record_Id == recordid)
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
                                   .Where(a => a.record_Id == recordid)
                                   .ToList();
                                if (pLLicenseModel.Count == 0 || pLLicenseModel == null) {
                                    for (int i = 0; i < mainModel.diligence.plLicenseList.Count; i++)
                                    {
                                        PLLicenseModel pLLicenseModel1 = new PLLicenseModel();
                                        pLLicenseModel1.General_PL_License = mainModel.diligence.plLicenseList[i].General_PL_License;
                                        pLLicenseModel1.PL_Confirmed = mainModel.diligence.plLicenseList[i].PL_Confirmed;
                                        pLLicenseModel1.PL_License_Type = mainModel.diligence.plLicenseList[i].PL_License_Type;
                                        pLLicenseModel1.PL_Location = mainModel.diligence.plLicenseList[i].PL_Location;
                                        pLLicenseModel1.PL_Number = mainModel.diligence.plLicenseList[i].PL_Number;
                                        pLLicenseModel1.PL_Organization = mainModel.diligence.plLicenseList[i].PL_Organization;
                                        pLLicenseModel1.record_Id = recordid;
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
                                        pLLicenseModel[i].record_Id = recordid;
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
                                                if (pLLicenseModel[0].General_PL_License == "No")
                                                {
                                                    TempData["message"] = "Cannot add additional details in PLLicense section as the General PL License is selected as No in first row.";
                                                }
                                                else
                                                {
                                                    PLLicenseModel pLLicenseModel1 = new PLLicenseModel();
                                                    pLLicenseModel1.General_PL_License = mainModel.diligence.plLicenseList[i].General_PL_License;
                                                    pLLicenseModel1.PL_Confirmed = mainModel.diligence.plLicenseList[i].PL_Confirmed;
                                                    pLLicenseModel1.PL_License_Type = mainModel.diligence.plLicenseList[i].PL_License_Type;
                                                    pLLicenseModel1.PL_Location = mainModel.diligence.plLicenseList[i].PL_Location;
                                                    pLLicenseModel1.PL_Number = mainModel.diligence.plLicenseList[i].PL_Number;
                                                    pLLicenseModel1.PL_Organization = mainModel.diligence.plLicenseList[i].PL_Organization;
                                                    pLLicenseModel1.record_Id = recordid;
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
                                    }
                                    catch { }
                                }
                                mainModel.pllicenseModels = _context.DbPLLicense
                                   .Where(u => u.record_Id == recordid)
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
                                    pLLicenseModel.record_Id = recordid;
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
                                  .Where(u => u.record_Id == recordid)
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
                       DiligenceInputModel diligence = _context.DbPersonalInfo
                       .Where(a => a.record_Id == recordid)
                       .FirstOrDefault();
                                diligence.Country = mainModel.diligenceInputModel.Country;
                                _context.DbPersonalInfo.Update(diligence);
                        _context.SaveChanges();
                    }
                    catch
                    {
                        DiligenceInputModel diligence = new DiligenceInputModel();
                        diligence.record_Id = recordid;
                        diligence.Country = mainModel.diligenceInputModel.Country;
                        _context.DbPersonalInfo.Add(diligence);
                        _context.SaveChanges();
                    }
                    CountrySpecificModel countrySpecific = _context.CSComment
                           .Where(u => u.record_Id == recordid)
                                          .FirstOrDefault();
                    try
                    {                        
                        countrySpecific.record_Id = recordid;
                    }
                    catch
                    {
                        countrySpecific = new CountrySpecificModel();
                        countrySpecific.record_Id = recordid;
                    }
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
                                countrySpecific.record_Id = recordid;
                                _context.CSComment.Add(countrySpecific);
                                _context.SaveChanges();
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
                           .Where(u => u.record_Id == recordid)
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
                            else {
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
                            .Where(u => u.record_Id == recordid)
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
                            mainModel.educationModels = _context.DbEducation
                            .Where(u => u.record_Id == recordid)
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
                            .Where(u => u.record_Id == recordid)
                            .ToList();
                        }
                    }
                    catch { }
                    break;
            }
            HttpContext.Session.SetString("recordid", recordid);
            
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
            HttpContext.Session.SetString("recordid", recordid);            
            ViewBag.CaseNumber = casenumber;
            ViewBag.LastName = lastname;
            return View(mainModel);
        }
        public IActionResult TemplateSelection()
        {            
            return View();
        }
        public IActionResult TemplateredirectionPage(TemplateSelection templateSelection)
        {
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
                ReportModel report1;
                try
                {
                    report1 = _context.reportModel.Where(x => x.lastname == templateSelection.LastName && x.casenumber == templateSelection.CaseNumber).FirstOrDefault();
                }
                catch
                {
                    report1 = null;
                }
                if (report1 == null)
                {

                }
                else
                {
                    // _logger.LogInformation("Login failed");
                    ModelState.AddModelError("LogOnError", "Filename :" + templateSelection.LastName.ToString() + "_SterlingDiligenceReport(" + templateSelection.CaseNumber.ToString() + ")_DRAFT.docx already exists");
                    TempData["message"] = "Filename: " + templateSelection.LastName.ToString() + "_SterlingDiligenceReport(" + templateSelection.CaseNumber.ToString() + ")_DRAFT.docx already exists";
                    return RedirectToAction("TemplateSelection", "Diligence");
                }
            //ViewBag.CaseNumber = templateSelection.CaseNumber;
            var recordid = Guid.NewGuid().ToString();
            recordid = recordid.Replace("-", "");
                ReportModel report = new ReportModel();
                report.record_Id = recordid;                
                report.lastname = templateSelection.LastName;
                report.casenumber = templateSelection.CaseNumber;
                report.TemplateType = templateSelection.templatetype;
                report.createdby = Environment.UserName;
                _context.reportModel.Add(report);
                _context.SaveChanges();
            string straction = "";
            string strcontroller = "";
            switch (templateSelection.templatetype.ToString())
            {
                case "International DD Individual":
                    straction = "DiligenceForm_Page1";
                    strcontroller = "Diligence";                    
                    break;
                case "US DD Individual":
                    straction = "US_DD_Individual_page1";
                    strcontroller = "USDDIndividual";
                    break;
                case "US PE Individual":
                    straction= "US_PE_Individual_page1";
                    strcontroller = "USPEIndividual";
                    break;
                case "International Citizenship Individual":
                    straction = "INTL_Citizen_Individualpage1";
                    strcontroller = "IN_CITI_Individual";
                    break;
                case "US Company":
                    straction = "US_COMPANY_page1";
                    strcontroller = "USCOMPANY";
                    break;
                case "International Company":
                    straction = "IN_COMPANY_page1";
                    strcontroller = "INCOMPANY";
                    break;
                case "International Citizenship Family":
                    straction = "INTL_Citizen_Family_Newpage";
                    strcontroller = "IN_CITI_Family";
                    FamilyModel family = new FamilyModel();
                    family.record_Id = recordid;
                    family.last_name = templateSelection.LastName;                    
                    family.case_number = templateSelection.CaseNumber;
                    family.applicant_type = "Main Applicant";
                    family.Family_record_id = Guid.NewGuid().ToString();
                    _context.familyModels.Add(family);
                    _context.SaveChanges();
                    break;
                case "International PE Individual":
                    straction = "IN_PE_Individual_Edit";
                    strcontroller = "IN_PE_Individual";
                    break;
                case "US Citizenship Family":
                    straction = "US_Citizen_Family_Newpage";
                    strcontroller = "US_CITI_Family";
                    FamilyModel family1 = new FamilyModel();
                    family1.record_Id = recordid;
                    family1.last_name = templateSelection.LastName;
                    family1.case_number = templateSelection.CaseNumber;
                    family1.applicant_type = "Main Applicant";
                    family1.Family_record_id = Guid.NewGuid().ToString();
                    _context.familyModels.Add(family1);
                    _context.SaveChanges();
                    break;
            }
            HttpContext.Session.SetString("recordid", recordid);
            return RedirectToAction(straction, strcontroller);
        }
        [HttpGet]
        public IActionResult DiligenceForm_Page1()
        {            
            //Country List
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
            string recordid= HttpContext.Session.GetString("recordid");
            ReportModel comment12 = _context.reportModel
                               .Where(u => u.record_Id == recordid)
                               .FirstOrDefault();
            ViewBag.CaseNumber = comment12.casenumber;
            ViewBag.LastName = comment12.lastname;            
            HttpContext.Session.SetString("recordid", recordid);
            HttpContext.Session.SetString("lastname", comment12.lastname);            
            HttpContext.Session.SetString("casenumber", comment12.casenumber);            
            return View();
        }
        [HttpPost]     
        public IActionResult DiligenceForm_Page1(MainModel mainModel, string SaveData,string Submit)
        {
            string recordid = HttpContext.Session.GetString("recordid");
            if (Submit == "SubmitData")
            {
                MainModel main = new MainModel();
                DiligenceInputModel dinput = new DiligenceInputModel();
                Otherdatails otherdatails = new Otherdatails();
                CountrySpecificModel countrySpecificModel = new CountrySpecificModel();
                EmployerModel employer = new EmployerModel();
                EducationModel education = new EducationModel();
                PLLicenseModel pL = new PLLicenseModel();
                SummaryResulttableModel summary = new SummaryResulttableModel();
                try
                {
                    dinput = _context.DbPersonalInfo
                                      .Where(u => u.record_Id == recordid)
                                      .FirstOrDefault();
                    if (dinput == null)
                    {
                        TempData["message"] = "Please enter the details in personal details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in personal details section";
                }
                try
                {
                    otherdatails = _context.othersModel
                               .Where(u => u.record_Id == recordid)
                               .FirstOrDefault();
                    if (otherdatails == null)
                    {
                        TempData["message"] = "Please enter the details in other details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in other details section";
                }
                try
                {
                    countrySpecificModel = _context.CSComment
                                 .Where(u => u.record_Id == recordid)
                                 .FirstOrDefault();
                    if (countrySpecificModel == null)
                    {
                        TempData["message"] = "Please enter the details in country specific section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in country specific section";
                }
                try
                {
                    employer = _context.DbEmployer
                                .Where(u => u.record_Id == recordid)
                                .FirstOrDefault();
                    if (employer == null)
                    {
                        TempData["message"] = "Please enter the details in Employee details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in Employee details section";
                }
                try
                {
                    education = _context.DbEducation
                               .Where(u => u.record_Id == recordid)
                               .FirstOrDefault();
                    if (education == null)
                    {
                        TempData["message"] = "Please enter the details in Education details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in Education details section";
                }
                try
                {
                    pL = _context.DbPLLicense
                               .Where(u => u.record_Id == recordid)
                                .FirstOrDefault();
                    if (pL == null)
                    {
                        TempData["message"] = "Please enter the details in Pllicence details section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in Pllicence details section";
                }
                try
                {
                    summary = _context.summaryResulttableModels
                                      .Where(u => u.record_Id == recordid)
                                      .FirstOrDefault();
                    if (summary == null)
                    {
                        TempData["message"] = "Please enter the details in summary result table section";
                    }
                }
                catch
                {
                    TempData["message"] = "Please enter the details in summary result table section";
                }
                if (TempData["message"] == null)
                {
                    main.diligenceInputModel = dinput;
                    main.educationModels = _context.DbEducation
                             .Where(u => u.record_Id == recordid)
                             .ToList();
                    main.otherdetails = otherdatails;
                    main.EmployerModel = _context.DbEmployer
                                .Where(u => u.record_Id == recordid)
                                .ToList();
                    main.pllicenseModels = _context.DbPLLicense
                               .Where(u => u.record_Id == recordid)
                                .ToList();
                    main.csModel = countrySpecificModel;
                    main.summarymodel = summary;
                    return Save_Page1(main);
                }
            }
            string lastname = HttpContext.Session.GetString("lastname");
            //string firstname = HttpContext.Session.GetString("firstname");
            string casenumber = HttpContext.Session.GetString("casenumber");
            //string middleinitial = HttpContext.Session.GetString("middleinitial");
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
                           .Where(u => u.record_Id == recordid)
                                          .FirstOrDefault();
                                inputModel.record_Id = recordid;
                                inputModel.ClientName = mainModel.diligenceInputModel.ClientName;
                                inputModel.Dob = mainModel.diligenceInputModel.Dob;
                                inputModel.FirstName = mainModel.diligenceInputModel.FirstName;
                                inputModel.LastName = lastname;
                                inputModel.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                                inputModel.MiddleName = mainModel.diligenceInputModel.MiddleName;
                                inputModel.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                                inputModel.CaseNumber = casenumber;
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
                                inputModel.CurrentCity = mainModel.diligenceInputModel.CurrentCity;
                                inputModel.CurrentCountry = mainModel.diligenceInputModel.CurrentCountry;
                                inputModel.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                                inputModel.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
                                _context.Entry(inputModel).State = EntityState.Modified;
                                _context.SaveChanges();
                            }
                            catch
                            {
                                DiligenceInputModel inputModel = new DiligenceInputModel();
                                inputModel.record_Id = recordid;
                                inputModel.ClientName = mainModel.diligenceInputModel.ClientName;
                                inputModel.Dob = mainModel.diligenceInputModel.Dob;
                                inputModel.FirstName = mainModel.diligenceInputModel.FirstName;
                                inputModel.LastName = lastname;
                                inputModel.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                                inputModel.MiddleName = mainModel.diligenceInputModel.MiddleName;
                                inputModel.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                                inputModel.CaseNumber = casenumber;
                                inputModel.MaidenName = mainModel.diligenceInputModel.MaidenName;
                                inputModel.FullaliasName = mainModel.diligenceInputModel.FullaliasName;
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Dob = mainModel.diligenceInputModel.Dob;
                                inputModel.Pob = mainModel.diligenceInputModel.Pob;
                                inputModel.City = mainModel.diligenceInputModel.City;
                                inputModel.Nationality = mainModel.diligenceInputModel.Nationality;
                                inputModel.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                                inputModel.Natlidissuedate = mainModel.diligenceInputModel.Natlidissuedate;
                                inputModel.Natlidexpirationdate = mainModel.diligenceInputModel.Natlidexpirationdate;
                                inputModel.CurrentCity = mainModel.diligenceInputModel.CurrentCity;
                                inputModel.CurrentCountry = mainModel.diligenceInputModel.CurrentCountry;
                                inputModel.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                                inputModel.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                                inputModel.Country = mainModel.diligenceInputModel.Country;
                                inputModel.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
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
                               .Where(u => u.record_Id == recordid)
                                              .FirstOrDefault();
                                strupdate.record_Id = recordid;
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
                                strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                                strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                                strupdate.Has_Companion_Report = mainModel.otherdetails.Has_Companion_Report;
                                strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
                                strupdate.Media_Based_Hits = mainModel.otherdetails.Media_Based_Hits;

                                _context.Entry(strupdate).State = EntityState.Modified;
                                _context.SaveChanges();
                            }
                            catch
                            {
                                Otherdatails strupdate = new Otherdatails();
                                strupdate.record_Id = recordid;
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
                                strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                                strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                                strupdate.Has_Companion_Report = mainModel.otherdetails.Has_Companion_Report;
                                strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
                                strupdate.Media_Based_Hits = mainModel.otherdetails.Media_Based_Hits;
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
                                          .Where(a => a.record_Id == recordid)
                                          .ToList();
                                if (employerModel1.Count == 0 || employerModel1 == null)
                                {
                                    for (int i = 0; i < mainModel.diligence.employerList.Count; i++)
                                    {
                                        EmployerModel employerModel = new EmployerModel();
                                        employerModel.record_Id = recordid;
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
                                        employerModel1[i].record_Id = recordid;
                                        employerModel1[i].Emp_Employer = mainModel.EmployerModel[i].Emp_Employer;
                                        employerModel1[i].Emp_Location = mainModel.EmployerModel[i].Emp_Location;
                                        employerModel1[i].Emp_Position = mainModel.EmployerModel[i].Emp_Position;
                                        employerModel1[i].Emp_Confirmed = mainModel.EmployerModel[i].Emp_Confirmed;
                                        employerModel1[i].Emp_Status = mainModel.EmployerModel[i].Emp_Status;
                                        try
                                        {
                                            if (mainModel.EmployerModel[i].Emp_StartDateDay == "")
                                            {
                                                employerModel1[i].Emp_StartDateDay = "Day";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_StartDateDay = mainModel.EmployerModel[i].Emp_StartDateDay;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel1[i].Emp_StartDateDay = "Day";
                                        }
                                        employerModel1[i].Emp_StartDateMonth = mainModel.EmployerModel[i].Emp_StartDateMonth;
                                        try
                                        {
                                            if (mainModel.EmployerModel[i].Emp_StartDateYear == "")
                                            {
                                                employerModel1[i].Emp_StartDateYear = "Year";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_StartDateYear = mainModel.EmployerModel[i].Emp_StartDateYear;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel1[i].Emp_StartDateYear = "Year";
                                        }
                                        try
                                        {
                                            if (mainModel.EmployerModel[i].Emp_EndDateDay == "")
                                            {
                                                employerModel1[i].Emp_EndDateDay = "Day";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_EndDateDay = mainModel.EmployerModel[i].Emp_EndDateDay;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel1[i].Emp_EndDateDay = "Day";
                                        }
                                        employerModel1[i].Emp_EndDateMonth = mainModel.EmployerModel[i].Emp_EndDateMonth;
                                        try
                                        {
                                            if (mainModel.EmployerModel[i].Emp_EndDateYear == "")
                                            {
                                                employerModel1[i].Emp_EndDateYear = "Year";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_EndDateYear = mainModel.EmployerModel[i].Emp_EndDateYear;
                                            }
                                        }
                                        catch
                                        {
                                            employerModel1[i].Emp_EndDateYear = "Year";
                                        }
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
                                                    EmployerModel employerModel = new EmployerModel();
                                                    employerModel.record_Id = recordid;
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
                                    catch { }
                                }
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.employerList.Count; i++)
                                {
                                    EmployerModel employerModel = new EmployerModel();
                                    employerModel.record_Id = recordid;
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
                    .Where(u => u.record_Id == recordid)
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
                                   .Where(a => a.record_Id == recordid)
                                   .ToList();
                                if (educationModel1.Count == 0 || educationModel1 == null)
                                {
                                    for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                    {
                                        EducationModel educationModel = new EducationModel();
                                        educationModel.record_Id = recordid;
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
                                        educationModel1[i].record_Id = recordid;
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
                                                if (educationModel1[0].Edu_History == "No")
                                                {
                                                    TempData["message"] = "Cannot add additional details in Educational section as the Education history is selected as No in first row.";
                                                }
                                                else
                                                {
                                                    EducationModel educationModel = new EducationModel();
                                                    educationModel.record_Id = recordid;
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
                                    }
                                    catch { }
                                }
                            }
                            catch
                            {
                                for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                {
                                    EducationModel educationModel = new EducationModel();
                                    educationModel.record_Id = recordid;
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
                              .Where(u => u.record_Id == recordid)
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
                                   .Where(a => a.record_Id == recordid)
                                   .ToList();
                                if (pLLicenseModel.Count == 0 || pLLicenseModel == null)
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
                                        pLLicenseModel1.record_Id = recordid;
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
                                        pLLicenseModel[i].record_Id = recordid;
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
                                                if (pLLicenseModel[0].General_PL_License == "No")
                                                {
                                                    TempData["message"] = "Cannot add additional details in PLLicense section as the General PL License is selected as No in first row.";
                                                }
                                                else
                                                {
                                                    PLLicenseModel pLLicenseModel1 = new PLLicenseModel();
                                                    pLLicenseModel1.General_PL_License = mainModel.diligence.plLicenseList[i].General_PL_License;
                                                    pLLicenseModel1.PL_Confirmed = mainModel.diligence.plLicenseList[i].PL_Confirmed;
                                                    pLLicenseModel1.PL_License_Type = mainModel.diligence.plLicenseList[i].PL_License_Type;
                                                    pLLicenseModel1.PL_Location = mainModel.diligence.plLicenseList[i].PL_Location;
                                                    pLLicenseModel1.PL_Number = mainModel.diligence.plLicenseList[i].PL_Number;
                                                    pLLicenseModel1.PL_Organization = mainModel.diligence.plLicenseList[i].PL_Organization;
                                                    pLLicenseModel1.record_Id = recordid;
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
                                    }
                                    catch { }
                                }
                                mainModel.pllicenseModels = _context.DbPLLicense
                                   .Where(u => u.record_Id == recordid)
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
                                    pLLicenseModel.record_Id = recordid;
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
                                  .Where(u => u.record_Id == recordid)
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
                                DiligenceInputModel diligence = _context.DbPersonalInfo
                                .Where(a => a.record_Id == recordid)
                                .FirstOrDefault();
                                diligence.Country = mainModel.diligenceInputModel.Country;
                                _context.DbPersonalInfo.Update(diligence);
                                _context.SaveChanges();
                            }
                            catch
                            {
                                DiligenceInputModel diligence = new DiligenceInputModel();
                                diligence.record_Id = recordid;
                                diligence.Country = mainModel.diligenceInputModel.Country;
                                _context.DbPersonalInfo.Add(diligence);
                                _context.SaveChanges();
                            }
                            CountrySpecificModel countrySpecific = _context.CSComment
                                   .Where(u => u.record_Id == recordid)
                                                  .FirstOrDefault();
                            try
                            {
                                countrySpecific.record_Id = recordid;
                            }
                            catch
                            {
                                countrySpecific = new CountrySpecificModel();
                                countrySpecific.record_Id = recordid;
                            }
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
                                countrySpecific.record_Id = recordid;
                                _context.CSComment.Add(countrySpecific);
                                _context.SaveChanges();
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
                           .Where(u => u.record_Id == recordid)
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
                            .Where(u => u.record_Id == recordid)
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
                            .Where(u => u.record_Id == recordid)
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
                            .Where(u => u.record_Id == recordid)
                            .ToList();
                        }
                    }
                    catch { }
                    break;
            }
            HttpContext.Session.SetString("recordid", recordid);       
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
            HttpContext.Session.SetString("recordid", recordid);
            ViewBag.CaseNumber = casenumber;
            ViewBag.LastName = lastname;
            return View(mainModel);
        }
        public IActionResult DiligenceForm_Page2()
        {
            string record_Id = HttpContext.Session.GetString("record_Id");            
            return View();
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
        public IActionResult Save_Page1(MainModel diligenceInput)
        {                      
            string templatePath;
            string savePath = _config.GetValue<string>("ReportPath");
            if(diligenceInput.diligenceInputModel.Country.ToString().Equals("France") || diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom") || diligenceInput.diligenceInputModel.Country.ToString().Equals("Canada") || diligenceInput.diligenceInputModel.Country.ToString().Equals("Australia")) 
            {
                templatePath = _config.GetValue<string>("templatePath");
            }
            else
            {                
                templatePath = _config.GetValue<string>("templatePath2");
            }            
            string pathTo = _config.GetValue<string>("OlderReport"); // the destination file name would be appended later
            savePath = string.Concat(savePath,diligenceInput.diligenceInputModel.LastName.ToString(), "_SterlingDiligenceReport(", diligenceInput.diligenceInputModel.CaseNumber.ToString(), ")_DRAFT.docx");
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
            Document doc = new Document(templatePath);
            string current_full_address = "";
            string currentstreet;
            string currentcity;
            string currentcountry;
            string strblnres = "";
            if(diligenceInput.diligenceInputModel.Country.ToString().Equals("Hong Kong SAR"))
            {
                diligenceInput.diligenceInputModel.Country = "Hong Kong";
            }
            try
            {
                if (diligenceInput.diligenceInputModel.FullaliasName.ToString().Equals("") || diligenceInput.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("NA") || diligenceInput.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("N/A"))
                {
                }
                else
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
                                        // Find the word "Spire.Doc" in paragraph1
                                        if (abc.ToString().Equals("“the subject”)") || abc.ToString().EndsWith("subject”)"))
                                        {
                                            Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                            footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                            footnote1.MarkerCharacterFormat.FontSize = 11;
                                            //Insert footnote1 after the word "Spire.Doc"
                                            para.ChildObjects.Insert(i + 1, footnote1);
                                            TextRange _text = footnote1.TextBody.AddParagraph().AppendText("It is noted that the subject was identified in connection with the name variation FullAliasName and investigative efforts were undertaken in connection with the same, as appropriate.");
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
                    }catch{ }
                    doc.Replace("FullAliasName", diligenceInput.diligenceInputModel.FullaliasName, true, true);
                }
            }
            catch
            {

            }
            try
            {
                if (diligenceInput.diligenceInputModel.foreignlanguage.ToString().Equals(""))
                {
                    doc.Replace("[FORSUB]", "", true, false);
                    doc.Replace("[FORSUBJECT]", "", true, false);
                }
                else
                {
                    string strforsub = string.Concat("\n", diligenceInput.diligenceInputModel.foreignlanguage);
                    doc.Replace("[FORSUB]", strforsub, true, false);
                    strforsub = string.Concat(diligenceInput.diligenceInputModel.foreignlanguage, " ");
                    doc.Replace("[FORSUBJECT]", strforsub, true, false);
                }
            }
            catch
            {
                doc.Replace("[FORSUB]", "", true, false);
                doc.Replace("[FORSUBJECT]", "", true, false);
            }
            doc.Replace("First_Name", diligenceInput.diligenceInputModel.FirstName.ToUpper().TrimEnd().ToString(), true, true);            
            doc.Replace("Last_Name", diligenceInput.diligenceInputModel.LastName.ToUpper().ToString().TrimEnd(), true, true);
            doc.Replace("FirstName", diligenceInput.diligenceInputModel.FirstName.TrimEnd(), true, true);
           
            doc.Replace("ClientName", diligenceInput.diligenceInputModel.ClientName.TrimEnd(), true, true);
            doc.Replace("LastName", diligenceInput.diligenceInputModel.LastName.TrimEnd(), true, true);
            doc.Replace("Position1", diligenceInput.EmployerModel[0].Emp_Position.TrimEnd(), true, true);
            doc.Replace("Employer1", diligenceInput.EmployerModel[0].Emp_Employer.TrimEnd(), true, true);                        
            doc.Replace("ClientName", diligenceInput.diligenceInputModel.ClientName.TrimEnd(), true, true);
            try
            {
                doc.Replace("DateofBirth", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Month) + " " + Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Year, true, true);
            }
            catch
            {
                doc.Replace("DateofBirth", "<not provided>", true, true);
            }
            try
            {
                if (diligenceInput.diligenceInputModel.City.ToString().ToUpper().Equals("NA") || diligenceInput.diligenceInputModel.City.ToString().ToUpper().Equals("N/A") || diligenceInput.diligenceInputModel.City.ToString().Equals(""))
                {
                    doc.Replace("[City], ", "", true, false);
                    doc.SaveToFile(savePath);
                    doc.Replace("[City]", "<not provided>", true, true);
                }
                else
                {
                    doc.Replace("[City]", diligenceInput.diligenceInputModel.City.ToString(), true, true);
                }
            }
            catch
            {
                doc.Replace("[City], ", "", true, false);
                doc.SaveToFile(savePath);
                doc.Replace("[City]", "<not provided>", true, true);
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
            try
            {
                if (diligenceInput.diligenceInputModel.CurrentStreet.ToString().Equals(""))
                {
                    currentstreet = "";
                }
                else
                {
                    if (diligenceInput.diligenceInputModel.CurrentCity.ToString().Equals("") && diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("") && diligenceInput.diligenceInputModel.CurrentZipcode.ToString().Equals(""))
                    {
                        currentstreet = string.Concat(diligenceInput.diligenceInputModel.CurrentStreet.ToString());
                    }
                    else
                    {
                        currentstreet = string.Concat(diligenceInput.diligenceInputModel.CurrentStreet.ToString(), ", ");
                    }

                }
            }
            catch
            {
                currentstreet = "";
            }
            try
            {
                if (diligenceInput.diligenceInputModel.CurrentCity.ToString().Equals(""))
                {
                    currentcity = "";
                }
                else
                {
                    try
                    {
                        if (diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("") && diligenceInput.diligenceInputModel.CurrentZipcode.ToString().Equals(""))
                        {
                            currentcity = diligenceInput.diligenceInputModel.CurrentCity.ToString();
                        }
                        else
                        {
                            currentcity = string.Concat(diligenceInput.diligenceInputModel.CurrentCity.ToString(), ", ");
                        }
                    }
                    catch
                    {
                        currentcity = diligenceInput.diligenceInputModel.CurrentCity.ToString();
                    }
                }
            }
            catch
            {
                currentcity = "";
            }
            if (diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals(""))
            {
                currentcountry = "";
            }
            else
            {
                if (diligenceInput.diligenceInputModel.CurrentZipcode.ToString().Equals("")) { currentcountry = diligenceInput.diligenceInputModel.CurrentCountry.ToString(); }
                else
                {
                    currentcountry = string.Concat(diligenceInput.diligenceInputModel.CurrentCountry.ToString(), ", ");
                }
            }
            current_full_address = string.Concat(currentstreet, currentcity, currentcountry, diligenceInput.diligenceInputModel.CurrentZipcode.ToString());
            if (current_full_address.ToString().Equals("")) { doc.Replace("CurrentFullAddress", "<not provided>", true, true); }
            else
            {
                doc.Replace("CurrentFullAddress", current_full_address, true, true);
            }
            //doc.Replace("FullAliasName", diligenceInput.FullaliasName, true, true);            
            doc.Replace("[CaseNumber]", diligenceInput.diligenceInputModel.CaseNumber.TrimEnd(), true, true);
            try
            {
                if (diligenceInput.diligenceInputModel.MiddleName.ToString() == "") 
                {
                    doc.Replace(" MiddleName","", false, false);
                }
                else
                {
                    doc.Replace("MiddleName", diligenceInput.diligenceInputModel.MiddleName.ToString().TrimEnd(), true, true);
                }
            }
            catch {
                doc.Replace(" MiddleName", "", false, false);
            }
            doc.Replace("Nationality", diligenceInput.diligenceInputModel.Nationality.ToString().TrimEnd(), true, true);
            doc.SaveToFile(savePath);

            //PROPERTY RECORDS
            switch (diligenceInput.otherdetails.Has_Property_Records)
            {
                case "Yes – Single Record":
                    doc.Replace("PROPERTYRECHITSDESCRIPTION", "Research efforts conducted through records maintained by the [Property Registry] revealed that <investigator to insert results here>.", true, true);
                    doc.Replace("PROPRECRESULT", "Record", true, true);
                    doc.Replace("PROPRECCOMMENT", "[LastName] was identified as an owner of at least <number> of properties in [Country]", true, true);
                    break;
                case "Yes – Multiple Records":
                    doc.Replace("PROPERTYRECHITSDESCRIPTION", "Research efforts conducted through records maintained by the [Property Registry] revealed that <investigator to insert results here>.", true, true);
                    doc.Replace("PROPRECRESULT", "Records", true, true);
                    doc.Replace("PROPRECCOMMENT", "[LastName] was identified as an owner of at least <number> of properties in [Country]", true, true);
                    break;
                case "N/A":
                    doc.Replace("PROPERTYRECHITSDESCRIPTION", "Information relating to real property records in [Country] is unavailable.", true, true);
                    doc.Replace("PROPRECRESULT", "N/A", true, true);
                    doc.Replace("PROPRECCOMMENT", "", true, true);
                    break;
                case "No":
                    doc.Replace("PROPERTYRECHITSDESCRIPTION", "Research efforts conducted through records maintained by the [Property Registry] did not reveal any property records in connection with the subject.", true, true);
                    doc.Replace("PROPRECRESULT", "No Records", true, true);
                    doc.Replace("PROPRECCOMMENT", "", true, true);
                    break;
            }           
            doc.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString().TrimEnd(),true,true);
            doc.SaveToFile(savePath);
            //Employee details
            string emp_Comment = "";
            string strempstartdate = "";
            string strempenddate = "";
            for (int i = 0; i < diligenceInput.EmployerModel.Count; i++)
            {               

                if (!diligenceInput.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && !diligenceInput.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                {                                      
                   strempstartdate = diligenceInput.EmployerModel[i].Emp_StartDateMonth + " " + diligenceInput.EmployerModel[i].Emp_StartDateDay + ", " + diligenceInput.EmployerModel[i].Emp_StartDateYear;                    
                }
                if (diligenceInput.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && diligenceInput.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                {
                    strempstartdate = "<not provided>";
                }
                if (diligenceInput.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && !diligenceInput.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year") )
                {
                    strempstartdate = diligenceInput.EmployerModel[i].Emp_StartDateMonth + " " + diligenceInput.EmployerModel[i].Emp_StartDateYear;
                    //employerModel.Emp_StartDate = strempstartdate;
                }
                if (diligenceInput.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && diligenceInput.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year") )
                {
                    strempstartdate = diligenceInput.EmployerModel[i].Emp_StartDateYear;
                    
                }               
                if (i == 0 && diligenceInput.EmployerModel[0].Emp_Status.ToString().Equals("Current") ) { }
                else
                {
                    if (diligenceInput.EmployerModel[i].Emp_Status.ToString().Equals("Concurrent")) { }
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
                }                                
                if (i == 0)
                {
                    CommentModel comment = _context.DbComment
                          .Where(u => u.Comment_type == "Emp1")
                          .FirstOrDefault();
                    if (diligenceInput.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                    {
                       emp_Comment = comment.confirmed_comment.ToString();
                       emp_Comment = emp_Comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
                       emp_Comment = emp_Comment.Replace("[Position1]", diligenceInput.EmployerModel[0].Emp_Position.ToString());
                       emp_Comment = emp_Comment.Replace("[Employer1]", diligenceInput.EmployerModel[0].Emp_Employer.ToString());
                       emp_Comment = emp_Comment.Replace("[EmpLocation1]", diligenceInput.EmployerModel[0].Emp_Location.ToString());
                       emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                       emp_Comment = string.Concat(emp_Comment, "\n");
                    }
                    else
                    {
                       emp_Comment = comment.unconfirmed_comment.ToString();
                       emp_Comment = emp_Comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
                       emp_Comment = emp_Comment.Replace("[Position1]", diligenceInput.EmployerModel[0].Emp_Position.ToString());
                       emp_Comment = emp_Comment.Replace("[Employer1]", diligenceInput.EmployerModel[0].Emp_Employer.ToString());
                       emp_Comment = emp_Comment.Replace("[EmpLocation1]", diligenceInput.EmployerModel[0].Emp_Location.ToString());
                       emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                       emp_Comment = string.Concat(emp_Comment, "\n");
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
                                            TextRange _text = footnote1.TextBody.AddParagraph().AppendText("<Investigator to insert company registry detail here as well as indicate whether a Companion Report has been prepared for the firm OR that additional research efforts would be required in connection with this firm>.");
                                            footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                            _text.CharacterFormat.FontName = "Calibri (Body)";
                                            _text.CharacterFormat.FontSize = 9;
                                            //Append the line
                                            string stremp1textappended = "";
                                            if (diligenceInput.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                                            {
                                                stremp1textappended = " in [EmpLocation1], since [EmpStartDate1]. EMPLOYEEDESCRIPTION";
                                            }
                                            else
                                            {
                                                stremp1textappended = " in [EmpLocation1], since [EmpStartDate1]. EMPLOYEEDESCRIPTION";
                                            }
                                            doc.Replace("[Employer1Footnote]", diligenceInput.EmployerModel[0].Emp_Employer.ToString(), true, true);
                                            stremp1textappended = stremp1textappended.Replace("[EmpLocation1]", diligenceInput.EmployerModel[0].Emp_Location.ToString());
                                            stremp1textappended = stremp1textappended.Replace("[EmpStartDate1]", strempstartdate);
                                            TextRange tr = para.AppendText(stremp1textappended);
                                            tr.CharacterFormat.FontName = "Calibri (Body)";
                                            tr.CharacterFormat.FontSize = 11;
                                            doc.SaveToFile(savePath);
                                            if (diligenceInput.EmployerModel[0].Emp_Status.ToString().Equals("Current"))
                                            {
                                                doc.Replace("EMPLOYEEDESCRIPTION", "<investigator to insert brief summary of company here> EMPLOYEEDESCRIPTION", true, true);
                                            }
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
                    if (diligenceInput.EmployerModel[i].Emp_Status.ToString().Equals("Concurrent"))
                    {
                        if (diligenceInput.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                        {
                            emp_Comment = string.Concat(emp_Comment, "\n", "In addition to the above, [LastName] is also a [Position1] of [Employer1Footnote]");
                            emp_Comment = emp_Comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
                            emp_Comment = emp_Comment.Replace("[Position1]", diligenceInput.EmployerModel[i].Emp_Position.ToString());
                            emp_Comment = emp_Comment.Replace("[Employer1]", diligenceInput.EmployerModel[i].Emp_Employer.ToString());
                            emp_Comment = emp_Comment.Replace("[EmpLocation1]", diligenceInput.EmployerModel[i].Emp_Location.ToString());
                            emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                            emp_Comment = string.Concat("\n\n",emp_Comment);
                        }
                        else
                        {
                            emp_Comment = string.Concat(emp_Comment, "\n", "In addition to the above, [LastName] is also a [Position1] of [Employer1Footnote]");
                            emp_Comment = emp_Comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
                            emp_Comment = emp_Comment.Replace("[Position1]", diligenceInput.EmployerModel[i].Emp_Position.ToString());
                            emp_Comment = emp_Comment.Replace("[Employer1]", diligenceInput.EmployerModel[i].Emp_Employer.ToString());
                            emp_Comment = emp_Comment.Replace("[EmpLocation1]", diligenceInput.EmployerModel[i].Emp_Location.ToString());
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
                                                TextRange _text = footnote1.TextBody.AddParagraph().AppendText("<Investigator to insert company registry detail here as well as indicate whether a Companion Report has been prepared for the firm OR that additional research efforts would be required in connection with this firm>.");
                                                footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                                _text.CharacterFormat.FontName = "Calibri (Body)";
                                                _text.CharacterFormat.FontSize = 9;
                                                //Append the line
                                                string stremp1textappended = "";
                                                if (diligenceInput.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                                                {
                                                    stremp1textappended = " in [EmpLocation1], since [EmpStartDate1]. EMPLOYEEDESCRIPTION";
                                                }
                                                else
                                                {
                                                    stremp1textappended = " in [EmpLocation1], since [EmpStartDate1]. EMPLOYEEDESCRIPTION";
                                                }
                                                doc.Replace("[Employer1Footnote]", diligenceInput.EmployerModel[i].Emp_Employer.ToString(), true, true);
                                                stremp1textappended = stremp1textappended.Replace("[EmpLocation1]", diligenceInput.EmployerModel[i].Emp_Location.ToString());
                                                stremp1textappended = stremp1textappended.Replace("[EmpStartDate1]", strempstartdate);
                                                TextRange tr = para.AppendText(stremp1textappended);
                                                tr.CharacterFormat.FontName = "Calibri (Body)";
                                                tr.CharacterFormat.FontSize = 11;
                                                doc.SaveToFile(savePath);
                                                //if (diligenceInput.EmployerModel[i].Emp_Status.ToString().Equals("Current"))
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
                        emp_Comment = emp_Comment.Replace("[Employer]", diligenceInput.EmployerModel[i].Emp_Employer.ToString());
                        emp_Comment = emp_Comment.Replace("[EmpLocation]", diligenceInput.EmployerModel[i].Emp_Location.ToString());
                        emp_Comment = emp_Comment.Replace("[EmpStartDate]", strempstartdate);
                        emp_Comment = emp_Comment.Replace("[EmpEndDate]", strempenddate);
                        if (i != diligenceInput.EmployerModel.Count - 1)
                        {
                            emp_Comment = string.Concat(emp_Comment, "\n");
                        }
                    }
                }
            }
            doc.Replace("EMPLOYEEDESCRIPTION", emp_Comment, true, true);
            doc.Replace("LastName", emp_Comment, true, false);
            doc.SaveToFile(savePath);
            string bnres = "";
            for (int i = 1; i < diligenceInput.EmployerModel.Count; i++)
            {
                bnres = "";
                if (diligenceInput.EmployerModel[i].Emp_Confirmed.ToString().Equals("Result Pending"))
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
            string edu_comment = "";            
            string edu_header = "";
            string edu_summcomment = "";
            //Education details
            for (int i = 0; i < diligenceInput.educationModels.Count; i++)
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
                string edustartdate = "";
                string eduenddate = "";
                string edugraddate = "";
                string edustartyr = "";
                string eduendyr = "";
                string edugradyr = "";
                if (diligenceInput.educationModels[i].Edu_History.ToString().Equals("Yes"))
                {
                    if (!diligenceInput.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && !diligenceInput.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !diligenceInput.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                    {
                        edustartdate = diligenceInput.educationModels[i].Edu_StartDateMonth +" "+ diligenceInput.educationModels[i].Edu_StartDateDay + ", " + diligenceInput.educationModels[i].Edu_StartDateYear;
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

                    if (i == diligenceInput.educationModels.Count-1) {
                        switch (diligenceInput.educationModels[i].Edu_Confirmed.ToString())
                        {
                            case "Yes":
                                edu_comment = string.Concat(edu_comment, comment1.confirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, "Confirmed");
                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.confirmed_comment.ToString(), "\n\n");
                                break;
                            case "No":
                                edu_comment = string.Concat(edu_comment, comment1.unconfirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, "Unconfirmed");
                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.unconfirmed_comment.ToString(), "\n\n");
                                break;
                            case "Result Pending":
                                edu_comment = string.Concat(edu_comment, edurescomment.confirmed_comment.ToString(), "APPENDEDURESULTPEND", i.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, "Results Pending");
                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.confirmed_comment.ToString(), " APPENDRESPENDING", "", "\n\n");
                                break;
                            case "Attendance Confirmed":
                                edu_comment = string.Concat(edu_comment, eduattcomment.confirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, "Confirmed");
                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.confirmed_comment.ToString(), "\n\n");
                                break;
                            case "Attendance Unconfirmed":
                                edu_comment = string.Concat(edu_comment, eduattcomment.unconfirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, "Unconfirmed");
                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.unconfirmed_comment.ToString(), "\n\n");
                                break;
                            case "Attendance Result pending":
                                edu_comment = string.Concat(edu_comment, edurescomment.unconfirmed_comment.ToString(), "APPENDEDUATTRESULTPEND", i.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, "Results Pending");
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
                                edu_header = string.Concat(edu_header,"Results Pending", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.confirmed_comment.ToString(), " APPENDRESPENDING", "", "\n\n");
                                break;
                            case "Attendance Confirmed":
                                edu_comment = string.Concat(edu_comment, eduattcomment.confirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header,"Confirmed", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.confirmed_comment.ToString(), "\n\n");
                                break;
                            case "Attendance Unconfirmed":
                                edu_comment = string.Concat(edu_comment, eduattcomment.unconfirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, "Unconfirmed", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.unconfirmed_comment.ToString(), "\n\n");
                                break;
                            case "Attendance Result pending":
                                edu_comment = string.Concat(edu_comment, edurescomment.unconfirmed_comment.ToString(), "APPENDEDUATTRESULTPEND", i.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header,"Results Pending", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.unconfirmed_comment.ToString(), " ATTRESPENDING", "\n\n");
                                break;
                        }
                    }
                    try
                    {
                        edu_comment = edu_comment.Replace("[Degree]", diligenceInput.educationModels[i].Edu_Degree.ToString());
                    }
                    catch {
                        edu_comment = edu_comment.Replace("[Degree]", "not provided");
                    } try
                    {
                        edu_comment = edu_comment.Replace("[School]", diligenceInput.educationModels[i].Edu_School.ToString());
                    }
                    catch {
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
                    catch {
                        edu_summcomment = edu_summcomment.Replace("[Degree]", "not provided");
                    }
                    try
                    {
                        edu_summcomment = edu_summcomment.Replace("[School]", diligenceInput.educationModels[i].Edu_School.ToString());
                    }
                    catch {
                        edu_summcomment = edu_summcomment.Replace("[School]", "not provided");
                    }
                    try
                    {
                        edu_summcomment = edu_summcomment.Replace("[EduLocation]", diligenceInput.educationModels[i].Edu_Location.ToString());
                    }
                    catch {
                        edu_summcomment = edu_summcomment.Replace("[EduLocation]", "not provided");
                    }
                    edu_summcomment = edu_summcomment.Replace("[GradDate]", edugradyr);
                    edu_summcomment = edu_summcomment.Replace("[EDUFROMDATE]", edustartyr);
                    edu_summcomment = edu_summcomment.Replace("[EDUTODATE]", eduendyr);
                }
                else
                {
                    edu_comment = "No reported educational credentials were identified in connection with [Last Name], and additional information from the subject -- if any -- would be required in order to pursue confirmation of the same.\n\n";                    
                    edu_summcomment = comment1.other_comment.ToString();
                    edu_comment = edu_comment.Replace("[Last Name]", diligenceInput.diligenceInputModel.LastName.ToString());
                    edu_summcomment = edu_summcomment.Replace("[Last Name]", diligenceInput.diligenceInputModel.LastName.ToString());
                    edu_header = "N/A";
                }
            }
            edu_comment = edu_comment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
            edu_summcomment = edu_summcomment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
            doc.Replace("SUMMDEGDESCRIPTION", edu_summcomment.TrimEnd(), true, true);
            doc.Replace("DEGREEDESCRIPTION", edu_comment, true, true);            
            doc.Replace("DEGREEHEADER", edu_header, true, true);
            doc.SaveToFile(savePath); ;
            Table table = doc.Sections[1].Tables[0] as Table;
            TableCell cell1 = table.Rows[5].Cells[2];
            TableCell cell2 = table.Rows[5].Cells[1];
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
                                textRange = p1.AppendText("DONEefforts to independently verify the same are currently ongoing (the results of which will be provided under separate cover -- if and when received)");
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
                        catch { }
                    }
                }
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

                string plstartdate = "";
                string plenddate = "";
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
                    
                    if(plenddate == "<not provided>" && plstartdate == "<not provided>")
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
                        }
                        catch { }
                        if(i== diligenceInput.pllicenseModels.Count - 1)
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
            //EDUCATIONALANDLICENSINGHITS
            if (diligenceInput.educationModels[0].Edu_History.ToString().Equals("No") )
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
            if (diligenceInput.summarymodel.bankruptcy_filings.ToString().Equals("Results Pending")|| diligenceInput.summarymodel.civil_court_Litigation.ToString().Equals("Results Pending")|| diligenceInput.summarymodel.criminal_records.ToString().Equals("Results Pending"))
            {
                doc.Replace("[DESCRESULTPENDINGCRIMM]", "\n\n<Search type> searches are currently ongoing in [Country], the results of which will be provided under separate cover upon receipt.", true, false);
            }
            else
            {
                doc.Replace("[DESCRESULTPENDINGCRIMM]", "", true, false);
            }
            doc.SaveToFile(savePath);
            //Additional Country
            string country_comment="";
            if (diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Equals(""))
            {
                country_comment = "";               
            }
            else
            {
                CommentModel comment2 = _context.DbComment
                              .Where(u => u.Comment_type == "NonScopeCountry")
                              .FirstOrDefault();
                if (diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Contains(",") || diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Contains(" and ")) {
                    country_comment = comment2.confirmed_comment.ToString();
                    country_comment = string.Concat("\n", country_comment, "\n");
                    country_comment = country_comment.Replace("[Non-ScopeCountries]", diligenceInput.diligenceInputModel.Nonscopecountry1.ToString());
                }
                else
                {
                    country_comment = comment2.unconfirmed_comment.ToString();
                    country_comment = string.Concat("\n", country_comment, "\n");
                    country_comment = country_comment.Replace("[Non-ScopeCountry]", diligenceInput.diligenceInputModel.Nonscopecountry1.ToString());
                }
            }
            doc.Replace("ADDITIONALCOUNTRIES", country_comment, true, true);
            doc.SaveToFile(savePath);
            //Business Affiliation
            string business_comment;
            CommentModel busi_comment2 = _context.DbComment
                              .Where(u => u.Comment_type == "Business_Affiliation")
                              .FirstOrDefault();
            string strbusumcomm = "";
            if (diligenceInput.otherdetails.Has_Business_Affiliations.ToString().Equals("Yes"))
            {
                doc.Replace("COMPRESULT", "Records", true, true);
                strbusumcomm = "[LastName] is a [Position1] of [Employer1] in [Employer1Location], since [EmpStartDate1][COMPCOMMENT]\n\nIn addition, [LastName] is identified in connection with additional entities in <investigator to add countries>";                             
                business_comment = busi_comment2.confirmed_comment.ToString();                
                business_comment = string.Concat(business_comment, "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>");
            }
            else
            {
                business_comment = busi_comment2.unconfirmed_comment.ToString();
                doc.Replace("COMPRESULT", "Clear", true, true);
                strbusumcomm = "[LastName] is a [Position1] of [Employer1] in [Employer1Location], since [EmpStartDate1]";
            }
            doc.Replace("BUSINESSAFFILIATIONSIDENTIFIED", business_comment, true, true);
            try {
                if (diligenceInput.EmployerModel[0].Emp_Position.Equals(""))
                {
                    strbusumcomm = strbusumcomm.Replace("[Position1]", "<not provided>");
                }
                else
                {
                    strbusumcomm = strbusumcomm.Replace("[Position1]", diligenceInput.EmployerModel[0].Emp_Position);
                }
            } catch
            {
                strbusumcomm = strbusumcomm.Replace("[Position1]", "<not provided>");
            }
            try
            {
                if (diligenceInput.EmployerModel[0].Emp_Employer.Equals(""))
                {
                    strbusumcomm = strbusumcomm.Replace("[Employer1]", "<not provided>");
                }
                else
                {
                    strbusumcomm = strbusumcomm.Replace("[Employer1]", diligenceInput.EmployerModel[0].Emp_Employer);
                }
            }
            catch
            {
                strbusumcomm = strbusumcomm.Replace("[Employer1]", "<not provided>");
            }
            try
            {
                if (diligenceInput.EmployerModel[0].Emp_Location.Equals(""))
                {
                    strbusumcomm = strbusumcomm.Replace("[Employer1Location]", "<not provided>");
                }
                else
                {
                    strbusumcomm = strbusumcomm.Replace("[Employer1Location]", diligenceInput.EmployerModel[0].Emp_Location);
                }
            }
            catch
            {
                strbusumcomm = strbusumcomm.Replace("[Employer1Location]", "<not provided>");
            }            
            try
            {
                if (diligenceInput.EmployerModel[0].Emp_StartDateYear.ToString().Equals("Year"))
                {
                    strbusumcomm = strbusumcomm.Replace("[EmpStartDate1]", "<not provided>");
                }
                else
                {
                    strbusumcomm = strbusumcomm.Replace("[EmpStartDate1]", diligenceInput.EmployerModel[0].Emp_StartDateYear);
                }
            }
            catch
            {
                strbusumcomm = strbusumcomm.Replace("[EmpStartDate1]", "<not provided>");
            }
            doc.Replace("COMPCOMMENT", strbusumcomm, true, true);
            doc.SaveToFile(savePath);
            string strcommentconcurrent = "";
            for (int i = 1;  i< diligenceInput.EmployerModel.Count; i++)
            {
                if (diligenceInput.EmployerModel[i].Emp_Status == "Concurrent")
                {
                    strcommentconcurrent = string.Concat(strcommentconcurrent, "\n\n[LastName] is a [Position1] of [Employer1] in [Employer1Location], since [EmpStartDate1]");
                    strcommentconcurrent = strcommentconcurrent.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName);
                    try
                    {
                        if (diligenceInput.EmployerModel[i].Emp_Position.Equals("")) { strcommentconcurrent = strcommentconcurrent.Replace("[Position1]", "<not provided>"); }
                        else { strcommentconcurrent = strcommentconcurrent.Replace("[Position1]", diligenceInput.EmployerModel[i].Emp_Position); }
                    }
                    catch
                    {
                        strcommentconcurrent = strcommentconcurrent.Replace("[Position1]", "<not provided>");
                    }
                    try
                    {
                        if (diligenceInput.EmployerModel[i].Emp_Employer.Equals("")) { strcommentconcurrent = strcommentconcurrent.Replace("[Employer1]", "<not provided>"); }
                        else { strcommentconcurrent = strcommentconcurrent.Replace("[Employer1]", diligenceInput.EmployerModel[i].Emp_Employer); }
                    }
                    catch
                    {
                        strcommentconcurrent = strcommentconcurrent.Replace("[Employer1]", "<not provided>");
                    }
                    try
                    {
                        if (diligenceInput.EmployerModel[i].Emp_StartDateYear.ToString().Equals(""))
                        {
                            strcommentconcurrent = strcommentconcurrent.Replace("[EmpStartDate1]", "<not provided>");
                        }
                        else
                        {
                            strcommentconcurrent = strcommentconcurrent.Replace("[EmpStartDate1]", diligenceInput.EmployerModel[i].Emp_StartDateYear);
                        }
                    }
                    catch
                    {
                        strcommentconcurrent = strcommentconcurrent.Replace("[EmpStartDate1]", "<not provided>");
                    }
                    try
                    {
                        if (diligenceInput.EmployerModel[i].Emp_Location.ToString().Equals(""))
                        {
                            strcommentconcurrent = strcommentconcurrent.Replace("[Employer1Location]", "<not provided>");
                        }
                        else
                        {
                            strcommentconcurrent = strcommentconcurrent.Replace("[Employer1Location]", diligenceInput.EmployerModel[i].Emp_Location);
                        }
                    }
                    catch
                    {
                        strcommentconcurrent = strcommentconcurrent.Replace("[Employer1Location]", "<not provided>");
                    }
                }
            }
            doc.Replace("[COMPCOMMENT]", strcommentconcurrent, true, false);
            doc.SaveToFile(savePath);
            //Intellectual Hits
            string intellectual_comment;
            CommentModel intellec_comment2 = _context.DbComment
                               .Where(u => u.Comment_type == "Intellectual_hits")
                               .FirstOrDefault();
            if (diligenceInput.otherdetails.Has_Intellectual_Hits.ToString().Equals("Yes"))
            {
                intellectual_comment = intellec_comment2.confirmed_comment.ToString();
                intellectual_comment=string.Concat("\n",intellectual_comment.ToString(), "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n");
            }
            else
            {
                intellectual_comment = "";                
            }
            doc.Replace("INTELLECTUALPROPERTYIDENTIFIED", intellectual_comment, true, true);
            doc.SaveToFile(savePath);
            //Global security hits
            string globalhit_comment;
            CommentModel globalhit_comment2 = _context.DbComment
                              .Where(u => u.Comment_type == "Global_Security")
                              .FirstOrDefault();
            if (diligenceInput.otherdetails.Global_Security_Hits.ToString().Equals("Yes"))
            {
                globalhit_comment = "Research was undertaken in connection with various governmental, quasi-governmental and private sector list repositories maintained for purposes of international security, fraud prevention, anti-terrorism and anti-money laundering, including the List of Specially Designated Nationals and Blocked Persons of the United States Office of Foreign Assets Control (“OFAC”), the United States Bureau of Industry and Security, and other agencies,";
                globalhit_comment = string.Concat(globalhit_comment, "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here> \n");
            }
            else
            {
                globalhit_comment = "Research was undertaken in connection with various governmental, quasi-governmental and private sector list repositories maintained for purposes of international security, fraud prevention, anti-terrorism and anti-money laundering, including the List of Specially Designated Nationals and Blocked Persons of the United States Office of Foreign Assets Control (“OFAC”), the United States Bureau of Industry and Security, and other agencies,";
                globalhit_comment = string.Concat(globalhit_comment, "\n");
            }
            doc.Replace("GLOBALSECURITYHITSDESCRIPTION", globalhit_comment, true, true);
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
            string PMCommentModel="";
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
            //US_SEC
            string usseccommentmodel ="";
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
                usseccommentmodel = string.Concat("\nUnited States Securities and Exchange Commission\n\n", usseccommentmodel,"\n");
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
                ukfcacommentmodel = string.Concat("\nUnited Kingdom’s Financial Conduct Authority\n\n", ukfcacommentmodel,"\n");
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
                finracommentmodel = string.Concat("\nUnited States Financial Industry Regulatory Authority\n\n", finracommentmodel,"\n");
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
                nfacommentmodel = string.Concat("\nUnited States National Futures Association\n\n", nfacommentmodel,"\n");
            }            
            string hksfccommentmodel = "";
            string holdslicensecommentmodel = "";
            string regflag = "";
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
                    hksfccommentmodel = string.Concat("\nHong Kong Securities and Futures Commission\n\n", hksfccommentmodel,"\n");                    
                }    
            //Holds Any License 
            string strotherheader = "";              
                if (diligenceInput.otherdetails.Registered_with_HKSFC.StartsWith("Yes") || diligenceInput.otherdetails.Has_Reg_US_SEC.StartsWith("Yes") || diligenceInput.otherdetails.Has_Reg_US_NFA.StartsWith("Yes") || diligenceInput.otherdetails.Has_Reg_FINRA.StartsWith("Yes") || diligenceInput.otherdetails.Has_Reg_UK_FCA.StartsWith("Yes") || diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                {
                    regflag = "Records";
                }
                else
                {
                        regflag = "Clear";     
                    
                }
                if(regflag == "Records")                
                {
                holdslicensecommentmodel = "Investigative efforts did not reveal any additional professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction and license-type basis to confirm the same.  As a precaution, research efforts included, but were not limited to: a Broker Check through the United States Financial Industry Regulatory Authority (for the past two years); a search of the United States Commodity Futures Trading Commission and National Futures Association membership registration; an Investment Adviser Firm Representative search through the United States Securities and Exchange Commission; a search through United Kingdom’s Financial Conduct Authority; as well as a search through the Hong Kong Securities and Futures Commission.  <Investigator to add or remove any sources that came back with hits above or are needed for this jurisdiction>";
                     //holdslicensecommentmodel = "Investigative efforts did not reveal any additional professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction and license-type basis to confirm the same. ";
                    strotherheader = "\nOther Professional Licensures and/or Designations\n\n";
                }
                else
                {
                holdslicensecommentmodel = "Investigative efforts did not reveal any professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction and license-type basis to confirm the same.  As a precaution, research efforts included, but were not limited to: a Broker Check through the United States Financial Industry Regulatory Authority (for the past two years); a search of the United States Commodity Futures Trading Commission and National Futures Association membership registration; an Investment Adviser Firm Representative search through the United States Securities and Exchange Commission; a search through United Kingdom’s Financial Conduct Authority; as well as a search through the Hong Kong Securities and Futures Commission.  <Investigator to add or remove any sources that came back with hits above or are needed for this jurisdiction>";
                strotherheader = "\nOther Professional Licensures and/or Designations\n\n";
                }
                if(diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Not Registered") && diligenceInput.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered") && diligenceInput.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered") && diligenceInput.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered") && diligenceInput.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered"))
                {
                strotherheader = "\n";
                }
                holdslicensecommentmodel = string.Concat(strotherheader, holdslicensecommentmodel,"\n");
           
            if (diligenceInput.pllicenseModels[0].General_PL_License.ToString().ToUpper().Equals("YES")) { doc.Replace("[PLLICENSEDESC]", string.Concat("\n",usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false); }
            else
            {
                doc.Replace("[PLLICENSEDESC]", string.Concat(usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false);
            }
            doc.SaveToFile(savePath);

            if (diligenceInput.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered")) { }
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
                catch { }
            }

            if (diligenceInput.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered")) { }
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
                catch
                {

                }
            }

            if (diligenceInput.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered")) { }
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
                catch
                {

                }
            }

            if (diligenceInput.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered")) { }
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
                catch { }
            }
            string holdresult = "";
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
            }
            catch { }
            holdresult = "";
            if (hksfccommentmodel.ToString().Equals("")) { }
            else
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
                catch { }
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
            regredflagommentmodel = string.Concat("\n",regredflagommentmodel, "\n");
            doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", regredflagommentmodel, true, true);
            doc.SaveToFile(savePath);
            string blnredresultfound="";
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
                                    if (diligenceInput.otherdetails.RegulatoryFlag.Equals("Yes"))
                                    {
                                        strredflagtextappended = " and the following information was identified in connection with [LastName]:   <Investigator to insert results here>";
                                    }
                                    else
                                    {
                                        strredflagtextappended = " and it is noted that [LastName] was not identified in any of these records.";
                                    }
                                    strredflagtextappended = strredflagtextappended.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName);
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
            }
            catch { }
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
            if (diligenceInput.summarymodel.bankruptcy_filings.ToString().Equals("Results Pending")|| diligenceInput.summarymodel.civil_court_Litigation.ToString().Equals("Results Pending")|| diligenceInput.summarymodel.criminal_records.ToString().Equals("Results Pending"))
            {               
                doc.Replace("[CountryHEADER]", string.Concat(strcountry, "\n\n<Search type> searches are currently ongoing through <source> in [Country], the results of which will be provided under separate cover upon receipt."),true,true);
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
            if (diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom")) {
                doc.Replace("[CountryHEADER]", "", true, true);
            }
            else
            {
                switch (strlegalrechitdesc)
                {
                    case "Yes":
                        doc.Replace("[COURTDESC]", "\n\nSearches of all available legal records in [Country], which include the [Courts], identified the subject, personally, as a party to the following bankruptcy filings, civil litigation matters, judgments and/or criminal convictions:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n[COURTDESC]", true, false);
                        break;
                    case "No":
                        doc.Replace("[COURTDESC]", "\n\nSearches of all available legal records in [Country], which include the [Courts], did not identify the subject, personally, as a party to any bankruptcy filings, civil litigation matters, judgments or criminal convictions.\n[COURTDESC]", true, false);
                        break;
                    case "Yes + Consent Needed for Crim":
                        doc.Replace("[COURTDESC]", "\n\nWhile the subject’s express written authorization would be required in order to conduct a criminal records search in [Country], searches of all available legal records in [Country], which include the [Courts], identified the subject, personally, as a party to the following bankruptcy filings, civil litigation matters and/or judgments:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n[COURTDESC]", true, false);
                        break;
                    case "No + Consent Needed for Crim":
                        doc.Replace("[COURTDESC]", "\n\nWhile the subject’s express written authorization would be required in order to conduct a criminal records search in [Country], searches of all available legal records in [Country], which include the [Courts], did not identify the subject, personally, as a party to any bankruptcy filings, civil litigation matters or judgments.\n[COURTDESC]", true, false);
                        break;
                    case "Yes + Crim Not Available in Country":
                        doc.Replace("[COURTDESC]", "\n\nWhile third-party access to criminal records information is restricted in [Country], searches of all available legal records in [Country], which include the [Courts], identified the subject, personally, as a party to the following bankruptcy filings, civil litigation matters and/or judgments:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n[COURTDESC]", true, false);
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
            if (diligenceInput.diligenceInputModel.Country.ToString() == "United Kingdom") { }
            else
            {
                if (diligenceInput.summarymodel.civil_court_Litigation1 == true)
                {
                    legrechitcommon = string.Concat("\nIt should be noted that one or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
                }
            }
            if (diligenceInput.summarymodel.criminal_records1 == true)
            {
                legrechitcommon = string.Concat(legrechitcommon, "\nIt should be noted that one or more individuals known only as [FirstName] [LastName] were identified as <party type> in at least <number> criminal records, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
            }
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
                        if (diligenceInput.summarymodel.civil_court_Litigation1 == true)
                        {
                            legrechitcommon = string.Concat(legrechitcommon, "\nIt should be noted that one or more individuals known only as [FirstName] [LastName] were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
                        }
                        string fr_lrhcommentmodel = "";
                        fr_lrhcommentmodel = legrechitcommon;
                        strlegrechit = string.Concat(fr_lrhcommentmodel, "\n[MEDIABASEDLEGALDESCRIPTION]");
                        doc.Replace("[COURTDESC]", strlegrechit, true, true);
                        break;
                    case "United Kingdom":
                        string civilpossible = "";
                        if (diligenceInput.summarymodel.civil_court_Litigation1 == true)
                        {
                            civilpossible = string.Concat("\nIt should be noted that one or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
                        }
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
                        uk_ihcommentmodel = "Civil Litigation, Insolvency and Other Filings in the United Kingdom[HEADER]\n\nRecords maintained by the Insolvency Service in the United Kingdom[INSERTFOOTNOTE]\n";
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
                        strregtrust = string.Concat("\n", uk_rthcommentmodel, "\n", civilpossible);                        
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
                                uk_ecrcommentmodel = "Pursuant to [LastName]'s express written authorization, criminal court records searches were conducted through the Crown, High and Magistrate Courts associated with the subject's address.  In addition, research efforts were also undertaken in all levels of the roughly 300 or more United Kingdom courts to ensure that no other records were identified with the subject's linked addresses.  In that regard, no such records were identified in connection with [LastName].";
                                break;
                            case "Manual Local with Hits":
                                uk_ecrcommentmodel = "Pursuant to [LastName]'s express written authorization, criminal court records searches were conducted through the Crown, High and Magistrate Courts associated with the subject's address.  In addition, research efforts were also undertaken in all levels of the roughly 300 or more United Kingdom courts to ensure that no other records were identified with the subject's linked addresses.  In that regard, the following records were identified in connection with [LastName]:\n\n<Investigator to add summary of hits>\n<Investigator to add summary of hits>";
                                break;
                            case "No Consent":
                                uk_ecrcommentmodel = uk_ecrcommentmodel1.other_comment.ToString();
                                break;
                        }
                        uk_ecrcommentmodel = uk_ecrcommentmodel.Replace("*n ", "\n");
                        uk_ecrcommentmodel = uk_ecrcommentmodel.Replace("*n", "\n");
                        uk_ecrcommentmodel = uk_ecrcommentmodel.Replace("*t", "\t");
                        strcriminal = string.Concat("\n", "Criminal Records[CRIMHEADER]\n\n", uk_ecrcommentmodel);
                        strcourtdesc = string.Concat(legrechitcommon,strinsolvency, strcivil, strregtrust, strcriminal,"\n");
                        doc.Replace("[COURTDESC]", strcourtdesc, true, false);
                        doc.SaveToFile(savePath);
                        //doc.Replace("Basic Disclosures for individuals throughout England, Wales, the Channel Islands and the Isle of Man.", "Basic Disclosures for individuals throughout England, Wales, the Channel Islands and the Isle of Man", true, true);
                        //doc.Replace("Basic Disclosures for individuals throughout Scotland.", "Basic Disclosures for individuals throughout Scotland", true, true);
                        //doc.Replace("Basic Disclosures for individuals throughout Northern Ireland.", "Basic Disclosures for individuals throughout Northern Ireland", true, true);
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
                                                if (abc.ToString().EndsWith("Basic Disclosures for individuals throughout England, Wales, the Channel Islands and the Isle of Man") || abc.ToString().EndsWith("Basic Disclosures for individuals throughout Scotland") || abc.ToString().EndsWith("Basic Disclosures for individuals throughout Northern Ireland"))
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
                                                    doc.SaveToFile(savePath);
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
                            catch{ }
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
                        civilpossible = "";
                        if (diligenceInput.summarymodel.civil_court_Litigation1 == true)
                        {
                            civilpossible = string.Concat("\nIt should be noted that one or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
                        }
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
                        strcivil = string.Concat("\n", can_civilcommentmodel, "\n",civilpossible);                        
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
                        strcourtdesc = string.Concat(legrechitcommon, strinsolvency, strcivil, strppr, "\nCriminal Records[CRIMHEADER]\n", strcriminal,"[MEDIABASEDLEGALDESCRIPTION]");
                        doc.Replace("[COURTDESC]", strcourtdesc, true, false);
                        doc.SaveToFile(savePath);
                        break;
                    case "Australia":                                                
                        string aus_inhitcommentmodel = "";                        
                        string aus_civilcommentmodel = "";
                        string aus_pprcommentmodel = "";
                        string aus_cricommentmodel = "";
                        civilpossible = "";
                        if (diligenceInput.summarymodel.civil_court_Litigation1 == true)
                        {
                            civilpossible = string.Concat("\nIt should be noted that one or more individuals known only as “[FirstName] [LastName]” were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
                        }
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
                        strcivil=string.Concat("\nAdditionally, searches conducted on all available legal records in Australia, including the Federal Court, District Courts, County Courts, Magistrate Courts and Tribunals\n",civilpossible);
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
                        strppr = string.Concat("\n", aus_pprcommentmodel, "\n");
                        strcourtdesc = string.Concat(legrechitcommon,"\n", strinsolvency, strcivil, "\nCriminal Records[CRIMHEADER]\n", strcriminal, strppr, "[MEDIABASEDLEGALDESCRIPTION]");
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
                        catch { }
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
                strcourtdesc = string.Concat(mediacommentmodel,uslegcommentmodel);
                doc.Replace("[MEDIABASEDLEGALDESCRIPTION]", strcourtdesc, true, true);
                doc.SaveToFile(savePath);

                if (strcivil.ToString().Equals("")) { }
                else {
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
                                        if (abc.ToString().Equals("Civil Litigation, Insolvency and Other Filings in the United Kingdom[HEADER]"))
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
                else {
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
                strcourtsection = string.Concat(strlegrechit, mediacommentmodel, uslegcommentmodel);
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
            //CountrySpecific
            CountrySpecificModel CS1 = new CountrySpecificModel();
            //CS1.record_Id = record_Id;
            //string strdrivingComment = "";
            //string strcreditComment = "";
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
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].";
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
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].";
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
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].";
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
                    aus_asiccommentmodel = string.Concat("\n\n", aus_asiccommentmodel, "\n");
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
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which identified the following information in connection with [LastName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [LastName].";
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
                        doc.Replace("DRIVINGRESULT", "N/A", true, true);
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
                strdrivingComment = strdrivingComment.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
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
                        doc.Replace("DRIVINGRESULT", "N/A", true, true);
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
            if (diligenceInput.diligenceInputModel.Country.ToString().Equals("United Kingdom")) {
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
                        doc.Replace("PERCREDITRESULT", "N/A", true, true);
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
                        doc.Replace("PERCREDITRESULT", "N/A", true, true);
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

                if (diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse")) {
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
            doc.Replace("PROFCOMMENT",strPLComment, true, true);            
            
            //Legal_Record_Judgments_Liens_Hits
            string Legrechitcommentmodel = "";
            string strexecutive_sumlegrec = "";

            CommentModel Legrechitcommentmodel1 = _context.DbComment
                            .Where(u => u.Comment_type == "Legal_Record_Judgments_Liens_Hits")
                            .FirstOrDefault();
            if (diligenceInput.summarymodel.criminal_records.StartsWith("Record")|| diligenceInput.summarymodel.bankruptcy_filings.StartsWith("Record")|| diligenceInput.summarymodel.civil_court_Litigation.StartsWith("Record"))
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
            if(regflag=="Records" || strexecutive_sumlegrec == "Yes")
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
                    hitcompcommentmodel = "In sum, no issues of potential relevance were identified in connection with [FirstName] [MiddleInitial] [LastName] in [Country].";
                    searchtext = "In sum, no issues of potential relevance were identified in connection with";
                }
                else
                {
                    // hitcompcommentmodel = nohitcompcommentmodel1.unconfirmed_comment.ToString();
                    hitcompcommentmodel = "In sum, no issues of potential relevance were identified in connection with [FirstName] [MiddleInitial] [LastName] in [Country].";
                    searchtext = "In sum, no issues of potential relevance were identified in connection with";
                }

            }
            hitcompcommentmodel = hitcompcommentmodel.Replace("[FirstName]", diligenceInput.diligenceInputModel.FirstName.ToString());
            //hitcompcommentmodel = hitcompcommentmodel.Replace("[MiddleInitial]", diligenceInput.diligenceInputModel.MiddleInitial.ToString());
            doc.Replace("HasHitsAboveAndCompanionReport", hitcompcommentmodel, true, true);
            doc.SaveToFile(savePath);
            if (searchtext.ToString().Equals("")) { }
            else {
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
            doc.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName.ToString(), true, true);           
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
                                if (abc.ToString().Contains("(“OFAC”),") || abc.ToString().EndsWith("(“OFAC”),"))
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
                                    strglobaltextappended = strglobaltextappended.Replace("[LastName]", diligenceInput.diligenceInputModel.LastName);
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
            }
            catch { }
            doc.SaveToFile(savePath);
            doc.Replace("[CaseNumber]", diligenceInput.diligenceInputModel.CaseNumber.TrimEnd(), true, true);
            doc.SaveToFile(savePath);
            HttpContext.Session.SetString("first_name",diligenceInput.diligenceInputModel.FirstName);
            HttpContext.Session.SetString("last_name", diligenceInput.diligenceInputModel.LastName);
            HttpContext.Session.SetString("country", diligenceInput.diligenceInputModel.Country);
            HttpContext.Session.SetString("case_number", diligenceInput.diligenceInputModel.CaseNumber);
            HttpContext.Session.SetString("middleinitial", diligenceInput.diligenceInputModel.MiddleInitial);
            HttpContext.Session.SetString("city", diligenceInput.diligenceInputModel.City);
            HttpContext.Session.SetString("regflag", regflag);
            HttpContext.Session.SetString("employer1", diligenceInput.EmployerModel[0].Emp_Employer.ToString());
            HttpContext.Session.SetString("emp_location1", diligenceInput.EmployerModel[0].Emp_Location.ToString());
            HttpContext.Session.SetString("emp_position1", diligenceInput.EmployerModel[0].Emp_Position.ToString());
            if (diligenceInput.EmployerModel[0].Emp_StartDateYear.ToString().Equals(""))
            {
                HttpContext.Session.SetString("emp_startdate1", "<not provided>");
            }
            else
            {
                HttpContext.Session.SetString("emp_startdate1", diligenceInput.EmployerModel[0].Emp_StartDateYear.ToString());
            }
            HttpContext.Session.SetString("pl_generallicense", diligenceInput.pllicenseModels[0].General_PL_License.ToString());
            if (diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
            {
                string plstartdate = "";
                string plenddate = "";
                HttpContext.Session.SetString("pl_licensetype", diligenceInput.pllicenseModels[0].PL_License_Type.ToString());
                HttpContext.Session.SetString("pl_organization", diligenceInput.pllicenseModels[0].PL_Organization.ToString());
                if (diligenceInput.pllicenseModels[0].PL_StartDateYear.ToString().Equals(""))
                {
                    plstartdate = "<not provided>";
                }
                else
                {
                    plstartdate = diligenceInput.pllicenseModels[0].PL_StartDateYear.ToString();
                }
                if (diligenceInput.pllicenseModels[0].PL_EndDateYear.ToString().Equals(""))
                {
                    plenddate = "<not provided>";
                }
                else
                {
                    plenddate = diligenceInput.pllicenseModels[0].PL_EndDateYear.ToString();
                }
                HttpContext.Session.SetString("pl_startdate", plstartdate);
                HttpContext.Session.SetString("pl_enddate", plenddate);
            }
            //return RedirectToAction("DiligenceForm_Page2", "Diligence");
            return Save_Page2(diligenceInput.summarymodel);
        }
        public IActionResult Save_Page2(SummaryResulttableModel ST)
        {
            string record_Id = HttpContext.Session.GetString("record_Id");
            string first_name = HttpContext.Session.GetString("first_name");
            string last_name = HttpContext.Session.GetString("last_name");
            string country = HttpContext.Session.GetString("country");
            string case_number = HttpContext.Session.GetString("case_number");
            string city = HttpContext.Session.GetString("city");
            string regflag = HttpContext.Session.GetString("regflag");
            string employer1 = HttpContext.Session.GetString("employer1");
            string emp_position1 = HttpContext.Session.GetString("emp_position1");
            string emp_startdate1 = HttpContext.Session.GetString("emp_startdate1");
            string middleinitial = HttpContext.Session.GetString("middleinitial");
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
                pl_enddate = HttpContext.Session.GetString("pl_enddate");
             
            }

            string savePath = _config.GetValue<string>("ReportPath");
            savePath = string.Concat(savePath,last_name.ToString(), "_SterlingDiligenceReport(", case_number.ToString(), ")_DRAFT.docx");
            Document doc = new Document(savePath);
            string strcomment;
            Table table = doc.Sections[1].Tables[0] as Table;
            switch (ST.personal_Identification.ToString())
            {
                case "Clear":
                    doc.Replace("PICOMMENT", "Confirmed", true, true);
                    doc.Replace("PLRESULT", "Clear", true, true);
                    break;
                case "Discrepancy Identified":
                    doc.Replace("PICOMMENT", "<Investigator to insert discrepancy here>", true, true);
                    doc.Replace("PLRESULT", "Discrepancy Identified", true, true);
                    break;
            }
            doc.SaveToFile(savePath);
            switch (ST.name_add_Ver_History.ToString())
            {
                case "Clear":
                    strcomment = "LastName has jurisdictional ties to [Country]";
                    strcomment = strcomment.Replace("LastName", last_name);                    
                    doc.Replace("NAMECOMMENT", strcomment, true, true);
                    doc.Replace("NAMERESULT", "Clear", true, true);
                    break;             
                case "Discrepancy Identified":
                    strcomment = "LastName has jurisdictional ties to [Country]\n\nHowever, while not reported by the subject, LastName was also identified as having current ties to < Investigator to insert additional jurisdictions here>";
                    strcomment = strcomment.Replace("LastName", last_name);                
                    doc.Replace("NAMECOMMENT", strcomment, true, true);
                    doc.Replace("NAMERESULT", "Discrepancy Identified", true, true);
                    break;
            }            
            switch (regflag.ToString())
            {
                case "Clear":
                    doc.Replace("REGCOMMENT", "", true, true);
                    doc.Replace("REGRESULT", "Clear", true, true);
                    break;
                case "Records":
                    strcomment = "<investigator to insert regulatory hits here>";                                        
                    doc.Replace("REGCOMMENT", strcomment, true, true);
                    doc.Replace("REGRESULT", "Records", true, true);
                    break;
            }
            string bankrupcomment = "";
            string bankrupresult = "";
            switch (ST.bankruptcy_filings.ToString())
            {
                case "Clear":
                    bankrupcomment = "";
                    bankrupresult = "Clear";
                    break;
                case "Records":
                    bankrupcomment = "LastName was identified as a <party type> in connection with at least <number> bankruptcy filings in [Country], filed between <date> and <date>, which are currently <status> ";
                    bankrupcomment = bankrupcomment.Replace("LastName", last_name);                                        
                    bankrupresult= "Records";
                    break;
                case "Record":
                    bankrupcomment = "[LastName] was identified as a <party type> in connection with a bankruptcy filing in [Country], which were recorded in <date> and <date>, and is currently <status>";
                    bankrupresult = "Record";
                    break;
                case "Results Pending":
                    bankrupcomment = "Bankruptcy court records searches are currently pending in [Country] the results of which will be provided under separate cover upon receipt";
                    bankrupresult = "Results Pending";
                    break;
            }
            if(ST.bankruptcy_filings1 == true)
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
                    civilcourtResult = "Clear";
                    break;
                case "Records":
                    civilcourtcomment = "LastName was identified as a <party type> in connection with at least <number> civil litigation filings in [Country], filed between <date> and <date>, which are currently <status>";
                    civilcourtcomment = civilcourtcomment.Replace("LastName", last_name);
                    civilcourtResult = "Records";
                    break;
                case "Record":
                    civilcourtcomment = "[Last Name] was identified as a <party type> in connection with a civil litigation filing in [Country], which was recorded in <date>, which is <status>";
                    civilcourtResult = "Record";
                    break;
                case "Results Pending":
                    civilcourtcomment = "A civil court records search is currently pending in [Country] the results of which will be provided under separate cover upon receipt";
                    civilcourtResult = "Results Pending";
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
                    doc.Replace("CIVILJUDGERESULT", "N/A CIVILJUDGERESULT", true, true);                    
                    break;
                case "Clear":
                    doc.Replace("CIVILJUDGECOMMENT", "CIVILJUDGECOMMENT", true, true);
                    doc.Replace("CIVILJUDGERESULT", "Clear CIVILJUDGERESULT", true, true);
                    break;
                case "Records":
                    strcomment = "[LastName] was identified as a <party type> in connection with at least <number> judgments filed in [Country] between <date> and <date>, which are currently <status> CIVILJUDGECOMMENT";
                    strcomment = strcomment.Replace("[LastName]", last_name);                    
                    doc.Replace("CIVILJUDGECOMMENT", strcomment, true, true);
                    doc.Replace("CIVILJUDGERESULT", "Records CIVILJUDGERESULT", true, true);
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
            string criminalcomment = "";
            string criminalresult = "";
            switch (ST.criminal_records.ToString())
            {
                case "Clear":
                    criminalcomment= "";
                    criminalresult= "Clear";
                    break;
                case "Records":
                    criminalcomment = "LastName was identified as a Defendant in connection with at least <number> criminal records in [Country] , filed between <date> and <date>, which pertain to <type of charge> and are currently <status>";
                    criminalcomment = criminalcomment.Replace("LastName", last_name);
                    criminalresult= "Records";
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
            switch (ST.news_media_searches.ToString())
            {
                case "Clear":
                    doc.Replace("[NewsRES]", "", true, false);
                    doc.Replace("NEWSCOMMENT", "No adverse or materially-significant information was identified ", true, true);
                    doc.Replace("NEWSRES", "Clear", true, false);
                    break;
                case "Records":
                    doc.Replace("NEWSCOMMENT", "<investigator to insert press summaries here>", true, true);
                    doc.Replace("NEWSRES", "Records", true, false);
                    doc.Replace("[NewsRES]", "", true, false);
                    break;
                case "Potentially-Relevant Information":
                    doc.Replace("NEWSCOMMENT", "<investigator to insert press summaries here>", true, true);
                    doc.Replace("NEWSRES", "", true, false);
                    doc.Replace("[NewsRES]", "Potentially-Relevant Information", true, false);
                    break;
            }
            switch (ST.department_foreign.ToString())
            {
                case "Clear":
                    doc.Replace("DEPTFORCOMMENT", "", true, true);
                    doc.Replace("DEPTFORRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("DEPTFORCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("DEPTFORRESULT", "Records", true, true);
                    break;
            }
            switch (ST.european_union.ToString())
            {
                case "Clear":
                    doc.Replace("EURUNCOMMENT", "", true, true);
                    doc.Replace("EURUNRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("EURUNCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("EURUNRESULT", "Records", true, true);
                    break;
            }
            switch (ST.HM_treasury.ToString())
            {
                case "Clear":
                    doc.Replace("HMTRECOMMENT", "", true, true);
                    doc.Replace("HMTRERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("HMTRECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("HMTRERESULT", "Records", true, true);
                    break;
            }
            switch (ST.US_bureau.ToString())
            {
                case "Clear":
                    doc.Replace("USBEUCOMMENT", "", true, true);
                    doc.Replace("USBEURESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USBEUCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USBEURESULT", "Records", true, true);
                    break;
            }
            switch (ST.US_department.ToString())
            {
                case "Clear":
                    doc.Replace("USBDEPCOMMENT", "", true, true);
                    doc.Replace("USBDEPRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USBDEPCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USBDEPRESULT", "Records", true, true);
                    break;
            }
            switch (ST.US_Directorate.ToString())
            {
                case "Clear":
                    doc.Replace("USBDIRCOMMENT", "", true, true);
                    doc.Replace("USBDIRRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USBDIRCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USBDIRRESULT", "Records", true, true);
                    break;
            }
            switch (ST.US_general.ToString())
            {
                case "Clear":
                    doc.Replace("USGENCOMMENT", "", true, true);
                    doc.Replace("USGENRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USGENCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USGENRESULT", "Records", true, true);
                    break;
            }
            try
            {
                switch (ST.US_office.ToString())
                {
                    case "Clear":
                        doc.Replace("USOFFCOMMENT", "", true, true);
                        doc.Replace("USOFFRESULT", "Clear", true, true);
                        break;
                    case "Records":
                        doc.Replace("USOFFCOMMENT", "<investigator to insert summary here>", true, true);
                        doc.Replace("USOFFRESULT", "Records", true, true);
                        break;
                }
            }
            catch
            {
                doc.Replace("USOFFCOMMENT", "", true, true);
                doc.Replace("USOFFRESULT", "Clear", true, true);
            }
            switch (ST.UN_consolidated.ToString())
            {
                case "Clear":
                    doc.Replace("UNCONSOCOMMENT", "", true, true);
                    doc.Replace("UNCONSORESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("UNCONSOCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("UNCONSORESULT", "Records", true, true);
                    break;
            }
            switch (ST.world_bank_list.ToString())
            {
                case "Clear":
                    doc.Replace("WORLDBANKCOMMENT", "", true, true);
                    doc.Replace("WORLDBANKRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("WORLDBANKCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("WORLDBANKRESULT", "Records", true, true);
                    break;
            }
            switch (ST.interpol.ToString())
            {
                case "Clear":
                    doc.Replace("INTERPOLCOMMENT", "", true, true);
                    doc.Replace("INTERPOLRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("INTERPOLCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("INTERPOLRESULT", "Records", true, true);
                    break;
            }
            switch (ST.US_Federal.ToString())
            {
                case "Clear":
                    doc.Replace("USFEDERALCOMMENT", "", true, true);
                    doc.Replace("USFEDERALRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USFEDERALCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USFEDERALRESULT", "Records", true, true);
                    break;
            }
            switch (ST.US_secret_service.ToString())
            {
                case "Clear":
                    doc.Replace("USSECRETCOMMENT", "", true, true);
                    doc.Replace("USSECRETRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USSECRETCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USSECRETRESULT", "Records", true, true);
                    break;
            }
            switch (ST.commodity_futures.ToString())
            {
                case "Clear":
                    doc.Replace("COMMODITYCOMMENT", "", true, true);
                    doc.Replace("COMMODITYRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("COMMODITYCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("COMMODITYRESULT", "Records", true, true);
                    break;
            }
            switch (ST.federal_deposit.ToString())
            {
                case "Clear":
                    doc.Replace("FDEPOCOMMENT", "", true, true);
                    doc.Replace("FDEPORESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("FDEPOCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("FDEPORESULT", "Records", true, true);
                    break;
            }
            switch (ST.federal_reserve.ToString())
            {
                case "Clear":
                    doc.Replace("FRESCOMMENT", "", true, true);
                    doc.Replace("FRESRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("FRESCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("FRESRESULT", "Records", true, true);
                    break;
            }
            switch (ST.financial_crimes.ToString())
            {
                case "Clear":
                    doc.Replace("FCRIMECOMMENT", "", true, true);
                    doc.Replace("FCRIMERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("FCRIMECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("FCRIMERESULT", "Records", true, true);
                    break;
            }                      
            switch (ST.national_credit.ToString())
            {
                case "Clear":
                    doc.Replace("NATCRECOMMENT", "", true, true);
                    doc.Replace("NATCRERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("NATCRECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("NATCRERESULT", "Records", true, true);
                    break;
            }
            switch (ST.new_york_stock.ToString())
            {
                case "Clear":
                    doc.Replace("NYCOMMENT", "", true, true);
                    doc.Replace("NYRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("NYCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("NYRESULT", "Records", true, true);
                    break;
            }
            switch (ST.Office_of_comptroller.ToString())
            {
                case "Clear":
                    doc.Replace("OFFCPTCOMMENT", "", true, true);
                    doc.Replace("OFFCPTRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("OFFCPTCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("OFFCPTRESULT", "Records", true, true);
                    break;
            }
            switch (ST.Office_of_superintendent.ToString())
            {
                case "Clear":
                    doc.Replace("OFFSUPCOMMENT", "", true, true);
                    doc.Replace("OFFSUPRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("OFFSUPCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("OFFSUPRESULT", "Records", true, true);
                    break;
            }
            switch (ST.resolution_trust.ToString())
            {
                case "Clear":
                    doc.Replace("RESTCOMMENT", "", true, true);
                    doc.Replace("RESTRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("RESTCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("RESTRESULT", "Records", true, true);
                    break;
            }
           
            switch (ST.US_court.ToString())
            {
                case "Clear":
                    doc.Replace("USCOURCOMMENT", "", true, true);
                    doc.Replace("USCOURESULT", "Clear", true, true);
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
                    doc.Replace("USDPJSRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USDPJSCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USDPJSRESULT", "Records", true, true);
                    break;
            }
            switch (ST.US_federal_trade.ToString())
            {
                case "Clear":
                    doc.Replace("USFEDCOMMENT", "", true, true);
                    doc.Replace("USFEDRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USFEDCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USFEDRESULT", "Records", true, true);
                    break;
            }
           
            switch (ST.US_office_thrifts.ToString())
            {
                case "Clear":
                    doc.Replace("USOTCOMMENT", "", true, true);
                    doc.Replace("USOTRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("USOTCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USOTRESULT", "Records", true, true);
                    break;
            }
            switch (ST.central_intelligence.ToString())
            {
                case "Clear":
                    doc.Replace("CICOMMENT", "", true, true);
                    doc.Replace("CIRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("CICOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("CIRESULT", "Records", true, true);
                    break;
            }
            switch (ST.city_london_police.ToString())
            {
                case "Clear":
                    doc.Replace("CITYCOMMENT", "", true, true);
                    doc.Replace("CITYRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("CITYCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("CITYRESULT", "Records", true, true);
                    break;
            }
            switch (ST.constabularies_cheshire.ToString())
            {
                case "Clear":
                    doc.Replace("COSTABCOMMENT", "", true, true);
                    doc.Replace("COSTABRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("COSTABCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("COSTABRESULT", "Records", true, true);
                    break;
            }
            switch (ST.hampshire_police.ToString())
            {
                case "Clear":
                    doc.Replace("HAMPSPCOMMENT", "", true, true);
                    doc.Replace("HAMPSPRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("HAMPSPCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("HAMPSPRESULT", "Records", true, true);
                    break;
            }
            switch (ST.hong_kong_police.ToString())
            {
                case "Clear":
                    doc.Replace("HONGKONGCOMMENT", "", true, true);
                    doc.Replace("HONGKONGRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("HONGKONGCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("HONGKONGRESULT", "Records", true, true);
                    break;
            }
            switch (ST.metropolitan_police.ToString())
            {
                case "Clear":
                    doc.Replace("METROPOLICOMMENT", "", true, true);
                    doc.Replace("METROLIRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("METROPOLICOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("METROLIRESULT", "Records", true, true);
                    break;
            }
            switch (ST.national_crime.ToString())
            {
                case "Clear":
                    doc.Replace("NATLSQADCOMMENT", "", true, true);
                    doc.Replace("NATLSQADRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("NATLSQADCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("NATLSQADRESULT", "Records", true, true);
                    break;
            }
            switch (ST.north_yorkshire_polic.ToString())
            {
                case "Clear":
                    doc.Replace("NOYOKCOMMENT", "", true, true);
                    doc.Replace("NOYOKRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("NOYOKCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("NOYOKRESULT", "Records", true, true);
                    break;
            }
            switch (ST.nottinghamshire_police.ToString())
            {
                case "Clear":
                    doc.Replace("NOTTINGCOMMENT", "", true, true);
                    doc.Replace("NOTTINGRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("NOTTINGCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("NOTTINGRESULT", "Records", true, true);
                    break;
            }
            switch (ST.surrey_police.ToString())
            {
                case "Clear":
                    doc.Replace("SURREYCOMMENT", "", true, true);
                    doc.Replace("SURREYRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("SURREYCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("SURREYRESULT", "Records", true, true);
                    break;
            }
            switch (ST.thames_valley_police.ToString())
            {
                case "Clear":
                    doc.Replace("THAMCOMMENT", "", true, true);
                    doc.Replace("THAMRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("THAMCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("THAMRESULT", "Records", true, true);
                    break;
            }
            switch (ST.warwickshire_police.ToString())
            {
                case "Clear":
                    doc.Replace("WARWCOMMENT", "", true, true);
                    doc.Replace("WARWRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("WARWCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("WARWRESULT", "Records", true, true);
                    break;
            }
            switch (ST.alberta_securities_commission.ToString())
            {
                case "Clear":
                    doc.Replace("ALBCOMMENT", "", true, true);
                    doc.Replace("ALBRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("ALBCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("ALBRESULT", "Records", true, true);
                    break;
            }
            switch (ST.asset_recovery_agency.ToString())
            {
                case "Clear":
                    doc.Replace("ASSECOMMENT", "", true, true);
                    doc.Replace("ASSERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("ASSECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("ASSERESULT", "Records", true, true);
                    break;
            }
            switch (ST.australian_prudential.ToString())
            {
                case "Clear":
                    doc.Replace("AUSPRCOMMENT", "", true, true);
                    doc.Replace("AUSPRRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("AUSPRCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("AUSPRRESULT", "Records", true, true);
                    break;
            }
            switch (ST.australian_securities.ToString())
            {
                case "Clear":
                    doc.Replace("AUSSECCOMMENT", "", true, true);
                    doc.Replace("AUSSECRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("AUSSECCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("AUSSECRESULT", "Records", true, true);
                    break;
            }
            switch (ST.banque_de_CECEI.ToString())
            {
                case "Clear":
                    doc.Replace("BAQUECOMMENT", "", true, true);
                    doc.Replace("BAQUERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("BAQUECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("BAQUERESULT", "Records", true, true);
                    break;
            }
            switch (ST.banque_de_commission.ToString())
            {
                case "Clear":
                    doc.Replace("BACOMCOMMENT", "", true, true);
                    doc.Replace("BACOMRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("BACOMCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("BACOMRESULT", "Records", true, true);
                    break;
            }
            switch (ST.british_virgin_islands.ToString())
            {
                case "Clear":
                    doc.Replace("BIRVIRCOMMENT", "", true, true);
                    doc.Replace("BRIVIRRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("BIRVIRCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("BRIVIRRESULT", "Records", true, true);
                    break;
            }
            switch (ST.cayman_islands_monetary.ToString())
            {
                case "Clear":
                    doc.Replace("CAYCOMMENT", "", true, true);
                    doc.Replace("CAYRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("CAYCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("CAYRESULT", "Records", true, true);
                    break;
            }
            switch (ST.commission_de_surveillance.ToString())
            {
                case "Clear":
                    doc.Replace("COMDECOMMENT", "", true, true);
                    doc.Replace("COMDERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("COMDECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("COMDERESULT", "Records", true, true);
                    break;
            }
            switch (ST.council_financial_activities.ToString())
            {
                case "Clear":
                    doc.Replace("COUNFINCOMMENT", "", true, true);
                    doc.Replace("COUNFINRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("COUNFINCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("HONGKONGRESULT", "Records", true, true);
                    break;
            }
            switch (ST.departamento_de_investigacoes.ToString())
            {
                case "Clear":
                    doc.Replace("MENTODECOMMENT", "", true, true);
                    doc.Replace("MENTODERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("MENTODECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("MENTODERESULT", "Records", true, true);
                    break;
            }
            switch (ST.department_labour_inspection.ToString())
            {
                case "Clear":
                    doc.Replace("DEPLABCOMMENT", "", true, true);
                    doc.Replace("DEPLABRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("DEPLABCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("DEPLABRESULT", "Records", true, true);
                    break;
            }
            switch (ST.financial_action_task.ToString())
            {
                case "Clear":
                    doc.Replace("FINACCOMMENT", "", true, true);
                    doc.Replace("FINACRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("FINACCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("FINACRESULT", "Records", true, true);
                    break;
            }
            switch (ST.financial_regulator_ireland.ToString())
            {
                case "Clear":
                    doc.Replace("REGIRECOMMENT", "", true, true);
                    doc.Replace("REGIRERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("REGIRECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("REGIRERESULT", "Records", true, true);
                    break;
            }
            switch (ST.hongkong_monetary_authority.ToString())
            {
                case "Clear":
                    doc.Replace("KONMONCOMMENT", "", true, true);
                    doc.Replace("KONMONRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("KONMONCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("KONMONRESULT", "Records", true, true);
                    break;
            }            
            switch (ST.investment_association_Canada.ToString())
            {
                case "Clear":
                    doc.Replace("INDEASCOMMENT", "", true, true);
                    doc.Replace("INDEASRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("INDEASCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("INDEASRESULT", "Records", true, true);
                    break;
            }
            switch (ST.investment_management_regulatory.ToString())
            {
                case "Clear":
                    doc.Replace("INMARECOMMENT", "", true, true);
                    doc.Replace("INMARERESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("INMARECOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("INMARERESULT", "Records", true, true);
                    break;
            }
            switch (ST.isle_financial_supervision.ToString())
            {
                case "Clear":
                    doc.Replace("ISMASUCOMMENT", "", true, true);
                    doc.Replace("ISMASURESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("ISMASUCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("ISMASURESULT", "Records", true, true);
                    break;
            }
            switch (ST.jersey_financial_commission.ToString())
            {
                case "Clear":
                    doc.Replace("JESECOCOMMENT", "", true, true);
                    doc.Replace("JESECORESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("JESECOCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("JESECORESULT", "Records", true, true);
                    break;
            }
            switch (ST.lloyd_insurance_arimbolaet.ToString())
            {
                case "Clear":
                    doc.Replace("LIARCOMMENT", "", true, true);
                    doc.Replace("LIARRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("LIARCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("LIARRESULT", "Records", true, true);
                    break;
            }
            switch (ST.monetary_authority_singapore.ToString())
            {
                case "Clear":
                    doc.Replace("MOSICOMMENT", "", true, true);
                    doc.Replace("MOSIRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("MOSICOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("MOSIRESULT", "Records", true, true);
                    break;
            }
            switch (ST.securities_exchange_commission.ToString())
            {
                case "Clear":
                    doc.Replace("SECOBRCOMMENT", "", true, true);
                    doc.Replace("SECOBRRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("SECOBRCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("SECOBRRESULT", "Records", true, true);
                    break;
            }
            switch (ST.securities_futuresauthority.ToString())
            {
                case "Clear":
                    doc.Replace("SEFUAUCOMMENT", "", true, true);
                    doc.Replace("SEFUAURESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("SEFUAUCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("SEFUAURESULT", "Records", true, true);
                    break;
            }
            switch (ST.swedish_financial_supervisory.ToString())
            {
                case "Clear":
                    doc.Replace("SWFICOMMENT", "", true, true);
                    doc.Replace("SWFIRESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("SWFICOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("SWFIRESULT", "Records", true, true);
                    break;
            }
            switch (ST.swiss_federal_banking.ToString())
            {
                case "Clear":
                    doc.Replace("SWBACOMMENT", "", true, true);
                    doc.Replace("SWBARESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("SWBACOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("SWBARESULT", "Records", true, true);
                    break;
            }
            switch (ST.U_K_companies_disqualified.ToString())
            {
                case "Clear":
                    doc.Replace("UKCOHOCOMMENT", "", true, true);
                    doc.Replace("UKCOHORESULT", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("UKCOHOCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("UKCOHORESULT", "Records", true, true);
                    break;
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
            doc.Replace("[Last Name]", last_name.ToString(), true, true);
            if (country.ToString().Equals("United Kingdom")) {
                doc.Replace("the [Country]", "[Country]", true, true);
                doc.SaveToFile(savePath);
                doc.Replace("[Country]", string.Concat("the ",country), true, false);
                doc.SaveToFile(savePath);
                doc.Replace("the The United Kingdom", "the United Kingdom", true, true);
            }
            else {
                doc.Replace("[Country]", country, true, false);
            }
            doc.SaveToFile(savePath);
            doc.Replace("[FirstName]", first_name.ToString(), true, true);
            doc.Replace("[LastName]", last_name.ToString(), true, true);
            doc.Replace("[HeaderCountry]", country, true, true);
            doc.Replace("  ", " ", true, false);
            doc.Replace("  (", " (", true, false);
            doc.Replace("united states", "United States", true, false);
            doc.Replace("investigator", "Investigator", true, false);
            doc.Replace(" .", ".", true, false);
            doc.Replace(" ,", ",", true, false);
            doc.SaveToFile(savePath);
            doc.Replace(".  ", ". ", true, false);
            doc.Replace(". ", ".  ", true, false);
            doc.SaveToFile(savePath);
            try
            {
                if (middleinitial == "") {
                    doc.Replace(" [MiddleInitial]", "", true, false);
                    doc.Replace(" Middle_Initial", "", true, false);
                    doc.Replace(" MiddleInitial", "", true, false);
                    
                }
                else
                {
                    doc.Replace("[MiddleInitial]", middleinitial, true, true);
                    doc.Replace("Middle_Initial", middleinitial.ToUpper().ToString().TrimEnd(), true, true);
                    doc.Replace("MiddleInitial", middleinitial.TrimEnd(), true, true);
                    
                }
            }
            catch
            {
                doc.Replace(" [MiddleInitial]", "", true, false);
                doc.Replace(" Middle_Initial", "", true, false);
                doc.Replace(" MiddleInitial", "", true, false);
                
            }
            doc.SaveToFile(savePath);
            doc.Replace("U.S.  ", "U.S. ", true, false);
            doc.Replace("U.K.  ", "U.K. ", true, false);
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
            savePath = string.Concat(savePath,last_name, "_SterlingDiligenceReport(", case_number, ")_DRAFT.docx");
            // string path = AppDomain.CurrentDomain.BaseDirectory + "FolderName/";
            byte[] fileBytes = System.IO.File.ReadAllBytes(savePath);
            string fileName = string.Concat(last_name, "_SterlingDiligenceReport(", case_number, ")_DRAFT.docx");
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }
    }
}   