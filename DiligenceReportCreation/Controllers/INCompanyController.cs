using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DiligenceReportCreation.Data;
using DiligenceReportCreation.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace DiligenceReportCreation.Controllers
{
    public class INCOMPANYController : Controller
    {
        private readonly DataBaseContext _context;
        private readonly IConfiguration _config;
        public INCOMPANYController(IConfiguration config, Data.DataBaseContext context)
        {
            _config = config;
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult IN_COMPANY_page1()
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
            ReportModel comment12 = _context.reportModel
                              .Where(u => u.record_Id == recordid)
                              .FirstOrDefault();
            ViewBag.CaseNumber = comment12.casenumber;
            ViewBag.LastName = comment12.lastname;
            HttpContext.Session.SetString("recordid", recordid);
            HttpContext.Session.SetString("FullEntityName", comment12.lastname);
            HttpContext.Session.SetString("casenumber", comment12.casenumber);
            return View();
        }
        [HttpPost]
        public IActionResult IN_COMPANY_page1(MainModel mainModel, string SaveData, string Submit)
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
            string FullEntityName = HttpContext.Session.GetString("FullEntityName");
            string casenumber = HttpContext.Session.GetString("casenumber");
            string recordid = HttpContext.Session.GetString("recordid");
            ReportModel report = _context.reportModel
               .Where(u => u.record_Id == recordid)
                                 .FirstOrDefault();
            if (Submit == "SubmitData")
            {
                MainModel main = new MainModel();
                DiligenceInputModel dinput = new DiligenceInputModel();
                Otherdatails otherdatails = new Otherdatails();
                CountrySpecificModel countrySpecificModel = new CountrySpecificModel();
                //Additional_statesModel additional = new Additional_statesModel();
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
                if (TempData["message"] == null)
                {
                    main.diligenceInputModel = dinput;
                    main.otherdetails = otherdatails;
                    main.pllicenseModels = _context.DbPLLicense
                               .Where(u => u.record_Id == recordid)
                                .ToList();
                    main.csModel = countrySpecificModel;
                    main.summarymodel = summary;
                    return Save_Page1(main);
                }
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
                            DiligenceInputModel DI1 = new DiligenceInputModel();

                            try
                            {
                                DI1 = _context.DbPersonalInfo
                              .Where(a => a.record_Id == recordid)
                              .FirstOrDefault();
                                DI1.record_Id = recordid.ToString();
                            }
                            catch
                            {
                                DI1 = new DiligenceInputModel();
                                DI1.record_Id = recordid.ToString();
                            }
                            DI1.ClientName = mainModel.diligenceInputModel.ClientName;
                            DI1.LastName = report.lastname;
                            DI1.CaseNumber = report.casenumber;
                            DI1.FullaliasName = mainModel.diligenceInputModel.FullaliasName;
                            DI1.Dob = mainModel.diligenceInputModel.Dob;
                            DI1.CurrentCity = mainModel.diligenceInputModel.CurrentCity;
                            DI1.CurrentState = mainModel.diligenceInputModel.CurrentState;
                            DI1.CurrentCountry = mainModel.diligenceInputModel.CurrentCountry;
                            DI1.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                            DI1.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                            //DI1.Employer1State = mainModel.diligenceInputModel.Employer1State;
                            DI1.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
                            DI1.CommonNameSubject = mainModel.diligenceInputModel.CommonNameSubject;
                            try
                            {
                                if (DI1 == null)
                                {
                                    _context.DbPersonalInfo.Add(DI1);
                                    _context.SaveChanges();
                                }
                                else
                                {
                                    _context.DbPersonalInfo.Update(DI1);
                                    _context.SaveChanges();
                                }
                            }
                            catch
                            {
                                _context.DbPersonalInfo.Add(DI1);
                                _context.SaveChanges();
                            }
                            TempData["PI"] = "Done";
                        }
                    }
                    catch (Exception e)
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
                            Otherdatails strupdate = new Otherdatails();
                            try
                            {
                                strupdate = _context.othersModel
                                   .Where(a => a.record_Id == recordid)
                                   .FirstOrDefault();
                                strupdate.record_Id = recordid;
                            }
                            catch
                            {
                                strupdate = new Otherdatails();
                                strupdate.record_Id = recordid;
                            }
                            
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
                            try
                            {
                                if (strupdate == null)
                                {
                                    _context.othersModel.Add(strupdate);
                                    _context.SaveChanges();
                                }
                                else
                                {
                                    _context.othersModel.Update(strupdate);
                                    _context.SaveChanges();
                                }
                            }
                            catch
                            {
                                _context.othersModel.Add(strupdate);
                                _context.SaveChanges();
                            }
                            TempData["OD"] = "Done";
                        }
                    }
                    catch (Exception e)
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
                                    
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.pprrecordhits = mainModel.csModel.pprrecordhits;
                                    break;
                                case "Canada":
                                    countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;                                                                                                         
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.pprrecordhits = mainModel.csModel.pprrecordhits;
                                    break;
                                case "United Kingdom":                                    
                                    
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.reg_trust_hits = mainModel.csModel.reg_trust_hits;
                                    countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                    break;
                                case "France":
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;

                                    break;
                                case "India":
                                    
                                    countrySpecific.india_corpregistry = mainModel.csModel.india_corpregistry;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;

                                    break;
                                case "Germany":
                                    
                                    countrySpecific.germany_regsearches = mainModel.csModel.germany_regsearches;
                                    
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "Switzerland":
                                    
                                    
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "United Arab Emirates":
                                    
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                default:
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    
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
                        if (SaveData.Contains("PL"))
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
          
            HttpContext.Session.SetString("FullEntityName", FullEntityName);
            HttpContext.Session.SetString("casenumber", casenumber);
            HttpContext.Session.SetString("recordid", recordid);
            ViewBag.CaseNumber = casenumber;
            ViewBag.LastName = FullEntityName;          
            return View(mainModel);
        }
        [HttpGet]
        public IActionResult IN_COMPANY_Edit()
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
            MainModel main = new MainModel();
            DiligenceInputModel diligence = _context.DbPersonalInfo
                .Where(a => a.record_Id == recordid)
                .FirstOrDefault();
            Otherdatails otherdatails = _context.othersModel
                                  .Where(u => u.record_Id == recordid)
                                  .FirstOrDefault();
            CountrySpecificModel countrySpecificModel = _context.CSComment
                       .Where(u => u.record_Id == recordid)
                       .FirstOrDefault();
            main.pllicenseModels = _context.DbPLLicense
                       .Where(u => u.record_Id == recordid)
                        .ToList();
            SummaryResulttableModel summary = _context.summaryResulttableModels
                 .Where(u => u.record_Id == recordid)
                                   .FirstOrDefault();
            main.summarymodel = summary;
            main.diligenceInputModel = diligence;
            main.otherdetails = otherdatails;
            main.csModel = countrySpecificModel;
            HttpContext.Session.SetString("recordid", recordid);
            ReportModel report = _context.reportModel
                .Where(u => u.record_Id == recordid)
                                  .FirstOrDefault();
            ViewBag.CaseNumber = report.casenumber;
            ViewBag.LastName = report.lastname;
            HttpContext.Session.SetString("FullEntityName", report.lastname);
            HttpContext.Session.SetString("casenumber", report.casenumber);
            return View(main);
        }
        [HttpPost]
        public IActionResult IN_COMPANY_Edit(MainModel mainModel, string SaveData, string Submit)
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
            string FullEntityName = HttpContext.Session.GetString("FullEntityName");
            string casenumber = HttpContext.Session.GetString("casenumber");
            string recordid = HttpContext.Session.GetString("recordid");
            ReportModel report = _context.reportModel
               .Where(u => u.record_Id == recordid)
                                 .FirstOrDefault();
            if (Submit == "SubmitData")
            {
                MainModel main = new MainModel();
                DiligenceInputModel dinput = new DiligenceInputModel();
                Otherdatails otherdatails = new Otherdatails();
                CountrySpecificModel countrySpecificModel = new CountrySpecificModel();
                //Additional_statesModel additional = new Additional_statesModel();
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
                if(TempData["message"] == null)
                {
                    main.diligenceInputModel = dinput;
                    main.otherdetails = otherdatails;
                    main.pllicenseModels = _context.DbPLLicense
                               .Where(u => u.record_Id == recordid)
                               .ToList();
                    main.csModel = countrySpecificModel;
                    main.summarymodel = summary;
                    return Save_Page1(main);
                }
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
                            DiligenceInputModel DI1 = new DiligenceInputModel();

                            try
                            {
                                DI1 = _context.DbPersonalInfo
                              .Where(a => a.record_Id == recordid)
                              .FirstOrDefault();
                                DI1.record_Id = recordid.ToString();
                            }
                            catch
                            {
                                DI1 = new DiligenceInputModel();
                                DI1.record_Id = recordid.ToString();
                            }
                            DI1.ClientName = mainModel.diligenceInputModel.ClientName;
                            DI1.LastName = report.lastname;
                            DI1.CaseNumber = report.casenumber;
                            DI1.FullaliasName = mainModel.diligenceInputModel.FullaliasName;
                            DI1.Dob = mainModel.diligenceInputModel.Dob;
                            DI1.CurrentCity = mainModel.diligenceInputModel.CurrentCity;
                            DI1.CurrentState = mainModel.diligenceInputModel.CurrentState;
                            DI1.CurrentCountry = mainModel.diligenceInputModel.CurrentCountry;
                            DI1.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                            DI1.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                            DI1.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
                          //  DI1.Employer1State = mainModel.diligenceInputModel.Employer1State;
                            DI1.CommonNameSubject = mainModel.diligenceInputModel.CommonNameSubject;
                            try
                            {
                                if (DI1 == null)
                                {
                                    _context.DbPersonalInfo.Add(DI1);
                                    _context.SaveChanges();
                                }
                                else
                                {
                                    _context.DbPersonalInfo.Update(DI1);
                                    _context.SaveChanges();
                                }
                            }
                            catch
                            {
                                _context.DbPersonalInfo.Add(DI1);
                                _context.SaveChanges();
                            }
                            TempData["PI"] = "Done";
                        }
                    }
                    catch (Exception e)
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
                            Otherdatails strupdate = new Otherdatails();
                            try
                            {
                                strupdate = _context.othersModel
                                   .Where(a => a.record_Id == recordid)
                                   .FirstOrDefault();
                                strupdate.record_Id = recordid;
                            }
                            catch
                            {
                                strupdate = new Otherdatails();
                                strupdate.record_Id = recordid;
                            }
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
                            try
                            {
                                if (strupdate == null)
                                {
                                    _context.othersModel.Add(strupdate);
                                    _context.SaveChanges();
                                }
                                else
                                {
                                    _context.othersModel.Update(strupdate);
                                    _context.SaveChanges();
                                }
                            }
                            catch
                            {
                                _context.othersModel.Add(strupdate);
                                _context.SaveChanges();
                            }
                            TempData["OD"] = "Done";
                        }
                    }
                    catch (Exception e)
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
                                    
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.pprrecordhits = mainModel.csModel.pprrecordhits;
                                    break;
                                case "Canada":
                                    countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.pprrecordhits = mainModel.csModel.pprrecordhits;
                                    break;
                                case "United Kingdom":                                                                       
                                    countrySpecific.insolvency_hits = mainModel.csModel.insolvency_hits;
                                    countrySpecific.reg_trust_hits = mainModel.csModel.reg_trust_hits;
                                    countrySpecific.civil_record_hits = mainModel.csModel.civil_record_hits;
                                    break;
                                case "France":
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "India":                                   
                                    countrySpecific.india_corpregistry = mainModel.csModel.india_corpregistry;
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "Germany":
                                    
                                    countrySpecific.germany_regsearches = mainModel.csModel.germany_regsearches;
                                    
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "Switzerland":
                                    
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                case "United Arab Emirates":
                                    
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    break;
                                default:
                                    countrySpecific.country_specific_reg_hits = mainModel.csModel.country_specific_reg_hits;
                                    
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
                        if (SaveData.Contains("PL"))
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
                    catch
                    {

                    }
                    break;
            }
            HttpContext.Session.SetString("FullEntityName", FullEntityName);
            HttpContext.Session.SetString("casenumber", casenumber);
            HttpContext.Session.SetString("recordid", recordid);
            ViewBag.CaseNumber = casenumber;
            ViewBag.LastName = FullEntityName;          
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
        public IActionResult Save_Page1(MainModel diligenceInput)
        {
            string templatePath;
            string savePath = _config.GetValue<string>("ReportPath");
            templatePath = _config.GetValue<string>("INCompanytemplatePath");           
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
            Document doc = new Document(templatePath);
            string current_full_address = "";
            string currentstreet;
            string currentcity;
            string currentcountry;
            string strblnres = "";            
            if (diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("Hong Kong SAR"))
            {
                diligenceInput.diligenceInputModel.CurrentCountry = "Hong Kong";
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
                    }
                    catch { }
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
            //doc.Replace("First_Name",diligenceInput.diligenceInputModel.FullaliasName.ToUpper().TrimEnd().ToString(), true, true);
           
            //doc.Replace("FirstName",diligenceInput.diligenceInputModel.FullaliasName.TrimEnd(), true, true);
            doc.Replace("ClientName", diligenceInput.diligenceInputModel.ClientName.TrimEnd(), true, false);
         
                     
            doc.Replace("ClientName", diligenceInput.diligenceInputModel.ClientName.TrimEnd(), true, false);
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
                if (diligenceInput.diligenceInputModel.CurrentCity.ToString().ToUpper().Equals("NA") || diligenceInput.diligenceInputModel.CurrentCity.ToString().ToUpper().Equals("N/A") || diligenceInput.diligenceInputModel.CurrentCity.ToString().Equals(""))
                {
                    doc.Replace("[City], ", "", true, false);
                    doc.SaveToFile(savePath);
                    doc.Replace("[City]", "<not provided>", true, true);
                }
                else
                {
                    doc.Replace("[City]", diligenceInput.diligenceInputModel.CurrentCity.ToString(), true, true);
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
            if (current_full_address.ToString().Equals("")) { doc.Replace("[FullEntityAddress]", "<not provided>", true, true); }
            else
            {
                doc.Replace("[FullEntityAddress]", current_full_address, true, true);
            }
            //doc.Replace("FullAliasName", diligenceInput.FullaliasName, true, true);            
            doc.Replace("[CaseNumber]", diligenceInput.diligenceInputModel.CaseNumber.TrimEnd(), true, true);
            try
            {
                if (diligenceInput.diligenceInputModel.MiddleName.ToString() == "")
                {
                    doc.Replace(" MiddleName", "", false, false);
                }
                else
                {
                    doc.Replace("MiddleName", diligenceInput.diligenceInputModel.MiddleName.ToString().TrimEnd(), true, true);
                }
            }
            catch
            {
                doc.Replace(" MiddleName", "", false, false);
            }
            
        
            
            
            doc.SaveToFile(savePath);
            string bnres = "";
            //PL License details
            string pl_comment = "";
            for (int i = 0; i < diligenceInput.pllicenseModels.Count; i++)
            {
               string plstartdate = "<not provided>";
               string plenddate = "<not provided>";
                try
                {
                    if (diligenceInput.pllicenseModels[i].General_PL_License.ToString().Equals("Yes"))
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
                        string strplorgfont = string.Concat(diligenceInput.pllicenseModels[i].PL_Organization.ToString(), " CHANGEFONTHEADER");
                        if (i == 0) { pl_comment = "\n"; }
                        if (diligenceInput.pllicenseModels[i].PL_Confirmed.Equals("Yes"))
                        {
                            if (diligenceInput.pllicenseModels.Count - 1 == i)
                            {
                                pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.confirmed_comment.ToString(), "\n\nThere was no disciplinary history identified in connection with the subject entity’s license. <investigator to modify if disciplinary history exists>");
                            }
                            else
                            {
                                pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.confirmed_comment.ToString(), "\n\nThere was no disciplinary history identified in connection with the subject entity’s license. <investigator to modify if disciplinary history exists>", "\n\n");
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
                        if (diligenceInput.diligenceInputModel.FullaliasName.ToString().Equals(""))
                        {
                            pl_comment = pl_comment.Replace("[Last Name]", "<not provided>");
                        }
                        else
                        {
                            pl_comment = pl_comment.Replace("[Last Name]", diligenceInput.diligenceInputModel.FullaliasName.ToString());
                        }
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
                catch
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
            
            if (diligenceInput.summarymodel.bankruptcy_filings.ToString().Equals("Results Pending") || diligenceInput.summarymodel.civil_court_Litigation.ToString().Equals("Results Pending") )
            {
                doc.Replace("[DESCRESULTPENDINGCRIMM]", "\n\n<Search type> searches are currently ongoing in <jurisdiction(s)>, the results of which will be provided under separate cover upon receipt.", true, false);
            }
            else
            {
                doc.Replace("[DESCRESULTPENDINGCRIMM]", "", true, false);
            }
            doc.SaveToFile(savePath);
            //Additional Country
            string country_comment = "";
            if (diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Equals(""))
            {
                country_comment = "";
            }
            else
            {
                CommentModel comment2 = _context.DbComment
                              .Where(u => u.Comment_type == "NonScopeCountry")
                              .FirstOrDefault();
                if (diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Contains(",") || diligenceInput.diligenceInputModel.Nonscopecountry1.ToString().Contains(" and "))
                {
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
                strbusumcomm = "[FullEntityName] was formed in [IncorporationState] in [IncorporationYear]\n\nFurther, other affiliated business entities were also identified";
                business_comment = "In addition to the above, research efforts conducted through information maintained by [Corp Registry] identified [ShortEntityName] in connection with the following other business entities in [Country]: <investigator to modify the below list as needed (i.e. footnote any relevant registration details, delineate Active v. Inactive entities  and/or add language/asterisks symbols for entities – as appropriate>\n\n\t•\t<Affiliate Name>[addbusinessfootnote]\n\t•\t<Affiliate Name>\n\t•\t<Affiliate Name>";
            }
            else
            {
                business_comment = "In addition to the above, research efforts conducted through information maintained by [Corp Registry] did not identify [ShortEntityName] in connection with any business entities in [Country].";
                doc.Replace("COMPRESULT", "Clear", true, true);
                strbusumcomm = "[FullEntityName] was formed in [IncorporationState] in [IncorporationYear]";
            }
            doc.Replace("BUSINESSAFFILIATIONSIDENTIFIED", business_comment, true, true);          
            doc.Replace("COMPCOMMENT", strbusumcomm, true, true);
            doc.SaveToFile(savePath);
                               
            try
            {
                string blnconfound = "";
                if (diligenceInput.otherdetails.Has_Business_Affiliations.ToString().Equals("Yes"))
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
                                    if (abc.ToString().EndsWith("[addbusinessfootnote]"))
                                    {
                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                        //Insert footnote1 after the word "Spire.Doc"
                                        para.ChildObjects.Insert(i + 1, footnote1);
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("According to records maintained by the [Corp Registry], <Affiliate Name> was formed in <Country> on <date>, where it is <status>.");
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        doc.SaveToFile(savePath);
                                        doc.Replace("[addbusinessfootnote]", "", true, false);
                                        doc.SaveToFile(savePath);
                                        blnconfound = "true";
                                        break;
                                    }

                                }
                            }
                            if (blnconfound == "true")
                            {
                                break;
                            }
                        }
                        if (blnconfound == "true")
                        {
                            break;
                        }
                    }
                }
            }
            catch { }
            doc.SaveToFile(savePath);

            //Intellectual Hits
            string intellectual_comment;
            CommentModel intellec_comment2 = _context.DbComment
                               .Where(u => u.Comment_type == "Intellectual_hits")
                               .FirstOrDefault();
            if (diligenceInput.otherdetails.Has_Intellectual_Hits.ToString().Equals("Yes"))
            {
                intellectual_comment = "Moreover, research of records maintained by the Patent and Trademark Office in [Country] and the World Intellectual Property Organization (“WIPO”) identified the subject entity as a Registrant/Owner in connection with the following trademarks and/or as an Applicant/Assignee/Inventor in connection with the following patents:";
                intellectual_comment = string.Concat("\n", intellectual_comment.ToString(), "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n");
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
            globalhit_comment = globalhit_comment.Replace("[LastName]", "[ShortEntityName]");
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
            icij_comment = icij_comment.Replace("subject", "subject entity");
            doc.Replace("ICIJHITSDESCRIPTION", icij_comment, true, true);
            doc.SaveToFile(savePath);
            //Press and media
            string PMCommentModel = "";
            //PRESSANDMEDIASEARCHDESCRIPTION
            switch (diligenceInput.otherdetails.Press_Media.ToString())
            {
                case "Common name with adverse Hits":                    
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the commonality of the subject entity’s nomenclature, searches were structured and designed to identify specific reports relating to [ShortEntityName] with a variety of negative search terms and derivatives thereof within the past <timeframe>, and a thorough review of the same revealed the following adverse and/or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);
                    break;
                case "Common name without adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the commonality of the subject entity’s nomenclature, searches were structured and designed to identify specific reports relating to [ShortEntityName] with a variety of negative search terms and derivatives thereof within the past <timeframe>, and a thorough review of the same did not reveal any adverse or materially-significant information in connection with [ShortEntityName].", true, true);
                    break;
                case "High volume with adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the volume of press and other media sources identified in connection with the subject entity, searches were structured and designed to identify specific reports relating to [ShortEntityName] with a variety of negative search terms and derivatives thereof within the past <timeframe>, and a thorough review of the same revealed the following adverse and/or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);
                    break;
                case "High volume without adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the volume of press and other media sources identified in connection with the subject entity, searches were structured and designed to identify specific reports relating to [ShortEntityName] with a variety of negative search terms and derivatives thereof within the past <timeframe>, and a thorough review of the same did not reveal any adverse or materially-significant information in connection with [ShortEntityName].", true, true);
                    break;
                case "Standard search with adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, identified various articles and/or media references in connection with [ShortEntityName], and a thorough review of the same revealed the following adverse and/or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);
                    break;
                case "Standard search without adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, identified various articles and/or media references in connection with [ShortEntityName], and a thorough review of the same did not identify any adverse or materially-significant information.", true, true);
                    break;
                case "No Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, did not identify any articles and/or media references in connection with [ShortEntityName].", true, true);
                    break;
            }
            doc.SaveToFile(savePath);
            doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", PMCommentModel.ToString(), true, true);
            //US_SEC
            string usseccommentmodel = "";
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
                    doc.Replace("SECCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("SECRESULT", "Records", true, true);
                }
                else
                {
                    doc.Replace("SECCOMMENT", "", true, true);
                    doc.Replace("SECRESULT", "Clear", true, true);
                }
                // doc.Replace("US_SECHEADER", " \nUnited States Securities and Exchange Commission \n", true, true);
                usseccommentmodel = string.Concat("\nUnited States Securities and Exchange Commission\n\n", "According to a Uniform Application for Investment Advisor Registration (“Form ADV”)[ussecfootnote]", "\n");
            }
            //UK_FCA
            string ukfcacommentmodel = "";
            if (diligenceInput.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered"))
            {
                ukfcacommentmodel = "";
                doc.Replace("UKFICOCOMMENT", "", true, true);
                doc.Replace("UKFICORESULT", "Clear", true, true);
            }
            else
            {
                if (diligenceInput.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Yes-With Adverse"))
                {
                    ukfcacommentmodel = "According to records maintained by the United Kingdom’s Financial Conduct Authority (“FCA”), [ShortEntityName], with a registration number of <number>, has been registered as an authorized firm since <date – Investigator should include FSA and predecessor dates/organizations, as applicable>. The firm has permission to conduct the following investment activities: <Investigator to list investment activities>.\n\nFurthermore, [ShortEntityName] is able to exercise passporting rights in order to issue Markets in Financial Instruments Directive (“MiFID”) Outward Service to: <Investigator to list countries – as appropriate>.\n\nAdditionally, the FCA has approved the following: <Investigator to summarize firm/individuals with current and/or previous involvements with the firm>.\n\nFurther, the following supervisory, disciplinary and/or civil regulatory actions are on file with the FCA against the subject entity:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                    doc.Replace("UKFICOCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("UKFICORESULT", "Records", true, true);
                }
                else
                {
                    ukfcacommentmodel = "According to records maintained by the United Kingdom’s Financial Conduct Authority (“FCA”), [ShortEntityName], with a registration number of <number>, has been registered as an authorized firm since <date – Investigator should include FSA and predecessor dates/organizations, as applicable>. The firm has permission to conduct the following investment activities: <Investigator to list investment activities>.\n\nFurthermore, [ShortEntityName] is able to exercise passporting rights in order to issue Markets in Financial Instruments Directive (“MiFID”) Outward Service to: <Investigator to list countries – as appropriate>.\n\nAdditionally, the FCA has approved the following: <Investigator to summarize firm/individuals with current and/or previous involvements with the firm>.\n\nThere are no records of any supervisory, disciplinary or civil regulatory actions on file with the FCA against the subject entity.";
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
            if (diligenceInput.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered"))
            {
                finracommentmodel = "";
                doc.Replace("FININDCOMMENT", "", true, true);
                doc.Replace("FININDRESULT", "Clear", true, true);
            }
            else
            {
                if (diligenceInput.otherdetails.Has_Reg_FINRA.ToString().Equals("Yes-With Adverse"))
                {
                    finracommentmodel = "According to records maintained by the United States Financial Industry Regulatory Authority (“FINRA”), [ShortEntityName] has been registered, with a Central Registration Depository (“CRD”) number of <number>, as a Brokerage firm (effective as of <date(s)>).  The subject entity is engaged in the following types of business: <Investigator to include types of business listed>.\n\nMoreover, the following Direct Owners and/or Executive Officers of the subject entity were identified:   <Investigator to add Direct Owner/Officer information>. <Investigator to add/indicate whether Indirect Direct Owner information exists>.\n\nFurther, the following customer disputes, disciplinary actions, and/or regulatory events were identified in connection with this registration:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                    doc.Replace("FININDCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("FININDRESULT", "Records", true, true);
                }
                else
                {
                    finracommentmodel = "According to records maintained by the United States Financial Industry Regulatory Authority (“FINRA”), [ShortEntityName] has been registered, with a Central Registration Depository (“CRD”) number of <number>, as a Brokerage firm (effective as of <date(s)>).  The subject entity is engaged in the following types of business: <Investigator to include types of business listed>.\n\nMoreover, the following Direct Owners and/or Executive Officers of the subject entity were identified:   <Investigator to add Direct Owner/Officer information>.  <Investigator to add/indicate whether Indirect Direct Owner information exists>.\n\nNo customer disputes, disciplinary actions, and/or regulatory events were identified in connection with this registration.";
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
            if (diligenceInput.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered"))
            {
                nfacommentmodel = "";
                doc.Replace("USNFCOMMENT", "", true, true);
                doc.Replace("USNFRESULT", "Clear", true, true);
            }
            else
            {
                if (diligenceInput.otherdetails.Has_Reg_US_NFA.ToString().Equals("Yes-With Adverse"))
                {
                    nfacommentmodel = "According to information maintained by the United States National Futures Association (“NFA”), [ShortEntityName], with the identification number <number>, has been registered as <Investigator to insert registration types and dates>.  Further, the following pool exemptions were connection with this registration:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\nAdditionally, the following regulatory actions, NFA Arbitration Awards and/or Commodity Futures Trading Commission (“CFTC”) Reparations Cases were reported in connection with the subject entity:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                    doc.Replace("USNFCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("USNFRESULT", "Records", true, true);
                }
                else
                {
                    doc.Replace("USNFCOMMENT", "", true, true);
                    doc.Replace("USNFRESULT", "Clear", true, true);
                    nfacommentmodel = "According to information maintained by the United States National Futures Association (“NFA”), [ShortEntityName], with the identification number <number>, has been registered as < Investigator to insert registration types and dates>.  Further, the following pool exemptions were connection with this registration:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\nNo regulatory actions, NFA Arbitration Awards or Commodity Futures Trading Commission (“CFTC”) Reparations Cases were reported in connection with the subject entity.";
                }
                nfacommentmodel = nfacommentmodel.Replace("*n ", "\n");
                nfacommentmodel = nfacommentmodel.Replace("*n", "\n");
                nfacommentmodel = nfacommentmodel.Replace("*t", "\t");
                //doc.Replace("US_NFAHEADER", "\n United States National Futures Association  \n", true, true);
                nfacommentmodel = string.Concat("\nUnited States National Futures Association\n\n", nfacommentmodel, "\n");
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
                    hksfccommentmodel = "[ShortEntityName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>), in the regulated activity of <Investigator to insert activities and dates>.\n\n<Investigator to summarize other license record information, prior regulated activities and dates, previous names, Responsible Officers / Representatives, conditions, etc.>\n\nThere are no public disciplinary actions on file against the subject within the past five years.";
                    doc.Replace("KONSECCOMMENT", "", true, true);
                    doc.Replace("KONSECRESULT", "Clear", true, true);
                }
                else
                {
                    hksfccommentmodel = "[ShortEntityName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>)), in the regulated activity of <Investigator to insert activities and dates>.\n\n<Investigator to summarize other license record information, prior regulated activities and dates, previous names, Responsible Officers / Representatives, conditions, etc.>\n\nAdditionally, the following public disciplinary actions were on file in connection with the subject entity for the past five years:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                    doc.Replace("KONSECCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("KONSECRESULT", "Records", true, true);
                }
                hksfccommentmodel = string.Concat("\nHong Kong Securities and Futures Commission\n\n", hksfccommentmodel, "\n");
            }

            //Holds Any License 
            string strotherheader = "";
            if (diligenceInput.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || diligenceInput.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || diligenceInput.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
            {
                regflag = "Records";
            }
            else
            {
                regflag = "Clear";

            }
            if (regflag == "Records")
            {
                holdslicensecommentmodel = "\nOther Professional Licensures and/or Designations\n\nInvestigative efforts did not reveal any additional licensure and/or registration information in connection with [ShortEntityName], however, further efforts would be required on a jurisdiction and license-type basis to confirm the same.  As a precaution, research efforts included, but were not limited to: a Brokerage Firm Check through the United States Financial Industry Regulatory Authority (for the past two years); a search of the United States Commodity Futures Trading Commission registration; an Investment Adviser Firm search through the United States Securities and Exchange Commission; a search through the United Kingdom’s Financial Conduct Authority; a search through the Hong Kong Securities and Futures Commission; as well as searches through the Bermuda Monetary Authority, the Cayman Islands Monetary Authority and the British Virgin Islands Financial Services Commission.  <Investigator to add or remove any sources that came back with hits above or are needed for this jurisdiction>\n";
                //holdslicensecommentmodel = "Investigative efforts did not reveal any additional professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction and license-type basis to confirm the same. ";
            }
            else
            {
                holdslicensecommentmodel = "\nOther Professional Licensures and/or Designations\n\nInvestigative efforts did not reveal any licensure and/or registration information in connection with [ShortEntityName], however, further efforts would be required on a jurisdiction and license-type basis to confirm the same.  As a precaution, research efforts included, but were not limited to: a Brokerage Firm Check through the United States Financial Industry Regulatory Authority (for the past two years); a search of the United States Commodity Futures Trading Commission registration and National Futures Association membership information; an Investment Adviser Firm search through the United States Securities and Exchange Commission; a search through United Kingdom’s Financial Conduct Authority; a search through the Hong Kong Securities and Futures Commission; as well as searches through the Bermuda Monetary Authority, the Cayman Islands Monetary Authority and the British Virgin Islands Financial Services Commission.  <Investigator to add or remove any sources that came back with hits above or are needed for this jurisdiction/scope/search type>\n";
            }                       

            if (diligenceInput.pllicenseModels[0].General_PL_License.ToString().ToUpper().Equals("YES")) { doc.Replace("[PLLICENSEDESC]", string.Concat("\n", usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false); }
            else
            {
                doc.Replace("[PLLICENSEDESC]", string.Concat(usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false);
            }
            doc.SaveToFile(savePath);
            if (diligenceInput.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered")) { }
            else
            {
                string blnredresultfou = "";
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
                                    if (abc.ToString().EndsWith("[ussecfootnote]"))
                                    {
                                        Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                        footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                        footnote1.MarkerCharacterFormat.FontSize = 11;
                                        //Insert footnote1 after the word "Spire.Doc"
                                        para.ChildObjects.Insert(i + 1, footnote1);
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("It is noted that SEC Registered Investment Advisors are required to file a Form ADV registration with the SEC on an annual basis, however, forms may be updated periodically to reflect certain firm structure changes.");
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        //Append the line
                                        string strredflagtextappended = "";
                                        if (diligenceInput.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse"))
                                        {
                                            strredflagtextappended = " filed with the United States Securities and Exchange Commission (“SEC”), [ShortEntityName]  has been registered, with a Central Registration Depository (“CRD”) number of <number>, as an Investment Adviser  since <date>, and manages <number> discretionary and/or non-discretionary accounts, totaling approximately $<amount from Item 5 F(2)> in assets under its management at the time this filing was captured on <filing date>.\n\nIn addition, [ShortEntityName] provided investment advisory services to <number> clients during their most recently-completed fiscal year, which consisted of <Investigator to list out client types from Item 5 C&D of the Form ADV>. Additionally, <percentage number>% of its clients are non-United States persons <Investigator to replace with ‘none of its clients…’ if percentage number is 0>.\n\nFurther, [ShortEntityName] is reportedly compensated for its investment advisory services by <Investigator to list out compensation methods checked in Item 5E>.\n\nThe Form ADV also revealed that as part of its services, the firm reportedly provides <Investigator to list out services provided in Item 5G of form>. Moreover, the firm does not sell products or provide services other than investment advice to their advisory clients <Investigator to remove or modify if untrue – see Item 6 B (3)>.\n\nAdditionally, the Form ADV also reports that [ShortEntityName] has approximately <number> employees, <number> of which <Investigator to summarize Item 5 A & B (1-6) in form>. Further, according to the filing, no “firms or other persons solicit advisory clients on [their] behalf,” and none of these employees are registered representatives of a broker-dealer. <Investigator to modify language as appropriate from the form>\n\nThe Form ADV also revealed the following Direct Owners and/or Executive Officers of the subject entity at the time of filing: <Investigator to add Direct Owner/Officer information from Schedule A>. <Investigator to add/indicate whether Indirect Direct Owner information exists in Schedule B>.\n\nFurther, the following disclosure events were reported on the subject entity’s SEC registration:\n\n\t•\t<Investigator to insert bulleted list of results from Schedule R - DRP Pages here>\n\t•\t<Investigator to insert bulleted list of results from Schedule R - DRP Pages here>\n\t•\t<Investigator to insert bulleted list of results from Schedule R - DRP Pages here>\n\nWith the exception of the above, no additional disclosure events, such as formal investigations, disciplinary actions, customer disputes and/or criminal charges or convictions, were on file with the SEC in connection with [ShortEntityName].";
                                        }
                                        else
                                        {
                                            strredflagtextappended = " filed with the United States Securities and Exchange Commission (“SEC”), [ShortEntityName] has been registered, with a Central Registration Depository (“CRD”) number of <number>, as an Investment Adviser since <date>, and manages <number> discretionary and/or non-discretionary accounts, totaling approximately $<amount from Item 5 F(2)> in assets under its management at the time this filing was captured on <filing date>.\n\nIn addition, [ShortEntityName] provided investment advisory services to <number> clients during their most recently-completed fiscal year, which consisted of <Investigator to list out client types from Item 5 C&D of the Form ADV>. Additionally, <percentage number>% of its clients are non-United States persons <Investigator to replace with ‘none of its clients…’ if percentage number is 0>.\n\nFurther, [ShortEntityName] is reportedly compensated for its investment advisory services by <Investigator to list out compensation methods checked in Item 5E>.\n\nThe Form ADV also revealed that as part of its services, the firm reportedly provides <Investigator to list out services provided in Item 5G of form>. Moreover, the firm does not sell products or provide services other than investment advice to their advisory clients <Investigator to remove or modify if untrue – see Item 6 B (3)>.\n\nAdditionally, the Form ADV also reports that [ShortEntityName] has approximately <number> employees, <number> of which <Investigator to summarize Item 5 A & B (1-6) in form>. Further, according to the filing, no “firms or other persons solicit advisory clients on [their] behalf,” and none of these employees are registered representatives of a broker-dealer. <Investigator to modify language as appropriate from the form>\n\nThe Form ADV also revealed the following Direct Owners and/or Executive Officers of the subject entity at the time of filing: <Investigator to add Direct Owner/Officer information from Schedule A>. <Investigator to add/indicate whether Indirect Direct Owner information exists in Schedule B>.\n\n<Investigator to add summary of other financial industry affiliate information, and/or listing of private funds in the Form ADV – as appropriate -- in Sections 7.A. and 7.B>\n\nNo disclosure events, such as formal investigations, disciplinary actions, customer disputes and/or criminal charges or convictions, were on file with the SEC in connection with the subject entity.";
                                        }
                                        //  strredflagtextappended = strredflagtextappended.Replace("[FullEntityName]", diligenceInput.diligenceInputModel.LastName);
                                        TextRange tr = para.AppendText(strredflagtextappended);
                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                        tr.CharacterFormat.FontSize = 11;
                                        doc.SaveToFile(savePath);
                                        doc.Replace("[ussecfootnote]", "", true, false);
                                        doc.SaveToFile(savePath);
                                        blnredresultfou = "true";
                                        break;
                                    }

                                }
                            }
                            if (blnredresultfou == "true")
                            {
                                break;
                            }
                        }
                        if (blnredresultfou == "true")
                        {
                            break;
                        }
                    }
                }
                catch { }
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

            if (diligenceInput.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered")) { }
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

            if (diligenceInput.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered")) { }
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

            if (diligenceInput.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered")) { }
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
            string holdresult = ""; try
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
                catch { }
            }
            //COMMONNAMESUBDESC
            if (diligenceInput.diligenceInputModel.CommonNameSubject == false)
            {
                doc.Replace("COMMONNAMESUBDESC", " Additionally, it is noted that searches of the United States District and Bankruptcy Courts were nearly national in scope.", true, true);
            }
            else
            {
                doc.Replace("COMMONNAMESUBDESC", "", true, true);
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
            regredflagommentmodel = regredflagommentmodel.Replace("[LastName]", "[ShortEntityName]");
            regredflagommentmodel = string.Concat("\n", regredflagommentmodel, "\n");
            doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", regredflagommentmodel, true, true);
            doc.SaveToFile(savePath);
            string blnredresultfound = "";
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
                                        strredflagtextappended = " and the following information was identified in connection with [ShortEntityName]:   <Investigator to insert results here>";
                                    }
                                    else
                                    {
                                        strredflagtextappended = " and it is noted that [ShortEntityName] was not identified in any of these records.";
                                    }
                                    //strredflagtextappended = strredflagtextappended.Replace("[ShortEntityName]", diligenceInput.diligenceInputModel.LastName);
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
            if (diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("United Kingdom"))
            {
                strcountry = "The United Kingdom";
            }
            else
            {
                strcountry = diligenceInput.diligenceInputModel.CurrentCountry.ToString();
            }
            if (diligenceInput.summarymodel.bankruptcy_filings.ToString().Equals("Results Pending") || diligenceInput.summarymodel.civil_court_Litigation.ToString().Equals("Results Pending"))
            {
                doc.Replace("[CountryHEADER]", string.Concat(strcountry, "\n\n<Search type> searches are currently ongoing through <source> in [Country], the results of which will be provided under separate cover upon receipt."), true, true);
            }
            else
            {
                if (diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("United Kingdom")) { }
                else
                {
                    doc.Replace("[CountryHEADER]", strcountry, true, true);
                }
            }
            doc.SaveToFile(savePath);
           
           
                    if (diligenceInput.summarymodel.bankruptcy_filings.ToString().StartsWith("Record") || diligenceInput.summarymodel.civil_court_Litigation.ToString().StartsWith("Record") )
                    {
                        strlegalrechitdesc = "Yes";
                    }
                    else
                    {
                        strlegalrechitdesc = "No";
                    }
                            
            if (diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("United Kingdom"))
            {
                doc.Replace("[CountryHEADER]", "", true, true);
            }
            else
            {
                switch (strlegalrechitdesc)
                {
                    case "Yes":
                        doc.Replace("[COURTDESC]", "\n\nSearches of all available legal records in [Country], which include the [Courts], identified the subject entity as a party to the following bankruptcy filings, civil litigation matters, liens and/or judgments:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n[COURTDESC]", true, false);
                        break;
                    case "No":
                        doc.Replace("[COURTDESC]", "\n\nSearches of all available legal records in [Country], which include the [Courts],  did not identify the subject entity as a party to any bankruptcy filings, civil litigation matters, liens or judgments.\n[COURTDESC]", true, false);
                        break;                   
                }
            }
            doc.SaveToFile(savePath);
            string legrechitcommon = "";

            if (diligenceInput.summarymodel.bankruptcy_filings1 == true || diligenceInput.summarymodel.civil_court_Litigation1 == true)
            {               
                legrechitcommon = string.Concat("\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <record type(i.e.civil litigation or bankruptcy)> filings, which were recorded in [Country] between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request -- if an expanded scope is warranted. <Investigator to separate out civil and bankruptcy cases -- if needed>\n");
            }    
                     
            CommentModel uslegcommentmodel1 = _context.DbComment
                            .Where(u => u.Comment_type == "U.S_Legal_Record")
                            .FirstOrDefault();

            if (diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("Australia") || diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("United Kingdom") || diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("Canada") || diligenceInput.diligenceInputModel.CurrentCountry.ToString().Equals("France"))
            {
                string strlegrechit = "";
                string strinsolvency = "";
                string strcivil = "";
                
                string strregtrust = "";
                string strppr = "";
                string strcourtdesc = "";
                switch (diligenceInput.diligenceInputModel.CurrentCountry.ToString())
                {
                    case "France":
                        if (diligenceInput.summarymodel.civil_court_Litigation1 == true)
                        {
                            legrechitcommon = string.Concat(legrechitcommon, "\nIt should be noted that one or more entities known only as [LastName] were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
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
                            civilpossible = string.Concat("\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
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
                        uk_ihcommentmodel = "Civil Litigation, Insolvency and Other Filings in Country[HEADER]\n\nRecords maintained by the Insolvency Service in the United Kingdom[INSERTFOOTNOTE]\n";
                        strinsolvency = uk_ihcommentmodel;
                        switch (diligenceInput.csModel.civil_record_hits.ToString())
                        {
                            case "Yes":
                                uk_crhcommentmodel = uk_crhcommentmodel1.confirmed_comment.ToString();
                                uk_crhcommentmodel = uk_crhcommentmodel.Replace("[LastName]", "[ShortEntityName]");
                                break;
                            case "No":
                                uk_crhcommentmodel = "Research efforts, including searches of records maintained by the Court of Appeal, Administrative Court and the High Court in the United Kingdom, did not reveal any litigation and/or judgments registered against [ShortEntityName].";
                                break;
                        }
                        uk_crhcommentmodel = uk_crhcommentmodel.Replace("*n ", "\n");
                        uk_crhcommentmodel = uk_crhcommentmodel.Replace("*n", "\n");
                        uk_crhcommentmodel = uk_crhcommentmodel.Replace("*t", "\t");
                        strcivil = string.Concat("\n", uk_crhcommentmodel, "\n");
                        switch (diligenceInput.csModel.reg_trust_hits.ToString())
                        {
                            case "N/A":
                                uk_rthcommentmodel = "Further, a search would be required through the Registry of Judgments, Orders and Fines within the past six years in connection with the subject entity.";
                                break;
                            case "Yes":
                                uk_rthcommentmodel = uk_rthcommentmodel1.confirmed_comment.ToString();
                                break;
                            case "No":
                                uk_rthcommentmodel = "Further, a search through the Registry of Judgments, Orders and Fines was condcuted for the subject entity within the past six years, and no such records were identified.";
                                break;
                        }
                        uk_rthcommentmodel = uk_rthcommentmodel.Replace("*n ", "\n");
                        uk_rthcommentmodel = uk_rthcommentmodel.Replace("*n", "\n");
                        uk_rthcommentmodel = uk_rthcommentmodel.Replace("*t", "\t");
                        strregtrust = string.Concat("\n", uk_rthcommentmodel, "\n", civilpossible);                                               
                        
                       
                       
                        strcourtdesc = string.Concat(legrechitcommon, strinsolvency, strcivil, strregtrust, "\n");
                        doc.Replace("[COURTDESC]", strcourtdesc, true, false);
                        doc.SaveToFile(savePath);                       
                        //Insolvency footnote
                       string blninvresultfound = "";
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
                                                    uk_ihcommentmodel = "identified the following bankruptcy filings in connection with [ShortEntityName]:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                                                }
                                                else
                                                {
                                                    uk_ihcommentmodel = "did not identify any bankruptcy filings in connection with [ShortEntityName].";
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
                            civilpossible = "\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <record type(i.e.civil litigation or bankruptcy)> filings, which were recorded in [Country] between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request -- if an expanded scope is warranted. <Investigator to separate out civil and bankruptcy cases -- if needed>\n";
                            //civilpossible = string.Concat("\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
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
                        strinsolvency = string.Concat("\n",can_Inhitcommentmodel, "\n");
                        strinsolvency = strinsolvency.Replace("[LastName]", "[ShortEntityName]");
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
                        strcivil = string.Concat("\n", can_civilcommentmodel, "\n", civilpossible);
                        strcivil = strcivil.Replace("[LastName]", "[ShortEntityName]");
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
                        can_pprcommentmodel = can_pprcommentmodel.Replace("[LastName]", "[ShortEntityName]");
                        strppr = string.Concat("\n", can_pprcommentmodel, "\n");
                        // strcourtdesc = string.Concat(legrechitcommon, strinsolvency, strcivil, strppr, "[MEDIABASEDLEGALDESCRIPTION]");
                        strcourtdesc = string.Concat(strinsolvency, strcivil, strppr, "[MEDIABASEDLEGALDESCRIPTION]");
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
                            civilpossible = string.Concat("\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> civil litigation filings, which were recorded in [Country]  between <date> and <date>, and manual court records retrieval efforts -- if available -- would be required in order to determine whether any of the same pertain to the subject of interest, which can be undertaken upon request if an expanded scope is warranted.\n");
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
                        
                        strinsolvency = "\nA search of the Australian National Bankruptcy Registrar\n";
                        strcivil = string.Concat("\nAdditionally, searches conducted on all available legal records in Australia, including the Federal Court, District Courts, County Courts, Magistrate Courts and Tribunals\n", civilpossible);
                        doc.SaveToFile(savePath);
                        string blnausinvresultfound = "";
                                                
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
                        aus_pprcommentmodel = aus_pprcommentmodel.Replace("[LastName]", "[ShortEntityName]");
                        strppr = string.Concat("\n", aus_pprcommentmodel, "\n");
                        strcourtdesc = string.Concat(legrechitcommon, "\n", strinsolvency, strcivil, strppr, "[MEDIABASEDLEGALDESCRIPTION]");
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
                    mediacommentmodel = "Further, efforts undertaken through official government announcements, published news stories, regulatory agency records and other available sources, identified the following references to bankruptcy filings, civil litigation filings and/or other legal records involving the subject in [Country]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                }
                else
                {
                    mediacommentmodel = "Further, efforts undertaken through official government announcements, published news stories, regulatory agency records and other available sources, did not locate any references to any bankruptcy filings, civil litigation filings or other legal records involving the subject in [Country].";
                }               
                mediacommentmodel = string.Concat("\n", mediacommentmodel, "\n");
                //U.S_Legal_Record            
                if (diligenceInput.otherdetails.USLegal_Record_Hits.ToString().Equals("Yes"))
                {
                    uslegcommentmodel = "In addition, as a precaution bankruptcy and civil court records were searched in the United States covering a period of at least 10 years or so through the federal-level United States District and Bankruptcy Courts at a nearly national level, which identified the subject entity in connection with the following records:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\nIt is noted that comprehensive statewide and local-level court research efforts would be required in connection with the subject entity in the United States, should an expanded scope be warranted.";                    
                    uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                }
                else
                {
                    if (diligenceInput.otherdetails.USLegal_Record_Hits.ToString().Equals("No"))
                    {
                        uslegcommentmodel = "In addition, as a precaution bankruptcy and civil court records were searched in the United States covering a period of at least 10 years or so through the federal-level United States District and Bankruptcy Courts at a nearly national level, which did not identify the subject entity in connection with any such records.\n\nIt is noted that comprehensive statewide and local-level court research efforts would be required in connection with the subject entity in the United States, should an expanded scope be warranted.";
                        uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                    }
                    else
                    {
                        uslegcommentmodel = "In addition to the above, given the commonality of [ShortEntityName]’s name, in order to undertake precautionary searches for bankruptcy and civil court records in the United States federal court system, additional information, such as any jurisdictions the subject entity may have ties to, would be required to complete the same, which can be conducted upon request --  if an expanded scope is warranted.\n\nIt is noted that comprehensive statewide and local-level court research efforts would be required in connection with the subject entity in the United States, should an expanded scope be warranted.";
                        uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                    }
                }              
                strcourtdesc = string.Concat(mediacommentmodel, uslegcommentmodel);
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
                                        if (abc.ToString().Equals("Civil Litigation, Insolvency and Other Filings in Country[HEADER]"))
                                        {
                                            textRange.CharacterFormat.Italic = true;
                                            textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                                            doc.Replace("Country[HEADER]", diligenceInput.diligenceInputModel.CurrentCountry.ToString(), true, true);
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
                    mediacommentmodel = "Further, efforts undertaken through official government announcements, published news stories, regulatory agency records and other available sources, identified the following references to bankruptcy filings, civil litigation filings and/or other legal records involving the subject in [Country]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                }
                else
                {
                    mediacommentmodel = "Further, efforts undertaken through official government announcements, published news stories, regulatory agency records and other available sources, did not locate any references to any bankruptcy filings, civil litigation filings or other legal records involving the subject in [Country].";
                }
                mediacommentmodel = mediacommentmodel.Replace("*n ", "\n");
                mediacommentmodel = mediacommentmodel.Replace("*n", "\n");
                mediacommentmodel = mediacommentmodel.Replace("*t", "\t");
                mediacommentmodel = string.Concat("\n", mediacommentmodel, "\n");
                //U.S_Legal_Record                
                if (diligenceInput.otherdetails.USLegal_Record_Hits.ToString().Equals("Yes"))
                {
                    uslegcommentmodel = "In addition, as a precaution bankruptcy and civil court records were searched in the United States covering a period of at least 10 years or so through the federal-level United States District and Bankruptcy Courts at a nearly national level, which identified the subject entity in connection with the following records:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\nIt is noted that comprehensive statewide and local-level court research efforts would be required in connection with the subject entity in the United States, should an expanded scope be warranted.";                    
                    uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                }
                else
                {
                    if (diligenceInput.otherdetails.USLegal_Record_Hits.ToString().Equals("No"))
                    {
                        uslegcommentmodel = "In addition, as a precaution bankruptcy and civil court records were searched in the United States covering a period of at least 10 years or so through the federal-level United States District and Bankruptcy Courts at a nearly national level, which did not identify the subject entity in connection with any such records.\n\nIt is noted that comprehensive statewide and local-level court research efforts would be required in connection with the subject entity in the United States, should an expanded scope be warranted.";
                        uslegcommentmodel = string.Concat("\nOther Legal Records Searches\n\n", uslegcommentmodel, "\n");
                    }
                    else
                    {
                        uslegcommentmodel = "In addition to the above, given the commonality of [ShortEntityName]’s name, in order to undertake precautionary searches for bankruptcy and civil court records in the United States federal court system, additional information, such as any jurisdictions the subject entity may have ties to, would be required to complete the same, which can be conducted upon request --  if an expanded scope is warranted.\n\nIt is noted that comprehensive statewide and local-level court research efforts would be required in connection with the subject entity in the United States, should an expanded scope be warranted.";
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
            switch (diligenceInput.diligenceInputModel.CurrentCountry.ToString())
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
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches],  which identified the following information in connection with [ShortEntityName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [ShortEntityName].";
                    }
                    strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "Monetary Authority of Singapore. It should be noted that the Monetary Authority of Singapore (“MAS”) regulates companies, and not entities, and while entities representing investment and fund management firms should have a Capital Market License or Financial Advice License, for which they have to pass an examination to get a certificate, the onus is on the companies that they represent to ensure that they have the necessary qualifications and credentials.  Further, individual fund managers and investment advisors are not regulated by the MAS. In addition, the subject’s MAS Representative Number -- if any -- would be required in order to confirm if the subject is registered with the MAS, and if any disciplinary actions were filed in connection with this registration, as well as to confirm if the subject has met MAS’s “fit and proper” criteria");
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                    doc.Replace("[Courts]", "the Insolvency & Public Trustee’s Office, the High Court, District Court, Magistrate’s Court and the Subordinate Courts of Singapore", true, true);
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches],  which identified the following information in connection with [ShortEntityName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [ShortEntityName].";
                    }
                    strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", strregulatortsearch);
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                    doc.Replace("[Courts]", "the Federal Court of Justice (Bundesgerichtshof), the Federal Labor Court (Bundesarbeitsgericht), the Federal Administrative Court (Bundesverwaltungsgericht), and the Federal Finance Court (Bundessozialgericht), <investigator to add other relevant courts as applicable>", true, true);
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches],  which identified the following information in connection with [ShortEntityName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [ShortEntityName].";
                    }
                    strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "It is noted that the Swiss Financial Market Supervisory Authority (“FINMA”) regulates certain companies, and not entities, and in this regard, the subject is not registered with FINMA");
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                    doc.Replace("[Courts]", "the Swiss Official Gazette of Commercial (“SOGC”), the Federal Supreme Court of Switzerland, the Federal Criminal, Patent and Administrative Courts of Switzerland, <investigator to add other relevant courts as applicable>", true, true);
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
                    doc.SaveToFile(savePath);
                    
                    break;
                case "India":
                    string strcorpreg = "";
                    if (diligenceInput.csModel.india_corpregistry.ToString().Equals("Has Business Affiliations"))
                    {
                        strcorpreg = "In addition to the above, while records maintained on entities by the Indian Registrar of Companies and the Ministry of Corporate Affairs are not available to third-party companies, research efforts identified [ShortEntityName] as an Officer, Director and/or Shareholder of the following business entities in [Country]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strcorpreg = "It should be noted that records maintained on entities by the Indian Registrar of Companies and the Ministry of Corporate Affairs are not available to third-party companies.";
                    }
                    doc.Replace("[Corp Registry]", strcorpreg, true, true);
                    doc.Replace("[Property Registry]", "Third-party access to property records is restricted in India.", true, true);
                    strothers_SpeRegHits = "";
                    if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches],  which identified the following information in connection with [ShortEntityName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [ShortEntityName].";
                    }
                    strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "Securities Exchange Board of India (“SEBI”)");
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                    doc.Replace("[Courts]", "the Local Police, Taluka/Small Causes Courts, the District/Session Courts, as well as the High Court and Supreme Court of India", true, true);
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
                    doc.SaveToFile(savePath);
                    break;
                case "United Arab Emirates":
                    doc.Replace("[Corp Registry]", "<Local Emirate> Chamber of Commerce ", true, true);
                    doc.Replace("[Property Registry]", "<Local Emirate> Land Department ", true, true);
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                    doc.Replace("[Courts]", "<Local Emirate> Ministry of Interior, the <Local Emirate> Court of First Instance,  searches of media and other information from sources in judicial circles in the U.A.E., and the U.A.E. Central Bank", true, true);
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", "", true, true);
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
                   
                    CommentModel fr_lrhcommentmodel1 = _context.DbComment
                           .Where(u => u.Comment_type == "fr_legalrechit")
                           .FirstOrDefault();
                    if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches],  which identified the following information in connection with [ShortEntityName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [ShortEntityName].";
                    }
                    strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "<insert Country Regulatory Body or delete section if not applicable>");
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
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
                    doc.Replace("[Corp Registry] did not identify", "Companies House did not identify", false, false);
                    doc.Replace("[Corp Registry] identified", "Companies House identified", false, false);
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
                    if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches],  which identified the following information in connection with [ShortEntityName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [ShortEntityName].";
                    }
                    strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "<insert Country Regulatory Body or delete section if not applicable>");
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
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
                    if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches],  which identified the following information in connection with [ShortEntityName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [ShortEntityName].";
                    }
                    strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "<insert Country Regulatory Body or delete section if not applicable>");
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                    doc.Replace("[Courts]", "<investigator to add relevant courts>", true, true);
                    doc.SaveToFile(savePath);
                    break;
                default:
                    doc.Replace("[Corp Registry]", "<Investigator to add source for corp records>", true, true);
                    doc.Replace("[Property Registry]", "<Investigator to add source for property records>", true, true);
                    strothers_SpeRegHits = "";
                    if (diligenceInput.csModel.country_specific_reg_hits.ToString().Equals("Yes"))
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches],  which identified the following information in connection with [ShortEntityName]:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strothers_SpeRegHits = "\nAdditionally, research efforts were conducted in [Country] through records maintained by the [Regulatory Searches], which did not identify any information in connection with [ShortEntityName].";
                    }
                    strothers_SpeRegHits = strothers_SpeRegHits.Replace("[Regulatory Searches]", "<insert Country Regulatory Body or delete section if not applicable>");
                    doc.Replace("COUNTRYSPECIFICREGHITDISCRIPTION", strothers_SpeRegHits, true, true);
                    doc.SaveToFile(savePath);
                    doc.Replace("[Courts]", "<insert relevant courts>", true, true);
                    doc.Replace("[FirstName]",diligenceInput.diligenceInputModel.FullaliasName, true, true);
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
            if (diligenceInput.summarymodel.bankruptcy_filings.StartsWith("Record") || diligenceInput.summarymodel.civil_court_Litigation.StartsWith("Record"))
            {
                Legrechitcommentmodel = Legrechitcommentmodel1.confirmed_comment.ToString();
                strexecutive_sumlegrec = "Yes";
            }
            else
            {
                Legrechitcommentmodel = Legrechitcommentmodel1.unconfirmed_comment.ToString();
                strexecutive_sumlegrec = "No";
            }
            Legrechitcommentmodel = Legrechitcommentmodel.Replace("criminal charges or ", "");
            Legrechitcommentmodel = Legrechitcommentmodel.Replace(", personally", "");
            Legrechitcommentmodel = Legrechitcommentmodel.Replace("subject", "subject entity");
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
            hasreghitcommentmodel= hasreghitcommentmodel.Replace("subject", "subject entity");
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
                    searchtext = "In sum, no issues of potential relevance were identified in connection with";
                }
                else
                {
                    // hitcompcommentmodel = nohitcompcommentmodel1.unconfirmed_comment.ToString();
                    hitcompcommentmodel = "In sum, no issues of potential relevance were identified in connection with [FirstName] [MiddleInitial] [LastName] in [Country].";
                    searchtext = "In sum, no issues of potential relevance were identified in connection with";
                }

            }
            hitcompcommentmodel = hitcompcommentmodel.Replace("[FirstName] [MiddleInitial] [LastName]", diligenceInput.diligenceInputModel.LastName.ToString());
            //hitcompcommentmodel = hitcompcommentmodel.Replace("[MiddleInitial]", diligenceInput.diligenceInputModel.MiddleInitial.ToString());
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
                                        strglobaltextappended = " and the following information was identified in connection with [ShortEntityName]: ";
                                    }
                                    else
                                    {
                                        strglobaltextappended = " and it is noted that the subject entity was not identified on any of these lists.";
                                    }
                                    //strglobaltextappended = strglobaltextappended.Replace("[ShortEntityName]", diligenceInput.diligenceInputModel.FullaliasName);
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
            try
            {
                doc.Replace("[IncorporationYear]", Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Year.ToString(), true, false);
            }
            catch
            {
                doc.Replace("[IncorporationYear]", "<not provided>", true, false);
            }
            if (diligenceInput.diligenceInputModel.FullaliasName.ToString().Equals(""))
            {
                doc.Replace("[ShortEntityName]", "<not provided>", true, false);
            }
            else
            {
                doc.Replace("[ShortEntityName]", diligenceInput.diligenceInputModel.FullaliasName.ToString(), true, false);
            }
            doc.SaveToFile(savePath);
            doc.Replace("[CaseNumber]", diligenceInput.diligenceInputModel.CaseNumber.TrimEnd(), true, true);
            doc.SaveToFile(savePath);
            doc.Replace("[Country]", diligenceInput.diligenceInputModel.CurrentCountry.ToString().TrimEnd(), true, true);
            doc.Replace("[EntityCountry]", diligenceInput.diligenceInputModel.CurrentCountry.ToString().TrimEnd(), true, true);
            doc.Replace("[MiddleInitial] ","", true, false);
            try
            {
                doc.Replace("[IncorporationDate]", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Month) + " " + Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Day + ", " + Convert.ToDateTime(diligenceInput.diligenceInputModel.Dob.ToString()).Year, true, true);
            }
            catch
            {
                doc.Replace("[IncorporationDate]", "<not provided>", true, true);
            }
            doc.SaveToFile(savePath);
            switch (regflag.ToString())
            {
                case "Clear":
                    doc.Replace("REGCOMMENT", "", true, true);
                    doc.Replace("REGRESULT", "Clear", true, true);
                    break;
                case "Records":                    
                    doc.Replace("REGCOMMENT", "<investigator to insert regulatory hits here>", true, true);
                    doc.Replace("REGRESULT", "Records", true, true);
                    break;
            }
            doc.SaveToFile(savePath);
            HttpContext.Session.SetString("first_name",diligenceInput.diligenceInputModel.FullaliasName);
            HttpContext.Session.SetString("last_name", diligenceInput.diligenceInputModel.LastName);
            HttpContext.Session.SetString("country", diligenceInput.diligenceInputModel.CurrentCountry);
            HttpContext.Session.SetString("case_number", diligenceInput.diligenceInputModel.CaseNumber);
            //HttpContext.Session.SetString("middleinitial", diligenceInput.diligenceInputModel.MiddleInitial);
           // HttpContext.Session.SetString("city", diligenceInput.diligenceInputModel.City);
            //HttpContext.Session.SetString("regflag", regflag);
            //HttpContext.Session.SetString("employer1", diligenceInput.EmployerModel[0].Emp_Employer.ToString());
            //HttpContext.Session.SetString("emp_location1", diligenceInput.EmployerModel[0].Emp_Location.ToString());
           // HttpContext.Session.SetString("employer1State", diligenceInput.diligenceInputModel.Employer1State.ToString());
           
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
        public IActionResult US_DD_Individual_page2()
        {
            return View();
        }
        public IActionResult Save_Page2(SummaryResulttableModel ST)
        {
            string last_name = HttpContext.Session.GetString("last_name");
            string first_name = HttpContext.Session.GetString("first_name");
            string case_number = HttpContext.Session.GetString("case_number");
            string country = HttpContext.Session.GetString("country");
            string savePath = _config.GetValue<string>("ReportPath");
            savePath = string.Concat(savePath, last_name.ToString(), "_SterlingDiligenceReport(", case_number.ToString(), ")_DRAFT.docx");
            Document doc = new Document(savePath);
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
            switch (ST.news_media_searches.ToString())
            {
                case "Clear":
                    doc.Replace("[NewsRES]", "", true, false);
                    doc.Replace("NEWSCOMMENT", "No adverse or materially-significant information was identified ", true, true);
                    doc.Replace("NEWSRES", "Clear", true, true);
                    break;
                case "Records":
                    doc.Replace("NEWSCOMMENT", "<investigator to insert press summaries here>", true, true);
                    doc.Replace("NEWSRES", "Records", true, true);
                    doc.Replace("[NewsRES]", "", true, false);
                    break;
                case "Potentially-Relevant Information":
                    doc.Replace("NEWSCOMMENT", "<investigator to insert press summaries here>", true, true);
                    doc.Replace("NEWSRES", "", true, true);
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
                  //  bankrupcomment = bankrupcomment.Replace("LastName", last_name);
                    bankrupresult = "Records";
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
            if (ST.bankruptcy_filings1 == true)
            {
                bankrupresult = string.Concat(bankrupresult, "\n\nPossible Records");
                bankrupcomment = string.Concat(bankrupcomment, "\n\nOne or more entities known only as “[FirstName] [LastName]” were identified as Petitioners to at least <number> bankruptcy filings in [Country], which were recorded between <date> and <date>, and are currently <status>");
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
                   // civilcourtcomment = civilcourtcomment.Replace("LastName", last_name);
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
                civilcourtcomment = string.Concat(civilcourtcomment, "\n\nOne or more entities known only as “[FirstName] [LastName]” were identified as <Party Type> in at least <number> civil litigation filings in [Country], which were recorded between <date> and <date>, and are <status>");
            }
            doc.Replace("CIVILCOMMENT", civilcourtcomment, true, true);
            doc.Replace("CIVILRESULT", civilcourtResult, true, true);
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
            //if (employer1State.ToString().Equals(""))
            //{
            //    doc.Replace("[Employer1State], ", "", true, true);
            //}
            //else
            //{
            //    doc.Replace("[Employer1State]", employer1State, true, true);
            //}

            doc.Replace("the United Kingdom’s Financial Conduct Authority", "United Kingdom’s Financial Conduct Authority", true, true);
            doc.Replace("< investigator to remove if inaccurate >", "<investigator to remove if inaccurate>", false, true);
            doc.Replace("<investigator to insert regulatory hits here >", "<investigator to insert regulatory hits here>", true, true);
            //doc.Replace("[FullEntityName]", last_name.ToString(), false, false);
            doc.Replace("[FirstName] [LastName]", "[ShortEntityName]", false, false);
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
            TextSelection[] tex1 = doc.FindAllString("•	<Affiliate Name>", false, false);
            if (tex1 != null)
            {
                foreach (TextSelection seletion in tex1)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] tex12 = doc.FindAllString("•	<Investigator to insert bulleted list of results from Schedule R - DRP Pages here>", false, false);
            if (tex12 != null)
            {
                foreach (TextSelection seletion in tex12)
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
            TextSelection[] tex33 = doc.FindAllString("•	<Affiliate Name>", false, false);
            if (tex33 != null)
            {
                foreach (TextSelection seletion in tex33)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] tex22 = doc.FindAllString("•	<Investigator to insert bulleted list of results from Schedule R - DRP Pages here>", false, false);
            if (tex22 != null)
            {
                foreach (TextSelection seletion in tex22)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] textSelections11 = doc.FindAllString("•	", false, false);
            if (textSelections11 != null)
            {
                foreach (TextSelection seletion in textSelections11)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.White;
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
          
            doc.Replace("[IncorporationState]", country, true, true);
                               
            doc.SaveToFile(savePath);
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
            doc.Replace("investigator", "Investigator", true, true);
            doc.Replace("  ", " ", true, false);
            doc.SaveToFile(savePath);
            doc.Replace(" ,", ",", true, false);
            doc.SaveToFile(savePath);
            doc.Replace("  (", " (", true, false);
            doc.Replace("subject entity", "subject", true, false);
            doc.Replace("subject", "subject entity", true, false);
            doc.Replace("Result Pending", "Results Pending", true, false);
            doc.Replace("united states", "United States", true, false);
            doc.Replace("investigator", "Investigator", true, false);
            doc.SaveToFile(savePath);
            TextSelection[] textresult1 = doc.FindAllString("Results Pending", false, false);
            if (textresult1 != null)
            {
                foreach (TextSelection seletion in textresult1)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                }
            }
            TextSelection[] textresult2 = doc.FindAllString("Possible Records", false, false);
            if (textresult2 != null)
            {
                foreach (TextSelection seletion in textresult2)
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
            doc.Replace("[LastName]", last_name, true, true);            
            doc.Replace("[FullEntityName]", last_name.ToString(), false, false);
            doc.Replace("[Last Name]", last_name.ToString(), true, true);
            doc.Replace("[Full_Entity_Name]", last_name.TrimEnd().ToUpper().ToString(), true, true);
            doc.Replace("[FullEntityName]", last_name.ToString(), false, false);
            doc.Replace("FirstName LastName", last_name.ToString(), false, false);
            doc.Replace("LastName", last_name, false, false);
            doc.Replace("the District of Columbia", "District of Columbia", false, false);
            doc.SaveToFile(savePath);
            doc.Replace("District of Columbia", "the District of Columbia", false, false);
            doc.SaveToFile(savePath);
            doc.Replace("U.S.  ", "U.S. ", true, false);
            doc.Replace("U.K.  ", "U.K. ", true, false);
            doc.Replace(".  (“Client”)", ". (“Client”)", false, false);
            doc.SaveToFile(savePath);
            return RedirectToAction("GenerateFile", "Diligence");
        }
    }
}