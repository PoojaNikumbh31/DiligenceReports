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
    public class USCOMPANYController : Controller
    {
        private readonly DataBaseContext _context;
        private readonly IConfiguration _config;
        public USCOMPANYController(IConfiguration config, Data.DataBaseContext context)
        {
            _config = config;
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult US_COMPANY_page1()
        {
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
        public IActionResult US_COMPANY_page1(MainModel mainModel, string SaveData, string Submit)
        {
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
                Additional_statesModel additional = new Additional_statesModel();               
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
                if (TempData["message"] == null)
                {
                    main.diligenceInputModel = dinput;                  
                    main.otherdetails = otherdatails;                  
                    main.pllicenseModels = _context.DbPLLicense
                               .Where(u => u.record_Id == recordid)
                                .ToList();
                    main.additional_States = _context.Dbadditionalstates
                               .Where(u => u.record_Id == recordid)
                                .ToList();
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
                            DI1.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                            DI1.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;                                                                                    
                            DI1.Employer1State = mainModel.diligenceInputModel.Employer1State;
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
                            strupdate.HasBankruptcyRecHits = mainModel.otherdetails.HasBankruptcyRecHits;
                            strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;                            
                            strupdate.Has_Civil_Records = mainModel.otherdetails.Has_Civil_Records;
                            strupdate.Has_Name_Only = mainModel.otherdetails.Has_Name_Only;
                            strupdate.Has_US_Tax_Court_Hit = mainModel.otherdetails.Has_US_Tax_Court_Hit;
                            strupdate.Has_Tax_Liens = mainModel.otherdetails.Has_Tax_Liens;
                            strupdate.Has_Name_Only_Tax_Lien = mainModel.otherdetails.Has_Name_Only_Tax_Lien;                                                                                
                            strupdate.Press_Media = mainModel.otherdetails.Press_Media;
                            strupdate.RegulatoryFlag = mainModel.otherdetails.RegulatoryFlag;
                            strupdate.Has_Reg_FINRA = mainModel.otherdetails.Has_Reg_FINRA;
                            strupdate.Has_Reg_UK_FCA = mainModel.otherdetails.Has_Reg_UK_FCA;
                            strupdate.Has_Reg_US_NFA = mainModel.otherdetails.Has_Reg_US_NFA;
                            strupdate.Has_Reg_US_SEC = mainModel.otherdetails.Has_Reg_US_SEC;
                            strupdate.ICIJ_Hits = mainModel.otherdetails.ICIJ_Hits;
                            strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                            strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                            strupdate.Has_Companion_Report = mainModel.otherdetails.Has_Companion_Report;
                            strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
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
                            resulttableModel.personal_Identification = mainModel.summarymodel.personal_Identification;
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
                        if (SaveData.Contains("State"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            Additional_statesModel additional_States = _context.Dbadditionalstates
                                      .Where(a => a.id == id_c).FirstOrDefault();

                            _context.Dbadditionalstates.Remove(additional_States);
                            _context.SaveChanges();

                            mainModel.additional_States = _context.Dbadditionalstates
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
            try
            {
                DiligenceInputModel dinput = _context.DbPersonalInfo
                                   .Where(u => u.record_Id == recordid)
                                   .FirstOrDefault();
                if (dinput == null)
                { }
                else
                {
                    ViewBag.CurrentState = dinput.CurrentState;
                    ViewBag.Employer1State = dinput.Employer1State;
                }

            }
            catch
            {
            }
            return View(mainModel);
        }
        [HttpGet]
        public IActionResult US_COMPANY_Edit()
        {           
            string recordid = HttpContext.Session.GetString("recordid");
            MainModel main = new MainModel();
            DiligenceInputModel diligence = _context.DbPersonalInfo
                .Where(a => a.record_Id == recordid)
                .FirstOrDefault();
            Otherdatails otherdatails = _context.othersModel
                                  .Where(u => u.record_Id == recordid)
                                  .FirstOrDefault();
            main.additional_States = _context.Dbadditionalstates
                        .Where(u => u.record_Id == recordid)
                        .ToList();          
            main.pllicenseModels = _context.DbPLLicense
                       .Where(u => u.record_Id == recordid)
                        .ToList();
            SummaryResulttableModel summary = _context.summaryResulttableModels
                 .Where(u => u.record_Id == recordid)
                                   .FirstOrDefault();
            main.summarymodel = summary;
            main.diligenceInputModel = diligence;
            main.otherdetails = otherdatails;
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
        public IActionResult US_COMPANY_Edit(MainModel mainModel, string SaveData, string Submit)
        {
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
                Additional_statesModel additional = new Additional_statesModel();
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
                if (TempData["message"] == null)
                {
                    main.diligenceInputModel = dinput;
                    main.otherdetails = otherdatails;
                    main.pllicenseModels = _context.DbPLLicense
                               .Where(u => u.record_Id == recordid)
                                .ToList();
                    main.additional_States = _context.Dbadditionalstates
                               .Where(u => u.record_Id == recordid)
                                .ToList();
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
                            DI1.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                            DI1.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                            DI1.Employer1State = mainModel.diligenceInputModel.Employer1State;
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
                            strupdate.HasBankruptcyRecHits = mainModel.otherdetails.HasBankruptcyRecHits;
                            strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;
                            strupdate.Has_Civil_Records = mainModel.otherdetails.Has_Civil_Records;
                            strupdate.Has_Name_Only = mainModel.otherdetails.Has_Name_Only;
                            strupdate.Has_US_Tax_Court_Hit = mainModel.otherdetails.Has_US_Tax_Court_Hit;
                            strupdate.Has_Tax_Liens = mainModel.otherdetails.Has_Tax_Liens;
                            strupdate.Has_Name_Only_Tax_Lien = mainModel.otherdetails.Has_Name_Only_Tax_Lien;
                            strupdate.Press_Media = mainModel.otherdetails.Press_Media;
                            strupdate.RegulatoryFlag = mainModel.otherdetails.RegulatoryFlag;
                            strupdate.Has_Reg_FINRA = mainModel.otherdetails.Has_Reg_FINRA;
                            strupdate.Has_Reg_UK_FCA = mainModel.otherdetails.Has_Reg_UK_FCA;
                            strupdate.Has_Reg_US_NFA = mainModel.otherdetails.Has_Reg_US_NFA;
                            strupdate.Has_Reg_US_SEC = mainModel.otherdetails.Has_Reg_US_SEC;
                            strupdate.ICIJ_Hits = mainModel.otherdetails.ICIJ_Hits;
                            strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                            strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                            strupdate.Has_Companion_Report = mainModel.otherdetails.Has_Companion_Report;
                            strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
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
                            resulttableModel.personal_Identification = mainModel.summarymodel.personal_Identification;
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
                        if (SaveData.Contains("State"))
                        {
                            string[] str = new string[3];
                            str = SaveData.Split("_");
                            int id_c = Convert.ToInt32(str[2]);

                            Additional_statesModel additional_States = _context.Dbadditionalstates
                                      .Where(a => a.id == id_c).FirstOrDefault();

                            _context.Dbadditionalstates.Remove(additional_States);
                            _context.SaveChanges();

                            mainModel.additional_States = _context.Dbadditionalstates
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
            try
            {
                DiligenceInputModel dinput = _context.DbPersonalInfo
                                   .Where(u => u.record_Id == recordid)
                                   .FirstOrDefault();
                if (dinput == null)
                { }
                else
                {
                    ViewBag.CurrentState = dinput.CurrentState;
                    ViewBag.Employer1State = dinput.Employer1State;
                }

            }
            catch
            {
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
        public IActionResult Save_Page1(MainModel uSDDIndividual)
        {
            try
            {
                string recordid = HttpContext.Session.GetString("recordid");
                string templatePath;
                string savePath = _config.GetValue<string>("ReportPath");
                templatePath = _config.GetValue<string>("USCompanytemplatePath");
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
                if (uSDDIndividual.diligenceInputModel.FullaliasName.ToString().Equals("") || uSDDIndividual.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("NA") || uSDDIndividual.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("N/A"))
                {
                    doc.Replace("[ShortEntityName]", "<not provided>", true, false);
                }
                else
                {                                                            
                    doc.Replace("[ShortEntityName]", uSDDIndividual.diligenceInputModel.FullaliasName.TrimEnd(), true, false);
                }               
                doc.Replace("[Full_Entity_Name]", uSDDIndividual.diligenceInputModel.LastName.TrimEnd().ToUpper().ToString(), true, true);
                
                try
                {
                    doc.Replace("ClientName", uSDDIndividual.diligenceInputModel.ClientName.TrimEnd(), true, false);
                }
                catch
                {
                    doc.Replace("ClientName", "<not provided>", true, false);
                }
              try
                {
                    doc.Replace("ClientName", uSDDIndividual.diligenceInputModel.ClientName.TrimEnd(), true, false);
                }
                catch
                {
                    doc.Replace("ClientName", "<not provided>", true, false);
                }
                try
                {
                    doc.Replace("[IncorporationDate]", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Month) + " "+ Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Day + ", "+ Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Year, true, true);
                }
                catch
                {
                    doc.Replace("[IncorporationDate]", "<not provided>", true, true);
                }
                doc.Replace("[CaseNumber]", uSDDIndividual.diligenceInputModel.CaseNumber.TrimEnd(), true, true);
                string current_full_address = "";
                string currentstreet;
                string currentcity = "";
                string currentstate = "";
                if (uSDDIndividual.diligenceInputModel.CurrentStreet.ToString().Equals(""))
                {
                    currentstreet = "";
                }
                else
                {
                    if (uSDDIndividual.diligenceInputModel.CurrentCity.ToString().Equals("") && uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State") && uSDDIndividual.diligenceInputModel.CurrentZipcode.ToString().Equals(""))
                    {
                        currentstreet = string.Concat(uSDDIndividual.diligenceInputModel.CurrentStreet.ToString().TrimEnd());
                    }
                    else
                    {
                        currentstreet = string.Concat(uSDDIndividual.diligenceInputModel.CurrentStreet.ToString().TrimEnd(), ", ");
                    }

                }
                if (uSDDIndividual.diligenceInputModel.CurrentCity.ToString().Equals(""))
                {
                    currentcity = "";
                    doc.Replace("[EntityCity], ", "", true, false);
                }
                else
                {
                    doc.Replace("[EntityCity]", uSDDIndividual.diligenceInputModel.CurrentCity.ToString(), true, false);
                    if (uSDDIndividual.diligenceInputModel.CurrentZipcode.ToString().Equals("") && uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                    {
                        currentcity = uSDDIndividual.diligenceInputModel.CurrentCity.ToString().TrimEnd();
                    }
                    else
                    {
                        currentcity = string.Concat(uSDDIndividual.diligenceInputModel.CurrentCity.ToString().TrimEnd(), ", ");
                    }
                }
                if (uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                {
                    doc.Replace(", [EntityState]", "", true, false);
                    doc.Replace("[EntityState]", "", true, false);
                    if (uSDDIndividual.diligenceInputModel.CurrentZipcode.ToString().Equals(""))
                    {
                        currentstate = "";
                    }
                    else
                    {
                        currentstate = uSDDIndividual.diligenceInputModel.CurrentZipcode.ToString().TrimEnd();
                    }
                }
                else
                {
                    doc.Replace("[EntityState]", uSDDIndividual.diligenceInputModel.CurrentState.ToString(), true, false);
                    try
                    {
                        StateWiseFootnoteModel statecomment1 = _context.stateModel
                                         .Where(u => u.states.ToUpper().TrimEnd() == uSDDIndividual.diligenceInputModel.CurrentState.ToString().ToUpper())
                                         .FirstOrDefault();
                        currentstate = statecomment1.abbreviation;
                    }
                    catch
                    {
                        currentstate = uSDDIndividual.diligenceInputModel.CurrentState.ToString();
                    }
                    if (uSDDIndividual.diligenceInputModel.CurrentZipcode.ToString().Equals(""))
                    {

                    }
                    else
                    {
                        currentstate = string.Concat(currentstate, " ", uSDDIndividual.diligenceInputModel.CurrentZipcode.ToString().TrimEnd());
                    }
                }
                current_full_address = string.Concat(currentstreet, currentcity, currentstate);
                if (current_full_address.ToString().Equals("")) { doc.Replace("[FullEntityAddress]", "<not provided>", true, true); }
                else
                {
                    doc.Replace("[FullEntityAddress]", current_full_address, true, false);
                }              
                doc.SaveToFile(savePath);
                if (uSDDIndividual.diligenceInputModel.Employer1State.ToString().ToUpper().Equals("Select state") || uSDDIndividual.diligenceInputModel.Employer1State.ToString().Equals(""))
                {
                    doc.Replace("[IncorporationState]", "<not provided>", true, false);                    
                }
                else
                {
                    doc.Replace("[IncorporationState]", uSDDIndividual.diligenceInputModel.Employer1State.TrimEnd(), true, false);
                }
                doc.SaveToFile(savePath);
                //Business Affiliation                  
                if (uSDDIndividual.otherdetails.Has_Business_Affiliations.ToString().Equals("Yes"))
                {
                    doc.Replace("SODCOMMENT", "[FullEntityName] was formed in [IncorporationState] in [IncorporationYear]\n\nFurther, other affiliated business entities were also identified", true, true);
                    doc.Replace("SODRESULT", "Records", true, true);
                    doc.Replace("BUSINESSAFFILIATIONSIDENTIFIED", "[COMADDBUSINESSBEFORE]In addition to the above, research of corporate records maintained by the Secretary of State’s Office, as well as other sources, identified the following [ShortEntityName]-affiliated business entities: <investigator to modify the below list as needed (i.e. footnote any relevant SOS details, delineate Active v. Inactive entities and/or add language/asterisks symbols for entities outside of the US>\n\n\t•\t<Affiliate Name>[addbusinessfootnote]\n\t•\t<Affiliate Name>\n\t•\t<Affiliate Name>[COMADDBUSINESSAFTER]", true, true);
                }
                else
                {
                    doc.Replace("SODCOMMENT", "[FullEntityName] was formed in [IncorporationState] in [IncorporationYear]", true, true);
                    doc.Replace("SODRESULT", "Clear", true, true);
                    doc.Replace("BUSINESSAFFILIATIONSIDENTIFIED", "[COMADDBUSINESSBEFORE]In addition to the above, research of records maintained by the Secretary of State’s Office, as well as other sources, did not identify the subject entity in connection with any business entities.[COMADDBUSINESSAFTER]", true, true);
                }              
                doc.SaveToFile(savePath);
                if (uSDDIndividual.diligenceInputModel.CommonNameSubject == true)
                {
                    doc.Replace("[COMADDBUSINESSBEFORE]", "For the purposes of this investigation, research efforts were structured utilizing the “[FullEntityName]” nomenclature in [IncorporationState], where the subject entity was domiciled, and [EntityState], where it maintains its headquarters.  Additional research efforts on the firm’s affiliated entities and/or prior names, as well as other relevant jurisdictions, can be conducted upon request -- if an expanded scope is warranted. <Investigator to modify (or remove) as needed to include other registration states or other search parameters>\n\n", true, false);
                    doc.Replace("[COMADDBUSINESSAFTER]", "\n\nIt is noted that numerous other entities utilizing the “[ShortEntityName]” nomenclature were identified in connection with corporate filings throughout the United States, and as such, investigative efforts were focused to “[FullEntityName]”.  Significant additional research efforts would be required in connection with other affiliated entities utilizing the “[ShortEntityName]” name, which can be conducted upon request -- if an expanded scope is warranted. <Investigator to modify (or remove) as needed based on search parameters>", true, true);
                }
                else
                {
                    doc.Replace("[COMADDBUSINESSBEFORE]", "", true, false);
                    doc.Replace("[COMADDBUSINESSAFTER]", "", true, true);
                }
                doc.SaveToFile(savePath);
                try
                {
                    string blnconfound = "";
                    if (uSDDIndividual.otherdetails.Has_Business_Affiliations.ToString().Equals("Yes"))
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
                                            TextRange _text = footnote1.TextBody.AddParagraph().AppendText("According to records maintained by the Secretary of State’s Office, <Affiliate Name> was formed in <State> on <date>, where it is <status>.");
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
                //Global security hits
                string globalhit_comment;
                
                if (uSDDIndividual.otherdetails.Global_Security_Hits.ToString().Equals("Yes"))
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
                string icijsummcomment = "";
                string icijsummresult = "";
                if (uSDDIndividual.otherdetails.ICIJ_Hits.ToString().Equals("Yes"))
                {
                    icij_comment = icij_comment2.confirmed_comment.ToString().Replace("subject", "subject entity");
                    icij_comment = string.Concat(icij_comment, "\n\n\t", "•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here> \n");
                    icijsummcomment = "<investigator to insert summary here>";
                    icijsummresult = "Records";
                }
                else
                {
                    icij_comment = icij_comment2.unconfirmed_comment.ToString().Replace("subject", "subject entity");
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
                        doc.Replace("[SECCOMMENT]", "<investigator to insert summary here>", true, true);
                        doc.Replace("[SECRESULT]", "Records", true, true);
                    }
                    else
                    {                        
                        doc.Replace("[SECCOMMENT]", "", true, true);
                        doc.Replace("[SECRESULT]", "Clear", true, true);
                    }                  
                    // doc.Replace("US_SECHEADER", " \nUnited States Securities and Exchange Commission \n", true, true);
                    usseccommentmodel = string.Concat("\nUnited States Securities and Exchange Commission\n\n", "According to a Uniform Application for Investment Advisor Registration (“Form ADV”)[ussecfootnote]", "\n");
                }
                //UK_FCA
                string ukfcacommentmodel = "";                
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
                if (uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Not Registered"))
                {
                    hksfccommentmodel = "";
                }
                else
                {
                    if (uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes - Without Adverse"))
                    {
                        hksfccommentmodel = "[ShortEntityName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>), in the regulated activity of <Investigator to insert activities and dates>.\n\n<Investigator to summarize other license record information, prior regulated activities and dates, previous names, Responsible Officers / Representatives, conditions, etc.>\n\nThere are no public disciplinary actions on file against the subject within the past five years.";
                    }
                    else
                    {
                        hksfccommentmodel = "[ShortEntityName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>)), in the regulated activity of <Investigator to insert activities and dates>.\n\n<Investigator to summarize other license record information, prior regulated activities and dates, previous names, Responsible Officers / Representatives, conditions, etc.>\n\nAdditionally, the following public disciplinary actions were on file in connection with the subject entity for the past five years:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
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
                    holdslicensecommentmodel = "\nOther Professional Licensures and/or Designations\n\nInvestigative efforts did not reveal any additional licensure and/or registration information in connection with [ShortEntityName], however, further efforts would be required on a jurisdiction and license-type basis to confirm the same.  As a precaution, research efforts included, but were not limited to: a Brokerage Firm Check through the United States Financial Industry Regulatory Authority (for the past two years); a search of the United States Commodity Futures Trading Commission registration; an Investment Adviser Firm search through the United States Securities and Exchange Commission; a search through the United Kingdom’s Financial Conduct Authority; a search through the Hong Kong Securities and Futures Commission; as well as searches through the Bermuda Monetary Authority, the Cayman Islands Monetary Authority and the British Virgin Islands Financial Services Commission.  <Investigator to add or remove any sources that came back with hits above or are needed for this jurisdiction>\n";
                }
                else
                {
                    holdslicensecommentmodel = "\nInvestigative efforts did not reveal any licensure and/or registration information in connection with [ShortEntityName], however, further efforts would be required on a jurisdiction and license-type basis to confirm the same.  As a precaution, research efforts included, but were not limited to: a Brokerage Firm Check through the United States Financial Industry Regulatory Authority (for the past two years); a search of the United States Commodity Futures Trading Commission registration and National Futures Association membership information; an Investment Adviser Firm search through the United States Securities and Exchange Commission; a search through United Kingdom’s Financial Conduct Authority; a search through the Hong Kong Securities and Futures Commission; as well as searches through the Bermuda Monetary Authority, the Cayman Islands Monetary Authority and the British Virgin Islands Financial Services Commission.  <Investigator to add or remove any sources that came back with hits above or are needed for this jurisdiction/scope/search type>\n";
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
                                            if (uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse"))
                                            {
                                                strredflagtextappended = " filed with the United States Securities and Exchange Commission (“SEC”), [ShortEntityName]  has been registered, with a Central Registration Depository (“CRD”) number of <number>, as an Investment Adviser since <date>, and manages <number> discretionary and/or non-discretionary accounts, totaling approximately $<amount from Item 5 F(2)> in assets under its management at the time this filing was captured on <filing date>.\n\nIn addition, [ShortEntityName] provided investment advisory services to <number> clients during their most recently-completed fiscal year, which consisted of <Investigator to list out client types from Item 5 C&D of the Form ADV>. Additionally, <percentage number>% of its clients are non-United States persons <Investigator to replace with ‘none of its clients…’ if percentage number is 0>.\n\nFurther, [ShortEntityName] is reportedly compensated for its investment advisory services by <Investigator to list out compensation methods checked in Item 5E>.\n\nThe Form ADV also revealed that as part of its services, the firm reportedly provides <Investigator to list out services provided in Item 5G of form>. Moreover, the firm does not sell products or provide services other than investment advice to their advisory clients <Investigator to remove or modify if untrue – see Item 6 B (3)>.\n\nAdditionally, the Form ADV also reports that [ShortEntityName] has approximately <number> employees, <number> of which <Investigator to summarize Item 5 A & B (1-6) in form>. Further, according to the filing, no “firms or other persons solicit advisory clients on [their] behalf,” and none of these employees are registered representatives of a broker-dealer. <Investigator to modify language as appropriate from the form>\n\nThe Form ADV also revealed the following Direct Owners and/or Executive Officers of the subject entity at the time of filing: <Investigator to add Direct Owner/Officer information from Schedule A>. <Investigator to add/indicate whether Indirect Direct Owner information exists in Schedule B>.\n\nFurther, the following disclosure events were reported on the subject entity’s SEC registration:\n\n\t•\t<Investigator to insert bulleted list of results from Schedule R - DRP Pages here>\n\t•\t<Investigator to insert bulleted list of results from Schedule R - DRP Pages here>\n\t•\t<Investigator to insert bulleted list of results from Schedule R - DRP Pages here>\n\nWith the exception of the above, no additional disclosure events, such as formal investigations, disciplinary actions, customer disputes and/or criminal charges or convictions, were on file with the SEC in connection with [ShortEntityName].";
                                            }
                                            else
                                            {
                                                strredflagtextappended = " filed with the United States Securities and Exchange Commission (“SEC”), [ShortEntityName] has been registered, with a Central Registration Depository (“CRD”) number of <number>, as an Investment Adviser since <date>, and manages <number> discretionary and/or non-discretionary accounts, totaling approximately $<amount from Item 5 F(2)> in assets under its management at the time this filing was captured on <filing date>.\n\nIn addition, [ShortEntityName] provided investment advisory services to <number> clients during their most recently-completed fiscal year, which consisted of <Investigator to list out client types from Item 5 C&D of the Form ADV>. Additionally, <percentage number>% of its clients are non-United States persons <Investigator to replace with ‘none of its clients…’ if percentage number is 0>.\n\nFurther, [ShortEntityName] is reportedly compensated for its investment advisory services by <Investigator to list out compensation methods checked in Item 5E>.\n\nThe Form ADV also revealed that as part of its services, the firm reportedly provides <Investigator to list out services provided in Item 5G of form>. Moreover, the firm does not sell products or provide services other than investment advice to their advisory clients <Investigator to remove or modify if untrue – see Item 6 B (3)>.\n\nAdditionally, the Form ADV also reports that [ShortEntityName] has approximately <number> employees, <number> of which <Investigator to summarize Item 5 A & B (1-6) in form>. Further, according to the filing, no “firms or other persons solicit advisory clients on [their] behalf,” and none of these employees are registered representatives of a broker-dealer. <Investigator to modify language as appropriate from the form>\n\nThe Form ADV also revealed the following Direct Owners and/or Executive Officers of the subject entity at the time of filing: <Investigator to add Direct Owner/Officer information from Schedule A>. <Investigator to add/indicate whether Indirect Direct Owner information exists in Schedule B>.\n\n<Investigator to add summary of other financial industry affiliate information, and/or listing of private funds in the Form ADV – as appropriate -- in Sections 7.A. and 7.B>\n\nNo disclosure events, such as formal investigations, disciplinary actions, customer disputes and/or criminal charges or convictions, were on file with the SEC in connection with the subject entity.";
                                            }
                                          //  strredflagtextappended = strredflagtextappended.Replace("[FullEntityName]", uSDDIndividual.diligenceInputModel.LastName);
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
                }catch{ }
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
                                        if (uSDDIndividual.otherdetails.RegulatoryFlag.Equals("Yes"))
                                        {
                                            strredflagtextappended = " and the following information was identified in connection with [ShortEntityName]:   <Investigator to insert results here>";
                                        }
                                        else
                                        {
                                            strredflagtextappended = " and it is noted that [ShortEntityName] was not identified in any of these records.";
                                        }   
                                        
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
                }catch{ }
                //AdditionalStates
                string add_states = "";
                string add_states2 = "";
                string strAdditionalStatesComment = "";
                string arr = "";
                string state1 = uSDDIndividual.diligenceInputModel.CurrentState.ToString();
                string state2 = uSDDIndividual.diligenceInputModel.Employer1State.ToString();
                int statecount = 0;
                try
                {
                    if (uSDDIndividual.additional_States.Count > 0)
                    {
                        if (uSDDIndividual.additional_States[0].additionalstate.ToUpper().ToString().Equals("Select State".ToUpper()) && uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State") && uSDDIndividual.diligenceInputModel.Employer1State.ToString().Equals("Select State"))
                        {
                            doc.Replace("[ADDITIONALSTATES]", "", true, true);
                            add_states = "<not provides>";
                        }

                        if (state1 == "Select State") { }
                        else
                        {
                            add_states = string.Concat("<investigator to add relevant county/countries>, ", uSDDIndividual.diligenceInputModel.CurrentState.ToString());
                            arr = uSDDIndividual.diligenceInputModel.CurrentState.ToString();
                            statecount = statecount + 1;
                        }
                        if (state2 == "Select State" || state2 == state1) { }
                        else
                        {
                            add_states2 = string.Concat("<investigator to add relevant county/countries>, ", uSDDIndividual.diligenceInputModel.Employer1State.ToString());
                            arr = string.Concat(arr, ", ", uSDDIndividual.diligenceInputModel.Employer1State.ToString());
                            statecount = statecount + 1;
                        }

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
                                            strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, "<investigator to add relevant county/countries>, ", uSDDIndividual.additional_States[j].additionalstate, " and ");
                                            add_states = strAdditionalStatesComment;
                                        }
                                        else
                                        {
                                            strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, "<investigator to add relevant county/countries>, ", uSDDIndividual.additional_States[j].additionalstate, ", ");
                                            if (i == count) { add_states = string.Concat(add_states, "<investigator to add relevant county/countries>, ", uSDDIndividual.additional_States[j].additionalstate); }
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
                        if (uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State") && uSDDIndividual.diligenceInputModel.Employer1State.ToString().Equals("Select State"))
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
                        if (state2 == "Select State" || state2 == state1) { }
                        else
                        {
                            add_states2 = string.Concat("<investigator to add relevant county/counties>, ", uSDDIndividual.diligenceInputModel.Employer1State.ToString());
                            arr = string.Concat(arr, ", ", uSDDIndividual.diligenceInputModel.Employer1State.ToString());
                            statecount = statecount + 1;
                        }

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
                        //doc.Replace("[REGISTEREDWITHHKSFCDESC]", "\n\n[ShortEntityName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nAdditionally, the following public disciplinary actions were on file in connection with the subject:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>", true, false);
                    }
                    else
                    {
                        doc.Replace("KONSECCOMMENT", "", true, true);
                        doc.Replace("KONSECRESULT", "Clear", true, true);
                        doc.Replace("[HKSFCHEAD]", "\nHong Kong Securities and Futures Commission", true, false);
                      //  doc.Replace("[REGISTEREDWITHHKSFCDESC]", "\n\n[ShortEntityName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nThere are no public disciplinary actions on file against the subject within the past five years.\n\n", true, false);
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
                    strconcatBankrupty = " identified the subject entity as a <party type> in connection with the following bankruptcy filings:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n";
                }
                else
                {
                    if (uSDDIndividual.otherdetails.HasBankruptcyRecHits.ToString().Equals("No"))
                    {
                        strbank1 = "Clear";
                        doc.Replace("HASBANKRUPTYRECHITDESC", "Efforts pursued in [ADDITIONALBANKSTATES]ADDFOOTNOTE", true, true);
                        strconcatBankrupty = " did not identify the subject entity in connection with any bankruptcy filings.\n";
                    }
                    else
                    {
                        strbank1 = "Record";
                        doc.Replace("HASBANKRUPTYRECHITDESC", "Efforts pursued in [ADDITIONALBANKSTATES]ADDFOOTNOTE", true, true);
                        strconcatBankrupty = " identified the subject entity as a <party type> in connection with the following bankruptcy filing:\n\n\t•\t<Investigator to insert bulleted list of results here>\n";
                    }
                }
                //if (uSDDIndividual.diligenceInputModel.CommonNameSubject == true)
                //{
                //    strconcatBankrupty = string.Concat(strconcatBankrupty, "\nIn light of the commonality of [ShortEntityName]’s name, this portion of the investigation was conducted utilizing the subject’s Social Security number, in order to identify only those records conclusively relating to the subject of interest.\n");
                //}
                if (uSDDIndividual.otherdetails.HasBankruptcyRecHits1 == true)
                {
                    strconcatBankrupty = string.Concat(strconcatBankrupty, "\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>. <Investigator to add other relevant details (i.e. relevant bank cases and/or dispositions)>\n\nManual civil court records research efforts – as available – would be required to determine whether the same relate to the subject entity of this investigation, which can be undertaken upon request – if an expanded scope is warranted.\n");
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
                                 .Where(u => u.states.ToUpper().TrimEnd() == arrlist[j].TrimEnd())
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
                try
                {
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
                }
                catch { }
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
                    strbankresrs = "Bankruptcy court records searches are currently pending in <jurisdiction> the results of which will be provided under separate cover upon receipt\n\n";
                    strbankrs = "Results Pending\n\n";
                }
                if (uSDDIndividual.otherdetails.HasBankruptcyRecHits1 == true)
                {
                    strbankpr = "\n\nOne or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <case type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                    strbankrespr = "\n\nPossible Records";
                }
                switch (strbank1)
                {
                    case "Clear":
                        strbankressumm = "Clear";
                        strbanksumcomm = "";
                        break;
                    case "Record":
                        strbanksumcomm = "[FullEntityName] was identified as having been a <party type> in connection with a <case type>, which was recorded in <State> in <YYYY>, and is currently <status>";
                        strbankressumm = "Record";
                        break;
                    case "Records":
                        strbanksumcomm = "[FullEntityName] was identified as a <party type> in connection with at least <number> <case type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                        strbankressumm = "Records";
                        break;
                }
                strbanksumcomm = string.Concat(strbankresrs, strbanksumcomm.TrimEnd(), strbankpr);
                strbankressumm = string.Concat(strbankrs, strbankressumm, strbankrespr);
                doc.Replace("BANKRUPCOMMENT", strbanksumcomm, true, true);
                doc.Replace("BANKRUPRESULT", strbankressumm, true, true);
                doc.SaveToFile(savePath);

                TextSelection[] textSelection = doc.FindAllString("Bankruptcy court records searches are currently pending in <jurisdiction> the results of which will be provided under separate cover upon receipt", true, true);
                if (textSelection != null)
                {
                    foreach (TextSelection seletion in textSelection)
                    {
                        seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                        seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    }
                }

                //HASNAMECIVILLITIDESC 
                if (uSDDIndividual.otherdetails.Has_Name_Only == false)
                {
                    doc.Replace("[HASNAMECIVILLITIDESC]", "", true, false);
                }
                else
                {
                    doc.Replace("[HASNAMECIVILLITIDESC]", "\n\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.[ADDNAMEONLYFN]\n\nManual civil court records research efforts – as available – would be required to determine whether the same relate to the subject entity of this investigation, which can be undertaken upon request – if an expanded scope is warranted.", true, false);
                }
                doc.SaveToFile(savePath);
                strblnres = ""; try
                {
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
                }
                catch { }
                doc.SaveToFile(savePath);
                doc.Replace("[ADDNAMEONLYFN]", "", true, false);
                //USTAXCOURTHITDESC  [ADDNAMEONLYFN] 
                if (uSDDIndividual.otherdetails.Has_US_Tax_Court_Hit.ToString().StartsWith("Yes"))
                {
                    doc.Replace("[USTAXCOURTHITDESC]", "\n\nAdditionally, research efforts were conducted of filings through the United States Tax Court, and the following information was identified in connection with [ShortEntityName]:  <Investigator to insert results here>\n", true, false);
                }
                else
                {
                    doc.Replace("[USTAXCOURTHITDESC]", "\n\nAdditionally, research efforts were conducted of filings through the United States Tax Court, and no records were identified for the subject entity.\n", true, false);
                }
                //HASTAXLIENSCIVILUCCDESC
                string strtaxlien = "";
                if (uSDDIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("No"))
                {
                    strtaxlien = "Investigative efforts did not reveal any tax liens in connection with [ShortEntityName]. <Investigator to combine/modify as needed with other filing types (i.e. any tax liens or judgments).>";
                }
                else
                {
                    if (uSDDIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("Yes, Multiple Records"))
                    {
                        strtaxlien = "Investigative efforts revealed the subject entity as having been a debtor in connection with the following tax liens:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strtaxlien = "Investigative efforts revealed the subject entity as having been a debtor in connection with the following tax lien:\n\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                }
                //HAS CIVIL_JUDGEMENT
                string strciviljudge = "";
                if (uSDDIndividual.otherdetails.Has_Property_Records.Equals("Yes, Single Record"))
                {
                    strciviljudge = "\n\nAdditionally, investigative efforts revealed the subject entity as having been a <party type> in connection with the following civil judgment:\n\n\t•	<Investigator to insert bulleted list of results here>";
                }
                else
                {
                    if (uSDDIndividual.otherdetails.Has_Property_Records.Equals("Yes, Multiple Records"))
                    {
                        strciviljudge = "\n\nAdditionally, investigative efforts revealed the subject entity in connection with the following civil judgments:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                    }
                    else
                    {
                        strciviljudge = "\n\nAdditionally, investigative efforts did not reveal any civil judgments in connection with [ShortEntityName]. <Investigator to combine/modify as needed with other filing types (i.e. any tax liens or judgments).>";
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
                        strcommjudge = "[FullEntityName] was identified as a <party type> in connection with a <record type>, which was filed in <State> in <YYYY>, and is currently <status>";
                        strresjudge = "Record";
                        break;
                    case "Records":
                        strcommjudge = "[FullEntityName] was identified as a <party type> in connection with at least <number> <record types>, which were filed in <States> between <YYYY> and <YYYY>, and are currently <status>";
                        strresjudge = "Records";
                        break;
                }

                if (uSDDIndividual.otherdetails.Has_Name_Only_Tax_Lien == true)
                {
                    strresjudge = string.Concat(strresjudge, "\n\nPossible Records");
                    strcommjudge = string.Concat(strcommjudge, "\n\nOne or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>");
                }

                doc.Replace("TAXLIENCOMMENT", strcommjudge, true, true);
                doc.Replace("TAXLIENRESULT", strresjudge, true, true);
                doc.SaveToFile(savePath);                                           
                
                //HASNAMETAXLIENUCCDESC
                string strtaxliennameonly = "";
                if (uSDDIndividual.otherdetails.Has_Name_Only_Tax_Lien == true)
                {
                    strtaxliennameonly = "\n\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.  Manual records research efforts -- as available -- would be required to determine whether the same relate to the subject entity of this investigation, which can be undertaken upon request -- if an expanded scope is warranted.";
                }
                //HASUCC
                string strucccomment = "";
                if (uSDDIndividual.otherdetails.has_ucc_fillings.ToString().Equals("Yes, Single Record"))
                {
                    strucccomment = "\n\nFurther, investigative efforts revealed the subject entity as having been a <party type> in connection with a Uniform Commercial Code (“UCC”) filing, which was recorded in <State> in <YYYY>, and is currently <status>. <Investigator to insert addition party details and/or collateral information>";
                }
                else
                {
                    if (uSDDIndividual.otherdetails.has_ucc_fillings.ToString().Equals("Yes, Multiple Records"))
                    {
                        strucccomment = "\n\nFurther, investigative efforts revealed the subject entity in connection with the following Uniform Commercial Code (“UCC”) filings:  <Investigator to incorporate summary of results here>.";
                    }
                    else
                    {
                        strucccomment = "\n\nAdditionally, investigative efforts did not reveal [ShortEntityName] in connection with any Uniform Commercial Code filings (“UCC”). <Investigator to combine/modify as needed with other filing types (i.e. any tax liens or UCC filings).>";
                    }
                }
                string commomstr = "";
                //if (uSDDIndividual.diligenceInputModel.CommonNameSubject == true)
                //{
                //    commomstr = "\n\nIn light of the commonality of [ShortEntityName]’s name, this portion of the investigation was conducted utilizing the subject’s Social Security number, in order to identify only those records conclusively relating to the subject of interest.";
                //}
                string strucc1 = "";
                if (uSDDIndividual.otherdetails.has_ucc_fillings1 == true)
                {
                    strucc1 = "\n\nIt should be noted that one or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>.  Manual records research efforts -- as available -- would be required to determine whether the same relate to the subject entity of this investigation, which can be undertaken upon request -- if an expanded scope is warranted.";
                }
                if (uSDDIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("No")  && uSDDIndividual.otherdetails.has_ucc_fillings.ToString().Equals("No"))
                {
                    strciviljudge = "Investigative efforts did not reveal any tax liens, civil judgments or Uniform Commercial Code filings in connection with [ShortEntityName].";
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
                        struccsummcom = "[FullEntityName] was identified as a <party type> in connection with a UCC filing, which was recorded in <State> in <YYYY>, and is currently <status>";
                        struccsummres = "Record";
                        break;
                    case "Yes, Multiple Records":
                        struccsummcom = "[FullEntityName] was identified as a <party type> in connection with at least <number> UCC filings, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                        struccsummres = "Records";
                        break;
                }
                if (uSDDIndividual.otherdetails.has_ucc_fillings1 == true)
                {
                    struccsummres = string.Concat(struccsummres, "\n\nPossible Records");
                    struccsummcom = string.Concat(struccsummcom, "\n\nOne or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <record type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>");
                }
                doc.Replace("UNIFCOMMENT", struccsummcom, true, true);
                doc.Replace("UNIFRESULT", struccsummres, true, true);
                doc.SaveToFile(savePath);
              
           
                //PRESSANDMEDIASEARCHDESCRIPTION
                switch (uSDDIndividual.otherdetails.Press_Media.ToString())
                {
                    case "Common name with adverse Hits":
                        doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the commonality of the subject entity's name, searches were structured and designed to identify specific reports relating to [ShortEntityName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. In that regard, numerous articles were reviewed, and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);                       
                        break;
                    case "Common name without adverse Hits":
                        doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the commonality of the subject entity's name, searches were structured and designed to identify specific reports relating to [ShortEntityName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. In that regard, numerous articles were reviewed, and a thorough review of the same did not identify any adverse or materially-significant information in connection with [ShortEntityName].", true, true);
                        break;
                    case "High volume with adverse Hits":
                        doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the volume of press and other media sources identified in connection with the subject entity, searches were structured and designed to uncover specific reports relating to [ShortEntityName] with a variety of negative search terms and derivatives thereof within the past <timeframe>.  In that regard, numerous articles were reviewed, and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);
                        break;
                    case "High volume without adverse Hits":
                        doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the volume of press and other media sources identified in connection with the subject entity, searches were structured and designed to uncover specific reports relating to [ShortEntityName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. In that regard, numerous articles were reviewed, and a thorough review of the same did not identify any adverse or materially-significant information in connection with [ShortEntityName].", true, true);
                        break;
                    case "Standard search with adverse Hits":
                        doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, identified various articles and/or media references in connection with [ShortEntityName], and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>", true, true);
                        break;
                    case "Standard search without adverse Hits":
                        doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, identified various articles and/or media references in connection with [ShortEntityName], and a thorough review of the same did not identify any adverse or materially-significant information.", true, true);
                        break;
                    case "No Hits":
                        doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, did not identify any articles and/or media references in connection with [ShortEntityName].", true, true);
                        break;
                }
                doc.SaveToFile(savePath);
                          
                strblnres = "";              
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
                        strconcatcivilrec = " did not identify the subject entity in connection with any civil litigation filings.";
                        break;
                    case "Yes, Multiple Records":
                        strsummcivilres = "Records";
                        doc.Replace("[HASCIVILRECDESC]", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                        strconcatcivilrec = " identified the subject entity in connection with the following civil litigation filings:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                        break;
                    case "Yes, Single Record":
                        strsummcivilres = "Record";
                        doc.Replace("[HASCIVILRECDESC]", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                        strconcatcivilrec = " identified the subject entity in connection with the following civil litigation filing:\n\n\t•\t<Investigator to insert bulleted list of results here>";
                        break;
                }
                doc.SaveToFile(savePath);
                strblnres = "";
                try
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
                                    try
                                    {
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
                                    catch { }
                                }
                                else
                                {
                                    if (k == count)
                                    {
                                        strblnres = "";
                                        strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE]");
                                        doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                        doc.SaveToFile(savePath);
                                        try
                                        {
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
                                        catch { }
                                    }
                                    else
                                    {
                                        strblnres = "";
                                        strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE],");
                                        doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                        doc.SaveToFile(savePath);
                                        try
                                        {
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
                                        catch { }
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
                                  .Where(u => u.states == arrlist[0].ToString())
                                  .FirstOrDefault();
                        string strstatecourt = comment1.state_specific_district_courts;
                        string strcivilfootnote = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert county courts searched>.";
                        strcivilfootnote = strcivilfootnote.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                        doc.SaveToFile(savePath);
                        try
                        {
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
                        catch { }
                    }
                }
                catch { }
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
                        strsummcivilcom = "[FullEntityName] was identified as a <party type> in connection with a civil litigation filing, which were recorded in <State> in <YYYY>, and is currently <status>";
                        break;
                    case "Records":
                        strsummcivilcom = "[FullEntityName] was identified as a <party type> in connection with at least <number> civil litigation filings, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                        break;
                }
                if (uSDDIndividual.otherdetails.has_civil_resultpending == true)
                {
                    strsummcivilres = string.Concat("Result Pending\n\n", strsummcivilres);
                    strsummcivilcom = string.Concat("A civil court records search is currently pending through the <enter source/court>, the results of which will be provided under separate cover upon receipt\n\n", strsummcivilcom);
                }
                if (uSDDIndividual.otherdetails.Has_Name_Only == true)
                {
                    strsummcivilres = string.Concat(strsummcivilres, "\n\nPossible Records");
                    strsummcivilcom = string.Concat(strsummcivilcom, "\n\nOne or more entities known only as “[ShortEntityName]” were identified as <party type> in at least <number> <case type>, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>");
                }
                doc.Replace("CIVILRESULT", strsummcivilres, true, true);
                doc.Replace("CIVILCOMMENT", strsummcivilcom, true, true);
                doc.SaveToFile(savePath);
                TextSelection[] textresult1 = doc.FindAllString("A civil court records search is currently pending through the <enter source/court>, the results of which will be provided under separate cover upon receipt", false, false);
                if (textresult1 != null)
                {
                    foreach (TextSelection seletion in textresult1)
                    {
                        seletion.GetAsOneRange().CharacterFormat.Italic = true;
                        seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                    }
                }
                           
                //PL License details
                string pl_comment = "";
                string plstartdate = "";
                string plenddate = "";
                string bnres = "";
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
                                    pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.confirmed_comment.ToString(), "\n\nThere was no disciplinary history identified in connection with the subject entity’s license. <investigator to modify if disciplinary history exists>");
                                }
                                else
                                {
                                    pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.confirmed_comment.ToString(), "\n\nThere was no disciplinary history identified in connection with the subject entity’s license. <investigator to modify if disciplinary history exists>", "\n\n");
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
                            if (uSDDIndividual.diligenceInputModel.FullaliasName.ToString().Equals("")) {
                                pl_comment = pl_comment.Replace("[Last Name]", "<not provided>");
                            }
                            else
                            {
                                pl_comment = pl_comment.Replace("[Last Name]", uSDDIndividual.diligenceInputModel.FullaliasName.ToString());
                            }
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
                                    // Find the word "Civil" in paragraph1                             
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
                }
                catch { }
                doc.Replace("[CIVILFONTCHANGE]", "", true, false);
                doc.Replace("[BANKFONTCHANGE]", "", true, false);
                doc.SaveToFile(savePath);
                if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                {
                    try
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
                    catch { }
                }
                //EDUCATIONALANDLICENSINGHITS
                string eduhistory = ""; string plhistory = "";
                
                string strEDULICENcomment = "";
                try
                {
                    if (uSDDIndividual.educationModels[0].Edu_History.ToString().Equals("Yes") && uSDDIndividual.educationModels[0].Edu_Confirmed.ToString().Equals("Yes"))
                    {
                        eduhistory = "Additionally, efforts included verification of the subject entity’s educational credentials, where available";
                        strEDULICENcomment = eduhistory;
                    }
                }
                catch { }
                try
                {
                    if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes") && uSDDIndividual.pllicenseModels[0].PL_Confirmed.ToString().Equals("Yes"))
                    {
                        plhistory = "Additionally, efforts included verification of the subject entity’s licensing credentials, where available";
                        strEDULICENcomment = plhistory;
                    }
                }
                catch { }
                if (uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_FINRA.ToString().StartsWith("Yes"))
                {
                    plhistory = "Additionally, efforts included verification of the subject entity’s licensing credentials, where available";
                    strEDULICENcomment = plhistory;
                }
                if (eduhistory != "" && plhistory != "")
                {
                    strEDULICENcomment = "Additionally, efforts included verification of the subject entity’s educational and licensing credentials, where available";
                }                                               
                doc.Replace("EDUCATIONALANDLICENSINGHITS", strEDULICENcomment, true, false);

                //Results Pending in Legal Records Section?
                if (uSDDIndividual.otherdetails.HasBankruptcyRecHits_resultpending == true  || uSDDIndividual.otherdetails.has_civil_resultpending == true)
                {
                    doc.Replace("[RESULTSPENSECTIONS1]", "\n\n<Search type> searches are currently ongoing in <jurisdiction>, the results of which will be provided under separate cover upon receipt.", true, false);
                }
                else
                {
                    doc.Replace("[RESULTSPENSECTIONS1]", "", true, false);
                }
                TextSelection[] texttt = doc.FindAllString("<Search type> searches are currently ongoing in <jurisdiction>, the results of which will be provided under separate cover upon receipt.", true, true);
                if (texttt != null)
                {
                    foreach (TextSelection seletion in texttt)
                    {
                        seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                        seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    }
                }
                if (uSDDIndividual.otherdetails.HasBankruptcyRecHits.ToString().StartsWith("Yes")  || uSDDIndividual.otherdetails.Has_Civil_Records.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Name_Only == true || uSDDIndividual.otherdetails.Has_US_Tax_Court_Hit.ToString().StartsWith("Yes") )
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
              
                if (uSDDIndividual.otherdetails.Has_Legal_Records_Hits.ToString().Equals("Yes"))
                {
                    Legrechitcommentmodel = "Searches were conducted through legal records and other sources in the United States in order to locate any bankruptcy filings or other claims, evidence of government actions, civil litigation filings, or liens or judgments in connection with the subject entity, and the following records were identified:   <Investigator to insert summarized findings>";
                }
                else
                {
                    Legrechitcommentmodel = "Searches of legal records and other sources in the United States did not locate any bankruptcy filings or other claims, evidence of government actions, civil litigation filings, or liens or judgments in connection with the subject entity.";
                }
                doc.Replace("HasLegalRecordsJudgmentsorLiensHits", Legrechitcommentmodel, true, true);
                //Has Regulatory or Global Security Hits?
                string hasreghitcommentmodel = "";                
                if (regflag == "Records")
                {
                    hasreghitcommentmodel = "Searches also were conducted of relevant legal and regulatory agencies in the United States, while anti-terrorist, anti-money laundering and other compliance searches were global in scope, and the following records were identified in connection with the subject entity:\t<Investigator to insert summarized findings>";
                }
                else
                {
                    hasreghitcommentmodel = "Searches also were conducted of relevant legal and regulatory agencies in the United States, while anti-terrorist, anti-money laundering and other compliance searches were global in scope, and no such records were identified in connection with the subject entity.";
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
                        searchtext = "In sum, with the exception of the above, no other issues of potential relevance were identified in connection with";
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
               // hitcompcommentmodel = hitcompcommentmodel.Replace("[FirstName]", uSDDIndividual.diligenceInputModel.FirstName.ToString());
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
                doc.Replace("[FullEntityName]", uSDDIndividual.diligenceInputModel.LastName, true, true);
                doc.Replace("[Country]", "the United States", true, false);
                doc.SaveToFile(savePath);
                string blnresultfound = ""; try
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
                                        TextRange _text = footnote1.TextBody.AddParagraph().AppendText("It should be emphasized that updates are made to these lists on a periodic and irregular basis, and for purposes of preparing this memorandum, these lists were searched on <investigator to insert date of research>.  ");
                                        footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                        _text.CharacterFormat.FontName = "Calibri (Body)";
                                        _text.CharacterFormat.FontSize = 9;
                                        //Append the line
                                        string strglobaltextappended = "";
                                        if (uSDDIndividual.otherdetails.Global_Security_Hits.Equals("Yes"))
                                        {
                                            strglobaltextappended = " and the following information was identified in connection with [ShortEntityName]: ";
                                        }
                                        else
                                        {
                                            strglobaltextappended = " and it is noted that [ShortEntityName] was not identified on any of these lists.";
                                        }                                       
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
               // doc.Replace("[First Name]", uSDDIndividual.diligenceInputModel.FirstName, true, false);
                doc.SaveToFile(savePath);
                int plcommentcount = 0;
                string strPLComment = "";
                if (uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse") || uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse"))
                {
                    if (uSDDIndividual.pllicenseModels[0].General_PL_License.Equals("Yes"))
                    {
                        plcommentcount = 1;
                        strPLComment = "[FullEntityName] has been registered as a [ProfessionalLicenseType1] license with the [PLOrganization1]";
                    }
                    if (uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse"))
                    {
                        if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                        strPLComment = string.Concat(strPLComment, "[FullEntityName] has been registered with the U.S. SEC");
                        if (plcommentcount == 0)
                        { 
                            plcommentcount = 1;
                        } 
                        else
                        {
                            plcommentcount = plcommentcount + 1;
                        }
                    }
                    if (uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse"))
                    {
                        if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                        strPLComment = string.Concat(strPLComment, "[FullEntityName] has been registered with the U.K. FCA");
                        if (plcommentcount == 0)
                        {
                            plcommentcount = 1;
                        }
                        else
                        {
                            plcommentcount = plcommentcount + 1;
                        }
                    }
                    if (uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse"))
                    {
                        if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                        strPLComment = string.Concat(strPLComment, "[FullEntityName] has been registered with the U.S. FINRA");
                        if (plcommentcount == 0)
                        {
                            plcommentcount = 1;
                        }
                        else
                        {
                            plcommentcount = plcommentcount + 1;
                        }
                    }
                    if (uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse"))
                    {
                        if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                        strPLComment = string.Concat(strPLComment, "[FullEntityName] has been registered with the U.S. NFA");
                        if (plcommentcount == 0)
                        {
                            plcommentcount = 1;
                        }
                        else
                        {
                            plcommentcount = plcommentcount + 1;
                        }
                    }
                    if (uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse"))
                    {
                        if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                        strPLComment = string.Concat(strPLComment, "[FullEntityName] has been registered with HKSFC");
                        if (plcommentcount == 0)
                        {
                            plcommentcount = 1;
                        }
                        else
                        {
                            plcommentcount = plcommentcount + 1;
                        }
                    }
                    if (uSDDIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSDDIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse"))
                    {
                        strPLComment = string.Concat(strPLComment, "\n\nFurther, <investigator to insert regulatory hits here>");
                        if (plcommentcount == 0)
                        {
                            plcommentcount = 1;
                        }
                        else
                        {
                            plcommentcount = plcommentcount + 1;
                        }
                    }
                    strPLComment = strPLComment.Replace("[ProfessionalLicenseType1]", uSDDIndividual.pllicenseModels[0].PL_License_Type);
                    strPLComment = strPLComment.Replace("[PLOrganization1]", uSDDIndividual.pllicenseModels[0].PL_Organization);
                    if (plcommentcount == 1)
                    {
                        doc.Replace("PROFRESULT", "Record", true, true);
                    }
                    else
                    {
                        doc.Replace("PROFRESULT", "Records", true, true);
                    }
                }
                else
                {
                    if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                    {
                        doc.Replace("PROFRESULT", "Record", true, true);
                        strPLComment = "[FullEntityName] has been registered as a [ProfessionalLicenseType1] license with the [PLOrganization1]";
                        strPLComment = strPLComment.Replace("[ProfessionalLicenseType1]", uSDDIndividual.pllicenseModels[0].PL_License_Type);
                        strPLComment = strPLComment.Replace("[PLOrganization1]", uSDDIndividual.pllicenseModels[0].PL_Organization);
                    }
                    else
                    {
                        strPLComment = "";
                        doc.Replace("PROFRESULT", "Clear", true, true);

                    }
                }
                doc.Replace("PROFCOMMENT", strPLComment, true, true);
                doc.SaveToFile(savePath);                
                doc.Replace("[IncorporationState]", uSDDIndividual.diligenceInputModel.Employer1State,true,false);
                try
                {
                    doc.Replace("[IncorporationYear]", Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Year.ToString(), true, false);
                }
                catch
                {
                    doc.Replace("[IncorporationYear]","<not provided>", true, false);
                }
                if (uSDDIndividual.diligenceInputModel.FullaliasName.ToString().Equals(""))
                {
                    doc.Replace("[ShortEntityName]", "<not provided>", true, false);
                }
                else
                {
                    doc.Replace("[ShortEntityName]", uSDDIndividual.diligenceInputModel.FullaliasName.ToString(), true, false);
                }
                if (uSDDIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                {                    
                    doc.Replace("[EntityState]", "", true, false);                  
                }
                else
                {
                    doc.Replace("[EntityState]", uSDDIndividual.diligenceInputModel.CurrentState.ToString(), true, false);                 
                }
                doc.Replace("[FirstName] [MiddleInitial] [LastName]", uSDDIndividual.diligenceInputModel.LastName, true, true);
                doc.SaveToFile(savePath);
                HttpContext.Session.SetString("last_name", uSDDIndividual.diligenceInputModel.LastName);
                HttpContext.Session.SetString("case_number", uSDDIndividual.diligenceInputModel.CaseNumber);                
                HttpContext.Session.SetString("city", uSDDIndividual.diligenceInputModel.CurrentCity);                
                HttpContext.Session.SetString("employer1State", uSDDIndividual.diligenceInputModel.Employer1State);                                        
                HttpContext.Session.SetString("additionalstates", add_states);                
                HttpContext.Session.SetString("pl_generallicense", uSDDIndividual.pllicenseModels[0].General_PL_License.ToString());
                if (uSDDIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                {
                    HttpContext.Session.SetString("pl_licensetype", uSDDIndividual.pllicenseModels[0].PL_License_Type.ToString());
                    HttpContext.Session.SetString("pl_organization", uSDDIndividual.pllicenseModels[0].PL_Organization.ToString());
                    if (uSDDIndividual.pllicenseModels[0].PL_StartDateYear.ToString().Equals(""))
                    {
                        plstartdate = "<not provided>";
                    }
                    else
                    {
                        plstartdate = uSDDIndividual.pllicenseModels[0].PL_StartDateYear.ToString();
                    }
                    if (uSDDIndividual.pllicenseModels[0].PL_EndDateYear.ToString().Equals(""))
                    {
                        plenddate = "<not provided>";
                    }
                    else
                    {
                        plenddate = uSDDIndividual.pllicenseModels[0].PL_EndDateYear.ToString();
                    }
                    HttpContext.Session.SetString("pl_startdate", plstartdate);
                    HttpContext.Session.SetString("pl_enddate", plenddate);
                }
                doc.SaveToFile(savePath);
                SummaryResulttableModel ST = new SummaryResulttableModel();
                ST = uSDDIndividual.summarymodel;

                return Save_Page2(ST);
            }
            catch (IOException e)
            {
                string recordid = HttpContext.Session.GetString("recordid");
                HttpContext.Session.SetString("recordid", recordid);
                TempData["message"] = e;
                return US_COMPANY_Edit();
            }
        }
        public IActionResult US_DD_Individual_page2()
        {
            return View();
        }
        public IActionResult Save_Page2(SummaryResulttableModel ST)
        {                       
            string last_name = HttpContext.Session.GetString("last_name");            
            string add_states = HttpContext.Session.GetString("additionalstates");            
            string case_number = HttpContext.Session.GetString("case_number");            
            string employer1State = HttpContext.Session.GetString("employer1State");                                                               
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
            if (employer1State.ToString().Equals(""))
            {
                doc.Replace("[Employer1State], ", "", true, true);
            }
            else
            {
                doc.Replace("[Employer1State]", employer1State, true, true);
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
            doc.Replace("[Last Name]", last_name.ToString(), true, true);
            doc.Replace("[FullEntityName]", last_name.ToString(), false, false);
            //doc.Replace("The United Kingdom", "United Kingdom", true, true);                      
            doc.SaveToFile(savePath);
            if (employer1State.ToString().ToUpper().Equals("NA") || employer1State.ToString().Equals(""))
            {
                doc.Replace(", [Employer1State]", "", true, false);
                doc.Replace("[Employer1State]", "", true, false);
            }
            else
            {
                doc.Replace("[Employer1State]", employer1State, true, false);
            }
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
            doc.Replace("the District of Columbia", "District of Columbia", false, false);
            doc.SaveToFile(savePath);
            doc.Replace("District of Columbia", "the District of Columbia", false, false);
            doc.SaveToFile(savePath);
            doc.Replace("U.S.  ", "U.S. ", true, false);
            doc.Replace("U.K.  ", "U.K. ", true, false);
            doc.SaveToFile(savePath);
            return RedirectToAction("GenerateFile", "Diligence");
        }
    }
}