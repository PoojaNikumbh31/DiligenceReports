using DiligenceReportCreation.Data;
using DiligenceReportCreation.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace DiligenceReportCreation.Controllers
{
    public class USPEIndividualController : Controller
    {
        private readonly DataBaseContext _context;
        private readonly IConfiguration _config;
        public USPEIndividualController(IConfiguration config, DataBaseContext context)
        {
            _config = config;
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public IActionResult US_PE_Individual_Edit()
        {
            string recordid = HttpContext.Session.GetString("recordid");
            MainModel main = new MainModel();
            DiligenceInputModel diligence2 = _context.DbPersonalInfo
                .Where(a => a.record_Id == recordid)
                .FirstOrDefault();
            Otherdatails otherdatails = _context.othersModel
                                  .Where(u => u.record_Id == recordid)
                                  .FirstOrDefault();
            main.additional_States = _context.Dbadditionalstates
                        .Where(u => u.record_Id == recordid)
                        .ToList();

            main.EmployerModel = _context.DbEmployer
                        .Where(u => u.record_Id == recordid)
                        .ToList();

            main.educationModels = _context.DbEducation
                       .Where(u => u.record_Id == recordid)
                       .ToList();

            main.pllicenseModels = _context.DbPLLicense
                       .Where(u => u.record_Id == recordid)
                        .ToList();

            main.summarymodel = _context.summaryResulttableModels
                .Where(u => u.record_Id == recordid)
                .FirstOrDefault();            
           
            main.diligenceInputModel = diligence2;
            main.otherdetails = otherdatails;
            HttpContext.Session.SetString("recordid", recordid);
            ReportModel report = _context.reportModel
                .Where(u => u.record_Id == recordid)
                                  .FirstOrDefault();
            ViewBag.CaseNumber = report.casenumber;
            ViewBag.LastName = report.lastname;                     
            return View(main);
        }
        [HttpPost]
        public IActionResult US_PE_Individual_Edit(MainModel mainModel, string SaveData, string Submit)
        {
            string recordid = HttpContext.Session.GetString("recordid");
            ReportModel report = _context.reportModel
              .Where(u => u.record_Id == recordid)
                                .FirstOrDefault();
            if (Submit == "SubmitData")
            {
                if (Submit == "SubmitData")
                {
                    MainModel main = new MainModel();
                    DiligenceInputModel dinput = new DiligenceInputModel();
                    Otherdatails otherdatails = new Otherdatails();
                    Additional_statesModel additional = new Additional_statesModel();
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
                    //try
                    //{
                    //    additional = _context.Dbadditionalstates
                    //                 .Where(u => u.record_Id == recordid)
                    //                 .FirstOrDefault();
                    //    if (additional == null)
                    //    {
                    //        TempData["message"] = "Please enter the details in additional states section";
                    //    }
                    //}
                    //catch
                    //{
                    //    TempData["message"] = "Please enter the details in additional states section";
                    //}
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
                        main.additional_States = _context.Dbadditionalstates
                                   .Where(u => u.record_Id == recordid)
                                    .ToList();
                        main.summarymodel = summary;
                        return Save_Page1(main);
                    }
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

                            DI1 = _context.DbPersonalInfo
                                .Where(a => a.record_Id == recordid)
                                .FirstOrDefault();
                            if (DI1 == null)
                            {
                                DI1 = new DiligenceInputModel();
                            }
                            DI1.record_Id = recordid.ToString();
                            DI1.ClientName = mainModel.diligenceInputModel.ClientName;
                            DI1.FirstName = mainModel.diligenceInputModel.FirstName;
                            DI1.LastName = report.lastname;
                            DI1.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                            DI1.CaseNumber = report.casenumber;
                            DI1.FirstName = mainModel.diligenceInputModel.FirstName;
                            DI1.MiddleName = mainModel.diligenceInputModel.MiddleName;
                            DI1.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                            DI1.MaidenName = mainModel.diligenceInputModel.MaidenName;
                            DI1.FullaliasName = mainModel.diligenceInputModel.FullaliasName;
                            DI1.Dob = mainModel.diligenceInputModel.Dob;
                            DI1.Nationality = mainModel.diligenceInputModel.Nationality;
                            DI1.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                            DI1.Natlidissuedate = mainModel.diligenceInputModel.Natlidissuedate;
                            DI1.Natlidexpirationdate = mainModel.diligenceInputModel.Natlidexpirationdate;
                            DI1.CurrentCity = mainModel.diligenceInputModel.CurrentCity;
                            DI1.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                            DI1.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                            DI1.SSNInitials = mainModel.diligenceInputModel.SSNInitials;
                            DI1.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
                            DI1.CommonNameSubject = mainModel.diligenceInputModel.CommonNameSubject;
                            DI1.CurrentState = mainModel.diligenceInputModel.CurrentState;
                            if (DI1 == null)
                            {
                                _context.DbPersonalInfo.Add(DI1);
                                _context.SaveChanges();
                            }
                            else
                            {
                                try
                                {
                                    _context.DbPersonalInfo.Update(DI1);
                                    _context.SaveChanges();
                                }
                                catch
                                {
                                    _context.DbPersonalInfo.Add(DI1);
                                    _context.SaveChanges();
                                }
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
                            Otherdatails strupdate = new Otherdatails();
                            strupdate = _context.othersModel
                               .Where(a => a.record_Id == recordid)
                               .FirstOrDefault();
                            if (strupdate == null)
                            {
                                strupdate = new Otherdatails();
                            }
                            strupdate.record_Id = recordid;
                            strupdate.CurrentResidentialProperty = mainModel.otherdetails.CurrentResidentialProperty;
                            strupdate.OtherCurrentResidentialProperty = mainModel.otherdetails.OtherCurrentResidentialProperty;
                            strupdate.OtherPropertyOwnershipinfo = mainModel.otherdetails.OtherPropertyOwnershipinfo;
                            strupdate.PrevPropertyOwnershipRes = mainModel.otherdetails.PrevPropertyOwnershipRes;
                            strupdate.HasBankruptcyRecHits = mainModel.otherdetails.HasBankruptcyRecHits;
                            strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;
                            strupdate.Has_CriminalRecHit = mainModel.otherdetails.Has_CriminalRecHit;
                            strupdate.Has_Bureau_PrisonHit = mainModel.otherdetails.Has_Bureau_PrisonHit;
                            strupdate.Has_Sex_Offender_RegHit = mainModel.otherdetails.Has_Sex_Offender_RegHit;
                            strupdate.Has_Civil_Records = mainModel.otherdetails.Has_Civil_Records;
                            strupdate.Has_US_Tax_Court_Hit = mainModel.otherdetails.Has_US_Tax_Court_Hit;
                            strupdate.Has_Tax_Liens = mainModel.otherdetails.Has_Tax_Liens;
                            strupdate.Executivesummary = mainModel.otherdetails.Executivesummary;
                            strupdate.Has_Driving_Hits = mainModel.otherdetails.Has_Driving_Hits;
                            strupdate.Was_credited_obtained = mainModel.otherdetails.Was_credited_obtained;
                            strupdate.Press_Media = mainModel.otherdetails.Press_Media;
                            strupdate.Fama = mainModel.otherdetails.Fama;
                            strupdate.RegulatoryFlag = mainModel.otherdetails.RegulatoryFlag;
                            strupdate.Has_Reg_FINRA = mainModel.otherdetails.Has_Reg_FINRA;
                            strupdate.Has_Reg_UK_FCA = mainModel.otherdetails.Has_Reg_UK_FCA;
                            strupdate.Has_Reg_US_NFA = mainModel.otherdetails.Has_Reg_US_NFA;
                            strupdate.Has_Reg_US_SEC = mainModel.otherdetails.Has_Reg_US_SEC;
                            strupdate.ICIJ_Hits = mainModel.otherdetails.ICIJ_Hits;
                            strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                            strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;                            
                            strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
                            try
                            {
                                if (mainModel.otherdetails.PatentType == null)
                                {
                                    strupdate.PatentType = "patent";
                                }
                                else
                                {
                                    strupdate.PatentType = mainModel.otherdetails.PatentType;
                                }
                            }
                            catch
                            {
                                strupdate.PatentType = "patent";
                            }
                            try
                            {
                                if (mainModel.otherdetails.ApplicantType == null)
                                {
                                    strupdate.ApplicantType = "Applicant";
                                }
                                else
                                {
                                    strupdate.ApplicantType = mainModel.otherdetails.ApplicantType;
                                }
                            }
                            catch
                            {
                                strupdate.ApplicantType = "Applicant";
                            }
                            strupdate.Has_Property_Records = mainModel.otherdetails.Has_Property_Records;
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
                            if (strupdate == null)
                            {
                                _context.othersModel.Add(strupdate);
                                _context.SaveChanges();
                            }
                            else
                            {
                                try
                                {
                                    _context.othersModel.Update(strupdate);
                                    _context.SaveChanges();
                                }
                                catch
                                {
                                    _context.othersModel.Add(strupdate);
                                    _context.SaveChanges();
                                }
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

                                if (employerModel1 == null || employerModel1.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.employerList.Count; i++)
                                    {
                                        EmployerModel employerModel = new EmployerModel();
                                        employerModel.record_Id = recordid;
                                        employerModel.Emp_Employer = mainModel.diligence.employerList[i].Emp_Employer;
                                        employerModel.Emp_City = mainModel.diligence.employerList[i].Emp_City;
                                        employerModel.Emp_State = mainModel.diligence.employerList[i].Emp_State;
                                        employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                                        employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                                        employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                                        if (mainModel.diligence.employerList[i].Emp_StartDateDay == "")
                                        {
                                            employerModel.Emp_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                                        }
                                        employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                                        if (mainModel.diligence.employerList[i].Emp_StartDateYear == "")
                                        {
                                            employerModel.Emp_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                                        }
                                        if (mainModel.diligence.employerList[i].Emp_EndDateDay == "")
                                        {
                                            employerModel.Emp_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                                        }
                                        employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                                        if (mainModel.diligence.employerList[i].Emp_EndDateYear == "")
                                        {
                                            employerModel.Emp_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
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
                                            employerModel1[i].Emp_City = mainModel.EmployerModel[i].Emp_City;
                                            employerModel1[i].Emp_State = mainModel.EmployerModel[i].Emp_State;
                                            employerModel1[i].Emp_Position = mainModel.EmployerModel[i].Emp_Position;
                                            employerModel1[i].Emp_Confirmed = mainModel.EmployerModel[i].Emp_Confirmed;
                                            employerModel1[i].Emp_Status = mainModel.EmployerModel[i].Emp_Status;
                                            if (mainModel.EmployerModel[i].Emp_StartDateDay == "") { employerModel1[i].Emp_StartDateDay = "Day"; }
                                            else
                                            {
                                                employerModel1[i].Emp_StartDateDay = mainModel.EmployerModel[i].Emp_StartDateDay;
                                            }
                                            employerModel1[i].Emp_StartDateMonth = mainModel.EmployerModel[i].Emp_StartDateMonth;
                                            if (mainModel.EmployerModel[i].Emp_StartDateYear == "")
                                            {
                                                employerModel1[i].Emp_StartDateYear = "Year";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_StartDateYear = mainModel.EmployerModel[i].Emp_StartDateYear;
                                            }
                                            if (mainModel.EmployerModel[i].Emp_EndDateDay == "") { employerModel1[i].Emp_EndDateDay = "Day"; }
                                            else
                                            {
                                                employerModel1[i].Emp_EndDateDay = mainModel.EmployerModel[i].Emp_EndDateDay;
                                            }
                                            employerModel1[i].Emp_EndDateMonth = mainModel.EmployerModel[i].Emp_EndDateMonth;
                                            if (mainModel.EmployerModel[i].Emp_EndDateYear == "")
                                            {
                                                employerModel1[i].Emp_EndDateYear = "Year";
                                            }
                                            else
                                            {
                                                employerModel1[i].Emp_EndDateYear = mainModel.EmployerModel[i].Emp_EndDateYear;
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
                                                    employerModel.Emp_City = mainModel.diligence.employerList[i].Emp_City;
                                                    employerModel.Emp_State = mainModel.diligence.employerList[i].Emp_State;
                                                    employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                                                    employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                                                    employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                                                    if (mainModel.diligence.employerList[i].Emp_StartDateDay == "")
                                                    {
                                                        employerModel.Emp_StartDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                                                    }
                                                    employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                                                    if (mainModel.diligence.employerList[i].Emp_StartDateYear == "")
                                                    {
                                                        employerModel.Emp_StartDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                                                    }
                                                    if (mainModel.diligence.employerList[i].Emp_EndDateDay == "")
                                                    {
                                                        employerModel.Emp_EndDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                                                    }
                                                    employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                                                    if (mainModel.diligence.employerList[i].Emp_EndDateYear == "")
                                                    {
                                                        employerModel.Emp_EndDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
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
                            employerModel.Emp_City = mainModel.diligence.employerList[i].Emp_City;
                            employerModel.Emp_State = mainModel.diligence.employerList[i].Emp_State;
                            employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                            employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                            employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                            if (mainModel.diligence.employerList[i].Emp_StartDateDay == "")
                            {
                                employerModel.Emp_StartDateDay = "Day";
                            }
                            else
                            {
                                employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                            }
                            employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                            if (mainModel.diligence.employerList[i].Emp_StartDateYear == "")
                            {
                                employerModel.Emp_StartDateYear = "Year";
                            }
                            else
                            {
                                employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                            }
                            if (mainModel.diligence.employerList[i].Emp_EndDateDay == "")
                            {
                                employerModel.Emp_EndDateDay = "Day";
                            }
                            else
                            {
                                employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                            }
                            employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                            if (mainModel.diligence.employerList[i].Emp_EndDateYear == "")
                            {
                                employerModel.Emp_EndDateYear = "Year";
                            }
                            else
                            {
                                employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
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
                                if (educationModel1 == null || educationModel1.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                    {
                                        EducationModel educationModel = new EducationModel();
                                        educationModel.record_Id = recordid;
                                        educationModel.Edu_Degree = mainModel.diligence.educationList[i].Edu_Degree;
                                        educationModel.Edu_Location = mainModel.diligence.educationList[i].Edu_Location;
                                        educationModel.Edu_School = mainModel.diligence.educationList[i].Edu_School;
                                        educationModel.Edu_Major = mainModel.diligence.educationList[i].Edu_Major;
                                        educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                        educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                        educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                        if (mainModel.diligence.educationList[i].Edu_GradDateDay == "")
                                        {
                                            educationModel.Edu_GradDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                        }
                                        if (mainModel.diligence.educationList[i].Edu_GradDateYear == "")
                                        {
                                            educationModel.Edu_GradDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                        }
                                        if (mainModel.diligence.educationList[i].Edu_StartDateDay == "")
                                        {
                                            educationModel.Edu_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay;
                                        }
                                        educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                        if (mainModel.diligence.educationList[i].Edu_StartDateYear == "")
                                        {
                                            educationModel.Edu_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                        }
                                        if (mainModel.diligence.educationList[i].Edu_EndDateDay == "")
                                        {
                                            educationModel.Edu_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                        }
                                        educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                        if (mainModel.diligence.educationList[i].Edu_EndDateYear == "")
                                        {
                                            educationModel.Edu_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                        }
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
                                        educationModel1[i].Edu_Major = mainModel.educationModels[i].Edu_Major;
                                        educationModel1[i].Edu_History = mainModel.educationModels[i].Edu_History;
                                        educationModel1[i].Edu_Confirmed = mainModel.educationModels[i].Edu_Confirmed;
                                        educationModel1[i].Edu_GradDateMonth = mainModel.educationModels[i].Edu_GradDateMonth;
                                        if (mainModel.educationModels[i].Edu_StartDateYear == "")
                                        {
                                            educationModel1[i].Edu_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_StartDateYear = mainModel.educationModels[i].Edu_StartDateYear;
                                        }
                                        if (mainModel.educationModels[i].Edu_GradDateYear == "")
                                        {
                                            educationModel1[i].Edu_GradDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_GradDateYear = mainModel.educationModels[i].Edu_GradDateYear;
                                        }
                                        if (mainModel.educationModels[i].Edu_EndDateYear == "")
                                        {
                                            educationModel1[i].Edu_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_EndDateYear = mainModel.educationModels[i].Edu_EndDateYear;
                                        }
                                        if (mainModel.educationModels[i].Edu_StartDateDay == "")
                                        {
                                            educationModel1[i].Edu_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_StartDateDay = mainModel.educationModels[i].Edu_StartDateDay;
                                        }
                                        if (mainModel.educationModels[i].Edu_GradDateDay == "")
                                        {
                                            educationModel1[i].Edu_GradDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_GradDateDay = mainModel.educationModels[i].Edu_GradDateDay;
                                        }
                                        if (mainModel.educationModels[i].Edu_EndDateDay == "")
                                        {
                                            educationModel1[i].Edu_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_EndDateDay = mainModel.educationModels[i].Edu_EndDateDay;
                                        }
                                        educationModel1[i].Edu_StartDateMonth = mainModel.educationModels[i].Edu_StartDateMonth;
                                        educationModel1[i].Edu_EndDateMonth = mainModel.educationModels[i].Edu_EndDateMonth;
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
                                                if (mainModel.diligence.educationList[i].Edu_Degree.ToString().Equals("") && mainModel.diligence.educationList[i].Edu_School.ToString().Equals("")) { }
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
                                                        educationModel.record_Id = recordid;
                                                        educationModel.Edu_Degree = mainModel.diligence.educationList[i].Edu_Degree;
                                                        educationModel.Edu_Location = mainModel.diligence.educationList[i].Edu_Location;
                                                        educationModel.Edu_School = mainModel.diligence.educationList[i].Edu_School;
                                                        educationModel.Edu_Major = mainModel.diligence.educationList[i].Edu_Major;
                                                        educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                                        educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                                        educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                                        if (mainModel.diligence.educationList[i].Edu_GradDateDay == "")
                                                        {
                                                            educationModel.Edu_GradDateDay = "Day";
                                                        }
                                                        else
                                                        {
                                                            educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                                        }
                                                        if (mainModel.diligence.educationList[i].Edu_GradDateYear == "")
                                                        {
                                                            educationModel.Edu_GradDateYear = "Year";
                                                        }
                                                        else
                                                        {
                                                            educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                                        }
                                                        if (mainModel.diligence.educationList[i].Edu_StartDateDay == "")
                                                        {
                                                            educationModel.Edu_StartDateDay = "Day";
                                                        }
                                                        else
                                                        {
                                                            educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay;
                                                        }
                                                        educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                                        if (mainModel.diligence.educationList[i].Edu_StartDateYear == "")
                                                        {
                                                            educationModel.Edu_StartDateYear = "Year";
                                                        }
                                                        else
                                                        {
                                                            educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                                        }
                                                        if (mainModel.diligence.educationList[i].Edu_EndDateDay == "")
                                                        {
                                                            educationModel.Edu_EndDateDay = "Day";
                                                        }
                                                        else
                                                        {
                                                            educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                                        }
                                                        educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                                        if (mainModel.diligence.educationList[i].Edu_EndDateYear == "")
                                                        {
                                                            educationModel.Edu_EndDateYear = "Year";
                                                        }
                                                        else
                                                        {
                                                            educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                                        }
                                                        educationModel.CreatedBy = Environment.UserName;
                                                        _context.DbEducation.Add(educationModel);
                                                        _context.SaveChanges();
                                                    }
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
                                    educationModel.Edu_Major = mainModel.diligence.educationList[i].Edu_Major;
                                    educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                    educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                    educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                    if (mainModel.diligence.educationList[i].Edu_GradDateDay == "")
                                    {
                                        educationModel.Edu_GradDateDay = "Day";
                                    }
                                    else
                                    {
                                        educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                    }
                                    if (mainModel.diligence.educationList[i].Edu_GradDateYear == "")
                                    {
                                        educationModel.Edu_GradDateYear = "Year";
                                    }
                                    else
                                    {
                                        educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                    }
                                    if (mainModel.diligence.educationList[i].Edu_StartDateDay == "")
                                    {
                                        educationModel.Edu_StartDateDay = "Day";
                                    }
                                    else
                                    {
                                        educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay;
                                    }
                                    educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                    if (mainModel.diligence.educationList[i].Edu_StartDateYear == "")
                                    {
                                        educationModel.Edu_StartDateYear = "Year";
                                    }
                                    else
                                    {
                                        educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                    }
                                    if (mainModel.diligence.educationList[i].Edu_EndDateDay == "")
                                    {
                                        educationModel.Edu_EndDateDay = "Day";
                                    }
                                    else
                                    {
                                        educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                    }
                                    educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                    if (mainModel.diligence.educationList[i].Edu_EndDateYear == "")
                                    {
                                        educationModel.Edu_EndDateYear = "Year";
                                    }
                                    else
                                    {
                                        educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                    }
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
                                        if (mainModel.diligence.plLicenseList[i].PL_StartDateDay == "")
                                        {
                                            pLLicenseModel1.PL_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            pLLicenseModel1.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                        }
                                        pLLicenseModel1.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                        if (mainModel.diligence.plLicenseList[i].PL_StartDateYear == "")
                                        {
                                            pLLicenseModel1.PL_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            pLLicenseModel1.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                        }
                                        if (mainModel.diligence.plLicenseList[i].PL_EndDateDay == "")
                                        {
                                            pLLicenseModel1.PL_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            pLLicenseModel1.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                        }
                                        if (mainModel.diligence.plLicenseList[i].PL_EndDateYear == "")
                                        {
                                            pLLicenseModel1.PL_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            pLLicenseModel1.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                        }
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
                                        if (mainModel.pllicenseModels[i].PL_StartDateDay == "")
                                        {
                                            pLLicenseModel[i].PL_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            pLLicenseModel[i].PL_StartDateDay = mainModel.pllicenseModels[i].PL_StartDateDay;
                                        }
                                        if (mainModel.pllicenseModels[i].PL_StartDateYear == "")
                                        {
                                            pLLicenseModel[i].PL_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            pLLicenseModel[i].PL_StartDateYear = mainModel.pllicenseModels[i].PL_StartDateYear;
                                        }
                                        pLLicenseModel[i].PL_StartDateMonth = mainModel.pllicenseModels[i].PL_StartDateMonth;
                                        if (mainModel.pllicenseModels[i].PL_EndDateDay == "")
                                        {
                                            pLLicenseModel[i].PL_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            pLLicenseModel[i].PL_EndDateDay = mainModel.pllicenseModels[i].PL_EndDateDay;
                                        }
                                        if (mainModel.pllicenseModels[i].PL_EndDateYear == "")
                                        {
                                            pLLicenseModel[i].PL_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            pLLicenseModel[i].PL_EndDateYear = mainModel.pllicenseModels[i].PL_EndDateYear;
                                        }
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
                                                    if (mainModel.diligence.plLicenseList[i].PL_StartDateDay == "")
                                                    {
                                                        pLLicenseModel1.PL_StartDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        pLLicenseModel1.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                                    }
                                                    pLLicenseModel1.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                                    if (mainModel.diligence.plLicenseList[i].PL_StartDateYear == "")
                                                    {
                                                        pLLicenseModel1.PL_StartDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        pLLicenseModel1.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                                    }
                                                    if (mainModel.diligence.plLicenseList[i].PL_EndDateDay == "")
                                                    {
                                                        pLLicenseModel1.PL_EndDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        pLLicenseModel1.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                                    }
                                                    if (mainModel.diligence.plLicenseList[i].PL_EndDateYear == "")
                                                    {
                                                        pLLicenseModel1.PL_EndDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        pLLicenseModel1.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                                    }
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
                                    if (mainModel.diligence.plLicenseList[i].PL_StartDateDay == "")
                                    {
                                        pLLicenseModel.PL_StartDateDay = "Day";
                                    }
                                    else
                                    {
                                        pLLicenseModel.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                    }
                                    pLLicenseModel.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                    if (mainModel.diligence.plLicenseList[i].PL_StartDateYear == "")
                                    {
                                        pLLicenseModel.PL_StartDateYear = "Year";
                                    }
                                    else
                                    {
                                        pLLicenseModel.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                    }
                                    if (mainModel.diligence.plLicenseList[i].PL_EndDateDay == "")
                                    {
                                        pLLicenseModel.PL_EndDateDay = "Day";
                                    }
                                    else
                                    {
                                        pLLicenseModel.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                    }
                                    if (mainModel.diligence.plLicenseList[i].PL_EndDateYear == "")
                                    {
                                        pLLicenseModel.PL_EndDateYear = "Year";
                                    }
                                    else
                                    {
                                        pLLicenseModel.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                    }
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
                                if (additionaL == null || additionaL.Count == 0) {
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
                                    Additional_statesModel additional = new Additional_statesModel();
                                    if (mainModel.diligence.additional_States[i].additionalstate.ToString().Equals("Select State")) { }
                                    else
                                    {
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
                            catch { resulttableModel = new SummaryResulttableModel(); }
                            resulttableModel.record_Id = recordid;
                            resulttableModel.Summarytable = mainModel.summarymodel.Summarytable;
                            resulttableModel.name_add_Ver_History = mainModel.summarymodel.name_add_Ver_History;
                            resulttableModel.personal_Identification = mainModel.summarymodel.personal_Identification;
                            resulttableModel.social_securitytrace = mainModel.summarymodel.social_securitytrace;
                            resulttableModel.real_estate_prop = mainModel.summarymodel.real_estate_prop;
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
                            if (resulttableModel == null)
                            {
                                _context.summaryResulttableModels.Add(resulttableModel);
                                _context.SaveChanges();
                            }
                            else
                            {
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
                    }
                    catch(IOException e)
                    {
                        TempData["message"] = e;
                        return RedirectToAction("LoginFormView", "Home");
                    }
                    break;
                default:
                    try
                    {
                        if (SaveData.Contains("Emp"))
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
                        if (SaveData.Contains("Edu"))
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
            HttpContext.Session.SetString("recordid", recordid);            
            ViewBag.CaseNumber = report.casenumber;
            ViewBag.LastName = report.lastname;
            ViewBag.CurrentState = mainModel.diligenceInputModel.CurrentState;
            return View(mainModel);
        }
        [HttpGet]
        public IActionResult US_PE_Individual_page1()
        {
            string recordid = HttpContext.Session.GetString("recordid");
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
        public IActionResult US_PE_Individual_page1(MainModel mainModel, string SaveData, string Submit)
        {
            string lastname = HttpContext.Session.GetString("lastname");
            string casenumber = HttpContext.Session.GetString("casenumber");
            string recordid = HttpContext.Session.GetString("recordid");
            ReportModel report = _context.reportModel
             .Where(u => u.record_Id == recordid)
                               .FirstOrDefault();
            if (Submit == "SubmitData")
            {
                if (Submit == "SubmitData")
                {
                    MainModel main = new MainModel();
                    DiligenceInputModel dinput = new DiligenceInputModel();
                    Otherdatails otherdatails = new Otherdatails();
                    Additional_statesModel additional = new Additional_statesModel();
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
                    //try
                    //{
                    //    additional = _context.Dbadditionalstates
                    //                 .Where(u => u.record_Id == recordid)
                    //                 .FirstOrDefault();
                    //    if (additional == null)
                    //    {
                    //        TempData["message"] = "Please enter the details in additional states section";
                    //    }
                    //}
                    //catch
                    //{
                    //    TempData["message"] = "Please enter the details in additional states section";
                    //}
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
                        main.additional_States = _context.Dbadditionalstates
                                   .Where(u => u.record_Id == recordid)
                                    .ToList();
                        main.summarymodel = summary;
                        return Save_Page1(main);                        
                    }
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

                            DI1 = _context.DbPersonalInfo
                                .Where(a => a.record_Id == recordid)
                                .FirstOrDefault();
                            if (DI1 == null)
                            {
                                DI1 = new DiligenceInputModel();
                            }
                            DI1.record_Id = recordid.ToString();
                            DI1.ClientName = mainModel.diligenceInputModel.ClientName;
                            DI1.FirstName = mainModel.diligenceInputModel.FirstName;
                            DI1.LastName = report.lastname;
                            DI1.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                            DI1.CaseNumber = report.casenumber;
                            DI1.FirstName = mainModel.diligenceInputModel.FirstName;
                            DI1.MiddleName = mainModel.diligenceInputModel.MiddleName;
                            DI1.MiddleInitial = mainModel.diligenceInputModel.MiddleInitial;
                            DI1.MaidenName = mainModel.diligenceInputModel.MaidenName;
                            DI1.FullaliasName = mainModel.diligenceInputModel.FullaliasName;
                            DI1.Dob = mainModel.diligenceInputModel.Dob;
                            DI1.Nationality = mainModel.diligenceInputModel.Nationality;
                            DI1.Nationalidnumber = mainModel.diligenceInputModel.Nationalidnumber;
                            DI1.Natlidissuedate = mainModel.diligenceInputModel.Natlidissuedate;
                            DI1.Natlidexpirationdate = mainModel.diligenceInputModel.Natlidexpirationdate;
                            DI1.CurrentCity = mainModel.diligenceInputModel.CurrentCity;
                            DI1.CurrentStreet = mainModel.diligenceInputModel.CurrentStreet;
                            DI1.CurrentZipcode = mainModel.diligenceInputModel.CurrentZipcode;
                            DI1.SSNInitials = mainModel.diligenceInputModel.SSNInitials;
                            DI1.Nonscopecountry1 = mainModel.diligenceInputModel.Nonscopecountry1;
                            DI1.CommonNameSubject = mainModel.diligenceInputModel.CommonNameSubject;
                            DI1.CurrentState = mainModel.diligenceInputModel.CurrentState;
                            if (DI1 == null)
                            {
                                _context.DbPersonalInfo.Add(DI1);
                                _context.SaveChanges();
                            }
                            else
                            {
                                try
                                {
                                    _context.DbPersonalInfo.Update(DI1);
                                    _context.SaveChanges();
                                }
                                catch
                                {
                                    _context.DbPersonalInfo.Add(DI1);
                                    _context.SaveChanges();
                                }
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
                            Otherdatails strupdate = new Otherdatails();
                            strupdate = _context.othersModel
                               .Where(a => a.record_Id == recordid)
                               .FirstOrDefault();
                            if (strupdate == null)
                            {
                                strupdate = new Otherdatails();
                            }
                            strupdate.record_Id = recordid;
                            strupdate.CurrentResidentialProperty = mainModel.otherdetails.CurrentResidentialProperty;
                            strupdate.OtherCurrentResidentialProperty = mainModel.otherdetails.OtherCurrentResidentialProperty;
                            strupdate.OtherPropertyOwnershipinfo = mainModel.otherdetails.OtherPropertyOwnershipinfo;
                            strupdate.PrevPropertyOwnershipRes = mainModel.otherdetails.PrevPropertyOwnershipRes;
                            strupdate.HasBankruptcyRecHits = mainModel.otherdetails.HasBankruptcyRecHits;
                            strupdate.Registered_with_HKSFC = mainModel.otherdetails.Registered_with_HKSFC;
                            strupdate.Has_CriminalRecHit = mainModel.otherdetails.Has_CriminalRecHit;
                            strupdate.Has_Bureau_PrisonHit = mainModel.otherdetails.Has_Bureau_PrisonHit;
                            strupdate.Has_Sex_Offender_RegHit = mainModel.otherdetails.Has_Sex_Offender_RegHit;
                            strupdate.Has_Civil_Records = mainModel.otherdetails.Has_Civil_Records;
                            strupdate.Has_US_Tax_Court_Hit = mainModel.otherdetails.Has_US_Tax_Court_Hit;
                            strupdate.Has_Tax_Liens = mainModel.otherdetails.Has_Tax_Liens;
                            strupdate.Executivesummary = mainModel.otherdetails.Executivesummary;
                            strupdate.Has_Driving_Hits = mainModel.otherdetails.Has_Driving_Hits;
                            strupdate.Was_credited_obtained = mainModel.otherdetails.Was_credited_obtained;
                            strupdate.Press_Media = mainModel.otherdetails.Press_Media;
                            strupdate.Fama = mainModel.otherdetails.Fama;
                            strupdate.RegulatoryFlag = mainModel.otherdetails.RegulatoryFlag;
                            strupdate.Has_Reg_FINRA = mainModel.otherdetails.Has_Reg_FINRA;
                            strupdate.Has_Reg_UK_FCA = mainModel.otherdetails.Has_Reg_UK_FCA;
                            strupdate.Has_Reg_US_NFA = mainModel.otherdetails.Has_Reg_US_NFA;
                            strupdate.Has_Reg_US_SEC = mainModel.otherdetails.Has_Reg_US_SEC;
                            strupdate.ICIJ_Hits = mainModel.otherdetails.ICIJ_Hits;
                            strupdate.Global_Security_Hits = mainModel.otherdetails.Global_Security_Hits;
                            strupdate.Has_Business_Affiliations = mainModel.otherdetails.Has_Business_Affiliations;
                            strupdate.Has_Intellectual_Hits = mainModel.otherdetails.Has_Intellectual_Hits;
                            try
                            {
                                if (mainModel.otherdetails.PatentType == null)
                                {
                                    strupdate.PatentType = "patent";
                                }
                                else
                                {
                                    strupdate.PatentType = mainModel.otherdetails.PatentType;
                                }
                            }
                            catch
                            {
                                strupdate.PatentType = "patent";
                            }
                            try
                            {
                                if (mainModel.otherdetails.ApplicantType == null)
                                {
                                    strupdate.ApplicantType = "Applicant";
                                }
                                else
                                {
                                    strupdate.ApplicantType = mainModel.otherdetails.ApplicantType;
                                }
                            }
                            catch
                            {
                                strupdate.ApplicantType = "Applicant";
                            }
                            strupdate.Has_Property_Records = mainModel.otherdetails.Has_Property_Records;
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
                            if (strupdate == null)
                            {
                                _context.othersModel.Add(strupdate);
                                _context.SaveChanges();
                            }
                            else
                            {
                                try
                                {
                                    _context.othersModel.Update(strupdate);
                                    _context.SaveChanges();
                                }
                                catch
                                {
                                    _context.othersModel.Add(strupdate);
                                    _context.SaveChanges();
                                }
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

                                if (employerModel1 == null || employerModel1.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.employerList.Count; i++)
                                    {
                                        EmployerModel employerModel = new EmployerModel();
                                        employerModel.record_Id = recordid;
                                        employerModel.Emp_Employer = mainModel.diligence.employerList[i].Emp_Employer;
                                        employerModel.Emp_City = mainModel.diligence.employerList[i].Emp_City;
                                        employerModel.Emp_State = mainModel.diligence.employerList[i].Emp_State;
                                        employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                                        employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                                        employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                                        if (mainModel.diligence.employerList[i].Emp_StartDateDay == "")
                                        {
                                            employerModel.Emp_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                                        }
                                        employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                                        if (mainModel.diligence.employerList[i].Emp_StartDateYear == "")
                                        {
                                            employerModel.Emp_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                                        }
                                        if (mainModel.diligence.employerList[i].Emp_EndDateDay == "") {
                                            employerModel.Emp_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                                        }
                                        employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                                        if (mainModel.diligence.employerList[i].Emp_EndDateYear == "") {
                                            employerModel.Emp_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
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
                                        employerModel1[i].Emp_City = mainModel.EmployerModel[i].Emp_City;
                                        employerModel1[i].Emp_State = mainModel.EmployerModel[i].Emp_State;
                                        employerModel1[i].Emp_Position = mainModel.EmployerModel[i].Emp_Position;
                                        employerModel1[i].Emp_Confirmed = mainModel.EmployerModel[i].Emp_Confirmed;
                                        employerModel1[i].Emp_Status = mainModel.EmployerModel[i].Emp_Status;
                                        if (mainModel.EmployerModel[i].Emp_StartDateDay == "") { employerModel1[i].Emp_StartDateDay = "Day"; }
                                        else
                                        {
                                            employerModel1[i].Emp_StartDateDay = mainModel.EmployerModel[i].Emp_StartDateDay;
                                        }
                                        employerModel1[i].Emp_StartDateMonth = mainModel.EmployerModel[i].Emp_StartDateMonth;
                                        if (mainModel.EmployerModel[i].Emp_StartDateYear == "") {
                                            employerModel1[i].Emp_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            employerModel1[i].Emp_StartDateYear = mainModel.EmployerModel[i].Emp_StartDateYear;
                                        }
                                        if (mainModel.EmployerModel[i].Emp_EndDateDay == "") { employerModel1[i].Emp_EndDateDay = "Day"; }
                                        else
                                        {
                                            employerModel1[i].Emp_EndDateDay = mainModel.EmployerModel[i].Emp_EndDateDay;
                                        }                                        
                                        employerModel1[i].Emp_EndDateMonth = mainModel.EmployerModel[i].Emp_EndDateMonth;
                                        if (mainModel.EmployerModel[i].Emp_EndDateYear == "")
                                        {
                                            employerModel1[i].Emp_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            employerModel1[i].Emp_EndDateYear = mainModel.EmployerModel[i].Emp_EndDateYear;
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
                                                    employerModel.Emp_City = mainModel.diligence.employerList[i].Emp_City;
                                                    employerModel.Emp_State = mainModel.diligence.employerList[i].Emp_State;
                                                    employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                                                    employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                                                    employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                                                    if (mainModel.diligence.employerList[i].Emp_StartDateDay == "")
                                                    {
                                                        employerModel.Emp_StartDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                                                    }
                                                    employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                                                    if (mainModel.diligence.employerList[i].Emp_StartDateYear == "")
                                                    {
                                                        employerModel.Emp_StartDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                                                    }
                                                    if (mainModel.diligence.employerList[i].Emp_EndDateDay == "")
                                                    {
                                                        employerModel.Emp_EndDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                                                    }
                                                    employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                                                    if (mainModel.diligence.employerList[i].Emp_EndDateYear == "")
                                                    {
                                                        employerModel.Emp_EndDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
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
                                    employerModel.Emp_City = mainModel.diligence.employerList[i].Emp_City;
                                    employerModel.Emp_State = mainModel.diligence.employerList[i].Emp_State;
                                    employerModel.Emp_Position = mainModel.diligence.employerList[i].Emp_Position;
                                    employerModel.Emp_Confirmed = mainModel.diligence.employerList[i].Emp_Confirmed;
                                    employerModel.Emp_Status = mainModel.diligence.employerList[i].Emp_Status;
                                    if (mainModel.diligence.employerList[i].Emp_StartDateDay == "")
                                    {
                                        employerModel.Emp_StartDateDay = "Day";
                                    }
                                    else
                                    {
                                        employerModel.Emp_StartDateDay = mainModel.diligence.employerList[i].Emp_StartDateDay;
                                    }
                                    employerModel.Emp_StartDateMonth = mainModel.diligence.employerList[i].Emp_StartDateMonth;
                                    if (mainModel.diligence.employerList[i].Emp_StartDateYear == "")
                                    {
                                        employerModel.Emp_StartDateYear = "Year";
                                    }
                                    else
                                    {
                                        employerModel.Emp_StartDateYear = mainModel.diligence.employerList[i].Emp_StartDateYear;
                                    }
                                    if (mainModel.diligence.employerList[i].Emp_EndDateDay == "")
                                    {
                                        employerModel.Emp_EndDateDay = "Day";
                                    }
                                    else
                                    {
                                        employerModel.Emp_EndDateDay = mainModel.diligence.employerList[i].Emp_EndDateDay;
                                    }
                                    employerModel.Emp_EndDateMonth = mainModel.diligence.employerList[i].Emp_EndDateMonth;
                                    if (mainModel.diligence.employerList[i].Emp_EndDateYear == "")
                                    {
                                        employerModel.Emp_EndDateYear = "Year";
                                    }
                                    else
                                    {
                                        employerModel.Emp_EndDateYear = mainModel.diligence.employerList[i].Emp_EndDateYear;
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
                                if (educationModel1 == null || educationModel1.Count == 0)
                                {
                                    for (int i = 0; i < mainModel.diligence.educationList.Count; i++)
                                    {
                                        EducationModel educationModel = new EducationModel();
                                        educationModel.record_Id = recordid;
                                        educationModel.Edu_Degree = mainModel.diligence.educationList[i].Edu_Degree;
                                        educationModel.Edu_Location = mainModel.diligence.educationList[i].Edu_Location;
                                        educationModel.Edu_School = mainModel.diligence.educationList[i].Edu_School;
                                        educationModel.Edu_Major = mainModel.diligence.educationList[i].Edu_Major;
                                        educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                        educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                        educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                        if (mainModel.diligence.educationList[i].Edu_GradDateDay == "") {
                                            educationModel.Edu_GradDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                        }
                                        if (mainModel.diligence.educationList[i].Edu_GradDateYear == "") {
                                            educationModel.Edu_GradDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                        }
                                        if (mainModel.diligence.educationList[i].Edu_StartDateDay == "")
                                        {
                                            educationModel.Edu_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay ;
                                        }                                        
                                        educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                        if (mainModel.diligence.educationList[i].Edu_StartDateYear == "")
                                        {
                                            educationModel.Edu_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                        }                                       
                                        if (mainModel.diligence.educationList[i].Edu_EndDateDay == "")
                                        {
                                            educationModel.Edu_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                        }                                        
                                        educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                        if (mainModel.diligence.educationList[i].Edu_EndDateYear == "")
                                        {
                                            educationModel.Edu_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                        }                                        
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
                                        educationModel1[i].Edu_Major = mainModel.educationModels[i].Edu_Major;
                                        educationModel1[i].Edu_History = mainModel.educationModels[i].Edu_History;
                                        educationModel1[i].Edu_Confirmed = mainModel.educationModels[i].Edu_Confirmed;
                                        educationModel1[i].Edu_GradDateMonth = mainModel.educationModels[i].Edu_GradDateMonth;                                        
                                        if (mainModel.educationModels[i].Edu_StartDateYear == "")
                                        {
                                            educationModel1[i].Edu_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_StartDateYear = mainModel.educationModels[i].Edu_StartDateYear;
                                        }
                                        if (mainModel.educationModels[i].Edu_GradDateYear == "")
                                        {
                                            educationModel1[i].Edu_GradDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_GradDateYear = mainModel.educationModels[i].Edu_GradDateYear;
                                        }
                                        if (mainModel.educationModels[i].Edu_EndDateYear == "")
                                        {
                                            educationModel1[i].Edu_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_EndDateYear = mainModel.educationModels[i].Edu_EndDateYear;
                                        }
                                        if (mainModel.educationModels[i].Edu_StartDateDay == "")
                                        {
                                            educationModel1[i].Edu_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_StartDateDay = mainModel.educationModels[i].Edu_StartDateDay;
                                        }
                                        if (mainModel.educationModels[i].Edu_GradDateDay == "")
                                        {
                                            educationModel1[i].Edu_GradDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_GradDateDay = mainModel.educationModels[i].Edu_GradDateDay;
                                        }
                                        if (mainModel.educationModels[i].Edu_EndDateDay == "")
                                        {
                                            educationModel1[i].Edu_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            educationModel1[i].Edu_EndDateDay = mainModel.educationModels[i].Edu_EndDateDay;
                                        }
                                        educationModel1[i].Edu_StartDateMonth = mainModel.educationModels[i].Edu_StartDateMonth;                                                                                
                                        educationModel1[i].Edu_EndDateMonth = mainModel.educationModels[i].Edu_EndDateMonth;                                        
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
                                                    educationModel.record_Id = recordid;
                                                    educationModel.Edu_Degree = mainModel.diligence.educationList[i].Edu_Degree;
                                                    educationModel.Edu_Location = mainModel.diligence.educationList[i].Edu_Location;
                                                    educationModel.Edu_School = mainModel.diligence.educationList[i].Edu_School;
                                                    educationModel.Edu_Major = mainModel.diligence.educationList[i].Edu_Major;
                                                    educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                                    educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                                    educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                                    if (mainModel.diligence.educationList[i].Edu_GradDateDay == "")
                                                    {
                                                        educationModel.Edu_GradDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                                    }
                                                    if (mainModel.diligence.educationList[i].Edu_GradDateYear == "")
                                                    {
                                                        educationModel.Edu_GradDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                                    }
                                                    if (mainModel.diligence.educationList[i].Edu_StartDateDay == "")
                                                    {
                                                        educationModel.Edu_StartDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay;
                                                    }
                                                    educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                                    if (mainModel.diligence.educationList[i].Edu_StartDateYear == "")
                                                    {
                                                        educationModel.Edu_StartDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                                    }
                                                    if (mainModel.diligence.educationList[i].Edu_EndDateDay == "")
                                                    {
                                                        educationModel.Edu_EndDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                                    }
                                                    educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                                    if (mainModel.diligence.educationList[i].Edu_EndDateYear == "")
                                                    {
                                                        educationModel.Edu_EndDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                                    }
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
                                    educationModel.Edu_Major = mainModel.diligence.educationList[i].Edu_Major;
                                    educationModel.Edu_History = mainModel.diligence.educationList[i].Edu_History;
                                    educationModel.Edu_Confirmed = mainModel.diligence.educationList[i].Edu_Confirmed;
                                    educationModel.Edu_GradDateMonth = mainModel.diligence.educationList[i].Edu_GradDateMonth;
                                    if (mainModel.diligence.educationList[i].Edu_GradDateDay == "")
                                    {
                                        educationModel.Edu_GradDateDay = "Day";
                                    }
                                    else
                                    {
                                        educationModel.Edu_GradDateDay = mainModel.diligence.educationList[i].Edu_GradDateDay;
                                    }
                                    if (mainModel.diligence.educationList[i].Edu_GradDateYear == "")
                                    {
                                        educationModel.Edu_GradDateYear = "Year";
                                    }
                                    else
                                    {
                                        educationModel.Edu_GradDateYear = mainModel.diligence.educationList[i].Edu_GradDateYear;
                                    }
                                    if (mainModel.diligence.educationList[i].Edu_StartDateDay == "")
                                    {
                                        educationModel.Edu_StartDateDay = "Day";
                                    }
                                    else
                                    {
                                        educationModel.Edu_StartDateDay = mainModel.diligence.educationList[i].Edu_StartDateDay;
                                    }
                                    educationModel.Edu_StartDateMonth = mainModel.diligence.educationList[i].Edu_StartDateMonth;
                                    if (mainModel.diligence.educationList[i].Edu_StartDateYear == "")
                                    {
                                        educationModel.Edu_StartDateYear = "Year";
                                    }
                                    else
                                    {
                                        educationModel.Edu_StartDateYear = mainModel.diligence.educationList[i].Edu_StartDateYear;
                                    }
                                    if (mainModel.diligence.educationList[i].Edu_EndDateDay == "")
                                    {
                                        educationModel.Edu_EndDateDay = "Day";
                                    }
                                    else
                                    {
                                        educationModel.Edu_EndDateDay = mainModel.diligence.educationList[i].Edu_EndDateDay;
                                    }
                                    educationModel.Edu_EndDateMonth = mainModel.diligence.educationList[i].Edu_EndDateMonth;
                                    if (mainModel.diligence.educationList[i].Edu_EndDateYear == "")
                                    {
                                        educationModel.Edu_EndDateYear = "Year";
                                    }
                                    else
                                    {
                                        educationModel.Edu_EndDateYear = mainModel.diligence.educationList[i].Edu_EndDateYear;
                                    }
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
                                        if (mainModel.diligence.plLicenseList[i].PL_StartDateDay == "") {
                                            pLLicenseModel1.PL_StartDateDay = "Day";
                                        }
                                        else
                                        {
                                            pLLicenseModel1.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                        }
                                        pLLicenseModel1.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                        if (mainModel.diligence.plLicenseList[i].PL_StartDateYear == "")
                                        {
                                            pLLicenseModel1.PL_StartDateYear = "Year";
                                        }
                                        else
                                        {
                                            pLLicenseModel1.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                        }                                        
                                        if (mainModel.diligence.plLicenseList[i].PL_EndDateDay == "")
                                        {
                                            pLLicenseModel1.PL_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            pLLicenseModel1.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                        }
                                        if (mainModel.diligence.plLicenseList[i].PL_EndDateYear == "")
                                        {
                                            pLLicenseModel1.PL_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            pLLicenseModel1.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                        }                                        
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
                                        if (mainModel.pllicenseModels[i].PL_StartDateDay == "") {
                                            pLLicenseModel[i].PL_StartDateDay = "Day";
                                        } else
                                        {
                                            pLLicenseModel[i].PL_StartDateDay = mainModel.pllicenseModels[i].PL_StartDateDay;
                                        }
                                        if (mainModel.pllicenseModels[i].PL_StartDateYear == "") {
                                            pLLicenseModel[i].PL_StartDateYear = "Year";
                                        } else {
                                            pLLicenseModel[i].PL_StartDateYear = mainModel.pllicenseModels[i].PL_StartDateYear;
                                        }
                                        pLLicenseModel[i].PL_StartDateMonth = mainModel.pllicenseModels[i].PL_StartDateMonth;
                                        if (mainModel.pllicenseModels[i].PL_EndDateDay == "") {
                                            pLLicenseModel[i].PL_EndDateDay = "Day";
                                        }
                                        else
                                        {
                                            pLLicenseModel[i].PL_EndDateDay = mainModel.pllicenseModels[i].PL_EndDateDay;
                                        }
                                        if (mainModel.pllicenseModels[i].PL_EndDateYear == "") {
                                            pLLicenseModel[i].PL_EndDateYear = "Year";
                                        }
                                        else
                                        {
                                            pLLicenseModel[i].PL_EndDateYear = mainModel.pllicenseModels[i].PL_EndDateYear;
                                        }
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
                                                    if (mainModel.diligence.plLicenseList[i].PL_StartDateDay == "")
                                                    {
                                                        pLLicenseModel1.PL_StartDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        pLLicenseModel1.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                                    }
                                                    pLLicenseModel1.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                                    if (mainModel.diligence.plLicenseList[i].PL_StartDateYear == "")
                                                    {
                                                        pLLicenseModel1.PL_StartDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        pLLicenseModel1.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                                    }
                                                    if (mainModel.diligence.plLicenseList[i].PL_EndDateDay == "")
                                                    {
                                                        pLLicenseModel1.PL_EndDateDay = "Day";
                                                    }
                                                    else
                                                    {
                                                        pLLicenseModel1.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                                    }
                                                    if (mainModel.diligence.plLicenseList[i].PL_EndDateYear == "")
                                                    {
                                                        pLLicenseModel1.PL_EndDateYear = "Year";
                                                    }
                                                    else
                                                    {
                                                        pLLicenseModel1.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                                    }
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
                                    if (mainModel.diligence.plLicenseList[i].PL_StartDateDay == "")
                                    {
                                        pLLicenseModel.PL_StartDateDay = "Day";
                                    }
                                    else
                                    {
                                        pLLicenseModel.PL_StartDateDay = mainModel.diligence.plLicenseList[i].PL_StartDateDay;
                                    }
                                    pLLicenseModel.PL_StartDateMonth = mainModel.diligence.plLicenseList[i].PL_StartDateMonth;
                                    if (mainModel.diligence.plLicenseList[i].PL_StartDateYear == "")
                                    {
                                        pLLicenseModel.PL_StartDateYear = "Year";
                                    }
                                    else
                                    {
                                        pLLicenseModel.PL_StartDateYear = mainModel.diligence.plLicenseList[i].PL_StartDateYear;
                                    }
                                    if (mainModel.diligence.plLicenseList[i].PL_EndDateDay == "")
                                    {
                                        pLLicenseModel.PL_EndDateDay = "Day";
                                    }
                                    else
                                    {
                                        pLLicenseModel.PL_EndDateDay = mainModel.diligence.plLicenseList[i].PL_EndDateDay;
                                    }
                                    if (mainModel.diligence.plLicenseList[i].PL_EndDateYear == "")
                                    {
                                        pLLicenseModel.PL_EndDateYear = "Year";
                                    }
                                    else
                                    {
                                        pLLicenseModel.PL_EndDateYear = mainModel.diligence.plLicenseList[i].PL_EndDateYear;
                                    }
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
                                    Additional_statesModel additional = new Additional_statesModel();
                                    if (mainModel.diligence.additional_States[i].additionalstate.ToString().Equals("Select State")) { }
                                    else
                                    {
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
                            catch { resulttableModel = new SummaryResulttableModel(); }
                            resulttableModel.record_Id = recordid;
                            resulttableModel.Summarytable = mainModel.summarymodel.Summarytable;
                            resulttableModel.name_add_Ver_History = mainModel.summarymodel.name_add_Ver_History;
                            resulttableModel.personal_Identification = mainModel.summarymodel.personal_Identification;
                            resulttableModel.social_securitytrace = mainModel.summarymodel.social_securitytrace;
                            resulttableModel.real_estate_prop = mainModel.summarymodel.real_estate_prop;
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
                            if (resulttableModel == null)
                            {
                                _context.summaryResulttableModels.Add(resulttableModel);
                                _context.SaveChanges();
                            }
                            else
                            {
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
                        if (SaveData.Contains("Emp"))
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
                        if (SaveData.Contains("Edu"))
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
            HttpContext.Session.SetString("lastname", lastname);
            HttpContext.Session.SetString("casenumber", casenumber);
            HttpContext.Session.SetString("recordid", recordid);
            ViewBag.CaseNumber = casenumber;
            ViewBag.LastName = lastname;
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
        public IActionResult Save_Page1(MainModel uSPEIndividual)
        {
            string record_id = HttpContext.Session.GetString("record_Id");
            string templatePath;
            string savePath = _config.GetValue<string>("ReportPath");
            templatePath = _config.GetValue<string>("USPEIndividualtemplatePath");
            if (uSPEIndividual.summarymodel.Summarytable.ToString().Equals("Yes") && uSPEIndividual.otherdetails.Executivesummary.ToString().Equals("Yes"))
            {
                if (uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("California"))
                {
                    templatePath = string.Concat(templatePath, "US_PE_Individual_SterlingDiligenceReport(XXXXX)_DRAFT5.docx");
                }
                else
                {
                    templatePath = string.Concat(templatePath, "US_PE_Individual_SterlingDiligenceReport(XXXXX)_DRAFT1.docx");
                }
            }
            else
            {
                if (uSPEIndividual.summarymodel.Summarytable.ToString().Equals("No") && uSPEIndividual.otherdetails.Executivesummary.ToString().Equals("No"))
                {
                    if (uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("California"))
                    {
                        templatePath = string.Concat(templatePath, "US_PE_Individual_SterlingDiligenceReport(XXXXX)_DRAFT7.docx");
                    }
                    else
                    {
                        templatePath = string.Concat(templatePath, "US_PE_Individual_SterlingDiligenceReport(XXXXX)_DRAFT3.docx");
                    }
                }
                else
                {
                    if (uSPEIndividual.summarymodel.Summarytable.ToString().Equals("Yes") && uSPEIndividual.otherdetails.Executivesummary.ToString().Equals("No"))
                    {
                        if (uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("California"))
                        {
                            templatePath = string.Concat(templatePath, "US_PE_Individual_SterlingDiligenceReport(XXXXX)_DRAFT6.docx");
                        }
                        else
                        {
                            templatePath = string.Concat(templatePath, "US_PE_Individual_SterlingDiligenceReport(XXXXX)_DRAFT4.docx");
                        }
                    }
                    if (uSPEIndividual.summarymodel.Summarytable.ToString().Equals("No") && uSPEIndividual.otherdetails.Executivesummary.ToString().Equals("Yes"))
                    {
                        if (uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("California"))
                        {
                            templatePath = string.Concat(templatePath, "US_PE_Individual_SterlingDiligenceReport(XXXXX)_DRAFT8.docx");
                        }
                        else
                        {
                            templatePath = string.Concat(templatePath, "US_PE_Individual_SterlingDiligenceReport(XXXXX)_DRAFT2.docx");
                        }
                    }
                }
            }            
            string pathTo = _config.GetValue<string>("OlderReport"); // the destination file name would be appended later
            savePath = string.Concat(savePath, uSPEIndividual.diligenceInputModel.LastName.ToString(), "_SterlingDiligenceReport(", uSPEIndividual.diligenceInputModel.CaseNumber.ToString(), ")_DRAFT.docx");
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
                var newExtension = GetNewFileExtension(files); // will return .1, .2,....N             
                fileInfo.MoveTo(Path.Combine(pathTo, string.Format("{0}{1}", filename, newExtension)));
            }
            Document doc = new Document(templatePath);
            string strblnres = "";
            int sectionLimit = 0;
            string strfn = "";
            if (uSPEIndividual.summarymodel.Summarytable.ToString().Equals("No"))
            {
                sectionLimit = 5;
            }
            else
            {
                sectionLimit = 7;
            }
            if (uSPEIndividual.diligenceInputModel.FullaliasName.ToString().Equals("") || uSPEIndividual.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("NA") || uSPEIndividual.diligenceInputModel.FullaliasName.ToString().ToUpper().Equals("N/A"))
            {
                if (uSPEIndividual.diligenceInputModel.CommonNameSubject == true)
                {
                    doc.Replace(" (also known as [FullAliasName])", "[ADDCommannamecheckboxFN]", true, false);
                    doc.SaveToFile(savePath);
                    strfn = "Due to the commonality of [LastName]'s name, research efforts were focused to the subject's known jurisdictions and timeframe of [LastName]'s affiliation with the same.";
                }
                else
                {
                    doc.Replace("(also known as [FullAliasName])", "", true, false);
                }
            }
            else
            {
                             
                if (uSPEIndividual.diligenceInputModel.CommonNameSubject == true)
                {
                    strfn = "It is noted that the subject was identified in connection with the alternate name variations of [FullAliasName] and investigative efforts were undertaken in connection with the same, as appropriate. Additionally, due to the commonality of [LastName]'s name, research efforts were focused to the subject's known jurisdictions and timeframe of [LastName]'s affiliation with the same.";
                }
                else
                {
                    strfn = "It is noted that the subject was identified in connection with the alternate name variations of [FullAliasName] and investigative efforts were undertaken in connection with the same, as appropriate.";
                }
                doc.Replace("[FullAliasName])", "[FullAliasName])", true, false);
             
            }
            try
            {
                for (int j = 2; j < sectionLimit; j++)
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
                                if (abc.ToString().Equals("[FullAliasName]") || abc.ToString().EndsWith("([FullAliasName])") || abc.ToString().Contains("FullAliasName") || abc.ToString().EndsWith("[ADDCommannamecheckboxFN]") || abc.ToString().Contains("[ADDCommannamecheckboxFN]"))
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
            }
            catch { }
            doc.Replace("[ADDCommannamecheckboxFN]", "", false, false);
            doc.Replace("[FullAliasName]", uSPEIndividual.diligenceInputModel.FullaliasName, true, false);
            doc.Replace("First_Name", uSPEIndividual.diligenceInputModel.FirstName.ToUpper().ToString().TrimEnd(), true, true);           
            doc.Replace("Last_Name", uSPEIndividual.diligenceInputModel.LastName.ToUpper().ToString().TrimEnd(), true, true);
            doc.Replace("ClientName", uSPEIndividual.diligenceInputModel.ClientName.TrimEnd(), true, true);
            doc.Replace("Position1", uSPEIndividual.EmployerModel[0].Emp_Position.TrimEnd(), true, true);
            doc.Replace("Employer1", uSPEIndividual.EmployerModel[0].Emp_Employer.ToString().TrimEnd(), true, true);
            doc.Replace("ClientName", uSPEIndividual.diligenceInputModel.ClientName, true, true);
            try
            {
                doc.Replace("[DateofBirth]", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(uSPEIndividual.diligenceInputModel.Dob.ToString()).Month) + " " + Convert.ToDateTime(uSPEIndividual.diligenceInputModel.Dob.ToString()).Year, true, true);
            }
            catch
            {
                doc.Replace("[DateofBirth]", "<not provided>", true, true);
            }
            doc.Replace("[Initial_digits_of_SSN]", uSPEIndividual.diligenceInputModel.SSNInitials.ToString(), true, true);
            doc.Replace("[CaseNumber]", uSPEIndividual.diligenceInputModel.CaseNumber.TrimEnd(), true, true);
            string current_full_address = "";
            string currentstreet;
            string currentcity;
            string currentstate = "";

            if (uSPEIndividual.diligenceInputModel.CurrentStreet.ToString().Equals(""))
            {
                currentstreet = "";
            }
            else
            {
                if (uSPEIndividual.diligenceInputModel.CurrentCity.ToString().Equals("") && uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State") && uSPEIndividual.diligenceInputModel.CurrentZipcode.ToString().Equals(""))
                {
                    currentstreet = string.Concat(uSPEIndividual.diligenceInputModel.CurrentStreet.ToString().TrimEnd());
                }
                else
                {
                    currentstreet = string.Concat(uSPEIndividual.diligenceInputModel.CurrentStreet.ToString().TrimEnd(), ", ");
                }

            }
            if (uSPEIndividual.diligenceInputModel.CurrentCity.ToString().Equals(""))
            {
                currentcity = "";
            }
            else
            {
                if (uSPEIndividual.diligenceInputModel.CurrentZipcode.ToString().Equals("") && uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                {
                    currentcity = uSPEIndividual.diligenceInputModel.CurrentCity.ToString().TrimEnd();
                }
                else
                {
                    currentcity = string.Concat(uSPEIndividual.diligenceInputModel.CurrentCity.ToString().TrimEnd(), ", ");
                }
            }
            if (uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
            {
                if (uSPEIndividual.diligenceInputModel.CurrentZipcode.ToString().Equals(""))
                {
                    currentstate = "";
                }
                else
                {
                    currentstate = uSPEIndividual.diligenceInputModel.CurrentZipcode.ToString().TrimEnd();
                }
            }
            else
            {
                try
                {
                    StateWiseFootnoteModel statecomment1 = _context.stateModel
                                         .Where(u => u.states.ToUpper().TrimEnd() == uSPEIndividual.diligenceInputModel.CurrentState.ToString().ToUpper())
                                         .FirstOrDefault();
                    currentstate = statecomment1.abbreviation;
                }
                catch
                {
                    currentstate = uSPEIndividual.diligenceInputModel.CurrentState.ToString();
                }
                if (uSPEIndividual.diligenceInputModel.CurrentZipcode.ToString().Equals(""))
                {
                    
                }
                else
                {
                    currentstate = string.Concat(currentstate, " ", uSPEIndividual.diligenceInputModel.CurrentZipcode.ToString().TrimEnd());
                }
            }
            current_full_address = string.Concat(currentstreet, currentcity, currentstate);            
            if (current_full_address.ToString().Equals("")) { doc.Replace("[CurrentFullAddress]", "<not provided>", true, true); }
            else
            {
                doc.Replace("[CurrentFullAddress]", current_full_address, true, true);
            }
            if (uSPEIndividual.diligenceInputModel.CurrentCity.ToString().ToUpper().Equals("NA") || uSPEIndividual.diligenceInputModel.CurrentCity.ToString().Equals(""))
            {
                doc.Replace("[CurrentCity], ", "", true, false);
            }
            else
            {
                doc.Replace("[CurrentCity]", uSPEIndividual.diligenceInputModel.CurrentCity, true, false);
            }
            doc.SaveToFile(savePath);
            if (uSPEIndividual.diligenceInputModel.CurrentState.Equals("Select State"))
            {
                doc.Replace(", [CurrentState]", "", true, false);
                doc.Replace(" [CurrentState]", "", true, false);
            }
            else
            {
                doc.Replace("[CurrentState]", uSPEIndividual.diligenceInputModel.CurrentState, true, false);
            }
            doc.SaveToFile(savePath);
            //Business Affiliation           
            if (uSPEIndividual.otherdetails.Has_Business_Affiliations.ToString().Equals("Yes"))
            {
                doc.Replace("SODCOMMENT", "[LastName] is a [Position1] of [Employer1] in [Employer1City], [Employer1State], since [EmpStartDate1]\n\nFurther, [LastName] was identified in connection with additional business entities.", true, true);
                doc.Replace("SODRESULT", "Records", true, true);
                doc.Replace("BUSINESSAFFILIATIONSIDENTIFIED", "\nBusiness Affiliations\nIn addition to the above, <and while not believed to be a comprehensive listing of the same>, research of records maintained by the Secretary of State’s Office, well as other sources, identified the subject in connection with the following business entities:\n\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>", true, true);
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_Business_Affiliations.ToString().Equals("Yes – Board Member"))
                {
                    doc.Replace("SODCOMMENT", "[LastName] is a [Position1] of [Employer1] in [Employer1City], [Employer1State], since [EmpStartDate1]\n\nFurther, [LastName] was identified in connection with additional business entities.", true, true);
                    doc.Replace("SODRESULT", "Records", true, true);
                    doc.Replace("BUSINESSAFFILIATIONSIDENTIFIED", "\nBusiness Affiliations\n\nIn addition to the above, research of records maintained by the Secretary of State’s Office, as well as other sources, identified [LastName] in connection with the following business entities (those entities where the subject serves and/or has previously served as a Board Member are denoted with an asterisk (“*”)):\n	•	<Investigator to insert bulleted list of results here>", true, true);
                }
                else
                {
                    doc.Replace("SODCOMMENT", "[LastName] is a [Position1] of [Employer1] in [Employer1City], [Employer1State], since [EmpStartDate1]", true, true);
                    doc.Replace("SODRESULT", "Clear", true, true);
                    doc.Replace("BUSINESSAFFILIATIONSIDENTIFIED", "\nBusiness Affiliations\n\nIn addition to the above, research of records maintained by the Secretary of State’s Office, as well as other sources, did not identify the subject in connection with any business entities.", true, true);
                }
            }           
            try
            {
                if (uSPEIndividual.EmployerModel[0].Emp_Position.Equals("")) { doc.Replace("[Position1]", "<not provided>", true, true); }
                else { doc.Replace("[Position1]", uSPEIndividual.EmployerModel[0].Emp_Position, true, true); }
            }
            catch
            {
                doc.Replace("[Position1]", "<not provided>", true, true);
            }
            try
            {
                if (uSPEIndividual.EmployerModel[0].Emp_Employer.Equals("")) { doc.Replace("[Employer1]", "<not provided>", true, true); }
                else { doc.Replace("[Employer1]", uSPEIndividual.EmployerModel[0].Emp_Employer, true, true); }
            }
            catch
            {
                doc.Replace("[Employer1]", "<not provided>", true, true);
            }
            try
            {
                if (uSPEIndividual.EmployerModel[0].Emp_StartDateYear.ToString().Equals(""))
                {
                    doc.Replace("[EmpStartDate1]", "<not provided>", true, true);
                }
                else
                {
                    doc.Replace("[EmpStartDate1]", uSPEIndividual.EmployerModel[0].Emp_StartDateYear, true, true);
                }
            }
            catch
            {
                doc.Replace("[EmpStartDate1]", "<not provided>", true, true);
            }
            doc.SaveToFile(savePath);
            //if (uSPEIndividual.diligenceInputModel.CommonNameSubject == true)
            //{
            //    doc.Replace("[COMADDBUSINESSAFTER]", "\n\nIt should be noted that numerous individuals known only as “[First Name] [Last Name]” were identified in corporate records throughout the United States, and significant additional research efforts would be required to determine whether any of the same relate to the subject of this investigation, which can be undertaken upon request -- if an expanded scope is warranted.", true, false);
            //}
            //else
            //{
            //    doc.Replace("[COMADDBUSINESSAFTER]", "", true, false);
            //}
            doc.SaveToFile(savePath);
            if (uSPEIndividual.EmployerModel[0].Emp_State.Equals("Select State"))
            {
                doc.Replace("[Employer1State], ", "", true, true);
            }
            else
            {
                doc.Replace("[Employer1State], ", uSPEIndividual.EmployerModel[0].Emp_State, true, true);
            }
            if (uSPEIndividual.EmployerModel[0].Emp_City.Equals("") || uSPEIndividual.EmployerModel[0].Emp_City.ToUpper().Equals("NA")|| uSPEIndividual.EmployerModel[0].Emp_City.ToUpper().Equals("N/A"))
            {
                doc.Replace("[Employer1City], ", "", true, true);
            }
            else
            {
                doc.Replace("[Employer1City], ", uSPEIndividual.EmployerModel[0].Emp_City, true, true);
            }
            //Intellectual Hits
            string intellectual_comment;
            CommentModel intellec_comment2 = _context.DbComment
                               .Where(u => u.Comment_type == "Intellectual_hits")
                               .FirstOrDefault();
            if (uSPEIndividual.otherdetails.Has_Intellectual_Hits.ToString().Equals("Yes"))
            {
                intellectual_comment = string.Concat("Research of records maintained by the United States Patent and Trademark Office (“USPTO”) and/or World Intellectual Property Organization (“WIPO”) identified the [LastName] as an ",uSPEIndividual.otherdetails.ApplicantType.ToString() ," in connection with the following ", uSPEIndividual.otherdetails.PatentType.ToString(), " :");
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
            if (uSPEIndividual.otherdetails.Global_Security_Hits.ToString().Equals("Yes"))
            {
                globalhit_comment = globalhit_comment2.confirmed_comment.ToString();
                globalhit_comment = string.Concat(globalhit_comment, "\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n");
            }
            else
            {
                globalhit_comment = globalhit_comment2.unconfirmed_comment.ToString();
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
            if (uSPEIndividual.otherdetails.ICIJ_Hits.ToString().Equals("Yes"))
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
            if (uSPEIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered"))
            {
                usseccommentmodel = "";
                doc.Replace("SECCOMMENT", "", true, true);
                doc.Replace("SECRESULT", "Clear", true, true);
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Yes-With Adverse"))
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
                // doc.Replace("US_SECHEADER", " \n United States Securities and Exchange Commission \n", true, true);
                usseccommentmodel = string.Concat("\nUnited States Securities and Exchange Commission\n\n", usseccommentmodel, "\n");
            }

            //UK_FCA
            string ukfcacommentmodel = "";
            CommentModel ukfcacommentmodel1 = _context.DbComment
                            .Where(u => u.Comment_type == "Reg_UK_FCA")
                            .FirstOrDefault();
            if (uSPEIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered"))
            {
                ukfcacommentmodel = "";
                doc.Replace("UKFICOCOMMENT", "", true, true);
                doc.Replace("UKFICORESULT", "Clear", true, true);
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Yes-With Adverse"))
                {
                    ukfcacommentmodel = ukfcacommentmodel1.confirmed_comment.ToString();
                    doc.Replace("UKFICOCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("UKFICORESULT", "Records", true, true);
                }
                else
                {
                    doc.Replace("UKFICOCOMMENT", "", true, true);
                    doc.Replace("UKFICORESULT", "Clear", true, true);
                    ukfcacommentmodel = ukfcacommentmodel1.unconfirmed_comment.ToString();
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
            if (uSPEIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered"))
            {
                finracommentmodel = "";
                doc.Replace("FININDCOMMENT", "", true, true);
                doc.Replace("FININDRESULT", "Clear", true, true);
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Yes-With Adverse"))
                {
                    finracommentmodel = finracommentmodel1.confirmed_comment.ToString();
                    doc.Replace("FININDCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("FININDRESULT", "Records", true, true);
                }
                else
                {
                    doc.Replace("FININDCOMMENT", "", true, true);
                    doc.Replace("FININDRESULT", "Clear", true, true);
                    finracommentmodel = finracommentmodel1.unconfirmed_comment.ToString();
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
            if (uSPEIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered"))
            {
                nfacommentmodel = "";
                doc.Replace("USNFCOMMENT", "", true, true);
                doc.Replace("USNFRESULT", "Clear", true, true);
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Yes-With Adverse"))
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
                //doc.Replace("US_NFAHEADER", "\n United States National Futures Association  \n", true, true);
                nfacommentmodel = string.Concat("\nUnited States National Futures Association\n\n", nfacommentmodel, "\n");
            }
            string hksfccommentmodel = "";                      
                //HKFSC
                if (uSPEIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Not Registered"))
                {
                    hksfccommentmodel = "";
                }
                else
                {
                    if (uSPEIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes - Without Adverse"))
                    {
                        hksfccommentmodel = "[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nThere are no public disciplinary actions on file against the subject within the past five years.";
                    }
                    else
                    {
                        hksfccommentmodel = "[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nAdditionally, the following public disciplinary actions were on file in connection with the subject:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                    }
                    hksfccommentmodel = string.Concat("\nHong Kong Securities and Futures Commission[HKSFCHEADER]\n\n", hksfccommentmodel, "\n");
                }
            //Holds Any License   
            string regflag = "";
                CommentModel holdslicensecommentmodel1 = _context.DbComment
                                .Where(u => u.Comment_type == "Other_License_Language")
                                .FirstOrDefault();
            string holdslicensecommentmodel = "";
            if (uSPEIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse"))
            {
                regflag = "Records";
            }
            else
            {
                regflag = "Clear";

            }

            if (uSPEIndividual.otherdetails.Registered_with_HKSFC.StartsWith("Yes") || uSPEIndividual.otherdetails.Has_Reg_US_SEC.StartsWith("Yes") || uSPEIndividual.otherdetails.Has_Reg_US_NFA.StartsWith("Yes") || uSPEIndividual.otherdetails.Has_Reg_FINRA.StartsWith("Yes") || uSPEIndividual.otherdetails.Has_Reg_UK_FCA.StartsWith("Yes") || uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
            {
                holdslicensecommentmodel = "\nOther Professional Licensures and/or Designations\n\nInvestigative efforts did not reveal any additional professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction- and license-type basis to confirm the same.\n";
            }
            else
            {
                holdslicensecommentmodel = "\nInvestigative efforts did not reveal any professional licensure or registration information in connection with [LastName], personally, however, further efforts would be required on a jurisdiction- and license-type basis to confirm the same.\n";
            }
            if (uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().ToUpper().Equals("YES")) { doc.Replace("[PLLICENSEDESC]", string.Concat("\n",usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel,holdslicensecommentmodel), true, false); }
            else
            {
                doc.Replace("[PLLICENSEDESC]", string.Concat(usseccommentmodel, ukfcacommentmodel, finracommentmodel, nfacommentmodel, hksfccommentmodel, holdslicensecommentmodel), true, false);
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
            doc.SaveToFile(savePath);
            if (uSPEIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Not Registered")) { }
            else
            {
                string blnresult = "";
                try
                {
                    for (int j = 1; j < sectionLimit; j++)
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

            if (uSPEIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Not Registered")) { }
            else
            {
                string blnresult = ""; try
                {
                    for (int j = 1; j < sectionLimit; j++)
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
                catch { }
            }

            if (uSPEIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Not Registered")) { }
            else
            {
                string blnresult = ""; try {
                    for (int j = 1; j < sectionLimit; j++)
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
                catch { }
            }

            if (uSPEIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Not Registered")) { }
            else
            {
                string blnresult = ""; try
                {
                    for (int j = 1; j < sectionLimit; j++)
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
            catch
            { }
            holdresult = "";
            if (hksfccommentmodel.ToString().Equals("")) { }
            else
            {
                //for (int j = 1; j < sectionLimit; j++)
                //{
                //    Section section = doc.Sections[j];
                //    foreach (Paragraph para in section.Paragraphs)
                //    {
                //        DocumentObject obj = null;
                //        for (int i = 0; i < para.ChildObjects.Count; i++)
                //        {
                //            obj = para.ChildObjects[i];
                //            if (obj.DocumentObjectType == DocumentObjectType.TextRange)
                //            {
                //                TextRange textRange = obj as TextRange;
                //                string abc = textRange.Text;
                //                if (abc.ToString().Equals("Hong Kong Securities and Futures Commission"))
                //                {
                //                    textRange.CharacterFormat.Italic = true;
                //                    textRange.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                //                    doc.SaveToFile(savePath);
                //                    holdresult = "true";
                //                    break;
                //                }
                //            }
                //        }
                //        if (holdresult.Equals("true")) { break; }
                //    }
                //}
            }

            //Regulatory_Red_Flag    
            string strrrffootnote = "";
            string strrrfappendtext = "";
            if (uSPEIndividual.otherdetails.RegulatoryFlag.ToString().Equals("Yes"))
            {
                doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", "In addition to the above, national financial institutions sanctions and legal actions were searched covering  the banking, mortgage and securities industries,", true, true);
                //strrrffootnote = "Searches included sanctions and/or legal actions of the following federal agencies: the Securities and Exchange Commission; the Financial Industry Regulatory Authority (formerly the National Association of Securities Dealers) and the New York Stock Exchange’s member regulation, enforcement and arbitration functions (prior to the consolidation of the two in July 2007); the Federal Deposit Insurance Corporation; Resolution Trust Corporation; Federal Reserve Board; National Credit Union Administrative Actions; Office of the Comptroller of the Currency; Department of Justice; Department of Housing and Urban Development; and other federal agencies (through the General Services Administration).";
                strrrffootnote = "Searches included sanctions and/or legal actions of the following federal agencies: the Securities and Exchange Commission; the Financial Industry Regulatory Authority (formerly the National Association of Securities Dealers) and the New York Stock Exchange’s member regulation, enforcement and arbitration functions (prior to the consolidation of the two in July 2007); the Federal Deposit Insurance Corporation; the Resolution Trust Corporation; the Federal Reserve Board; the National Credit Union Administrative Actions; the Office of the Comptroller of the Currency; the Department of Justice; the Department of Housing and Urban Development; and other federal agencies (through the General Services Administration).  Further, state-level agencies searches included mortgage and real estate regulators and certain records of Secretaries of State.  These sources cover securities and other regulatory violations, disciplinary actions, company filings and similar records.";
                strrrfappendtext = " and the following information was identified in connection with [LastName]:   <Investigator to insert results here>";
            }
            else
            {
                if (uSPEIndividual.otherdetails.RegulatoryFlag.ToString().Equals("No – Medical"))
                {
                    doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", "In addition to the above, national financial institutions sanctions and legal actions were searched covering the banking, mortgage and securities industries, as well as in an effort to identify any medical compliance issues or other regulatory events,", true, true);
                    strrrffootnote = "Searches included various sanctions and/or legal actions through the United States Department of Health & Human Services’ Office of Inspector General, the Fraud and Abuse Control Information System, the Drug Enforcement Agency, and the Food and Drug Administration.  Further, searches included sanctions and/or legal actions of the following federal agencies: the Securities and Exchange Commission; the Financial Industry Regulatory Authority (formerly the National Association of Securities Dealers) and the New York Stock Exchange’s member regulation, enforcement and arbitration functions (prior to the consolidation of the two in July 2007); the Federal Deposit Insurance Corporation; the Resolution Trust Corporation; the Federal Reserve Board; the National Credit Union Administrative Actions; the Office of the Comptroller of the Currency; the Department of Justice; the Department of Housing and Urban Development; and other federal agencies (through the General Services Administration).  Moreover, state-level agencies searches included mortgage and real estate regulators and certain records of Secretaries of State.  These sources cover securities and other regulatory violations, disciplinary actions, company filings and similar records.";
                    strrrfappendtext = " and it is noted that [LastName] was not identified in any records.";
                }
                else
                {
                    doc.Replace("OTHERREGULATORYREDFLAGSDESCRIPTION", "In addition to the above, national financial institutions sanctions and legal actions were searched covering the banking, mortgage and securities industries,", true, true);
//                    strrrffootnote = "Searches included sanctions and/or legal actions of the following federal agencies: the Securities and Exchange Commission; the Financial Industry Regulatory Authority (formerly the National Association of Securities Dealers) and the New York Stock Exchange’s member regulation, enforcement and arbitration functions (prior to the consolidation of the two in July 2007); the Federal Deposit Insurance Corporation; Resolution Trust Corporation; Federal Reserve Board; National Credit Union Administrative Actions; Office of the Comptroller of the Currency; Department of Justice; Department of Housing and Urban Development; and other federal agencies (through the General Services Administration).";
                    strrrffootnote = "Searches included sanctions and/or legal actions of the following federal agencies: the Securities and Exchange Commission; the Financial Industry Regulatory Authority (formerly the National Association of Securities Dealers) and the New York Stock Exchange’s member regulation, enforcement and arbitration functions (prior to the consolidation of the two in July 2007); the Federal Deposit Insurance Corporation; the Resolution Trust Corporation; the Federal Reserve Board; the National Credit Union Administrative Actions; the Office of the Comptroller of the Currency; the Department of Justice; the Department of Housing and Urban Development; and other federal agencies (through the General Services Administration).  Further, state-level agencies searches included mortgage and real estate regulators and certain records of Secretaries of State.  These sources cover securities and other regulatory violations, disciplinary actions, company filings and similar records.";
                    strrrfappendtext = " and it is noted that [LastName] was not identified in any of these records.";
                }
            }
            doc.SaveToFile(savePath);
            string blnredresultfound = "";
            try
            {
                for (int j = 2; j < sectionLimit; j++)
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
                                if (abc.ToString().EndsWith("mortgage and securities industries,") || abc.ToString().Contains("mortgage and securities industries,"))
                                {
                                    Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                    footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                    footnote1.MarkerCharacterFormat.FontSize = 11;
                                    //Insert footnote1 after the word "Spire.Doc"
                                    para.ChildObjects.Insert(i + 1, footnote1);
                                    TextRange _text = footnote1.TextBody.AddParagraph().AppendText(strrrffootnote);
                                    footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                    _text.CharacterFormat.FontName = "Calibri (Body)";
                                    _text.CharacterFormat.FontSize = 9;
                                    //Append the line                                                                                                   
                                    strrrfappendtext = strrrfappendtext.Replace("[LastName]", uSPEIndividual.diligenceInputModel.LastName);
                                    TextRange tr = para.AppendText(strrrfappendtext);
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
            //AdditionalStates
            string add_states = "";
            string strAdditionalStatesComment = "";
            string arr = "";
            string state1 = uSPEIndividual.diligenceInputModel.CurrentState.ToString();
            int statecount = 0;
            try
            {
                if (uSPEIndividual.additional_States.Count > 0)
                {
                    if (uSPEIndividual.additional_States[0].additionalstate.ToUpper().ToString().Equals("Select State".ToUpper()) && uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                    {
                        doc.Replace("[ADDITIONALSTATES]", "", true, true);
                        add_states = "<not provides>";
                    }
                    if (state1 == "Select State") { }
                    else
                    {
                        add_states = string.Concat("<investigator to add relevant county/counties>, ", uSPEIndividual.diligenceInputModel.CurrentState.ToString());
                        arr = uSPEIndividual.diligenceInputModel.CurrentState.ToString();
                        statecount = statecount + 1;
                        strAdditionalStatesComment = add_states;
                    }
                    if (uSPEIndividual.additional_States[0].additionalstate.ToUpper().ToString().Equals("Select State".ToUpper()))
                    {
                    }
                    else
                    {
                        if (uSPEIndividual.additional_States.Count > 1)
                        {
                            int count = uSPEIndividual.additional_States.Count();
                            int i = 1;
                            for (int j = 0; j < count; j++)
                            {
                                if (uSPEIndividual.additional_States[j].additionalstate.ToUpper().ToString().Equals("Select State".ToUpper()))
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
                                        strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, "<investigator to add relevant county/counties>, ", uSPEIndividual.additional_States[j].additionalstate, " and ");
                                        add_states = strAdditionalStatesComment;
                                    }
                                    else
                                    {
                                        strAdditionalStatesComment = string.Concat(strAdditionalStatesComment, "<investigator to add relevant county/counties>, ", uSPEIndividual.additional_States[j].additionalstate, ", ");
                                        if (i == count) { add_states = string.Concat(add_states, "<investigator to add relevant county/counties>, ", uSPEIndividual.additional_States[j].additionalstate); }
                                        else
                                        {
                                            add_states = strAdditionalStatesComment;
                                        }
                                    }
                                    arr = string.Concat(arr, ", ", uSPEIndividual.additional_States[j].additionalstate.ToString());
                                    statecount = statecount + 1;
                                }
                                i++;
                            }
                        }
                        else
                        {
                            if (add_states == "") { } else { add_states = string.Concat(add_states, " and "); }
                            strAdditionalStatesComment = string.Concat(add_states, "<investigator to add relevant county/counties>, ", uSPEIndividual.additional_States[0].additionalstate.ToString());
                            add_states = strAdditionalStatesComment;
                            arr = string.Concat(arr, ", ", uSPEIndividual.additional_States[0].additionalstate.ToString());
                            statecount = statecount + 1;
                        }
                        doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, false);
                    }
                    doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, false);
                }
                else
                {
                    if (uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
                    {
                        doc.Replace("[ADDITIONALSTATES]", "", true, true);
                        add_states = "<not provides>";
                    }
                    if (state1 == "Select State") { }
                    else
                    {
                        add_states = string.Concat("<investigator to add relevant county/counties>, ", uSPEIndividual.diligenceInputModel.CurrentState.ToString());
                        arr = uSPEIndividual.diligenceInputModel.CurrentState.ToString();
                        statecount = statecount + 1;
                        strAdditionalStatesComment = add_states;
                    }
                    
                }
            }
            catch
            {

            }
            String[] arrlist = arr.Split(", ");
            //ADDITIONALJURIDICTIONS
            if (uSPEIndividual.diligenceInputModel.Nonscopecountry1.ToString().ToUpper().Equals("NA") || uSPEIndividual.diligenceInputModel.Nonscopecountry1.ToString().ToUpper().Equals("") || uSPEIndividual.diligenceInputModel.Nonscopecountry1.ToString().ToUpper().Equals("N/A"))
            {
                doc.Replace("ADDITIONALJURIDICTIONS", "", true, true);
            }
            else
            {
                if (uSPEIndividual.diligenceInputModel.Nonscopecountry1.Contains(", "))
                {
                    doc.Replace("ADDITIONALJURIDICTIONS", "\nAdditionally, the subject has historical and/or possible ties to [ADITIONALJURISCOMMENT], and additional research would be required in these jurisdictions, which can be undertaken upon request -- if an expanded scope is warranted. <investigator to alter language to account for only one jurisdiction -- as needed>\n", true, true);
                }
                else
                {
                    doc.Replace("ADDITIONALJURIDICTIONS", "\nAdditionally, the subject has historical and/or possible ties to [ADITIONALJURISCOMMENT], and additional research would be required in this jurisdiction, which can be undertaken upon request -- if an expanded scope is warranted.  <investigator to alter language to account for multiple jurisdictions as needed>\n", true, true);
                }
                doc.Replace("[ADITIONALJURISCOMMENT]", uSPEIndividual.diligenceInputModel.Nonscopecountry1, true, true);
            }
            //Real Estate
            if (uSPEIndividual.otherdetails.CurrentResidentialProperty.Equals("No") && uSPEIndividual.otherdetails.CurrentResidentialProperty.Equals("No") && uSPEIndividual.otherdetails.OtherPropertyOwnershipinfo.Equals("No") && uSPEIndividual.otherdetails.PrevPropertyOwnershipRes.Equals("No"))
            {
                doc.Replace("REALCOMMENT", "", true, true);
                doc.Replace("REALRESULT", "No Records", true, true);
            }
            else
            {
                if (uSPEIndividual.otherdetails.OtherCurrentResidentialProperty.Equals("Yes, multiple") || uSPEIndividual.otherdetails.PrevPropertyOwnershipRes.Equals("Yes, multiple") || uSPEIndividual.otherdetails.CurrentResidentialProperty.ToString().Equals("Yes") && uSPEIndividual.otherdetails.OtherPropertyOwnershipinfo.ToString().Equals("Yes") || uSPEIndividual.otherdetails.CurrentResidentialProperty.ToString().Equals("Yes") && uSPEIndividual.otherdetails.PrevPropertyOwnershipRes.ToString().StartsWith("Yes") || uSPEIndividual.otherdetails.OtherPropertyOwnershipinfo.ToString().Equals("Yes") && uSPEIndividual.otherdetails.PrevPropertyOwnershipRes.ToString().StartsWith("Yes"))
                {
                    doc.Replace("REALCOMMENT", "[LastName] was identified as the owner of properties located in <Investigator to insert Counties and States>", true, true);
                    doc.Replace("REALRESULT", "Records", true, true);
                }
                else
                {
                    doc.Replace("REALCOMMENT", "[LastName] was identified as the owner of a property located in <investigator to insert County, State>", true, true);
                    doc.Replace("REALRESULT", "Record", true, true);
                }
            }
            doc.SaveToFile(savePath);
            string strpopertyappend = "";
            //CURRENTRESIDENTIALPROPERTYDESC
            if (uSPEIndividual.otherdetails.CurrentResidentialProperty.ToString().Equals("Yes"))
            {
                doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "Current ownership records are identified in connection with a certain property located at [CurrentFullAddress].  According to <County> County, [CurrentState1] property records, the subject purchased this property for <purchase price> from <seller names> on <date>, <Investigator to insert other mortgage information or UCC details here>[ADDOLOFOOTNOTE].", true, true);
                strpopertyappend = "This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>[ADDOOFOOTNOTE]";

            }
            else
            {
                doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "While [LastName] currently resides at [CurrentFullAddress], the subject was not identified as the owner of the same. <Investigator to insert property record or rental information>.CURRENTRESIDENTIALPROPERTYDESC", true, true);
            }
            doc.SaveToFile(savePath);
            if (uSPEIndividual.diligenceInputModel.CurrentState.ToString().Equals("Select State"))
            {
                doc.Replace("[CurrentState1]", "<State>", true, true);
            }
            else
            {
                doc.Replace("[CurrentState1]", uSPEIndividual.diligenceInputModel.CurrentState.ToString(), true, true);
            }
            doc.SaveToFile(savePath);
            blnredresultfound = "";
            try
            {
                for (int j = 2; j < sectionLimit; j++)
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
                                //Find the word "Spire.Doc" in paragraph1
                                if (abc.ToString().Contains("[ADDOLOFOOTNOTE]"))
                                {
                                    Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                    footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                    footnote1.MarkerCharacterFormat.FontSize = 11;
                                    //Insert footnote1 after the word "Spire.Doc"
                                    para.ChildObjects.Insert(i + 1, footnote1);
                                    TextRange _text = footnote1.TextBody.AddParagraph().AppendText("According to the subject’s credit history report, <investigator to insert credit history in regard to mortgage amount>.");
                                    footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                    _text.CharacterFormat.FontName = "Calibri (Body)";
                                    _text.CharacterFormat.FontSize = 9;
                                    TextRange tr = para.AppendText(strpopertyappend);
                                    tr.CharacterFormat.FontName = "Calibri (Body)";
                                    tr.CharacterFormat.FontSize = 11;
                                    blnredresultfound = "true";
                                    doc.Replace("[ADDOLOFOOTNOTE]", "", true,false);
                                    doc.SaveToFile(savePath);
                                    // break;
                                }
                                if (abc.ToString().EndsWith("[ADDOOFOOTNOTE]") || abc.ToString().Contains("[ADDOOFOOTNOTE]") )
                                {
                                    Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                    footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                    footnote1.MarkerCharacterFormat.FontSize = 11;
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
            }
            catch
            {
            }
            doc.SaveToFile(savePath);
            //OtherCurrentResidentialProperty            
            try
            {
                if (uSPEIndividual.otherdetails.OtherCurrentResidentialProperty.ToString().Equals("Yes, only one"))
                {
                    doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "\n\nCurrent ownership records are also identified in connection with a certain property located at <Investigator to add other current property address>. According to <County> County, <State> property records, the subject purchased this property for <purchase price> from <seller names> on <date>, <Investigator to insert other mortgage information or UCC details here>[ADDOLO1FOOTNOTE].", true, false);
                    strpopertyappend = "This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>.[ADDOOFOOT1NOTE]\n";

                }
                else
                {
                    if (uSPEIndividual.otherdetails.OtherCurrentResidentialProperty.ToString().Equals("Yes, multiple"))
                    {
                        doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "\n\nIn addition, current ownership records also identified [LastName] in connection with the following properties, which are located in <Investigator to add all relevant County, State(s)>:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here> OTHERPROPOWNERSHIPDESC PREVIOUSPROPERTYOWNERSHIPDESC", true, false);
                    }
                    else
                    {
                        doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "OTHERPROPOWNERSHIPDESC PREVIOUSPROPERTYOWNERSHIPDESC", true, false);
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
                                if (abc.ToString().Contains("[ADDOLO1FOOTNOTE]"))
                                {
                                    Footnote footnote1 = para.AppendFootnote(FootnoteType.Footnote);
                                    footnote1.MarkerCharacterFormat.FontName = "Calibri (Body)";
                                    footnote1.MarkerCharacterFormat.FontSize = 11;
                                    //Insert footnote1 after the word "Spire.Doc"
                                    para.ChildObjects.Insert(i + 1, footnote1);
                                    TextRange _text = footnote1.TextBody.AddParagraph().AppendText("According to the subject’s credit history report, <investigator to insert credit history in regard to mortgage amount>.");
                                    footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                    _text.CharacterFormat.FontName = "Calibri (Body)";
                                    _text.CharacterFormat.FontSize = 9;
                                    TextRange tr = para.AppendText(strpopertyappend);
                                    tr.CharacterFormat.FontName = "Calibri (Body)";
                                    tr.CharacterFormat.FontSize = 11;
                                    blnredresultfound = "true";
                                    doc.Replace("[ADDOLO1FOOTNOTE]", "", true, false);
                                    doc.SaveToFile(savePath);
                                    // break;
                                }
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
                                    TextRange tr = para.AppendText("OTHERPROPOWNERSHIPDESC PREVIOUSPROPERTYOWNERSHIPDESC");
                                    tr.CharacterFormat.FontName = "Calibri (Body)";
                                    tr.CharacterFormat.FontSize = 11;
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
                doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "OTHERPROPOWNERSHIPDESC PREVIOUSPROPERTYOWNERSHIPDESC", true, false);
                doc.SaveToFile(savePath);
            }
            //OTHERPROPOWNERSHIPDESC
            if (uSPEIndividual.otherdetails.OtherPropertyOwnershipinfo.ToString().Equals("Yes"))
            {
                doc.Replace("OTHERPROPOWNERSHIPDESC", "\n\nCurrent ownership records are also identified in connection with a certain property located at <address>.  According to <County> County, <State> property records, the subject purchased this property for <purchase price> from <seller names> on <date>. <Investigator to insert other mortgage information or UCC details here>.  This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>.", true, true);
            }
            else
            {
                doc.Replace("OTHERPROPOWNERSHIPDESC", "", true, true);
            }
            //PREVIOUSPROPERTYOWNERSHIPDESC
            if (uSPEIndividual.otherdetails.PrevPropertyOwnershipRes.ToString().Equals("No"))
            {
                doc.Replace("PREVIOUSPROPERTYOWNERSHIPDESC", "", true, true);
            }
            else
            {
                if (uSPEIndividual.otherdetails.PrevPropertyOwnershipRes.ToString().Equals("Yes, only one"))
                {
                    doc.Replace("PREVIOUSPROPERTYOWNERSHIPDESC", "\n\nPrevious ownership records were identified in connection with a certain property located at <address>.  According to <County> County, <State> property records, the subject sold this property for <sale price> to <buyer names> on <date>.", true, true);
                }
                else
                {
                    doc.Replace("PREVIOUSPROPERTYOWNERSHIPDESC", "\n\nFurther, previous ownership records were identified in connection with the following properties, which are located in <add Counties/States>:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>", true, true);
                }
            }
            //REGISTEREDWITHHKSFCDESC
            if (uSPEIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Not Registered"))
            {
                doc.Replace("HKSFCHEAD", "", true, true);
                doc.Replace("REGISTEREDWITHHKSFCDESC", "", true, false);
                doc.Replace("KONSECCOMMENT", "", true, true);
                doc.Replace("KONSECRESULT", "Clear", true, true);
            }
            else
            {
                doc.Replace("HKSFCHEAD", "HKSFCHEAD".Trim(), true, true);
                doc.SaveToFile(savePath);
                if (uSPEIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes - With Adverse"))
                {
                    doc.Replace("KONSECCOMMENT", "<investigator to insert summary here>", true, true);
                    doc.Replace("KONSECRESULT", "Records", true, true);
                    doc.Replace("HKSFCHEAD", "\nHong Kong Securities and Futures Commission", true, true);
                    doc.Replace("REGISTEREDWITHHKSFCDESC", "\n\n[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>)), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nAdditionally, the following public disciplinary actions were on file in connection with the subject:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\n", true, false);
                }
                else
                {
                    doc.Replace("KONSECCOMMENT", "", true, true);
                    doc.Replace("KONSECRESULT", "Clear", true, true);
                    doc.Replace("HKSFCHEAD", "\nHong Kong Securities and Futures Commission", true, true);
                    doc.Replace("REGISTEREDWITHHKSFCDESC", "\n\n[LastName] has been registered with the Hong Kong Securities and Futures Commission (“HKSFC”), with the identification number of <number>, as a <role> (effective as of <dates>)), in the regulated activity of <capacity> (effective as of <dates>) with <entity>.\n\nThere are no public disciplinary actions on file against the subject within the past five years.\n\n", true, false);
                }
            }
            //COMMONNAMESUBDESC
            if (uSPEIndividual.diligenceInputModel.CommonNameSubject==false)
            {
                doc.Replace("COMMONNAMESUBDESC", "Additionally, it is noted that searches of the United States Bankruptcy and District Courts were nearly national in scope.", true, true);
            }
            else
            {
                doc.Replace("COMMONNAMESUBDESC", "", true, true);
            }
            doc.SaveToFile(savePath);
            //HASBANKRUPTYRECHITDESC       
            string strbankrupfootnt1 = "";
            string strconcatBankrupty = "";
            strAdditionalStatesComment = "";
            string strbank1 = "";
            string strbanksumcomm = "";
            string strbankressumm = "";
            if (uSPEIndividual.otherdetails.HasBankruptcyRecHits_resultpending == true)
            {
                doc.Replace("HASBANKRUPTYRECHITDESC", "Bankruptcy court records searches are currently pending in <jurisdiction> the results of which will be provided under separate cover upon receipt.\n\nHASBANKRUPTYRECHITDESC", true, true);
                doc.SaveToFile(savePath);
            }
            if (uSPEIndividual.otherdetails.HasBankruptcyRecHits.ToString().Equals("Yes, Multiple Records"))
            {
                strbank1 = "Records";
                doc.Replace("HASBANKRUPTYRECHITDESC", "Efforts pursued in [ADDITIONALBANKSTATES]ADDFOOTNOTE", true, true);
                strconcatBankrupty = " identified the subject, personally, as a <party type> in connection with the following bankruptcy filings:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n";
            }
            else
            {
                if (uSPEIndividual.otherdetails.HasBankruptcyRecHits.ToString().Equals("No"))
                {
                    strbank1 = "Clear";
                    doc.Replace("HASBANKRUPTYRECHITDESC", "Efforts pursued in [ADDITIONALBANKSTATES]ADDFOOTNOTE", true, true);
                    strconcatBankrupty = " did not identify the subject, personally, in connection with any bankruptcy filings.\n";
                }
                else
                {
                    strbank1 = "Record";
                    doc.Replace("HASBANKRUPTYRECHITDESC", "Efforts pursued in [ADDITIONALBANKSTATES]ADDFOOTNOTE", true, true);
                    strconcatBankrupty = " identified the subject, personally, as a <party type> in connection with the following bankruptcy filing:\n\n\t•\t<Investigator to insert bulleted list of results here>\n";
                }
            }
            //if (uSPEIndividual.diligenceInputModel.CommonNameSubject == true)
            //{
            //    strconcatBankrupty = string.Concat(strconcatBankrupty, "\nIn light of the commonality of [LastName]’s name, this portion of the investigation was conducted utilizing the subject’s Social Security number, in order to identify only those records conclusively relating to the subject of interest.\n");
            //}
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
                        if (arrlist[j].ToUpper().ToString().Equals("Select State".ToUpper()))
                        {
                        }
                        else
                        {
                            StateWiseFootnoteModel comment1 = _context.stateModel
                             .Where(u => u.states.ToUpper().TrimEnd() == arrlist[j].ToUpper().TrimEnd())
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
                for (int j = 2; j < sectionLimit; j++)
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
            string strbankresrs = "";            
            if (uSPEIndividual.otherdetails.HasBankruptcyRecHits_resultpending == true)
            {
                strbankresrs = "Bankruptcy court records searches are currently pending in <jurisdiction> the results of which will be provided under separate cover upon receipt\n\n";
                strbankrs = "Results Pending\n\n".ToUpper();
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
            strbanksumcomm = string.Concat(strbankresrs, strbanksumcomm.TrimEnd());
            strbankressumm = string.Concat(strbankrs, strbankressumm);
            doc.Replace("BANKRUPCOMMENT", strbanksumcomm, true, true);
            doc.Replace("BANKRUPRESULT", strbankressumm, true, true);
            doc.SaveToFile(savePath);
            //HASBUREAUPRISONHITDESC
            if (uSPEIndividual.otherdetails.Has_Bureau_PrisonHit.ToString().Equals("Yes"))
            {
                doc.Replace("HASBUREAUPRISONHITDESC", "\nSearches on the subject were also conducted of selected criminal records from 47 states and the District of Columbia. (Delaware, South Dakota and Wyoming do not release criminal information to these sources.) The available information varies considerably from state to state, but generally includes the records of state departments of corrections, some criminal courts, county arrests, and/or traffic violations. In most cases, coverage is statewide and goes back at least 10 years. No records were located on the subject in these sources.\n\nIn addition, research efforts are conducted of Federal Bureau of Prisons (“BOP”) incarceration records, which reports felony inmate records on a nearly-nationwide basis.  In this regard, the following information was identified in connection with [LastName]:  <Investigator to insert results here>\n", true, true);
            }
            else
            {
                doc.Replace("HASBUREAUPRISONHITDESC", "\nSearches on the subject were also conducted of selected criminal records from 47 states and the District of Columbia. (Delaware, South Dakota and Wyoming do not release criminal information to these sources.) The available information varies considerably from state to state, but generally includes the records of state departments of corrections, some criminal courts, county arrests, and/or traffic violations. In most cases, coverage is statewide and goes back at least 10 years. No records were located on the subject in these sources.\n\nIn addition, research efforts are conducted of Federal Bureau of Prisons (“BOP”) incarceration records, which reports felony inmate records on a nearly-nationwide basis.  In this regard, no such records were revealed relating to the subject.\n", true, true);
            }
            //HASSEXOFFENDERREGHITDESC
            if (uSPEIndividual.otherdetails.Has_Sex_Offender_RegHit.ToString().Equals("Yes"))
            {
                doc.Replace("HASSEXOFFENDERREGHITDESC", "Research efforts were further conducted through the United States Department of Justice’s National Sex Offender Registry, which contains sex offender registry information from the 50 states, the District of Columbia and the territories of Guam and Puerto Rico. The public availability of sex offender-related information varies by state. In this regard, the following information was identified in connection with [LastName]:  <Investigator to insert results here>\n", true, true);
            }
            else
            {
                doc.Replace("HASSEXOFFENDERREGHITDESC", "Research efforts were further conducted through the United States Department of Justice’s National Sex Offender Registry, which contains sex offender registry information from the 50 states, the District of Columbia and the territories of Guam and Puerto Rico. The public availability of sex offender-related information varies by state. No records were identified for the subject.\n", true, true);
            }
            //USTAXCOURTHITDESC  [ADDNAMEONLYFN]
            if (uSPEIndividual.otherdetails.Has_US_Tax_Court_Hit.ToString().Equals("Yes"))
            {
                doc.Replace("[USTAXCOURTHITDESC]", "\n\nAdditionally, research efforts were conducted of filings through the United States Tax Court, and the following information was identified in connection with [LastName]:  <Investigator to insert results here>\n", true, false);
            }
            else
            {
                doc.Replace("[USTAXCOURTHITDESC]", "\n\nAdditionally, research efforts were conducted of filings through the United States Tax Court, and no records were identified for the subject.\n", true, false);
            }
            //HASTAXLIENSCIVILUCCDESC
            string strtaxlien = "";
            if (uSPEIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("No"))
            {
                strtaxlien = "Investigative efforts did not reveal any tax liens in connection with [LastName].";
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("Yes, Multiple Records"))
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
            if (uSPEIndividual.otherdetails.Has_Property_Records.Equals("Yes, Single Record"))
            {
                strciviljudge = "\n\nAdditionally, investigative efforts revealed the subject as having been a <party type> in connection with the following civil judgment:\n\n\t•	<Investigator to insert bulleted list of results here>";
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_Property_Records.Equals("Yes, Multiple Records"))
                {
                    strciviljudge = "\n\nAdditionally, investigative efforts revealed the subject in connection with the following civil judgments:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>";
                }
                else
                {
                    strciviljudge = "\n\nAdditionally, investigative efforts did not reveal any civil judgments in connection with [LastName].";
                }
            }
            string strjudge = "";
            if (uSPEIndividual.otherdetails.Has_Property_Records.Equals("Yes, Single Record") && uSPEIndividual.otherdetails.Has_Tax_Liens.Equals("Yes, Single Record") || uSPEIndividual.otherdetails.Has_Property_Records.Equals("Yes, Multiple Records") || uSPEIndividual.otherdetails.Has_Tax_Liens.Equals("Yes, Multiple Records"))
            {
                strjudge = "Records";
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_Property_Records.Equals("Yes, Single Record") || uSPEIndividual.otherdetails.Has_Tax_Liens.Equals("Yes, Single Record"))
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
       
            doc.Replace("TAXLIENCOMMENT", strcommjudge, true, true);
            doc.Replace("TAXLIENRESULT", strresjudge, true, true);                        
            //HASUCC
            string strucccomment = "";
            if (uSPEIndividual.otherdetails.has_ucc_fillings.ToString().Equals("Yes, Single Record"))
            {
                strucccomment = "\n\nFurther, investigative efforts revealed the subject as having been a <party type> in connection with the following Uniform Commercial Code filing:\n\n\t•	<Investigator to insert bulleted list of results here>\n";
            }
            else
            {
                if (uSPEIndividual.otherdetails.has_ucc_fillings.ToString().Equals("Yes, Multiple Records"))
                {
                    strucccomment = "\n\nFurther, investigative efforts revealed the subject in connection with the following Uniform Commercial Code filings:\n\n\t•	<Investigator to insert bulleted list of results here>\n\t•	<Investigator to insert bulleted list of results here>\n";
                }
                else
                {
                    strucccomment = "\n\nAdditionally, investigative efforts did not reveal [LastName] in connection with any Uniform Commercial Code filings.\n";
                }
            }
            string commomstr = "";
            //if (uSPEIndividual.diligenceInputModel.CommonNameSubject == true)
            //{
            //    commomstr = "\nIn light of the commonality of [LastName]’s name, this portion of the investigation was conducted utilizing the subject’s Social Security number, in order to identify only those records conclusively relating to the subject of interest.";
            //}                    
            if (uSPEIndividual.otherdetails.Has_Tax_Liens.ToString().Equals("No") && uSPEIndividual.otherdetails.Has_Property_Records.Equals("No") && uSPEIndividual.otherdetails.has_ucc_fillings.ToString().Equals("No"))
            {
                if(uSPEIndividual.diligenceInputModel.CommonNameSubject == true)
                {
                    commomstr = string.Concat("\n", commomstr);
                }
                strciviljudge = "Investigative efforts did not reveal any tax liens, civil judgments or Uniform Commercial Code filings in connection with [LastName].\n";                
                doc.Replace("HASTAXLIENSCIVILUCCDESC", strciviljudge, true, true);
            }
            else
            {
                strciviljudge = string.Concat(strtaxlien, strciviljudge, strucccomment, commomstr,"\n");
                doc.Replace("HASTAXLIENSCIVILUCCDESC", strciviljudge, true, true);
            }

            string struccsummres = "";
            string struccsummcom = "";
            switch (uSPEIndividual.otherdetails.has_ucc_fillings)
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
            doc.Replace("UNIFCOMMENT", struccsummcom, true, true);
            doc.Replace("UNIFRESULT", struccsummres, true, true);

            //CREDITHISTORYDESCRIPTION 
            string credit_sumresult = "";
            string credit_summcomment = "";
            if (uSPEIndividual.otherdetails.Was_credited_obtained.ToString().Equals("No"))
            {
                doc.Replace("CREDITHISTORYDESCRIPTION", "", true, true);
                credit_sumresult = "N/A";
                credit_summcomment = "Credit history information cannot be obtained without the subject’s express written authorization";
            }
            else
            {
                if (uSPEIndividual.otherdetails.Was_credited_obtained.ToString().Equals("Credit with Adverse Hits"))
                {
                    doc.Replace("CREDITHISTORYDESCRIPTION", "\nWith [LastName]’s express written authorization, the subject’s personal credit history was retrieved, which revealed that <Investigator to insert summary of trade account information here>.\n\nFurther, <Investigator to insert summary of adverse results here>.\n\nThere were no additional payment delinquencies or collections accounts reported in the subject's credit file for the past seven years or so.\n", true, true);
                    credit_sumresult = "Records";
                    credit_summcomment = "With the subject’s express written authorization, [LastName]’s personal credit history was retrieved, which revealed <investigator to insert summary here>";
                }
                else
                {
                    doc.Replace("CREDITHISTORYDESCRIPTION", "\nWith [LastName]’s express written authorization, the subject’s personal credit history was retrieved, which revealed that <Investigator to insert summary of trade account information here>.\n\nThere were no payment delinquencies or collections accounts reported in the subject's credit file for the past seven years or so.\n", true, true);
                    credit_sumresult = "Clear";
                    credit_summcomment = "";
                }
            }
            doc.Replace("PERCREDITRESULT", credit_sumresult, true, true);
            doc.Replace("PERCREDITCOMMENT", credit_summcomment, true, true);
            //PRESSANDMEDIASEARCHDESCRIPTION
            switch (uSPEIndividual.otherdetails.Press_Media.ToString())
            {
                case "Common name with adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the commonality of the subject's name, searches were structured and designed to identify specific reports relating to [LastName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. Further, the subject’s name was also searched using a variety of other terms, focusing on the subject’s employment history, and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>[DESCFAMADESC]", true, true);
                    break;
                case "Common name without adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the commonality of the subject's name, searches were structured and designed to identify specific reports relating to [LastName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. Further, the subject’s name was also searched using a variety of other terms, focusing on the subject’s employment history, and a thorough review of the same did not identify any adverse or materially-significant information in connection with [LastName].[DESCFAMADESC]", true, true);
                    break;
                case "High volume with adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the volume of press and other media sources identified in connection with the subject, searches were structured and designed to identify specific reports relating to [LastName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. Further, the subject’s name was also searched using a variety of other terms, focusing on the subject’s employment history, and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>[DESCFAMADESC]", true, true);
                    break;
                case "High volume without adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Due to the volume of press and other media sources identified in connection with the subject, searches were structured and designed to identify specific reports relating to [LastName] with a variety of negative search terms and derivatives thereof within the past <timeframe>. Further, the subject’s name was also searched using a variety of other terms, focusing on the subject’s employment history, and a thorough review of the same did not identify any adverse or materially-significant information in connection with [LastName].[DESCFAMADESC]", true, true);
                    break;
                case "Standard search with adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, identified various articles and/or media references in connection with [LastName], and a thorough review of the same revealed the following adverse or potentially-relevant information:   <Investigator to insert article summaries here>[DESCFAMADESC]", true, true);
                    break;
                case "Standard search without adverse Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, identified various articles and/or media references in connection with [LastName], and a thorough review of the same did not identify any adverse or materially-significant information.[DESCFAMADESC]", true, true);
                    break;
                case "No Hits":
                    doc.Replace("PRESSANDMEDIASEARCHDESCRIPTION", "Extensive press and other media research, including searches of international, national, regional and local newspapers, magazines, trade and finance journals, as well as Internet sources, did not identify any articles and/or media references in connection with [LastName].[DESCFAMADESC]", true, true);
                    break;
            }
            doc.SaveToFile(savePath);
            //Fama
            switch (uSPEIndividual.otherdetails.Fama.ToString())
            {
                case "Yes - Clear":
                    doc.Replace("[DESCFAMADESC]", "\n\nIn addition to the above, a FAMA report was completed, covering a search of publicly available online content, and an in-depth review of such did not identify any adverse or materially-significant information in connection with the subject. (See Exhibit A for the subject’s FAMA Report.)", true, true);
                    break;
                case "Yes - Adverse":
                    doc.Replace("[DESCFAMADESC]", "\n\nIn addition to the above, a FAMA report was completed, covering a search of publicly available online content, and an in-depth review of such revealed that the subject <investigator to enter details.> (See Exhibit A for the subject’s FAMA Report.)", true, true);
                    break;
                case "No":
                    doc.Replace("[DESCFAMADESC]", "", true, true);
                    break;
            }
            doc.SaveToFile(savePath);           
            //DRIVINGHISTORYDESCRIPTION            
            string drivingcomm = "";
            string divingresult = "";
            switch (uSPEIndividual.otherdetails.Has_Driving_Hits.ToString())
            {
                case "No Records, No Consent":
                    drivingcomm = "\nIt is noted that [Last Name]’s express written authorization would be required to obtain the subject’s driving history information.\n";
                    divingresult = "N/A [No Consent or hits]";
                    break;
                case "Clear, With Consent":
                    drivingcomm = "\nIt is noted that [Last Name] has a valid <State> driver’s license, which is set to expire in <expiration date (month and year only)>, unless renewed.  A driving history report was requested from the <State driving agency>, covering the past three years or so, which did not reveal any traffic violations, license suspensions or motor vehicle incidents involving the subject.\n";
                    divingresult = "Clear";
                    break;
                case "Multiple Records, No Consent":
                    drivingcomm = "\nIt is noted that [Last Name]’s express written authorization would be required to obtain the subject’s driving history information.\n\nHowever, additional research efforts identified <Investigator to insert results here>.\n";
                    divingresult = "Records [Hit w/o Consent]";
                    break;
                case "Single Record, No Consent":
                    drivingcomm = "\nIt is noted that [Last Name]’s express written authorization would be required to obtain the subject’s driving history information.\n\nHowever, additional research efforts identified <Investigator to insert results here>.\n";
                    divingresult = "Record [Hit w/o Consent]";
                    break;
                case "Multiple Records, with Consent":
                    drivingcomm = "\nIt is noted that [Last Name] has a valid <State> driver’s license, which is set to expire in <expiration date (month and year only)>, unless renewed.  A driving history report was requested from the <State driving agency>, covering the past three years or so, which revealed that the subject was cited for the following:\n\n\t•\t<Investigator to insert results>\n\t•\t<Investigator to insert results>\n\nThere were no other traffic violations, license suspensions or motor vehicle incidents involving the subject.\n";
                    divingresult = "Records [Hits w/ Consent]";
                    break;
                case "Single Record, with Consent":
                    drivingcomm = "\nIt is noted that [Last Name] has a valid <State> driver’s license, which is set to expire in <expiration date (month and year only)>, unless renewed.  A driving history report was requested from the <State driving agency>, covering the past three years or so, which revealed that the subject was cited for <Investigator to insert results>.\n\nThere were no other traffic violations, license suspensions or motor vehicle incidents involving the subject.\n";
                    divingresult = "Record [Hit w/ Consent]";
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
                    strcomment = strcomment.Replace("[LastName]", uSPEIndividual.diligenceInputModel.LastName);
                    doc.Replace("DRIVINGCOMMENT", strcomment, true, true);
                    doc.Replace("DRIVINGRESULT", "Record", true, true);
                    break;
                case "Records [Hit w/o Consent]":
                    strcomment = "While [LastName]’s driving history record cannot be obtained without the subject’s express written authorization, research efforts conducted in relevant jurisdictions identified the subject <investigator to insert summary here>";
                    strcomment = strcomment.Replace("[LastName]", uSPEIndividual.diligenceInputModel.LastName);
                    doc.Replace("DRIVINGCOMMENT", strcomment, true, true);
                    doc.Replace("DRIVINGRESULT", "Records", true, true);
                    break;
                case "N/A [No Consent or hits]":
                    doc.Replace("DRIVINGCOMMENT", "Driving history records cannot be obtained without the subject’s express written authorization", true, true);
                    doc.Replace("DRIVINGRESULT", "N/A", true, true);
                    break;
            }
            //CRIMINALRECDES
            if (uSPEIndividual.otherdetails.Has_CriminalRecHit_resultpending == true)
            {
                doc.Replace("CRIMINALRECDES", "\n\nA criminal records search is currently pending through the <investigator to insert source/court/law enforcement agency>, the results of which will be provided under separate cover upon receipt. CRIMINALRECDES", true, true);
                doc.SaveToFile(savePath);
            }
            string strconcatcrim = "";
            string strcrimsumcom = "";
            string strcrimsumres = "";
            switch (uSPEIndividual.otherdetails.Has_CriminalRecHit.ToString())
            {
                case "Yes, Multiple Records":
                    strcrimsumres = "Records";
                    doc.Replace("CRIMINALRECDES", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, true);
                    strconcatcrim = " identified the subject as a Defendant in connection with the following criminal records:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                    break;
                case "No":
                    strcrimsumres = "Clear";
                    doc.Replace("CRIMINALRECDES", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, true);
                    strconcatcrim = " did not identify the subject in connection with any criminal records.";
                    break;
                case "Yes, Single Record":
                    strcrimsumres = "Record";
                    doc.Replace("CRIMINALRECDES", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, true);
                    strconcatcrim = " identified the subject as a Defendant in connection with the following criminal record:\n\n\t•\t<Investigator to insert bulleted list of results here>";
                    break;

            }
            //if (uSPEIndividual.diligenceInputModel.CommonNameSubject == true)
            //{
            //    strconcatcrim = string.Concat(strconcatcrim,"\n\nIn light of the commonality of [Last Name]’s name, this portion of the investigation was conducted utilizing the subject’s date of birth and Social Security number, in order to identify only those records conclusively relating to the subject of interest.  Records may exist that do not contain this type of identifying information, and an expanded scope would be required in this regard.");
            //}          
            doc.SaveToFile(savePath);                       
            strblnres = "";
            if (arrlist.Count() == 0)
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
                                for (int j = 2; j < sectionLimit; j++)
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
                                                    string strapendfootnotetext = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert County Courts searched, if applicable>.  [STATEWIDECRIMINALLANGUAGE]";
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
                                    try
                                    {
                                        for (int j = 2; j < sectionLimit; j++)
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
                                                            string strapendfootnotetext = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert County Courts searched, if applicable>.  [STATEWIDECRIMINALLANGUAGE]";
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
                                        for (int j = 2; j < sectionLimit; j++)
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
                                                            string strapendfootnotetext = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert County Courts searched, if applicable>.  [STATEWIDECRIMINALLANGUAGE]";
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
                                    catch { }

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
                    for (int j = 2; j < sectionLimit; j++)
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
                                        string strapendfootnotetext = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert County Courts searched, if applicable>.  [STATEWIDECRIMINALLANGUAGE]";
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
                    strcrimsumcom = "<investigator to specify whether Criminal Clearance Certificate was provided and authenticated>";
                    break;
                case "Record":
                    strcrimsumcom = "[LastName] was identified as a Defendant in connection with a <case type>, which was recorded in <State> in <YYYY>, and is currently <status>";
                    break;
                case "Records":
                    strcrimsumcom = "[LastName] was identified as a Defendant in connection with at least <record types>, which were recorded in <State> between <YYYY> and <YYYY>, and are currently <status>";
                    break;
            }
            if (uSPEIndividual.otherdetails.Has_CriminalRecHit_resultpending == true)
            {
                strcrimsumres = string.Concat("Result Pending\n\n", strcrimsumres);
                strcrimsumcom = string.Concat("Criminal records searches are currently pending through the <investigator to insert source/court/law enforcement agency>, the results of which will be provided under separate cover upon receipt\n\n", strcrimsumcom);
            }          
            doc.Replace("CRIMINALCOMMENT", strcrimsumcom, true, true);
            doc.Replace("CRIMINALRESULT", strcrimsumres, true, true);            
            doc.SaveToFile(savePath);
            TextSelection[] crimtext = doc.FindAllString("Criminal records searches are currently pending through the <investigator to insert source/court/law enforcement agency>, the results of which will be provided under separate cover upon receipt", true, true);
            if (crimtext != null)
            {
                foreach (TextSelection seletion in crimtext)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                }
            }            
            //HASCIVILRECDESC
            if (uSPEIndividual.otherdetails.has_civil_resultpending == true)
            {
                doc.Replace("[HASCIVILRECDESC]", "\n\nA civil court records search is currently pending through the <enter source/court> the results of which will be provided under separate cover upon receipt.[HASCIVILRECDESC]", true, false);
                doc.SaveToFile(savePath);
            }
            string strconcatcivilrec = "";
            strAdditionalStatesComment = "";
            string strsummcivilcom = "";
            string strsummcivilres = "";
            switch (uSPEIndividual.otherdetails.Has_Civil_Records.ToString())
            {
                case "No":
                    strsummcivilres = "Clear";
                    doc.Replace("[HASCIVILRECDESC]", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                    strconcatcivilrec = " did not identify the subject, personally, in connection with any civil filings.";
                    break;
                case "Yes, Multiple Records":
                    strsummcivilres = "Records";
                    doc.Replace("[HASCIVILRECDESC]", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                    strconcatcivilrec = " identified the subject, personally, in connection with the following civil filings:\n\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>\n\t•\t<Investigator to insert bulleted list of results here>";
                    break;
                case "Yes, Single Record":
                    strsummcivilres = "Record";
                    doc.Replace("[HASCIVILRECDESC]", "\n\nSpecific efforts pursued in relevant jurisdictions in [ADDITIONALSTATES]", true, false);
                    strconcatcivilrec = " identified the subject, personally, in connection with the following civil filing:\n\n\t•\t<Investigator to insert bulleted list of results here>";
                    break;
            }
            doc.SaveToFile(savePath);
            strblnres = "";

            if (uSPEIndividual.otherdetails.Has_Civil_Records.ToString().Contains("Results Pending"))
            {
                try
                {
                    for (int j = 2; j < sectionLimit; j++)
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
                                    if (abc.ToString().Equals("[CIVILRECORDSHITS]") || abc.ToString().Contains("CIVILRECORDSHITS"))
                                    {
                                        TextRange tr = para.AppendText("While a civil court records search is currently pending through the <enter source/court> (the results of which will be provided under separate cover upon receipt), [APPENDTEXTINNORMALFONT]");
                                        tr.CharacterFormat.Italic = true;
                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                        tr.CharacterFormat.FontSize = 11;
                                        tr.CharacterFormat.HighlightColor = Color.Aqua;
                                        doc.SaveToFile(savePath);
                                        doc.Replace("[CIVILRECORDSHITS]", "", true, false);
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
                strblnres = ""; try
                {
                    for (int j = 2; j < sectionLimit; j++)
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
                                    if (abc.ToString().Equals("[APPENDTEXTINNORMALFONT]") || abc.ToString().EndsWith("[APPENDTEXTINNORMALFONT]"))
                                    {
                                        TextRange tr = para.AppendText("specific efforts pursued in other relevant jurisdictions in [ADDITIONALSTATES]");
                                        tr.CharacterFormat.Italic = false;
                                        tr.CharacterFormat.FontName = "Calibri (Body)";
                                        tr.CharacterFormat.FontSize = 11;
                                        doc.SaveToFile(savePath);
                                        doc.Replace("[APPENDTEXTINNORMALFONT]", "", true, false);
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
            strblnres = "";
            if (arrlist.Count() == 0)
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
                            string strcivilfootnote = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert County Courts searched>.";
                            strcivilfootnote = strcivilfootnote.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                            if (k == count - 1)
                            {
                                strblnres = "";
                                strAdditionalStatesComment = string.Concat(arrlist[a], "[ADDITIONALSTATE]");
                                doc.Replace("[ADDITIONALSTATES]", strAdditionalStatesComment, true, true);
                                doc.SaveToFile(savePath);
                                for (int j = 2; j < sectionLimit; j++)
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
                                    for (int j = 2; j < sectionLimit; j++)
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
                                    for (int j = 2; j < sectionLimit; j++)
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
                    string strcivilfootnote = "Searches were pursued through the United States District [STATESPECIFICDISTRICTCOURTS], as well as the <Investigator to insert County Courts searched>.";
                    strcivilfootnote = strcivilfootnote.Replace("[STATESPECIFICDISTRICTCOURTS]", strstatecourt);
                    doc.SaveToFile(savePath);
                    for (int j = 2; j < sectionLimit; j++)
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

            if (uSPEIndividual.otherdetails.Has_US_Tax_Court_Hit.Equals("Yes, Single Record") && strsummcivilres == "Clear")
            {
                strsummcivilres = "Record";
            }
            else
            {
                if (uSPEIndividual.otherdetails.Has_US_Tax_Court_Hit.Equals("Yes, Single Record") && strsummcivilres == "Records" || strsummcivilres == "Record")
                {
                    strsummcivilres = "Records";
                }
                if (uSPEIndividual.otherdetails.Has_US_Tax_Court_Hit.Equals("Yes, Multiple Records"))
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
                    strsummcivilcom = "[LastName] was identified as a <party type> in connection with a civil filing, which were recorded in <State> in <YYYY>, and is currently <status>";
                    break;
                case "Records":
                    strsummcivilcom = "[LastName] was identified as a <party type> in connection with at least <number> civil filings, which were recorded in <States> between <YYYY> and <YYYY>, and are currently <status>";
                    break;
            }
            if (uSPEIndividual.otherdetails.has_civil_resultpending == true)
            {
                strsummcivilres = string.Concat("Result Pending\n\n", strsummcivilres);
                strsummcivilcom = string.Concat("A civil court records search is currently pending through the <enter source/court> the results of which will be provided under separate cover upon receipt\n\n", strsummcivilcom);
            }      
            doc.Replace("CIVILRESULT", strsummcivilres, true, true);
            doc.Replace("CIVILCOMMENT", strsummcivilcom, true, true);
            doc.SaveToFile(savePath);
            //EmployeeDetails
            string emp_Comment = "";
            string strempstartdate = "<not provided>";
            string strempenddate = "<not provided>";
            string empLocation = "";
           
            for (int i = 0; i < uSPEIndividual.EmployerModel.Count; i++)
            {
                strempstartdate = "<not provided>";
                strempenddate = "<not provided>";
                if (uSPEIndividual.EmployerModel[i].Emp_State.Equals("Select State"))
                {
                    if (uSPEIndividual.EmployerModel[i].Emp_City.Equals("") || uSPEIndividual.EmployerModel[i].Emp_City.ToUpper().Equals("NA") || uSPEIndividual.EmployerModel[i].Emp_City.ToUpper().Equals("N/A"))
                    {
                        empLocation = "<not provided>";
                    }
                    else
                    {
                        empLocation = uSPEIndividual.EmployerModel[i].Emp_City;
                    }
                }
                else
                {
                    if (uSPEIndividual.EmployerModel[i].Emp_City.Equals("") || uSPEIndividual.EmployerModel[i].Emp_City.ToUpper().Equals("NA") || uSPEIndividual.EmployerModel[i].Emp_City.ToUpper().Equals("N/A"))
                    {
                        empLocation = uSPEIndividual.EmployerModel[i].Emp_State;
                    }
                    else
                    {
                        empLocation = string.Concat(uSPEIndividual.EmployerModel[i].Emp_City,", ", uSPEIndividual.EmployerModel[i].Emp_State);
                    }
                }

                if (!uSPEIndividual.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && !uSPEIndividual.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                {
                    strempstartdate = uSPEIndividual.EmployerModel[i].Emp_StartDateMonth + " " + uSPEIndividual.EmployerModel[i].Emp_StartDateDay + ", " + uSPEIndividual.EmployerModel[i].Emp_StartDateYear;
                    // employerModel.Emp_StartDate = strempstartdate;
                }
                if (uSPEIndividual.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && uSPEIndividual.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && uSPEIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                {
                    strempstartdate = "<not provided>";
                }
                if (uSPEIndividual.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && !uSPEIndividual.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                {
                    strempstartdate = uSPEIndividual.EmployerModel[i].Emp_StartDateMonth + " " + uSPEIndividual.EmployerModel[i].Emp_StartDateYear;
                    //employerModel.Emp_StartDate = strempstartdate;
                }
                if (uSPEIndividual.EmployerModel[i].Emp_StartDateDay.ToString().Equals("Day") && uSPEIndividual.EmployerModel[i].Emp_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year"))
                {
                    strempstartdate = uSPEIndividual.EmployerModel[i].Emp_StartDateYear;
                    //employerModel.Emp_StartDate = strempstartdate;
                }
                if (i == 0 && uSPEIndividual.EmployerModel[0].Emp_Status.ToString().Equals("Current")) { }
                else
                {
                    try
                    {
                        if (!uSPEIndividual.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && !uSPEIndividual.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                        {
                            strempenddate = uSPEIndividual.EmployerModel[i].Emp_EndDateMonth + " " + uSPEIndividual.EmployerModel[i].Emp_EndDateDay + ", " + uSPEIndividual.EmployerModel[i].Emp_EndDateYear;
                            //  employerModel.Emp_EndDate = strempenddate;
                        }
                        if (uSPEIndividual.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && uSPEIndividual.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && uSPEIndividual.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                        {
                            strempenddate = "<not provided>";
                        }
                        if (uSPEIndividual.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && !uSPEIndividual.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                        {
                            strempenddate = uSPEIndividual.EmployerModel[i].Emp_EndDateMonth + " " + uSPEIndividual.EmployerModel[i].Emp_EndDateYear;
                            //employerModel.Emp_EndDate = strempenddate;
                        }
                        if (uSPEIndividual.EmployerModel[i].Emp_EndDateDay.ToString().Equals("Day") && uSPEIndividual.EmployerModel[i].Emp_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.EmployerModel[i].Emp_EndDateYear.ToString().Equals("Year"))
                        {
                            strempenddate = uSPEIndividual.EmployerModel[i].Emp_EndDateYear;
                            //employerModel.Emp_EndDate = strempenddate;
                        }
                    }
                    catch { }
                }
                if (i == 0)
                {
                    if (uSPEIndividual.EmployerModel[i].Emp_Confirmed.ToString().Equals("Not Reported By Subject")) {
                        emp_Comment = string.Concat("•	[Employer1]\n	Not Reported by Subject");
                    }
                    else
                    {
                        emp_Comment = string.Concat("•	[Employer1]\n	[EmpLocation1]\n	[Position1]\n	[EmpStartDate1] to Present");                      
                    }
                    switch (uSPEIndividual.EmployerModel[i].Emp_Confirmed.ToString())
                    {
                        case "Not Reported By Subject":
                            emp_Comment = string.Concat(emp_Comment, "\n\nWhile not reported on his application, resume or LinkedIn profile, research efforts utilizing the subject’s Social Security number through an automated third-party employment verification service housing numerous personnel records for thousands of companies revealed that the subject is <currently/previously> employed as a [Position] at [Employer], presumably in [EmpLocation], since [EmpStartDate].");
                            break;
                        case "No":
                            emp_Comment = string.Concat(emp_Comment, "\n\nIn order to preserve the confidentiality of this investigation, [Employer1] was not contacted to confirm the subject's tenure with the same, however, research efforts can be conducted at a later date upon request.  <Investigator to insert any additional information found here>.");
                            break;
                        case "Yes":
                            emp_Comment = string.Concat(emp_Comment, "\n\nIn order to preserve the confidentiality of this investigation, [Employer1] was not contacted directly to confirm the subject's tenure with the same.  However, <Investigator to insert verification information found here>.");
                            break;
                        case "No – Client is current employer":
                            emp_Comment = string.Concat(emp_Comment, "\n\nIn order to preserve the confidentiality of this investigation, the company was not contacted to verify the subject’s tenure with the same, however, research efforts can be conducted at a later date upon request.");
                            break;
                        case "No - Current Employer":
                            emp_Comment = string.Concat(emp_Comment, "\n\nIn order to preserve the confidentiality of this investigation, the company was not contacted to verify the subject’s tenure with the same, however, research efforts can be conducted at a later date upon request.");
                            break;
                        case "Result Pending":
                            emp_Comment = string.Concat(emp_Comment, "\n\n", "APPENDEMPTEXT", i.ToString());
                            break;
                    }
                    try
                    {
                        emp_Comment = emp_Comment.Replace("[Position1]", uSPEIndividual.EmployerModel[i].Emp_Position.ToString());
                    }
                    catch {
                        emp_Comment = emp_Comment.Replace("[Position1]", "<not provided>");
                    }
                    try
                    {
                        emp_Comment = emp_Comment.Replace("[Employer1]", uSPEIndividual.EmployerModel[i].Emp_Employer.ToString());
                    }
                    catch
                    {
                        emp_Comment = emp_Comment.Replace("[Employer1]", "<not provided>");
                    }
                    emp_Comment = emp_Comment.Replace("[EmpLocation1]", empLocation);
                    emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                }
                else
                {
                    if (uSPEIndividual.EmployerModel[i].Emp_Confirmed.ToString().Equals("Not Reported By Subject"))
                    {
                        emp_Comment = string.Concat(emp_Comment, "\n\n", "•	[Employer]\n	Not Reported by Subject");
                    }
                    else
                    {
                        emp_Comment = string.Concat(emp_Comment, "\n\n", "•	[Employer]\n	[EmpLocation]\n	[Position]\n	[EmpStartDate] to [EmpEndDate]");                        
                    }
                    switch (uSPEIndividual.EmployerModel[i].Emp_Confirmed)
                    {
                        case "Not Reported By Subject":
                            emp_Comment = string.Concat(emp_Comment, "\n\nWhile not reported on his application, resume or LinkedIn profile, research efforts utilizing the subject’s Social Security number through an automated third-party employment verification service housing numerous personnel records for thousands of companies revealed that the subject is <currently/previously> employed as a [Position] at [Employer], presumably in [EmpLocation], since [EmpStartDate].");
                            break;
                        case "Yes – By Representative":
                            emp_Comment = string.Concat(emp_Comment, "\n\nAs confirmed <Investigator to add method of confirmation>, by a representative of the company, [Last Name] was previously employed as a [Position] with [Employer] from [EmpStartDate] to [EmpEndDate].  <Investigator to insert any additional information found here>.");
                            break;
                        case "Yes – By TWN":
                            emp_Comment = string.Concat(emp_Comment, "\n\nAs confirmed <Investigator to add method of confirmation>, through an automated third-party verification service utilized by <investigator to insert company name>, [LastName] was previously employed as a [Position] from [EmpStartDate] to [EmpEndDate].");
                            break;
                        case "Yes":
                            emp_Comment = string.Concat(emp_Comment, "\n\nAs confirmed <Investigator to add method of confirmation>, the subject was employed as a [Position] with [Employer] from [EmpStartDate] to [EmpEndDate].  <Investigator to add additional detail here, if needed.>");
                            break;
                        case "No":
                            emp_Comment = string.Concat(emp_Comment, "\n\nEfforts to confirm the subject’s employment with [Employer] were unsuccessful.  <Investigator to insert reason for failed verification here>.");
                            break;
                        case "Result Pending":
                            emp_Comment = string.Concat(emp_Comment, "\n\n", "APPENDEMPTEXT", i.ToString());
                            break;
                        case "No – Client is current employer":
                            emp_Comment = string.Concat(emp_Comment, "\n\nIn order to preserve the confidentiality of this investigation, the company was not contacted to verify the subject’s tenure with the same, however, research efforts can be conducted at a later date upon request.");
                            break;
                        case "No - Current Employer":
                            emp_Comment = string.Concat(emp_Comment, "\n\nIn order to preserve the confidentiality of this investigation, the company was not contacted to verify the subject’s tenure with the same, however, research efforts can be conducted at a later date upon request.");
                            break;
                    }
                    emp_Comment = emp_Comment.Replace("[LastName]", uSPEIndividual.diligenceInputModel.LastName.ToString());
                    try
                    {
                        emp_Comment = emp_Comment.Replace("[Position]", uSPEIndividual.EmployerModel[i].Emp_Position.ToString());
                    }
                    catch
                    {
                        emp_Comment = emp_Comment.Replace("[Position]","<not provided>");
                    }
                    try
                    {
                        emp_Comment = emp_Comment.Replace("[Employer]", uSPEIndividual.EmployerModel[i].Emp_Employer.ToString());
                    }
                    catch
                    {
                        emp_Comment = emp_Comment.Replace("[Employer]", "<not provided>");
                    }
                    try
                    {
                        emp_Comment = emp_Comment.Replace("[EmpLocation]", empLocation);
                    }
                    catch {
                        emp_Comment = emp_Comment.Replace("[EmpLocation]", "<not provided>");
                    }
                    emp_Comment = emp_Comment.Replace("[EmpStartDate]", strempstartdate);
                    emp_Comment = emp_Comment.Replace("[EmpEndDate]", strempenddate);
                }
            }
            doc.Replace("EMPLOYEEDESCRIPTION", emp_Comment, true, true);
            doc.SaveToFile(savePath);
            for (int i = 0; i < uSPEIndividual.EmployerModel.Count; i++)
            {
                TextSelection[] texteemployer = doc.FindAllString(string.Concat("•	", uSPEIndividual.EmployerModel[i].Emp_Employer), false, false);
                if (texteemployer != null)
                {
                    foreach (TextSelection seletion in texteemployer)
                    {
                        seletion.GetAsOneRange().CharacterFormat.Bold = true;
                    }
                }
            }
            doc.SaveToFile(savePath);
            string bnres = "";
            for (int i = 0; i < uSPEIndividual.EmployerModel.Count; i++)
            {
                bnres = "";
                if (uSPEIndividual.EmployerModel[i].Emp_Confirmed.ToString().Equals("Result Pending"))
                {
                    try
                    {
                        for (int j = 1; j < sectionLimit; j++)
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
                                            textRange = para.AppendText(string.Concat("Efforts to confirm the subject’s employment with ", uSPEIndividual.EmployerModel[i].Emp_Employer.ToString(), " are currently ongoing, the results of which will be provided under separate cover, if and when received."));
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
                    catch { }
                }
            }
            doc.SaveToFile(savePath);

            bnres = "";
            string[] arrfind = { "While not reported on his application, resume or LinkedIn profile, research efforts", "In order to preserve the confidentiality of this investigation,", "As confirmed <Investigator to add method of confirmation>,","As confirmed through an automated third-party verification service utilized by <investigator to insert company name>,", "Efforts to confirm the subject’s employment with" };
            foreach (string item in arrfind)
            {
                try
                {
                    for (int j = 1; j < sectionLimit; j++)
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
                                    if (abc.ToString().Contains(item))
                                    {
                                        para.Format.LeftIndent = 72;
                                        doc.SaveToFile(savePath);
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
            //Education Details
            string edu_comment = "";            
            string edu_header = "";
            string edu_summcomment = "";
            
            bnres = "";
            for (int i = 0; i < uSPEIndividual.educationModels.Count; i++)
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
                if (uSPEIndividual.educationModels[i].Edu_History.ToString().Equals("Yes"))
                {
                    if (!uSPEIndividual.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && !uSPEIndividual.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                    {
                        edustartdate = uSPEIndividual.educationModels[i].Edu_StartDateMonth + " " + uSPEIndividual.educationModels[i].Edu_StartDateDay + ", " + uSPEIndividual.educationModels[i].Edu_StartDateYear;
                        edustartyr = uSPEIndividual.educationModels[i].Edu_StartDateYear;                        
                    }
                    if (uSPEIndividual.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && uSPEIndividual.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && uSPEIndividual.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                    {
                        edustartdate = "<not provided>";
                        edustartyr = "<not provided>";
                    }
                    if (uSPEIndividual.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && !uSPEIndividual.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                    {
                        edustartdate = uSPEIndividual.educationModels[i].Edu_StartDateMonth + " " + uSPEIndividual.educationModels[i].Edu_StartDateYear;
                        edustartyr = uSPEIndividual.educationModels[i].Edu_StartDateYear;                        
                    }
                    if (uSPEIndividual.educationModels[i].Edu_StartDateDay.ToString().Equals("Day") && uSPEIndividual.educationModels[i].Edu_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_StartDateYear.ToString().Equals("Year"))
                    {
                        edustartdate = uSPEIndividual.educationModels[i].Edu_StartDateYear;
                        edustartyr = uSPEIndividual.educationModels[i].Edu_StartDateYear;                        
                    }

                    if (!uSPEIndividual.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && !uSPEIndividual.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                    {
                        eduenddate = uSPEIndividual.educationModels[i].Edu_EndDateMonth + " " + uSPEIndividual.educationModels[i].Edu_EndDateDay + ", " + uSPEIndividual.educationModels[i].Edu_EndDateYear;
                        eduendyr = uSPEIndividual.educationModels[i].Edu_EndDateYear;                        
                    }
                    if (uSPEIndividual.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && uSPEIndividual.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && uSPEIndividual.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                    {
                        eduenddate = "<not provided>";
                        eduendyr = "<not provided>";
                    }
                    if (uSPEIndividual.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && !uSPEIndividual.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                    {
                        eduenddate = uSPEIndividual.educationModels[i].Edu_EndDateMonth + " " + uSPEIndividual.educationModels[i].Edu_EndDateYear;
                        eduendyr = uSPEIndividual.educationModels[i].Edu_EndDateYear;                        
                    }
                    if (uSPEIndividual.educationModels[i].Edu_EndDateDay.ToString().Equals("Day") && uSPEIndividual.educationModels[i].Edu_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_EndDateYear.ToString().Equals("Year"))
                    {
                        eduenddate = uSPEIndividual.educationModels[i].Edu_EndDateYear;
                        eduendyr = uSPEIndividual.educationModels[i].Edu_EndDateYear;                        
                    }

                    if (!uSPEIndividual.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && !uSPEIndividual.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                    {
                        edugraddate = uSPEIndividual.educationModels[i].Edu_GradDateMonth + " " + uSPEIndividual.educationModels[i].Edu_GradDateDay + ", " + uSPEIndividual.educationModels[i].Edu_GradDateYear;
                        edugradyr = uSPEIndividual.educationModels[i].Edu_GradDateYear;                        
                    }
                    if (uSPEIndividual.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && uSPEIndividual.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && uSPEIndividual.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                    {
                        edugraddate = "<not provided>";
                        edugradyr = "<not provided>";
                    }
                    if (uSPEIndividual.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && !uSPEIndividual.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                    {
                        edugraddate = uSPEIndividual.educationModels[i].Edu_GradDateMonth + " " + uSPEIndividual.educationModels[i].Edu_GradDateYear;
                        edugradyr = uSPEIndividual.educationModels[i].Edu_GradDateYear;                        
                    }
                    if (uSPEIndividual.educationModels[i].Edu_GradDateDay.ToString().Equals("Day") && uSPEIndividual.educationModels[i].Edu_GradDateMonth.ToString().Equals("0") && !uSPEIndividual.educationModels[i].Edu_GradDateYear.ToString().Equals("Year"))
                    {
                        edugraddate = uSPEIndividual.educationModels[i].Edu_GradDateYear;
                        edugradyr = uSPEIndividual.educationModels[i].Edu_GradDateYear;                        
                    }                   
                    if (i == uSPEIndividual.educationModels.Count - 1) {
                        switch (uSPEIndividual.educationModels[i].Edu_Confirmed.ToString())
                        {
                            case "Yes":
                                edu_comment = string.Concat(edu_comment, "As confirmed, <investigator to add method of confirmation>, [LastName] received a [Degree] from the [School] in [EduLocation] on [GradDate], having attended the same from [EduStartDate] to [EduEndDate].", "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Confirmed");
                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.confirmed_comment.ToString(), "\n\n");
                                break;
                            case "No":
                                edu_comment = string.Concat(edu_comment, comment1.unconfirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Unconfirmed");
                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.unconfirmed_comment.ToString(), "\n\n");
                                break;
                            case "Result Pending":
                                edu_comment = string.Concat(edu_comment, "Efforts to independently confirm the subject’s tenure at the same are currently ongoing, the results of which will be provided under separate cover, if and when received.  However, the subject self-reports having received a [Degree] degree in <major> from [School] in [EduLocation] on [GradDate], having reportedly attended the same from [EduStartDate] to [EduEndDate].", "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Results Pending");
                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.confirmed_comment.ToString(), " APPENDRESPENDING", "", "\n\n");
                                break;
                            case "Attendance Confirmed":
                                edu_comment = string.Concat(edu_comment, "As confirmed, <investigator to add method of confirmation>, [LastName] attended [School] in [EduLocation] from the [EDUFROMDATE] to [EDUTODATE], however, did not earn a degree", "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "[Attendance] - Confirmed");
                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.confirmed_comment.ToString(), "\n\n");
                                break;
                            case "Attendance Unconfirmed":
                                edu_comment = string.Concat(edu_comment, eduattcomment.unconfirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "[Attendance] - Unconfirmed");
                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.unconfirmed_comment.ToString(), "\n\n");
                                break;
                            case "Attendance Result pending":
                                edu_comment = string.Concat(edu_comment, "Efforts to independently verify the same are currently ongoing, the results of which will be provided under separate cover, if and when received.  However, the subject self-reports having attended a [Degree] from [School] in [EduLocation] from [EduStartDate] to [EduEndDate], however, did not earn a degree.", "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "[Attendance] - Results Pending".ToUpper());
                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.unconfirmed_comment.ToString(), " ATTRESPENDING", "\n\n");
                                break;
                        }
                    }
                    else
                    {
                        switch (uSPEIndividual.educationModels[i].Edu_Confirmed.ToString())
                        {
                            case "Yes":
                                edu_comment = string.Concat(edu_comment, "As confirmed, <investigator to add method of confirmation>, [LastName] received a [Degree] from the [School] in [EduLocation] on [GradDate], having attended the same from [EduStartDate] to [EduEndDate].", "\n\n");                                
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Confirmed", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.confirmed_comment.ToString(), "\n\n");
                                break;
                            case "No":
                                edu_comment = string.Concat(edu_comment, comment1.unconfirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Unconfirmed", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumcomment1.unconfirmed_comment.ToString(), "\n\n");
                                break;
                            case "Result Pending":
                                edu_comment = string.Concat(edu_comment, "Efforts to independently confirm the subject’s tenure at the same are currently ongoing, the results of which will be provided under separate cover, if and when received.  However, the subject self-reports having received a [Degree] degree in <major> from [School] in [EduLocation] on [GradDate], having reportedly attended the same from [EduStartDate] to [EduEndDate].", "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "Results Pending", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.confirmed_comment.ToString(), " APPENDRESPENDING", "", "\n\n");
                                break;
                            case "Attendance Confirmed":
                                edu_comment = string.Concat(edu_comment, "As confirmed, <investigator to add method of confirmation>, [LastName] attended [School] in [EduLocation] from the [EDUFROMDATE] to [EDUTODATE], however, did not earn a degree", "\n\n");
                                //edu_comment = string.Concat(edu_comment, eduattcomment.confirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "[Attendance] - Confirmed", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.confirmed_comment.ToString(), "\n\n");
                                break;
                            case "Attendance Unconfirmed":
                                edu_comment = string.Concat(edu_comment, eduattcomment.unconfirmed_comment.ToString(), "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "[Attendance] - Unconfirmed", "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumattcomment.unconfirmed_comment.ToString(), "\n\n");
                                break;
                            case "Attendance Result pending":
                                edu_comment = string.Concat(edu_comment, "Efforts to independently verify the same are currently ongoing, the results of which will be provided under separate cover, if and when received.  However, the subject self-reports having attended a [Degree] from [School] in [EduLocation] from [EduStartDate] to [EduEndDate], however, did not earn a degree.", "\n\n");
                                edu_header = string.Concat(edu_header, uSPEIndividual.educationModels[i].Edu_Degree.ToString(), " - ", "[Attendance] - Results Pending".ToUpper(), "\n\n");
                                edu_summcomment = string.Concat(edu_summcomment, edusumrescomment.unconfirmed_comment.ToString(), " ATTRESPENDING", "\n\n");
                                break;
                        }
                    }
                    try
                    {
                        if (uSPEIndividual.educationModels[i].Edu_Major.Equals("")) { }
                        else
                        {
                            edu_comment = edu_comment.Replace("[Degree]", string.Concat("[Degree] degree in ", uSPEIndividual.educationModels[i].Edu_Major));
                        }
                    }
                    catch { }
                    try
                    {
                        edu_comment = edu_comment.Replace("[Degree]", uSPEIndividual.educationModels[i].Edu_Degree.ToString());
                    }
                    catch
                    {
                        edu_comment = edu_comment.Replace("[Degree]", "<not provided>");
                    }
                    try
                    {
                        edu_comment = edu_comment.Replace("[School]", uSPEIndividual.educationModels[i].Edu_School.ToString());
                    }
                    catch
                    {
                        edu_comment = edu_comment.Replace("[School]", "<not provided>");
                    }
                    try
                    {
                        edu_comment = edu_comment.Replace("[EduLocation]", uSPEIndividual.educationModels[i].Edu_Location.ToString());
                    }
                    catch {
                        edu_comment = edu_comment.Replace("[EduLocation]", "<not provided>");
                    }
                    edu_comment = edu_comment.Replace("[GradDate]", edugraddate);
                    edu_comment = edu_comment.Replace("[EduStartDate]", edustartdate);
                    edu_comment = edu_comment.Replace("[EduEndDate]", eduenddate);
                    try
                    {
                        edu_summcomment = edu_summcomment.Replace("[Degree]", uSPEIndividual.educationModels[i].Edu_Degree.ToString());
                    }
                    catch
                    {
                        edu_summcomment = edu_summcomment.Replace("[Degree]", "<not provided>");
                    }
                    try
                    {
                        edu_summcomment = edu_summcomment.Replace("[School]", uSPEIndividual.educationModels[i].Edu_School.ToString());
                    }
                    catch
                    {
                        edu_summcomment = edu_summcomment.Replace("[School]", "<not provided>");
                    }
                    try
                    {
                        edu_summcomment = edu_summcomment.Replace("[EduLocation]", uSPEIndividual.educationModels[i].Edu_Location.ToString());
                    }
                    catch
                    {
                        edu_summcomment = edu_summcomment.Replace("[EduLocation]", "<not provided>");
                    }
                    edu_summcomment = edu_summcomment.Replace("[GradDate]", edugradyr);
                    edu_summcomment = edu_summcomment.Replace("[EDUFROMDATE]", edustartyr);
                    edu_summcomment = edu_summcomment.Replace("[EDUTODATE]", eduendyr);
                }
                else
                {
                    edu_comment = "";
                    edu_summcomment = comment1.other_comment.ToString();
                    edu_summcomment = edu_summcomment.Replace("[Last Name]", uSPEIndividual.diligenceInputModel.LastName.ToString());
                    edu_header = "N/A";
                }
            }
            edu_comment = edu_comment.Replace("[LastName]", uSPEIndividual.diligenceInputModel.LastName.ToString());
            doc.Replace("SUMMDEGDESCRIPTION", edu_summcomment, true, true);
            doc.Replace("DEGREEDESCRIPTION", edu_comment, true, true);            
            doc.Replace("DEGREEHEADER", edu_header, true, true);
            doc.SaveToFile(savePath);
            if (uSPEIndividual.summarymodel.Summarytable.ToString().Equals("Yes"))
            {
                Table table = doc.Sections[1].Tables[0] as Table;
                TableCell cell1 = table.Rows[6].Cells[2];
                TableCell cell2 = table.Rows[6].Cells[1];
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
            }

            TextSelection[] textedurespending = doc.FindAllString("Efforts to independently confirm the subject’s tenure at the same are currently ongoing, the results of which will be provided under separate cover, if and when received.", false, true);
            if (textedurespending != null)
            {
                foreach (TextSelection seletion in textedurespending)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                }
            }
            TextSelection[] textAttedurespending = doc.FindAllString("Efforts to independently verify the same are currently ongoing, the results of which will be provided under separate cover, if and when received.", false, true);
            if (textAttedurespending != null)
            {
                foreach (TextSelection seletion in textAttedurespending)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                }
            }          
            doc.SaveToFile(savePath);
            //PL License details
            string pl_comment = "";
            string plstartdate = "";
            string plenddate = "";
            for (int i = 0; i < uSPEIndividual.pllicenseModels.Count; i++)
            {
                plstartdate = "<not provided>";
                plenddate = "<not provided>";
                if (uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                {
                    if (!uSPEIndividual.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && !uSPEIndividual.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                    {
                        plstartdate = uSPEIndividual.pllicenseModels[i].PL_StartDateMonth + " " + uSPEIndividual.pllicenseModels[i].PL_StartDateDay + ", " + uSPEIndividual.pllicenseModels[i].PL_StartDateYear;
                        //pllicenseModel.PL_Start_Date = plstartdate;

                    }
                    if (uSPEIndividual.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && uSPEIndividual.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && uSPEIndividual.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                    {
                        plstartdate = "<not provided>";
                    }
                    if (uSPEIndividual.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && !uSPEIndividual.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                    {
                        plstartdate = uSPEIndividual.pllicenseModels[i].PL_StartDateMonth + " " + uSPEIndividual.pllicenseModels[i].PL_StartDateYear;
                        //pllicenseModel.PL_Start_Date = plstartdate;
                    }
                    if (uSPEIndividual.pllicenseModels[i].PL_StartDateDay.ToString().Equals("Day") && uSPEIndividual.pllicenseModels[i].PL_StartDateMonth.ToString().Equals("0") && !uSPEIndividual.pllicenseModels[i].PL_StartDateYear.ToString().Equals("Year"))
                    {
                        plstartdate = uSPEIndividual.pllicenseModels[i].PL_StartDateYear;
                        //pllicenseModel.PL_Start_Date = plstartdate;
                    }

                    if (!uSPEIndividual.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && !uSPEIndividual.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                    {
                        plenddate = uSPEIndividual.pllicenseModels[i].PL_EndDateMonth + " " + uSPEIndividual.pllicenseModels[i].PL_EndDateDay + ", " + uSPEIndividual.pllicenseModels[i].PL_EndDateYear;
                        //pllicenseModel.PL_End_Date = plenddate;

                    }
                    if (uSPEIndividual.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && uSPEIndividual.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && uSPEIndividual.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                    {
                        plenddate = "<not provided>";
                    }
                    if (uSPEIndividual.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && !uSPEIndividual.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                    {
                        plenddate = uSPEIndividual.pllicenseModels[i].PL_EndDateMonth + " " + uSPEIndividual.pllicenseModels[i].PL_EndDateYear;
                        //pllicenseModel.PL_End_Date = plenddate;
                    }
                    if (uSPEIndividual.pllicenseModels[i].PL_EndDateDay.ToString().Equals("Day") && uSPEIndividual.pllicenseModels[i].PL_EndDateMonth.ToString().Equals("0") && !uSPEIndividual.pllicenseModels[i].PL_EndDateYear.ToString().Equals("Year"))
                    {
                        plenddate = uSPEIndividual.pllicenseModels[i].PL_EndDateYear;
                        //pllicenseModel.PL_End_Date = plenddate;
                    }

                    CommentModel comment2 = _context.DbComment
                                .Where(u => u.Comment_type == "PL1")
                                .FirstOrDefault();
                    string strplorgfont = string.Concat(uSPEIndividual.pllicenseModels[i].PL_License_Type.ToString(), " CHANGEFONTHEADER");
                    if (i == 0) { pl_comment = "\n"; }
                    if (uSPEIndividual.pllicenseModels[i].PL_Confirmed.Equals("Yes"))
                    {
                        if (uSPEIndividual.pllicenseModels.Count - 1 == i)
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
                        if (uSPEIndividual.pllicenseModels[i].PL_Confirmed.Equals("No"))
                        {
                            if (uSPEIndividual.pllicenseModels.Count - 1 == i)
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
                            if (uSPEIndividual.pllicenseModels.Count - 1 == i)
                            {
                                pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.other_comment.ToString(), "APPENDPLTEXT", i.ToString());
                            }
                            else
                            {
                                pl_comment = string.Concat(pl_comment, strplorgfont, "\n\n", comment2.other_comment.ToString(), "APPENDPLTEXT", i.ToString(), "\n\n");
                            }
                        }
                    }
                    pl_comment = pl_comment.Replace("[Last Name]", uSPEIndividual.diligenceInputModel.LastName.ToString());
                    pl_comment = pl_comment.Replace("[PL Organization]", uSPEIndividual.pllicenseModels[i].PL_Organization);
                    pl_comment = pl_comment.Replace("[Professional License Type]", uSPEIndividual.pllicenseModels[i].PL_License_Type.ToString());
                    if (uSPEIndividual.pllicenseModels[i].PL_Number.Equals(""))
                    { pl_comment = pl_comment.Replace(", with a license number [PL Number]", ""); }
                    else
                    {
                        pl_comment = pl_comment.Replace("[PL Number]", uSPEIndividual.pllicenseModels[i].PL_Number.ToString());
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
            if (uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
            {
                try
                {
                    for (int j = 1; j < sectionLimit; j++)
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
            if (uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
            {
                for (int i = 0; i < uSPEIndividual.pllicenseModels.Count; i++)
                {
                    bnres = "";
                    if (uSPEIndividual.pllicenseModels[i].PL_Confirmed.ToString().Equals("Result Pending"))
                    {
                        try
                        {
                            for (int j = 1; j < sectionLimit; j++)
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

                        if (i == uSPEIndividual.pllicenseModels.Count - 1)
                        {
                            bnres = "";
                            try
                            {
                                for (int j = 1; j < sectionLimit; j++)
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
            if (uSPEIndividual.educationModels[0].Edu_History.ToString().Equals("No"))
            {
                if (uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("No"))
                {
                    doc.Replace("EDUCATIONALANDLICENSINGHITS", "[RESULTSPENSECTIONS1]", true, true);
                }
                else
                {
                    doc.Replace("EDUCATIONALANDLICENSINGHITS", "Additionally, efforts included verification of the subject’s licensing credentials, where available. [RESULTSPENSECTIONS1]", true, true);
                }
            }
            else
            {
                if (uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
                {
                    doc.Replace("EDUCATIONALANDLICENSINGHITS", "Additionally, efforts included verification of the subject’s educational and licensing credentials, where available. [RESULTSPENSECTIONS1]", true, true);
                }
                else
                {
                    doc.Replace("EDUCATIONALANDLICENSINGHITS", "Additionally, efforts included verification of the subject’s educational credentials, where available. [RESULTSPENSECTIONS1]", true, true);
                }
            }
            doc.SaveToFile(savePath);
            //Results Pending in Legal Records Section?
            if (uSPEIndividual.otherdetails.HasBankruptcyRecHits_resultpending == true || uSPEIndividual.otherdetails.Has_CriminalRecHit_resultpending == true || uSPEIndividual.otherdetails.has_civil_resultpending == true)
            {
                doc.Replace("[RESULTSPENSECTIONS1]", "\n\n<Search type> searches are currently ongoing in <jurisdiction>, the results of which will be provided under separate cover upon receipt.", true, false);
            }
            else
            {
                doc.Replace("[RESULTSPENSECTIONS1]", "", true, false);
            }
            //ExecutiveSummary
            if (uSPEIndividual.otherdetails.Executivesummary.ToString().Equals("Yes"))
            {
                if (uSPEIndividual.otherdetails.HasBankruptcyRecHits.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Has_CriminalRecHit.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Has_Bureau_PrisonHit.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Has_Sex_Offender_RegHit.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Has_Civil_Records.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Has_US_Tax_Court_Hit.ToString().Equals("Yes"))
                {
                    uSPEIndividual.otherdetails.Has_Legal_Records_Hits = "Yes";
                }
                else
                {
                    uSPEIndividual.otherdetails.Has_Legal_Records_Hits = "No";
                }
                if (uSPEIndividual.otherdetails.Has_Legal_Records_Hits.ToString().Equals("Yes") || uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Yes-With Adverse") || regflag == "Records" || uSPEIndividual.otherdetails.RegulatoryFlag.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Global_Security_Hits.ToString().Equals("Yes") || uSPEIndividual.otherdetails.ICIJ_Hits.ToString().Equals("Yes"))
                {
                    uSPEIndividual.otherdetails.Has_Regulatory_Hits = "Yes";
                }
                else
                {
                    uSPEIndividual.otherdetails.Has_Regulatory_Hits = "No";
                }
                if (uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Has_Reg_US_SEC.ToString().Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_UK_FCA.ToString().Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_NFA.ToString().Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_FINRA.ToString().Equals("Yes-With Adverse") || regflag == "Records" || uSPEIndividual.otherdetails.RegulatoryFlag.ToString().Equals("Yes") || uSPEIndividual.otherdetails.Global_Security_Hits.ToString().Equals("Yes") || uSPEIndividual.otherdetails.ICIJ_Hits.ToString().Equals("Yes"))
                {
                    uSPEIndividual.otherdetails.Has_Hits_Above = "Yes";
                }
                else
                {
                    uSPEIndividual.otherdetails.Has_Hits_Above = "No";
                }
                //Legal_Record_Judgments_Liens_Hits
                string Legrechitcommentmodel = "";
                CommentModel Legrechitcommentmodel1 = _context.DbComment
                                .Where(u => u.Comment_type == "Legal_Record_Judgments_Liens_Hits")
                                .FirstOrDefault();
                if (uSPEIndividual.otherdetails.Has_Legal_Records_Hits.ToString().Equals("Yes"))
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
                if (uSPEIndividual.otherdetails.Has_Regulatory_Hits.ToString().Equals("Yes"))
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
                if (uSPEIndividual.otherdetails.Has_Hits_Above.ToString().Equals("Yes"))
                {
                    hitcompcommentmodel = "In sum, with the exception of the above, no other issues of potential relevance were identified in connection with [FirstName] [MiddleInitial] [LastName] in the United States.";
                }
                else
                {
                    hitcompcommentmodel = "In sum, no other issues of potential relevance were identified in connection with [FirstName] [MiddleInitial] [LastName] in the United States.";
                }
                hitcompcommentmodel = hitcompcommentmodel.Replace("[FirstName]", uSPEIndividual.diligenceInputModel.FirstName.ToString());                
                doc.Replace("HasHitsAboveAndCompanionReport", hitcompcommentmodel, true, true);
                doc.SaveToFile(savePath);
            }
           
            if (current_full_address.ToString().Equals("")) { doc.Replace("[CurrentFullAddress]", "<not provided>", true, true); }
            else
            {
                doc.Replace("[CurrentFullAddress]", current_full_address, true, true);
            }
            doc.SaveToFile(savePath);            
            doc.Replace("[FirstName]", uSPEIndividual.diligenceInputModel.FirstName, true, true);
            try
            {
                if (uSPEIndividual.diligenceInputModel.MiddleName == "") { doc.Replace(" [MiddleName]", "", true, false); }
                else
                {
                    doc.Replace("[MiddleName]", uSPEIndividual.diligenceInputModel.MiddleName, true, true);
                }
            }
            catch {
                doc.Replace(" [MiddleName]", "", true, false);
            }
            doc.Replace("[LastName]", uSPEIndividual.diligenceInputModel.LastName, true, false);
            doc.Replace("[First Name]", uSPEIndividual.diligenceInputModel.FirstName, true, false);
            doc.Replace("[Last Name]", uSPEIndividual.diligenceInputModel.LastName, true, false);
            doc.Replace("[Country]", "the United States", true, false);
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
                                    TextRange _text = footnote1.TextBody.AddParagraph().AppendText("It should be emphasized that updates are made to these lists on a periodic and irregular basis, and for purposes of preparing this report searches were conducted on <Investigator to insert date of research>.");
                                    footnote1.TextBody.FirstParagraph.Format.HorizontalAlignment = HorizontalAlignment.Justify;
                                    _text.CharacterFormat.FontName = "Calibri (Body)";
                                    _text.CharacterFormat.FontSize = 9;
                                    //Append the line
                                    string strglobaltextappended = "";
                                    if (uSPEIndividual.otherdetails.Global_Security_Hits.Equals("Yes"))
                                    {
                                        strglobaltextappended = " and the following information was identified in connection with [LastName]: ";
                                    }
                                    else
                                    {
                                        strglobaltextappended = " and it is noted that [LastName] was not identified on any of these lists.";
                                    }
                                    strglobaltextappended = strglobaltextappended.Replace("[LastName]", uSPEIndividual.diligenceInputModel.LastName);
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
            catch
            {

            }
            doc.SaveToFile(savePath);
            string strPLComment = "";
            if (uSPEIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse") || uSPEIndividual.otherdetails.Registered_with_HKSFC.ToString().Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse") || uSPEIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse") || uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
            {
                if (uSPEIndividual.pllicenseModels[0].General_PL_License.Equals("Yes"))
                {
                    strPLComment = "[LastName] holds a [ProfessionalLicenseType1] license with the [PLOrganization1]";
                }
                if (uSPEIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-Without Adverse"))
                {
                    if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                    strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. SEC");
                }
                if (uSPEIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-Without Adverse"))
                {
                    if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                    strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.K. FCA");
                }
                if (uSPEIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-Without Adverse"))
                {
                    if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                    strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. FINRA");
                }
                if (uSPEIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-Without Adverse"))
                {
                    if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                    strPLComment = string.Concat(strPLComment, "[LastName] has been registered with the U.S. NFA");
                }
                if (uSPEIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-Without Adverse"))
                {
                    if (strPLComment != "") { strPLComment = string.Concat(strPLComment, "\n\n"); }
                    strPLComment = string.Concat(strPLComment, "[LastName] has been registered with HKSFC");
                }

                if (uSPEIndividual.otherdetails.Has_Reg_US_SEC.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_UK_FCA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_FINRA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Has_Reg_US_NFA.Equals("Yes-With Adverse") || uSPEIndividual.otherdetails.Registered_with_HKSFC.Equals("Yes-With Adverse"))
                {
                    strPLComment = string.Concat(strPLComment, "\n\nFurther, <investigator to insert regulatory hits here>");
                }
                strPLComment = strPLComment.Replace("[ProfessionalLicenseType1]", uSPEIndividual.pllicenseModels[0].PL_License_Type);
                strPLComment = strPLComment.Replace("[PLOrganization1]", uSPEIndividual.pllicenseModels[0].PL_Organization);
                doc.Replace("PROFRESULT", "Records", true, true);
            }
            else
            {
                strPLComment = "";
                doc.Replace("PROFRESULT", "Clear", true, true);
            }
            doc.Replace("PROFCOMMENT", strPLComment, true, true);
            doc.SaveToFile(savePath);
            TextSelection[] crimappend = doc.FindAllString("No records were located on the subject in these sources", false, false);
            if (crimappend != null)
            {
                foreach (TextSelection seletion in crimappend)
                {                    
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            doc.SaveToFile(savePath);
            HttpContext.Session.SetString("first_name", uSPEIndividual.diligenceInputModel.FirstName);
            HttpContext.Session.SetString("last_name", uSPEIndividual.diligenceInputModel.LastName);
            HttpContext.Session.SetString("case_number", uSPEIndividual.diligenceInputModel.CaseNumber);
            HttpContext.Session.SetString("middleinitial", uSPEIndividual.diligenceInputModel.MiddleInitial);            
            HttpContext.Session.SetString("regulatoryflag", uSPEIndividual.otherdetails.RegulatoryFlag);
            HttpContext.Session.SetString("employer1location", string.Concat(uSPEIndividual.EmployerModel[0].Emp_City.TrimEnd(),", ",uSPEIndividual.EmployerModel[0].Emp_State.ToString().TrimEnd()));
            //HttpContext.Session.SetString("employer1State", uSPEIndividual.Employer1City);
           
            if (uSPEIndividual.EmployerModel[0].Emp_StartDateYear.ToString().Equals(""))
            {
                strempstartdate = "<not provided>";
            }
            else
            {
                strempstartdate = uSPEIndividual.EmployerModel[0].Emp_StartDateYear.ToString();
            }

            HttpContext.Session.SetString("emp_startdate1", strempstartdate);
            HttpContext.Session.SetString("employer1", uSPEIndividual.EmployerModel[0].Emp_Employer.ToString());         
            HttpContext.Session.SetString("additionalstates", add_states);
            HttpContext.Session.SetString("reg_ussec", uSPEIndividual.otherdetails.Has_Reg_US_SEC);
            HttpContext.Session.SetString("pl_generallicense", uSPEIndividual.pllicenseModels[0].General_PL_License.ToString());
            if (uSPEIndividual.pllicenseModels[0].General_PL_License.ToString().Equals("Yes"))
            {
                HttpContext.Session.SetString("pl_licensetype", uSPEIndividual.pllicenseModels[0].PL_License_Type.ToString());
                HttpContext.Session.SetString("pl_organization", uSPEIndividual.pllicenseModels[0].PL_Organization.ToString());                           
                if (uSPEIndividual.pllicenseModels[0].PL_StartDateYear.ToString().Equals(""))
                {
                    plstartdate = "<not provided>";
                }
                else
                {
                    plstartdate = uSPEIndividual.pllicenseModels[0].PL_StartDateYear.ToString();
                }                                                                     
                if (uSPEIndividual.pllicenseModels[0].PL_EndDateYear.ToString().Equals(""))
                {
                    plenddate = "<not provided>";
                }
                else
                {
                    plenddate = uSPEIndividual.pllicenseModels[0].PL_EndDateYear.ToString();
                }              
                HttpContext.Session.SetString("pl_startdate", plstartdate);
                HttpContext.Session.SetString("pl_enddate", plenddate);
            }
            doc.SaveToFile(savePath);
            TextSelection[] crimrs = doc.FindAllString("A criminal records search is currently pending through the <Investigator to insert source/court/law enforcement agency>, the results of which will be provided under separate cover upon receipt", false, false);
            if (crimrs != null)
            {
                foreach (TextSelection seletion in crimrs)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                }
            }
            TextSelection[] civirs = doc.FindAllString("A civil court records search is currently pending through the <enter source/court> the results of which will be provided under separate cover upon receipt", false, false);
            if (civirs != null)
            {
                foreach (TextSelection seletion in civirs)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;                    
                }
            }
            TextSelection[] bankrs = doc.FindAllString("Bankruptcy court records searches are currently pending in <jurisdiction> the results of which will be provided under separate cover upon receipt", false, false);
            if (bankrs != null)
            {
                foreach (TextSelection seletion in bankrs)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Aqua;
                }
            }
            TextSelection[] hksfctext = doc.FindAllString("Hong Kong Securities and Futures Commission[HKSFCHEADER]", true, true);
            if (hksfctext != null)
            {
                foreach (TextSelection seletion in hksfctext)
                {                    
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                    seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                }
            }

            doc.Replace("[HKSFCHEADER]", "", true, false);
            doc.SaveToFile(savePath);
            TextSelection[] textSelections = doc.FindAllString("identified the subject as a Registrant/Owner in connection with the following trademarks and as an Applicant/Assignee/Inventor in connection with the following patents:", true, true);
            if (textSelections != null)
            {
                foreach (TextSelection seletion in textSelections)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }

            string straction = "";
            string strcontroller = "";
            doc.SaveToFile(savePath);
            if (uSPEIndividual.summarymodel.Summarytable.ToString().Equals("No"))
            {               
                straction = "GenerateFile";
                strcontroller = "Diligence";
                //Regex regex = new Regex(@"\<\w+\b\>");    
                Regex regex = new Regex(@"\<([^>]*)\>");
                TextSelection[] textSelections1 = doc.FindAllPattern(regex);
                if (textSelections1 != null)
                {
                    foreach (TextSelection seletion in textSelections1)
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

                TextSelection[] textexibit = doc.FindAllString("Exhibit A",false,true);
                if (textexibit != null)
                {
                    foreach (TextSelection seletion in textexibit)
                    {
                        seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                        seletion.GetAsOneRange().CharacterFormat.Bold = true;
                    }
                }
                TextSelection[] textSelections2 = doc.FindAllString("Business Affiliations", false, true);
                if(textSelections2!=null)
                {
                    foreach(TextSelection selection in textSelections2)
                    {
                        selection.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                        selection.GetAsOneRange().CharacterFormat.Italic = true;
                    }
                }
                String[] strhighlist = { "<Investigator to insert property record or rental information>", "<investigator to insert source/court/law enforcement agency>", "<jurisdiction>", "<Investigator to insert County Courts searched>", "<Investigator to insert County Courts searched, if applicable>", "<investigator to insert date of research>", "<investigator to remove if inaccurate>", "<not provided>", "<investigator to modify if disciplinary history exists>", "<Investigator to insert company registry detail here>", "<Investigator to insert article summaries here>", "<Investigator to insert results here>", "<investigator to remove any sources that came back with hits above>", "<Investigator to modify for multiple Companion Reports if applicable>", "<date>", "<number>", "<entity>", "<series name>", "<insert controlled functions, entities and effective dates>", "<start date>", "<end date>", "<insert provincial courts>", "<investigator to insert press summaries here>", "<investigator to insert summary here>", "<type of charge>", "<status>", "<party type>", "<investigator to add counties>", "<Investigator to insert discrepancy here>", "<role>", "<capacity>", "<previous role>", "<timeframe>", "<Companion Subject>", "<investigator to insert regulatory hits here>", "As confirmed", "<Investigator to insert summarized findings>", "<Investigator to insert reason for failed verification here>", "<investigator to add relevant county/counties>", "<investigator to add relevant county/counties>", "<investigator to insert County, State>", "<investigator to add all relevant historical and/or possible counties/jurisdictions>", "<State>", "<County>", "<purchase price>", "<seller names>", "<Investigator to insert other mortgage information or UCC details here>", "<market/assessed (choose applicable)>", "<amount>", "<YYYY>", "<address>", "<Investigator County, <State> property records, the subject to insert other mortgage information or UCC details here>", "<add Counties/States>", "<investigator to insert clear record type (i.e. civil judgments or Uniform Commercial Code filings)>", "<investigator to insert details here>", "<States>", "<Investigator to insert summary of trade account information here>", "<sale price>", "<buyer names>", "<record type>", "<expiration date (month and year only)>", "<State driving agency>", "<dates>", "<case type>", "<enter source/court>", "<investigator to add or remove any sources that came back with hits above or are needed for this jurisdiction>", "<Investigator to insert verification information found here>" };
                //foreach (String s in strhighlist)
                //{
                //    TextSelection[] textSelections = doc.FindAllString(s,false,false);
                //    if (textSelections != null)
                //    {
                //        foreach (TextSelection seletion in textSelections)
                //        {
                //            seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                //        }
                //    }
                //}
                doc.SaveToFile(savePath);
                doc.Replace("countries", "counties", true, false);
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
                    if (uSPEIndividual.diligenceInputModel.MiddleInitial == "")
                    {
                        doc.Replace(" [MiddleInitial]", "", true, false);
                        doc.Replace(" Middle_Initial", "", true, false);
                    }
                    else
                    {
                        doc.Replace("[MiddleInitial]", uSPEIndividual.diligenceInputModel.MiddleInitial, true, true);
                        doc.Replace("Middle_Initial", uSPEIndividual.diligenceInputModel.MiddleInitial.TrimEnd().ToUpper().ToString(), true, true);
                    }
                }
                catch
                {
                    doc.Replace(" [MiddleInitial]", "", true, false);
                    doc.Replace(" Middle_Initial", "", true, false);
                }
                doc.Replace("[ADDOOFOOTNOTE]", ".".TrimEnd(), true, false);
                doc.SaveToFile(savePath);
                doc.Replace("the District of Columbia", "District of Columbia", false, false);
                doc.SaveToFile(savePath);
                doc.Replace("District of Columbia", "the District of Columbia", false, false);
                doc.Replace("U.S.  ", "U.S. ", true, false);
                doc.Replace("U.K.  ", "U.K. ", true, false);
                doc.SaveToFile(savePath);
                doc.Replace("This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>.", "  This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>.", true, false);
                doc.SaveToFile(savePath);
                return RedirectToAction(straction, strcontroller);
            }
            else
            {
                SummaryResulttableModel summary = new SummaryResulttableModel();
                summary = uSPEIndividual.summarymodel;
                return Save_Page2(summary);
            }           
        }

        public IActionResult US_PE_Individual_page2()
        {
            return View();
        }
        public IActionResult Save_Page2(SummaryResulttableModel ST)
        {
            string middleinitial= HttpContext.Session.GetString("middleinitial");
            string record_Id = HttpContext.Session.GetString("record_Id");
            string first_name = HttpContext.Session.GetString("first_name");
            string last_name = HttpContext.Session.GetString("last_name");
            string country = "the United States";
            string add_states = HttpContext.Session.GetString("additionalstates");
            string strussec = HttpContext.Session.GetString("reg_ussec");
            //string country = HttpContext.Session.GetString("country");
            string case_number = HttpContext.Session.GetString("case_number");            
            string regulatoryflag = HttpContext.Session.GetString("regulatoryflag");
            string employer1location = HttpContext.Session.GetString("employer1location");            
            string employer1 = HttpContext.Session.GetString("employer1");
            string emp_position1 = HttpContext.Session.GetString("emp_position1");
            string emp_startdate1 = HttpContext.Session.GetString("emp_startdate1");         
            string pl_generallicense = HttpContext.Session.GetString("pl_generallicense");
            string pl_licensetype = "";
            string pl_organization = "";
            //string pl_startdate = "";
            //string pl_enddate = "";
            if (pl_generallicense.Equals("Yes"))
            {
                pl_licensetype = HttpContext.Session.GetString("pl_licensetype");
                pl_organization = HttpContext.Session.GetString("pl_organization");              
            }
            string savePath = _config.GetValue<string>("ReportPath");
            savePath = string.Concat(savePath, last_name.ToString(), "_SterlingDiligenceReport(", case_number.ToString(), ")_DRAFT.docx");
            Document doc = new Document(savePath);
            string strcomment;            
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
            ST.name_add_Ver_History = "Clear";
            switch (ST.name_add_Ver_History.ToString())
            {
                case "Clear":
                    strcomment = "LastName has jurisdictional ties to [ADDITIONALSTATES]";
                    strcomment = strcomment.Replace("LastName", last_name);
                    strcomment = strcomment.Replace("[ADDITIONALSTATES]", add_states);
                    doc.Replace("NAMECOMMENT", strcomment, true, true);
                    doc.Replace("NAMERESULT", "Clear", true, true);
                    break;
                case "Records":
                    strcomment = "LastName has jurisdictional ties to [ADDITIONALSTATES] \n\n Further, the subject was identified as an owner of at least <number> of properties in Country";
                    strcomment = strcomment.Replace("LastName", last_name);
                    strcomment = strcomment.Replace("[ADDITIONALSTATES]", add_states);
                    strcomment = strcomment.Replace("Country", country);
                    doc.Replace("NAMECOMMENT", strcomment, true, true);
                    doc.Replace("NAMERESULT", "Records", true, true);
                    break;
                case "Discrepancy Identified":
                    strcomment = "LastName has jurisdictional ties to [ADDITIONALSTATES] \n\n However, while not reported by the subject, LastName was also identified as having current ties to < Investigator to insert additional jurisdictions here>";
                    strcomment = strcomment.Replace("LastName", last_name);
                    strcomment = strcomment.Replace("[ADDITIONALSTATES]", add_states);
                    strcomment = strcomment.Replace("Country", country);
                    doc.Replace("NAMECOMMENT", strcomment, true, true);
                    doc.Replace("NAMERESULT", "Discrepancy Identified", true, true);
                    break;
            }
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
            Regex regex = new Regex(@"\<([^>]*)\>");
            TextSelection[] textSelections1 = doc.FindAllPattern(regex);
            if (textSelections1 != null)
            {
                foreach (TextSelection seletion in textSelections1)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] texsec3 = doc.FindAllString("	•", false, true);
            if (texsec3 != null)
            {
                foreach (TextSelection seletion in texsec3)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.White;
                }
            }
            TextSelection[] texsec4 = doc.FindAllString("•	<Investigator to insert bulleted list of results here>", false, false);
            if (texsec4 != null)
            {
                foreach (TextSelection seletion in texsec4)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
                }
            }
            TextSelection[] textexibit = doc.FindAllString("Exhibit A", false, true);
            if (textexibit != null)
            {
                foreach (TextSelection seletion in textexibit)
                {
                    seletion.GetAsOneRange().CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
                    seletion.GetAsOneRange().CharacterFormat.Bold = true;
                }
            }
            doc.Replace("< investigator to insert summary here >", "<investigator to insert summary here>", true, false);
            doc.Replace("the United Kingdom’s Financial Conduct Authority", "United Kingdom’s Financial Conduct Authority", true, true);
            doc.Replace("< investigator to remove if inaccurate >", "<investigator to remove if inaccurate>", false, true);
            doc.Replace("<investigator to insert regulatory hits here >", "<investigator to insert regulatory hits here>", true, true);
            doc.Replace("< ", "<", true, false);
            doc.Replace(" >", ">", true, false);
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
            TextSelection[] text31 = doc.FindAllString(" 	•", false, true);
            if (text31 != null)
            {
                foreach (TextSelection seletion in text31)
                {
                    seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.White;
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
            
         
            String[] strlist = { "<investigator to insert date of research>", "<investigator to remove if inaccurate>", "<not provided>", "<investigator to modify if disciplinary history exists>", "<Investigator to insert company registry detail here>", "<Investigator to insert article summaries here>", "<Investigator to insert results here>", "<investigator to remove any sources that came back with hits above>", "<Investigator to modify for multiple Companion Reports if applicable>", "<date>", "<number>", "<entity>", "<series name>", "<insert controlled functions, entities and effective dates>", "<start date>", "<end date>", "<insert provincial courts>", "<investigator to insert press summaries here>", "<investigator to insert summary here>", "<type of charge>", "<status>", "<party type>", "<investigator to add counties>", "<Investigator to insert discrepancy here>", "<role>", "<capacity>", "<previous role>", "<timeframe>", "<Companion Subject>", "<investigator to insert regulatory hits here>", "As confirmed", "<Investigator to insert summarized findings>", "<Investigator to insert reason for failed verification here>", "<investigator to add relevant county/counties>", "<investigator to add relevant county/counties>", "<investigator to insert County, State>", "<investigator to add all relevant historical and/or possible counties/jurisdictions>", "<State>", "<County>", "<purchase price>", "<seller names>", "<Investigator to insert other mortgage information or UCC details here>", "<market/assessed (choose applicable)>", "<amount>", "<YYYY>", "<address>", "<Investigator County, <State> property records, the subject to insert other mortgage information or UCC details here>", "<add Counties/States>", "<investigator to insert clear record type (i.e. civil judgments or Uniform Commercial Code filings)>", "<investigator to insert details here>", "<States>", "<Investigator to insert summary of trade account information here>", "<sale price>", "<buyer names>", "<record type>", "<expiration date (month and year only)>", "<State driving agency>", "<dates>", "<case type>", "<enter source/court>","<investigator to add or remove any sources that came back with hits above or are needed for this jurisdiction>" };
            //foreach (String s in strlist)
            //{
          
        
        doc.Replace("[Last Name]", last_name.ToString(), true, true);
            //doc.Replace("The United Kingdom", "United Kingdom", true, true);            
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
                if (middleinitial == "")
                {
                    doc.Replace(" [MiddleInitial]", "", false, false);
                    doc.Replace(" Middle_Initial", "", false, false);
                }
                else
                {
                    doc.Replace("[MiddleInitial]", middleinitial, true, true);
                    doc.Replace("Middle_Initial", middleinitial.TrimEnd().ToUpper().ToString(), true, true);
                }
            }
            catch
            {
                doc.Replace(" [MiddleInitial]", "", false, false);
                doc.Replace(" Middle_Initial", "", false, false);
            }
            doc.Replace("[ADDOOFOOTNOTE]", ".".TrimEnd(), true, false);
            doc.SaveToFile(savePath);
            doc.Replace("countries", "counties", true, false);
            doc.Replace("the District of Columbia", "District of Columbia", false, false);
            doc.SaveToFile(savePath);
            doc.Replace("District of Columbia", "the District of Columbia", false, false);
            doc.Replace("U.S.  ", "U.S. ", true, false);
            doc.Replace("U.K.  ", "U.K. ", true, false);            
            doc.SaveToFile(savePath);
            return RedirectToAction("GenerateFile", "Diligence");
        }
    }
}