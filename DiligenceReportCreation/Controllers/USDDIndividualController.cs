﻿using System;
            List<int> Years = new List<int>();
            DateTime startYear = DateTime.Now.AddYears(-80);
            while (startYear.Year <= DateTime.Now.AddYears(11).Year)
            {
                Years.Add(startYear.Year);
                startYear = startYear.AddYears(1);
            }
            ViewBag.Years = Years;
            string recordid = HttpContext.Session.GetString("recordid");
            DiligenceInputModel diligence = new DiligenceInputModel();
            try
            {
                diligence = _context.DbPersonalInfo
                    .Where(a => a.record_Id == recordid)
                    .FirstOrDefault();
            }
            catch { }
            Otherdatails otherdatails = new Otherdatails();
            try
            {
                otherdatails = _context.othersModel
                                      .Where(u => u.record_Id == recordid)
                                      .FirstOrDefault();
            }
            catch { }
            try
            {
                main.additional_States = _context.Dbadditionalstates
                            .Where(u => u.record_Id == recordid)
                            .ToList();
            }
            catch { }
            try
            {
                main.EmployerModel = _context.DbEmployer
                            .Where(u => u.record_Id == recordid)
                            .ToList();
            }
            catch { }
            try { 
            }
            catch { }
            try
            {
                main.pllicenseModels = _context.DbPLLicense
                           .Where(u => u.record_Id == recordid)
                            .ToList();
            }
            catch { }
            SummaryResulttableModel summary = new SummaryResulttableModel();
            try
            {
                summary = _context.summaryResulttableModels
                     .Where(u => u.record_Id == recordid)
                                       .FirstOrDefault();
            }
            catch { }
            HttpContext.Session.SetString("recordid", recordid);
            ReportModel report = _context.reportModel
                .Where(u => u.record_Id == recordid)
                                  .FirstOrDefault();
            ViewBag.CaseNumber = report.casenumber;
            ViewBag.LastName = report.lastname;
                .Where(u => u.record_Id == recordid)
                                  .FirstOrDefault();
                                   DI1 = new DiligenceInputModel();
                    //strupdate.OtherPropertyOwnershipinfo = mainModel.otherdetails.OtherPropertyOwnershipinfo;
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
                                                //if (mainModel.diligence.employerList[i].Emp_Status.ToString().Equals("Previous")) { }
                                                //else
                                                //{
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
                                                //}
                                            }
                                        }
                                    }
                                    catch { }
                                }
                                    _context.SaveChanges();
                                }
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
            ViewBag.CaseNumber = report.casenumber;
            ViewBag.LastName = report.lastname;
            ViewBag.Employer1State = mainModel.diligenceInputModel.Employer1State;
            return View(mainModel);
            ReportModel report = _context.reportModel
               .Where(u => u.record_Id == recordid)
                                 .FirstOrDefault();
            if (Submit == "SubmitData")
                    catch(Exception e)
                case "SaveOD":
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
                            strupdate.CurrentResidentialProperty = mainModel.otherdetails.CurrentResidentialProperty;
                            strupdate.OtherCurrentResidentialProperty = mainModel.otherdetails.OtherCurrentResidentialProperty;
                            //strupdate.OtherPropertyOwnershipinfo = mainModel.otherdetails.OtherPropertyOwnershipinfo;
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
                case "SaveEMPData":
                        {
                            try
                            {

                                List<EmployerModel> employerModel1 = _context.DbEmployer
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
                                                if (educationModel1[0].Edu_History == "No" ) {
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
                case "SaveCSData":
                    try
                                List<Additional_statesModel> additionaL = _context.Dbadditionalstates
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
                    break;
                case "SaveSummaryData":
                    try
                        {
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
            }
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
                        doc.Replace(" (also known as [FullAliasName])", "[ADDCommannamecheckboxFN]", true, false);
                        doc.SaveToFile(savePath);
//                        strfn = "Due to the commonality of [LastName]'s name, research efforts were focused to the subject's known jurisdictions and timeframe of [LastName]'s affiliation with the same.";
                        strfn = "Due to the commonality of [LastName]’s name, research efforts were strictly focused to those jurisdictions to which the subject has ties to and the timeframes of the subject's affiliation with the same.  Additionally, where possible, certain other identifying criteria, such as addresses, relatives’ names, middle initials, etc., were utilized to attempt to identify public records relating specifically to the subject of interest.";
                    }
                    else
                    {
                        doc.Replace("(also known as [FullAliasName])", "", true, false);
                    }
                        strfn = "It is noted that the subject was identified in connection with the alternate name variations of “[FullAliasName]” and investigative efforts were undertaken in connection with the same, as appropriate. Additionally, due to the commonality of [LastName]'s name, research efforts were focused to the subject's known jurisdictions and timeframe of [LastName]'s affiliation with the same.";
                try
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
                                    if (abc.ToString().Equals("[FullAliasName]") || abc.ToString().EndsWith("[FullAliasName])") || abc.ToString().Contains("FullAliasName") || abc.ToString().EndsWith("[ADDCommannamecheckboxFN]") || abc.ToString().Contains("[ADDCommannamecheckboxFN]"))
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
                doc.Replace("[FullAliasName]", uSDDIndividual.diligenceInputModel.FullaliasName.TrimEnd(), true, false);
                {
                    doc.Replace("ClientName", uSDDIndividual.diligenceInputModel.ClientName.TrimEnd(), true, true);
                }
                catch
                {
                    doc.Replace("ClientName", "<not provided>", true, false);
                }
                {
                    doc.Replace("ClientName", uSDDIndividual.diligenceInputModel.ClientName.TrimEnd(), true, true);
                }
                catch
                {
                    doc.Replace("ClientName","<not provided>", true, true);
                }
                {
                    doc.Replace("[DateofBirth]", CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Month) + " " + Convert.ToDateTime(uSDDIndividual.diligenceInputModel.Dob.ToString()).Year, true, true);
                }
                catch
                {
                    doc.Replace("[DateofBirth]", "<not provided>", true, true);
                }
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
                {
                    if (uSDDIndividual.EmployerModel[i].Emp_Status == "Concurrent") {
                        strcommentconcurrent = string.Concat(strcommentconcurrent, "\n\n[LastName] is a [Position1] of [Employer1] in [Employer1Location], since [EmpStartDate1]");
                        strcommentconcurrent = strcommentconcurrent.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName);
                        try
                        {
                            if (uSDDIndividual.EmployerModel[i].Emp_Position.Equals("")) { strcommentconcurrent = strcommentconcurrent.Replace("[Position1]", "<not provided>");}
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
                            if (uSDDIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals("Year") || uSDDIndividual.EmployerModel[i].Emp_StartDateYear.ToString().Equals(""))
                            {
                                strcommentconcurrent = strcommentconcurrent.Replace("[EmpStartDate1]", "<not provided>" );
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
                    //globalhit_comment = globalhit_comment2.confirmed_comment.ToString();
                    globalhit_comment = "Research was undertaken in connection with various governmental, quasi-governmental and private sector list repositories maintained for purposes of international security, fraud prevention, anti-terrorism and anti-money laundering, including the List of Specially Designated Nationals and Blocked Persons of the United States Office of Foreign Assets Control (“OFAC”), the United States Bureau of Industry and Security, and other agencies,";
                    //globalhit_comment = globalhit_comment2.unconfirmed_comment.ToString();
                    globalhit_comment = "Research was undertaken in connection with various governmental, quasi-governmental and private sector list repositories maintained for purposes of international security, fraud prevention, anti-terrorism and anti-money laundering, including the List of Specially Designated Nationals and Blocked Persons of the United States Office of Foreign Assets Control (“OFAC”), the United States Bureau of Industry and Security, and other agencies,";
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
                //CURRENTRESIDENTIALPROPERTYDESC
                if (uSDDIndividual.otherdetails.CurrentResidentialProperty.ToString().Equals("Yes"))
                    doc.Replace("[CurrentState1]", "<State>", true, true);
                }
                else
                {
                    doc.Replace("[CurrentState1]", uSDDIndividual.diligenceInputModel.CurrentState.ToString(),true,true);
                }
                //OtherCurrentResidentialProperty
                try
                {
                    if (uSDDIndividual.otherdetails.OtherCurrentResidentialProperty.ToString().Equals("Yes, only one"))
                    {
                        doc.Replace("CURRENTRESIDENTIALPROPERTYDESC", "\n\nCurrent ownership records are also identified in connection with a certain property located at <Investigator to add other current property address>. According to <County> County, <State> property records, the subject purchased this property for <purchase price> from <seller names> on <date>. <Investigator to insert other mortgage information or UCC details here>. This property had a total <market/assessed (choose applicable)> value of <amount> as of <YYYY>.[ADDOOFOOT1NOTE]\n", true, false);
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
                //if (uSDDIndividual.otherdetails.OtherPropertyOwnershipinfo.ToString().Equals("Yes"))
                
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
                        {
                            strcrimsumcom = "<investigator to specify which states came back clear>";
                        }
                        else { strcrimsumcom = ""; }
                    strempenddate = "<not provided>";
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
                                        // Find the word "Spire.Doc" in paragraph1
                                        if (abc.ToString().Contains("[Employer1Footnote]") || abc.ToString().EndsWith("[Employer1Footnote]"))
                                            //Insert footnote1 after the word "Spire.Doc"
                                            para.ChildObjects.Insert(k + 1, footnote1);
                                            //Append the line
                                            string stremp1textappended = "";
                                            if (uSDDIndividual.EmployerModel[0].Emp_Status.ToString().Equals("Current"))
                                            {
                                                doc.Replace("EMPLOYEEDESCRIPTION", "<investigator to insert brief summary of company here> EMPLOYEEDESCRIPTION", true, true);
                                            }
                                            doc.SaveToFile(savePath);
                        if (uSDDIndividual.EmployerModel[i].Emp_Status.ToString().Equals("Concurrent"))
                        {
                            if (uSDDIndividual.EmployerModel[i].Emp_Confirmed.Equals("Yes"))
                            {
                                emp_Comment = string.Concat(emp_Comment, "\n", "In addition to the above, [LastName] is also a [Position1] of [Employer1Footnote]");
                                emp_Comment = emp_Comment.Replace("[LastName]", uSDDIndividual.diligenceInputModel.LastName.ToString());
                                emp_Comment = emp_Comment.Replace("[Position1]", uSDDIndividual.EmployerModel[i].Emp_Position.ToString());
                                emp_Comment = emp_Comment.Replace("[Employer1]", uSDDIndividual.EmployerModel[i].Emp_Employer.ToString());
                                emp_Comment = emp_Comment.Replace("[EmpLocation1]", uSDDIndividual.EmployerModel[i].Emp_Location.ToString());
                                emp_Comment = emp_Comment.Replace("[EmpStartDate1]", strempstartdate);
                                emp_Comment = string.Concat("\n\n", emp_Comment);
                            }
                            else
                            {
                                emp_Comment = string.Concat(emp_Comment, "\n", "In addition to the above, [LastName] is also a [Position1] of [Employer1Footnote]");
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
                    }
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
                                            edu_header = string.Concat(edu_header,"Confirmed");
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
                            catch {
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
                        //else
                        //{
                        //    pl_comment = "";
                        //}
                    catch
                    {
                        //pl_comment = "";
                    }
                //EDUCATIONALANDLICENSINGHITS
                string eduhistory = "";string plhistory = "";                
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
                if(uSDDIndividual.otherdetails.Registered_with_HKSFC.ToString().StartsWith("Yes") || uSDDIndividual.otherdetails.Has_Reg_US_SEC.ToString().StartsWith("Yes")|| uSDDIndividual.otherdetails.Has_Reg_US_NFA.ToString().StartsWith("Yes")|| uSDDIndividual.otherdetails.Has_Reg_UK_FCA.ToString().StartsWith("Yes")|| uSDDIndividual.otherdetails.Has_Reg_FINRA.ToString().StartsWith("Yes"))
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
                        drivhistory = string.Concat(strEDULICENcomment,", as well as a review of [LastName]’s driving history.");                        
                    }
                    
                }
                if (uSDDIndividual.otherdetails.Was_credited_obtained.ToUpper().ToString() != "No".ToUpper())
                {
                    if (strEDULICENcomment=="")
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
                        if (eduhistory != "" || plhistory != "" )
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
                if(strEDULICENcomment.EndsWith("where available")) { strEDULICENcomment = string.Concat(strEDULICENcomment, "."); }
                doc.Replace("EDUCATIONALANDLICENSINGHITS", strEDULICENcomment, true, false);

                                    // Find the word "Civil" in paragraph1
                                    if (searchtext.ToString().Equals("")) { }
                                    else
                                    {
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
                                    else
                                    {
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
                catch { }
                {
                    if (uSDDIndividual.diligenceInputModel.MiddleName == "") { doc.Replace(" [MiddleName]", "", true, false); }
                    else
                    {
                        doc.Replace("[MiddleName]", uSDDIndividual.diligenceInputModel.MiddleName, true, true);
                    }
                catch
                {
                    doc.Replace(" [MiddleName]", "", true, false);
                }
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
            catch {
                doc.Replace("USOFFCOMMENT", "", true, true);
                doc.Replace("USOFFRESULT", "Clear", true, true);
            
            //Regex regex1 = new Regex(@"/\p{No}/ug");
            //TextSelection[] footnotesupSelections1 = doc.FindAllPattern(regex1);
            if (text03 != null)
            //TextSelection[] text21 = doc.FindAllString("<Investigator to insert bulleted list of results here>", false, true);
            //if (text21 != null)
            //{
            //    foreach (TextSelection seletion in text21)
            //    {
            //        seletion.GetAsOneRange().CharacterFormat.HighlightColor = Color.Yellow;
            //    }
            //}           
            doc.Replace("[Last Name]", last_name.ToString(), true, true);
            if (textresult1 != null)
            {
                foreach (TextSelection seletion in textresult1)
                {
                    seletion.GetAsOneRange().CharacterFormat.Italic = true;
                }
            }
            try
            {
                if (MiddleInitial=="") 
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
                doc.Replace(" [MiddleInitial]", "", true, false);
                doc.Replace(" Middle_Initial", "", true,false);
            doc.Replace("U.K.  ", "U.K. ", true, false);
            doc.SaveToFile(savePath);