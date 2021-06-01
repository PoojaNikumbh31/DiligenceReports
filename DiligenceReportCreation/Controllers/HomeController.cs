using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using DiligenceReportCreation.Models;
using DiligenceReportCreation.Data;
using Microsoft.AspNetCore.Http;
using System.DirectoryServices;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using System.Text;

namespace DiligenceReportCreation.Controllers
{
    public class HomeController : Controller
    {
        private readonly DataBaseContext _context;
        public HomeController(DataBaseContext context)
        {
            _context = context;
        }

        

        public IActionResult LoginFormView()
        {
            return View();
        }

        public IActionResult LoginAuth(LoginViewModel loginModel)
        {
            //DirectoryEntry directoryEntry;
            //DirectorySearcher directorySearcher;
            //SearchResult searchResult;
            UserModel user;


            try
            {
                //directoryEntry = new DirectoryEntry("LDAP://ST", loginModel.Username, loginModel.Password);
                //directorySearcher = new DirectorySearcher(directoryEntry);
                //directorySearcher.Filter = string.Format("(&(SAMAccountName={0}))", loginModel.Username);
                //searchResult = directorySearcher.FindOne();
               
                user = _context.DbUser.Where(x => x.Username == loginModel.Username).FirstOrDefault();

            }
            catch
            {
                
                user = null;
            }
            if (null == user)
            {                
                // _logger.LogInformation("Login failed");
                ModelState.AddModelError("LogOnError", "The username provided by you is incorrect");
                TempData["message"] = "The username provided by you is incorrect";
                return RedirectToAction("LoginFormView", "Home");
            }
            else
            {
                //string passwordFromDb = Convert.ToBase64String(Encoding.Unicode.GetBytes(loginModel.Password));
                string passwordFromDb = Encoding.Unicode.GetString(Convert.FromBase64String(user.Password));
               
               if (loginModel.Password == passwordFromDb)
                {
                    HttpContext.Session.SetString("UserName", loginModel.Username);
                    HttpContext.Session.SetString("Permissions", user.Permissions);
                    HttpContext.Session.SetString("Role", user.Role);
                    //user.Password = Convert.ToBase64String(Encoding.Unicode.GetBytes(loginModel.Password));
                    user.LastLogin = DateTime.UtcNow;
                    _context.DbUser.Update(user);
                    _context.SaveChanges();                 
                }
                else
                {
                    ModelState.AddModelError("LogOnError", "The password provided by you is incorrect");
                    TempData["message"] = "The password provided by you is incorrect";
                    return RedirectToAction("LoginFormView", "Home");

                }
                return RedirectToAction("TemplateSelection", "Diligence");
            }
        }

        public IActionResult Logout()
        {
            HttpContext.Session.Clear();
            return RedirectToAction("LoginFormView", "Home");
        }



        public IActionResult DiligenceForm()
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
            return View();
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
