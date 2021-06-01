using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using DiligenceReportCreation.Models;
using DiligenceReportCreation.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.Text;

namespace DiligenceReportCreation.Controllers
{
    public class AdminController : Controller
    {
        private readonly DataBaseContext _context;
        public AdminController(DataBaseContext context)
        {
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult AddNewUser(UserModel user)
        {
            
            UserModel userfromdb = _context.DbUser.Where(x => x.Username == user.Username).FirstOrDefault();
            if(userfromdb == null)
            {
                user.Permissions = "None";
                //user.Permissions = "request_rest_password";
                user.Email = user.Username.ToString().Trim() + "@SterlingCheck.com";
                user.Password = Convert.ToBase64String(Encoding.Unicode.GetBytes("Test@123"));
                user.LastLogin = DateTime.Now;
                _context.DbUser.Add(user);
                _context.SaveChanges();
                TempData["message"] = "New User Added: " + user.Username.ToString();
                // TempData["message"] = "New User Added Email with default password is sent on: "+user.Email.ToString();
                return RedirectToAction("AdminHome", "Admin");

            }
            else
            {
                TempData["message"] = "User with user name: "+user.Username +" already configured";
                return RedirectToAction("AddNewUserform", "Admin");

            }
               
        }
        public IActionResult AddNewUserform()
        {
            return View();
        }



        public IActionResult AdminHome()
        {
            List<UserModel> users = _context.DbUser.OrderBy(column => column.Role).ToList();
            return View(users);
        }


    }
}