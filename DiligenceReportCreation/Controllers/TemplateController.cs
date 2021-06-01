using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DiligenceReportCreation.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace DiligenceReportCreation.Controllers
{
    public class TemplateController : Controller
    {
        private readonly DataBaseContext _context;
        private readonly IConfiguration _config;
        public TemplateController(IConfiguration config, DataBaseContext context)
        {
            _config = config;
            _context = context;
        }
        public IActionResult Index()
        {
            return View();
        }       
    }
}