﻿using System.Web.Mvc;

namespace WebApiContrib.Formatting.Xlsx.NetStandard.Sample.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Sample web application";

            return View();
        }
    }
}
