using Microsoft.AspNetCore.Mvc;
using NibulonTest.Models;
using System.Diagnostics;

namespace NibulonTest.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return RedirectToAction("dataList", "GrainData");
        }

        public IActionResult Privacy()
        {
            return View();
        }     
    }
}