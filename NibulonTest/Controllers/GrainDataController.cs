using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis.Elfie.Serialization;
using NibulonTest.Models;
using NibulonTest.Services;

namespace NibulonTest.Controllers
{
    public class GrainDataController : Controller
    {
        private readonly GrainDataService _grainDataService;

        public GrainDataController()
        {
            _grainDataService = new GrainDataService();
        }
        public IActionResult DataList()
        {
            var data = _grainDataService.GetGrainDataRecords();

            return View(data);
        }

        [HttpGet]
        public IActionResult Edit(int id)
        {
            var data = _grainDataService.GetGrainDataRecordsById(id);

            if(data.Count == 0)
                return NotFound();

            return View(data.FirstOrDefault());
        }

        [HttpPost]
        public IActionResult Edit(GrainDataRecord dataRecord)
        {
            var res = _grainDataService.UpdateDataGrainRecord(dataRecord);

            if (!res)
                return NotFound();

            return RedirectToAction("DataList");
        }

        [HttpGet]
        public async Task<IActionResult> GetReport(DateTime begin, DateTime end)
        {
            if (end == DateTime.MinValue)
                end = begin;

            var file = await _grainDataService.GetReportFile(begin, end);

            return File(file, "application/vnd.ms-excel", $"Report { begin.ToShortDateString() } - { end.Date.ToShortDateString() } + { DateTime.Now.ToLocalTime() }.xlsx");
        }                
    }
}
