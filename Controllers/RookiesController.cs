using OfficeOpenXml;
using Asp_MVC1.Models;
using Microsoft.AspNetCore.Mvc;

namespace Asp_MVC1.Controllers
{
    [Route("NashTech/Rookies")]
    public class RookiesController : Controller
    {
        private static List<Person> persons = new List<Person>
        {
            new Person { FirstName = "Van A", LastName = "Nguyen", Gender = "Male", DateOfBirth = new DateTime(2000, 3, 15), PhoneNumber = "1234567890", BirthPlace = "Ha Noi", IsGraduated = true },
            new Person { FirstName = "Van B", LastName = "Nguyen", Gender = "Male", DateOfBirth = new DateTime(1999, 6, 20), PhoneNumber = "9876543210", BirthPlace = "Sai Gon", IsGraduated = false },
            new Person { FirstName = "Thi C", LastName = "Nguyen", Gender = "Female", DateOfBirth = new DateTime(2001, 6, 20), PhoneNumber = "9876543210", BirthPlace = "Sai Gon", IsGraduated = false },
        };
        [HttpGet]
        public IActionResult Index()
        {
            return View(persons);
        }
        [HttpGet("MaleMembers")]
        public IActionResult MaleMembers()
        {
            var maleMembers = persons.Where(p => p.Gender == "Male").ToList();
            return View("Result", maleMembers);
        }
        [HttpGet("OldestMember")]
        public IActionResult OldestMember()
        {
            var oldestMember = persons.OrderBy(p => p.DateOfBirth).FirstOrDefault();
            return View("Result", oldestMember);
        }
        [HttpGet("FullNames")]
        public IActionResult FullNames()
        {
            var fullNames = persons.Select(p => $"{p.LastName} {p.FirstName}").ToList();
            return View("Result", fullNames);
        }
        [HttpGet("FilterByBirthYear")]
        public IActionResult FilterByBirthYear(int year)
        {
            var filteredMembers = persons.Where(p => p.DateOfBirth.Year == year).ToList();
            return View("Result", filteredMembers);
        }
        [HttpGet("ExportToExcel")]
        public IActionResult ExportToExcel()
        {
            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Persons");
                worksheet.Cells["A1"].LoadFromCollection(persons.Select(p => new
                {
                    p.FirstName,
                    p.LastName,
                    p.Gender,
                    DateOfBirth = p.DateOfBirth.ToString("yyyy-MM-dd"),
                    p.PhoneNumber,
                    p.BirthPlace,
                    p.IsGraduated
                }), true);

                package.Save();
            }

            stream.Position = 0;
            var fileName = $"Persons_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        [HttpPost]
        public IActionResult HandleOption(int option)
        {
            switch (option)
            {
                case 1:
                    return RedirectToAction("MaleMembers");
                case 2:
                    return RedirectToAction("OldestMember");
                case 3:
                    return RedirectToAction("FullNames");
                case 4:
                    return RedirectToAction("FilterByBirthYear", new { year = 2000 });
                case 5:
                    return RedirectToAction("FilterByBirthYear", new { year = 2001 });
                case 6:
                    return RedirectToAction("FilterByBirthYear", new { year = 1999 });
                case 7:
                    return RedirectToAction("ExportToExcel");
                default:
                    return RedirectToAction("Index");
            }
        }
    }
}
