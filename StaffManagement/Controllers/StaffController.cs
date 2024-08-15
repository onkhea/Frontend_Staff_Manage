using Microsoft.AspNetCore.Mvc;
using StaffManagement.Models;
using System.Text.Json;
using OfficeOpenXml;
using System.Reflection.Metadata;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using Document = iText.Layout.Document;
using iText.Layout.Properties;


namespace StaffManagementMVC.Controllers
{
    public class StaffController : Controller
    {
        private readonly HttpClient _client;

        public StaffController()
        {
            _client = new HttpClient();
            _client.BaseAddress = new Uri("https://localhost:7052/api/");
        }

        public async Task<IActionResult> Index()
        {
            var response = await _client.GetAsync("Staffs");
            var staffList = await response.Content.ReadFromJsonAsync<IEnumerable<Staff>>();
            return View(staffList);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Create(Staff staff)
        {
            var response = await _client.PostAsJsonAsync("Staffs", staff);
            if (response.IsSuccessStatusCode)
            {
                return RedirectToAction("Index");
            }
            return View(staff);
        }

        public async Task<IActionResult> Edit(string id)
        {
            var response = await _client.GetAsync($"Staffs/{id}");
            var staff = await response.Content.ReadFromJsonAsync<Staff>();
            return View(staff);
        }

        [HttpPost]
        public async Task<IActionResult> Edit(Staff staff)
        {
            var response = await _client.PutAsJsonAsync($"Staffs/{staff.StaffId}", staff);
            if (response.IsSuccessStatusCode)
            {
                return RedirectToAction("Index");
            }
            return View(staff);
        }

        public async Task<IActionResult> Delete(string id)
        {
            var response = await _client.DeleteAsync($"Staffs/{id}");
            return RedirectToAction("Index");
        }

        public IActionResult AdvancedSearch()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> AdvancedSearch(string staffId, int? gender, DateTime? fromDate, DateTime? toDate)
        {
            var query = $"Staffs/Search?staffId={staffId}&gender={gender}&fromDate={fromDate}&toDate={toDate}";
            var response = await _client.GetAsync(query);
            var staffList = await response.Content.ReadFromJsonAsync<IEnumerable<Staff>>();
            return View("Index", staffList);
        }
        // export excel
 

public async Task<IActionResult> ExportToExcel(string staffId, int? gender, DateTime? fromDate, DateTime? toDate)
    {
        var query = $"Staffs/Search?staffId={staffId}&gender={gender}&fromDate={fromDate:yyyy-MM-dd}&toDate={toDate:yyyy-MM-dd}";
        var response = await _client.GetAsync(query);

        if (!response.IsSuccessStatusCode)
        {
            TempData["ErrorMessage"] = "Failed to fetch staff data.";
            return RedirectToAction("Index");
        }

        var responseContent = await response.Content.ReadAsStringAsync();
        IEnumerable<Staff> staffList;

        try
        {
            staffList = JsonSerializer.Deserialize<IEnumerable<Staff>>(responseContent, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                PropertyNameCaseInsensitive = true
            });
        }
        catch (JsonException ex)
        {
            TempData["ErrorMessage"] = "Error deserializing staff data.";
            return RedirectToAction("Index");
        }

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("StaffList");
            worksheet.Cells["A1"].Value = "Staff ID";
            worksheet.Cells["B1"].Value = "Full Name";
            worksheet.Cells["C1"].Value = "Birthday";
            worksheet.Cells["D1"].Value = "Gender";

            int row = 2;
            foreach (var staff in staffList)
            {
                worksheet.Cells[row, 1].Value = staff.StaffId;
                worksheet.Cells[row, 2].Value = staff.FullName;
                worksheet.Cells[row, 3].Value = staff.Birthday.ToString("yyyy-MM-dd");
                worksheet.Cells[row, 4].Value = staff.Gender == 1 ? "Male" : "Female";
                row++;
            }

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            var fileName = "StaffList.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }


    //export PDF

public async Task<IActionResult> ExportToPDF(string staffId, int? gender, DateTime? fromDate, DateTime? toDate)
    {
        var query = $"Staffs/Search?staffId={staffId}&gender={gender}&fromDate={fromDate:yyyy-MM-dd}&toDate={toDate:yyyy-MM-dd}";
        var response = await _client.GetAsync(query);

        if (!response.IsSuccessStatusCode)
        {
            TempData["ErrorMessage"] = "Failed to fetch staff data.";
            return RedirectToAction("Index");
        }

        var responseContent = await response.Content.ReadAsStringAsync();
        IEnumerable<Staff> staffList;

        try
        {
            staffList = JsonSerializer.Deserialize<IEnumerable<Staff>>(responseContent, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                PropertyNameCaseInsensitive = true
            });
        }
        catch (JsonException ex)
        {
            TempData["ErrorMessage"] = "Error deserializing staff data.";
            return RedirectToAction("Index");
        }

        using (var ms = new MemoryStream())
        {
            var writer = new PdfWriter(ms);
            var pdf = new PdfDocument(writer);
            var document = new Document(pdf);

            document.Add(new Paragraph("Staff List")
                .SetFontSize(18)
                .SetBold());

            Table table = new Table(UnitValue.CreatePercentArray(4)).UseAllAvailableWidth();
            table.AddHeaderCell("Staff ID");
            table.AddHeaderCell("Full Name");
            table.AddHeaderCell("Birthday");
            table.AddHeaderCell("Gender");

            foreach (var staff in staffList)
            {
                table.AddCell(staff.StaffId);
                table.AddCell(staff.FullName);
                table.AddCell(staff.Birthday.ToString("yyyy-MM-dd"));
                table.AddCell(staff.Gender == 1 ? "Male" : "Female");
            }

            document.Add(table);
            document.Close();

            var fileName = "StaffList.pdf";
            return File(ms.ToArray(), "application/pdf", fileName);
        }
    }


}
}
