using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ReportSys.DAL;
using ReportSys.DAL.Entities;
using ReportSys.Pages.Services;
using System.Globalization;
using System.Xml.Linq;

namespace ReportSys.Pages.PageAccess0
{
    public class IndexModel : ServicesPage
    {
        private readonly ReportSysContext _context;

        [BindProperty]
        public string StartDateString { get; set; }
        [BindProperty]
        public string EndDateString { get; set; }
        [BindProperty]
        public string _id { get; set; }
        public string _name { get; set; }

        public DateOnly StartDate
        {
            get
            {
                // ѕопытка преобразовать строку в DateOnly
                if (DateOnly.TryParse(StartDateString, out var date))
                {
                    return date;
                }
                // ¬озврат значени€ по умолчанию, если преобразование не удалось
                return DateOnly.FromDateTime(DateTime.Now);
            }
        }
        public DateOnly EndDate
        {
            get
            {
                // ѕопытка преобразовать строку в DateOnly
                if (DateOnly.TryParse(EndDateString, out var date))
                {
                    return date;
                }
                // ¬озврат значени€ по умолчанию, если преобразование не удалось
                return DateOnly.FromDateTime(DateTime.Now);
            }
        }
        public IndexModel(ReportSysContext context)
        {
            _context = context; 
        }

        public async Task<IActionResult> OnGetAsync(string myParameter)
        {
            _id = myParameter;
            // »нициализаци€ строки StartDateString текущей датой в формате yyyy-MM-dd
            StartDateString = DateOnly.FromDateTime(DateTime.Now).ToString("MM/dd/yyyy");
            var employee = await _context.Employees
                .Include(e => e.Department)
                .FirstOrDefaultAsync(e => e.Id.ToString() == _id);

            _name = employee.FirstName + " " + employee.SecondName + " " + employee.LastName;

            return Page();
        }
        public async Task<IActionResult> OnPostAsync()
        {
            string employeeNumber = _id;

            List<string> employeeNumbers = new List<string> { employeeNumber };

            return await CreateXlsxFirst(_context, employeeNumbers, StartDate, EndDate);
        }
    }
}
