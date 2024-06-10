using ElectronNET.API;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json.Linq;
using ReportSys.DAL;
using ReportSys.Pages.Services;
using System.Threading.Tasks;

namespace ReportSys.Pages.EntryAccess
{
    public class IndexModel : ServicesPage
    {
        private readonly ReportSysContext _context;

        public IndexModel(ReportSysContext context)
        {
            _context = context;
        }

        [BindProperty]
        public List<SelectListItem> AllEmployeeList { get; set; }


        [BindProperty]
        public int SelectedEmployeeId { get; set; }


        [BindProperty]
        public string EmployeeNumber { get; set; }

        public async Task<IActionResult> OnGetAsync()
        {


            await GetAllEmployeesAsync(_context);

            AllEmployeeList = AllEmployeesSL.ToList();

            return Page();

        }

        public async Task<IActionResult> OnPostAsync()
        {
            //if (string.IsNullOrEmpty(EmployeeNumber))
            //{
            //    ModelState.AddModelError(string.Empty, "��������� ����� ����������.");
            //    return Page();
            //}

            var employee = await _context.Employees
                .Include(e => e.Position)
                .FirstOrDefaultAsync(e => e.Id.ToString() == SelectedEmployeeId.ToString());

            if (employee == null)
            {
                ModelState.AddModelError(string.Empty, "��������� ����� �� ������.");
                return Page();
            }

            //HttpContext.Session.SetString("EmployeeNumber", EmployeeNumber);
            //HttpContext.Session.SetString("EmployeeNumber", SelectedEmployeeId.ToString());
          




            switch (employee.Position.AccessLevel)
            {
                case 0:
                    return RedirectToPage("/PageAccess0/Index", new { myParameter = SelectedEmployeeId.ToString() });
                case 1:
                    return RedirectToPage("/PageAccess1/Index", new { myParameter = SelectedEmployeeId.ToString() });
                default:
                    ModelState.AddModelError(string.Empty, "����������� ������.");
                    return Page();
            }
        }
    }
}
