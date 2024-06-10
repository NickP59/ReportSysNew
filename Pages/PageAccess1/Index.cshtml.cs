using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ReportSys.DAL;
using ReportSys.Pages.Services;
using System.Drawing;
using System.Runtime.Intrinsics.Arm;
//using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ReportSys.Pages.PageAccess1
{
    public class IndexModel : ServicesPage
    {


        private readonly ReportSysContext _context;

        public IndexModel(ReportSysContext context)
        {
            _context = context;
        }

        [BindProperty]
        public List<SelectListItem> EmployeeList { get; set; }


        [BindProperty]
        public List<int> SelectedEmployeeIds { get; set; }



        [BindProperty]
        public List<SelectListItem> DepartList { get; set; }


        [BindProperty]
        public List<int> SelectedDepartIds { get; set; }

        [BindProperty]
        public string _id { get; set; }
        public string _name { get; set; }

        [BindProperty]
        public string StartDateString { get; set; }
        [BindProperty]
        public string EndDateString { get; set; }

        [BindProperty]
        public List<DateOnly> Dates { get; set; }


        [BindProperty]
        public string Action { get; set; }



        public DateOnly StartDate
        {
            get
            {
                // Попытка преобразовать строку в DateOnly
                if (DateOnly.TryParse(StartDateString, out var date))
                {
                    return date;
                }
                // Возврат значения по умолчанию, если преобразование не удалось
                return DateOnly.FromDateTime(DateTime.Now);
            }
        }
        public DateOnly EndDate
        {
            get
            {
                // Попытка преобразовать строку в DateOnly
                if (DateOnly.TryParse(EndDateString, out var date))
                {
                    return date;
                }
                // Возврат значения по умолчанию, если преобразование не удалось
                return DateOnly.FromDateTime(DateTime.Now);
            }
        }

        public async Task<IActionResult> OnGetAsync(string myParameter)
        {
            var employeeNumber = myParameter;
            _id = myParameter;

            if (string.IsNullOrEmpty(employeeNumber))
            {
                return RedirectToPage("/Error"); // Перенаправление на страницу ошибки, если нет номера сотрудника
            }

            var employee = await _context.Employees
                .Include(e => e.Department)
                .FirstOrDefaultAsync(e => e.Id.ToString() == employeeNumber);


            _name = employee.FirstName + " " + employee.SecondName + " " + employee.LastName;

            if (employee == null)
            {
                return RedirectToPage("/Error"); // Перенаправление на страницу ошибки, если сотрудник не найден
            }

            await EmployeesFromDepartAsync(_context, employee);
            await DepartmentsFromDepartAsync(_context, employee.DepartmentId);

            // Заполнение свойств EmployeeList и DepartList
            EmployeeList = EmployeesSL.ToList();
            DepartList = DepartmentsSL.ToList();

            return Page();
        }




        public async Task<IActionResult> OnPostAsync()
        {
            if (Action == "Action1")
            {
                return await HandleAction1();
            }
            else if (Action == "Action2")
            {
                return await HandleAction2();
            }

            return Page();
        }

        private async Task<IActionResult> HandleAction1()
        {
            // Логика для Action1
            // Например, перенаправление на другую страницу или возврат данных
            var employeeNumbers = new List<string>();

            foreach(var empId in SelectedEmployeeIds)
            {
                employeeNumbers.Add(empId.ToString());
            }
            return await CreateXlsxFirst(_context, employeeNumbers, StartDate, EndDate);
        }

        private async Task<IActionResult> HandleAction2()
        {
         
            if (StartDate > EndDate)
            {
                TempData["Message"] = "Start date cannot be later than end date.";
                return Page();
            }

            Dates = new List<DateOnly>();
            for (var date = StartDate; date <= EndDate; date = date.AddDays(1))
            {
                if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
                {
                    continue; // Пропускаем субботу и воскресенье
                }
                Dates.Add(date);
            }

            var employeeNumber = _id;

            var employee = await _context.Employees.Include(e => e.WorkSchedule)
               .Include(e => e.Events).ThenInclude(s => s.EventType)
               .Include(e => e.Unavailabilitys).ThenInclude(s => s.UnavailabilityType)
               .FirstOrDefaultAsync(e => e.Id.ToString() == employeeNumber);

            var deps = await _context.Employees.Include(e => e.Department)
                                                    .Include(e => e.Position)
                                                    .Include(e => e.Events).ThenInclude(s => s.EventType)
                                                    .Where(e => e.DepartmentId == employee.DepartmentId).ToListAsync();

            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Departments");

                worksheet.Cells[1, 1].Value = $"Сведения о времени прихода на работу, уходе с работы и отсутствиях с {StartDate.ToString("dd-MM-yyyy")} по {EndDate.ToString("dd-MM-yyyy")}";
                worksheet.Cells[2, 1].Value = $"Дата составления: {DateOnly.FromDateTime(DateTime.Now).ToString("dd-MM-yyyy")} {TimeOnly.FromDateTime(DateTime.Now).ToString("HH:mm:ss")}";
                worksheet.Cells[3, 1].Value = "ПП";
                worksheet.Cells[3, 2].Value = "Наименование штатной должности";
                worksheet.Cells[3, 3].Value = "ФИО";
                worksheet.Cells[3, 4].Value = "Событие";
                worksheet.Cells[3, 5].Value = "Время прихода ухода";

                worksheet.Column(1).Width = 10;
                worksheet.Column(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Column(2).Width = 35;
                worksheet.Column(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Column(3).Width = 25;
                worksheet.Column(3).Style.WrapText = true;
                worksheet.Column(4).Width = 15;
                worksheet.Column(4).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                worksheet.Cells["A3:A4"].Merge = true; 
                worksheet.Cells["B3:B4"].Merge = true; 
                worksheet.Cells["C3:C4"].Merge = true; 
                worksheet.Cells["D3:D4"].Merge = true;

                

                for (int i = 0; i < Dates.Count(); i++)
                {
                    worksheet.Cells[4, i + 5].Value = Dates[i].ToString("dd.MM.yyyy");
                    worksheet.Column(i + 5).Width = 15;
                    worksheet.Column(i + 5).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Column(i + 5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                }
                worksheet.Cells[$"E3:{GetExcelColumnName(Dates.Count()+4)}3"].Merge = true;
                for (var i = 6; i <= Dates.Count() + 4; i++)
                {
                    worksheet.Column(i).OutlineLevel = 1;
                    worksheet.Column(i).Collapsed = true;
                }


                int baseColumnIndex = 5 + Dates.Count();

                worksheet.Cells[3, baseColumnIndex].Value = "Кол-во \"-\" откл.";
                worksheet.Cells[$"{GetExcelColumnName(baseColumnIndex)}3:{GetExcelColumnName(baseColumnIndex + 1)}3"].Merge = true;

                worksheet.Cells[3, baseColumnIndex + 2].Value = "Общее время";
                worksheet.Cells[$"{GetExcelColumnName(baseColumnIndex + 2)}3:{GetExcelColumnName(baseColumnIndex + 3)}3"].Merge = true;

                worksheet.Cells[3, baseColumnIndex + 4].Value = "Кол-во \"+\" откл.";
                worksheet.Cells[$"{GetExcelColumnName(baseColumnIndex + 4)}3:{GetExcelColumnName(baseColumnIndex + 5)}3"].Merge = true;

                worksheet.Cells[3, baseColumnIndex + 6].Value = "Общее время";
                worksheet.Cells[$"{GetExcelColumnName(baseColumnIndex + 6)}3:{GetExcelColumnName(baseColumnIndex + 7)}3"].Merge = true;

                worksheet.Cells[4, 5 + Dates.Count()].Value = "ед";
                worksheet.Column(5 + Dates.Count()).Width = 10;
                worksheet.Cells[4, 6 + Dates.Count()].Value = "%";
                worksheet.Column(6 + Dates.Count()).Width = 10;
                worksheet.Cells[4, 7 + Dates.Count()].Value = "ч";
                worksheet.Column(7 + Dates.Count()).Width = 10;
                worksheet.Cells[4, 8 + Dates.Count()].Value = "%";
                worksheet.Column(8 + Dates.Count()).Width = 10;
                worksheet.Cells[4, 9 + Dates.Count()].Value = "ед";
                worksheet.Column(9 + Dates.Count()).Width = 10;
                worksheet.Cells[4, 10 + Dates.Count()].Value = "%";
                worksheet.Column(10 + Dates.Count()).Width = 10;
                worksheet.Cells[4, 11 + Dates.Count()].Value = "ч";
                worksheet.Column(11+ Dates.Count()).Width = 10;
                worksheet.Cells[4, 12 + Dates.Count()].Value = "%";
                worksheet.Column(12+ Dates.Count()).Width = 10;

                // Форматирование ячеек заголовков
                worksheet.Cells[$"A1:{GetExcelColumnName(12 + Dates.Count())}2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[$"A1:{GetExcelColumnName(12 + Dates.Count())}2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                // Добавление границ к заголовкам
                worksheet.Cells[$"A3:{GetExcelColumnName(12 + Dates.Count())}4"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A3:{GetExcelColumnName(12 + Dates.Count())}4"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A3:{GetExcelColumnName(12 + Dates.Count())}4"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[$"A3:{GetExcelColumnName(12 + Dates.Count())}4"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                //var UpDepIds = FindTopLevelDepartments(SelectedDepartIds, _context);
                var UpDepIds = await _context.Hierarchies
                         .Where(e => SelectedDepartIds.Contains(e.UpperDepartmentId))
                         .Select(e => e.LowerDepartmentId)
                         .ToListAsync();


                int row = 5;
                var MinSumS = 0;
                var PlusSumS = 0;
                var MinSumE = 0;
                var PlusSumE = 0;
                var MinTimeS = new TimeSpan();
                var MinTimeE = new TimeSpan();
                var PlusTimeS = new TimeSpan();
                var PlusTimeE = new TimeSpan();
                if (UpDepIds.Count() == 0)
                {
                    foreach(var depId in SelectedDepartIds)
                    {
                        var reportData = new ReportData
                        {
                            Worksheet = worksheet,
                            Row = row,
                            MinSumE = MinSumE,
                            MinSumS = MinSumS,
                            PlusSumE = PlusSumE,
                            PlusSumS = PlusSumS,
                            MinTimeS = MinTimeS,
                            MinTimeE = MinTimeE,
                            PlusTimeS = PlusTimeS,
                            PlusTimeE = PlusTimeE
                        };

                        reportData = await Smth(_context, Dates, depId, reportData);

                        // Обновите значения после вызова метода
                        worksheet = reportData.Worksheet;
                        row = reportData.Row;
                        MinSumE = reportData.MinSumE;
                        MinSumS = reportData.MinSumS;
                        PlusSumE = reportData.PlusSumE;
                        PlusSumS = reportData.PlusSumS;
                        MinTimeS = reportData.MinTimeS;
                        MinTimeE = reportData.MinTimeE;
                        PlusTimeS = reportData.PlusTimeS;
                        PlusTimeE = reportData.PlusTimeE;

                    }
                }
                else
                {
                    var datesSet = new HashSet<DateOnly>(Dates); // Преобразуем список дат в HashSet для быстрой проверки
                    
                    var dep = await _context.Departments
                        .Include(d => d.Employees).ThenInclude(e => e.Unavailabilitys)
                        .Include(d => d.Employees).ThenInclude(e => e.Position)
                        .Include(d => d.Employees).ThenInclude(e => e.WorkSchedule)
                        .Include(d => d.Employees).ThenInclude(e => e.Events.Where(e => datesSet.Contains(e.Date)))
                        .FirstOrDefaultAsync(d => d.Id == employee.DepartmentId);
                   

                    worksheet.Cells[row, 1].Value = dep.Name;
                    worksheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Row(row).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    row++;
                    worksheet.Cells[$"A{row - 1}:{GetExcelColumnName(12 + Dates.Count())}{row - 1}"].Merge = true;
                    worksheet.Cells[$"A{row - 1}:{GetExcelColumnName(12 + Dates.Count())}{row - 1}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[$"A{row - 1}:{GetExcelColumnName(12 + Dates.Count())}{row - 1}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[$"A{row - 1}:{GetExcelColumnName(12 + Dates.Count())}{row - 1}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[$"A{row - 1}:{GetExcelColumnName(12 + Dates.Count())}{row - 1}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    var numPP = 0;
                    var numWorkDays = 0;
                    var numNegDevsS = 0;
                    var numPosDevsS = 0;
                    var numNegDevsE = 0;
                    var numPosDevsE = 0;
                    var timNegDevsS = new TimeSpan();
                    var timPosDevsS = new TimeSpan();
                    var timNegDevsE = new TimeSpan();
                    var timPosDevsE = new TimeSpan();
                    // Создание временного интервала в 8 часов
                    TimeSpan eightHours = TimeSpan.FromHours(8);

                    foreach (var emp in dep.Employees)
                    {

                        var startTime = emp.WorkSchedule.Arrival;
                        var endTime = emp.WorkSchedule.Exit;


                        worksheet.Cells[row, 1].Value = numPP;
                        row++;
                        worksheet.Cells[$"A{row - 1}:A{row}"].Merge = true;

                        numPP++;


                        worksheet.Cells[row - 1, 2].Value = emp.Position.Name;

                        worksheet.Cells[$"B{row - 1}:B{row}"].Merge = true;


                        worksheet.Cells[row - 1, 3].Value = emp.FirstName + " " + emp.SecondName + " " + emp.LastName;

                        worksheet.Cells[$"C{row - 1}:C{row}"].Merge = true;

                        worksheet.Cells[row - 1, 4].Value = "приход";
                        worksheet.Cells[row, 4].Value = "уход";
                        var i = 5;
                        numWorkDays = 0;
                        numWorkDays = 0;
                        numWorkDays = 0;
                        numWorkDays = 0;
                        numPosDevsE = 0;
                        timNegDevsS = new TimeSpan();
                        timPosDevsS = new TimeSpan();
                        timNegDevsE = new TimeSpan();
                        timPosDevsE = new TimeSpan();
                        foreach (var date in Dates)
                        {
                            var events = await _context.Events
                                        .Include(e => e.EventType)
                                        .Where(e => e.EmployeeId == emp.Id && e.Date == date)
                                        .ToListAsync();



                            var unf = await _context.Unavailabilitys
                                          .Include(e => e.UnavailabilityType)
                                           .Include(e => e.Employee)
                                          .Where(e => e.Date == date && e.Employee == emp)
                                          .FirstOrDefaultAsync();


                            if (events != null && events.Count != 0)
                            {
                                // Найти первый евент с EventTypeId == 0
                                var firstEventType0 = events.FirstOrDefault(e => e.EventType.Id == 1);

                                // Найти последний евент с EventTypeId == 1
                                var lastEventType1 = events.LastOrDefault(e => e.EventType.Id == 2);



                                if (firstEventType0 != null)
                                {
                                    worksheet.Cells[row - 1, i].Value = firstEventType0.Time.ToString("HH:mm:ss");


                                    if (firstEventType0.Time - startTime > TimeSpan.FromMinutes(3) && firstEventType0.Time > startTime)
                                    {

                                        if (unf != null)
                                        {
                                            if (unf.UnavailabilityTypeId != 4)
                                            {
                                                colorCell(worksheet, row - 1, i, Color.SkyBlue);

                                            }
                                            else if (firstEventType0.Time < unf.UnavailabilityFrom || firstEventType0.Time > unf.UnavailabilityBefore)
                                            {
                                                colorCell(worksheet, row - 1, i, Color.SandyBrown);
                                                numNegDevsS++;
                                                timNegDevsS = timNegDevsS.Add(firstEventType0.Time - startTime);
                                            }
                                            else
                                            {
                                                colorCell(worksheet, row - 1, i, Color.Khaki);
                                            }
                                        }
                                        else
                                        {
                                            colorCell(worksheet, row - 1, i, Color.SandyBrown);
                                            numNegDevsS++;
                                            timNegDevsS = timNegDevsS.Add(firstEventType0.Time - startTime);
                                        }

                                    }
                                    if (startTime - firstEventType0.Time > TimeSpan.FromMinutes(3) && startTime > firstEventType0.Time)
                                    {
                                        if (unf != null)
                                        {
                                            if (unf.UnavailabilityTypeId != 4)
                                            {
                                                colorCell(worksheet, row - 1, i, Color.SkyBlue);

                                            }

                                        }
                                        else
                                        {
                                            colorCell(worksheet, row - 1, i, Color.LightGreen);
                                            numPosDevsS++;
                                            timPosDevsS = timPosDevsS.Add(startTime - firstEventType0.Time);
                                        }
                                    }
                                }
                                else
                                {
                                    colorCell(worksheet, row - 1, i, Color.Pink);
                                }
                                if (lastEventType1 != null)
                                {
                                    worksheet.Cells[row, i].Value = lastEventType1.Time.ToString("HH:mm:ss");

                                    if (endTime - lastEventType1.Time > TimeSpan.FromMinutes(3) && endTime > lastEventType1.Time)
                                    {

                                        if (unf != null)
                                        {
                                            if (unf.UnavailabilityTypeId != 4)
                                            {
                                                colorCell(worksheet, row, i, Color.SkyBlue);

                                            }
                                            else if (lastEventType1.Time < unf.UnavailabilityFrom || lastEventType1.Time > unf.UnavailabilityBefore)
                                            {
                                                colorCell(worksheet, row, i, Color.SandyBrown);
                                                numNegDevsE++;
                                                timNegDevsE = timNegDevsE.Add(endTime - lastEventType1.Time);
                                            }
                                            else
                                            {
                                                colorCell(worksheet, row, i, Color.Khaki);

                                            }
                                        }
                                        else
                                        {
                                            colorCell(worksheet, row, i, Color.SandyBrown);
                                            numNegDevsE++;
                                            timNegDevsE = timNegDevsE.Add(endTime - lastEventType1.Time);
                                        }

                                    }
                                    if (lastEventType1.Time - endTime > TimeSpan.FromMinutes(3) && lastEventType1.Time > endTime)
                                    {
                                        if (unf != null)
                                        {
                                            if (unf.UnavailabilityTypeId != 4)
                                            {
                                                colorCell(worksheet, row, i, Color.SkyBlue);

                                            }

                                        }
                                        else
                                        {
                                            colorCell(worksheet, row, i, Color.LightGreen);
                                            numPosDevsE++;
                                            timPosDevsE = timPosDevsE.Add(lastEventType1.Time - endTime);
                                        }

                                    }

                                }
                                else
                                {
                                    colorCell(worksheet, row, i, Color.Pink);
                                }
                                if (lastEventType1 != null || firstEventType0 != null)
                                {
                                    numWorkDays++;
                                }
                                // Добавление бордера к каждой заполненной ячейке
                                for (int col = 1; col <= 12 + Dates.Count(); col++)
                                {
                                    worksheet.Cells[i - 1, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[i - 1, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[i - 1, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[i - 1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[i, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[i, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[i, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                    worksheet.Cells[i, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                }
                                i++;
                            }
                            else
                            {
                                
                                if (unf != null)
                                {
                                    if (unf.UnavailabilityTypeId != 4)
                                    {
                                        colorCell(worksheet, row - 1, i, Color.Khaki);
                                        colorCell(worksheet, row, i, Color.Khaki);

                                    }

                                }
                                else
                                {
                                    colorCell(worksheet, row - 1, i, Color.SandyBrown);
                                    colorCell(worksheet, row, i, Color.SandyBrown);
                                }
                                i++;
                            }

                        }

                        if (numWorkDays != 0)
                        {
                            worksheet.Cells[row - 1, 5 + Dates.Count()].Value = numNegDevsS;
                            worksheet.Cells[row - 1, 6 + Dates.Count()].Value = CalculatePercentageDevs(numNegDevsS, numWorkDays);
                           

                            worksheet.Cells[row - 1, 7 + Dates.Count()].Value = FormatTimeSpan(timNegDevsS);
                            worksheet.Cells[row - 1, 8 + Dates.Count()].Value = CalculatePercentageTimeDevs(timNegDevsS, numWorkDays);

                            worksheet.Cells[row - 1, 9 + Dates.Count()].Value = numPosDevsS;
                            worksheet.Cells[row - 1, 10 + Dates.Count()].Value = CalculatePercentageDevs(numPosDevsS, numWorkDays);

                            worksheet.Cells[row - 1, 11 + Dates.Count()].Value = FormatTimeSpan(timPosDevsS);
                            worksheet.Cells[row - 1, 12 + Dates.Count()].Value = CalculatePercentageTimeDevs(timPosDevsS, numWorkDays);



                            worksheet.Cells[row, 5 + Dates.Count()].Value = numNegDevsE;
                            worksheet.Cells[row, 6 + Dates.Count()].Value = CalculatePercentageDevs(numNegDevsE, numWorkDays);

                            worksheet.Cells[row, 7 + Dates.Count()].Value = FormatTimeSpan(timNegDevsE);
                            worksheet.Cells[row, 8 + Dates.Count()].Value = CalculatePercentageTimeDevs(timNegDevsE, numWorkDays);

                            worksheet.Cells[row, 9 + Dates.Count()].Value = numPosDevsE;
                            worksheet.Cells[row, 10 + Dates.Count()].Value = CalculatePercentageDevs(numPosDevsE, numWorkDays);

                            worksheet.Cells[row, 11 + Dates.Count()].Value = FormatTimeSpan(timPosDevsE);
                            worksheet.Cells[row, 12 + Dates.Count()].Value = CalculatePercentageTimeDevs(timPosDevsE, numWorkDays);

                            MinSumS += numNegDevsS;
                            MinSumE += numNegDevsE;
                            PlusSumS += numPosDevsS;
                            PlusSumE += numPosDevsE;
                            MinTimeS = MinTimeS.Add(timNegDevsS);
                            MinTimeE = MinTimeE.Add(timNegDevsE);
                            PlusTimeS = PlusTimeS.Add(timPosDevsS);
                            PlusTimeE = PlusTimeE.Add(timPosDevsE);
                        }
                        for (int col = 1; col <= 12 + Dates.Count(); col++)
                        {
                            worksheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }
                        row++;
                    }



                    foreach (var depId in UpDepIds)
                    {
                        var reportData = new ReportData
                        {
                            Worksheet = worksheet,
                            Row = row,
                            MinSumE = MinSumE,
                            MinSumS = MinSumS,
                            PlusSumE = PlusSumE,
                            PlusSumS = PlusSumS,
                            MinTimeS = MinTimeS,
                            MinTimeE = MinTimeE,
                            PlusTimeS = PlusTimeS,
                            PlusTimeE = PlusTimeE
                        };

                        reportData = await Smth(_context, Dates, depId, reportData);

                        // Обновите значения после вызова метода
                        worksheet = reportData.Worksheet;
                        row = reportData.Row;
                        MinSumE = reportData.MinSumE;
                        MinSumS = reportData.MinSumS;
                        PlusSumE = reportData.PlusSumE;
                        PlusSumS = reportData.PlusSumS;
                        MinTimeS = reportData.MinTimeS;
                        MinTimeE = reportData.MinTimeE;
                        PlusTimeS = reportData.PlusTimeS;
                        PlusTimeE = reportData.PlusTimeE;

                    }
                }


                worksheet.Cells[row, 4 + Dates.Count()].Value = "Итого приход: ";
                worksheet.Cells[row, 5 + Dates.Count()].Value = MinSumS;
                worksheet.Cells[row, 7 + Dates.Count()].Value = FormatTimeSpan(MinTimeS);
                worksheet.Cells[row, 9 + Dates.Count()].Value = PlusSumS;
                worksheet.Cells[row, 11 + Dates.Count()].Value = FormatTimeSpan(PlusTimeS);

                worksheet.Cells[row+1, 4 + Dates.Count()].Value = "Итого уход: ";
                worksheet.Cells[row+1, 5 + Dates.Count()].Value = MinSumE;
                worksheet.Cells[row+1, 7 + Dates.Count()].Value = FormatTimeSpan(MinTimeE);
                worksheet.Cells[row+1, 9 + Dates.Count()].Value = PlusSumE;
                worksheet.Cells[row+1, 11 + Dates.Count()].Value = FormatTimeSpan(PlusTimeE);

                worksheet.Cells[row + 2, 4 + Dates.Count()].Value = "Всего: ";
                worksheet.Cells[row + 2, 5 + Dates.Count()].Value = MinSumE + MinSumS;
                worksheet.Cells[row + 2, 7 + Dates.Count()].Value = FormatTimeSpan(MinTimeE + MinTimeS);
                worksheet.Cells[row + 2, 9 + Dates.Count()].Value = PlusSumE + PlusSumS;
                worksheet.Cells[row + 2, 11 + Dates.Count()].Value = FormatTimeSpan(PlusTimeE + PlusTimeS);

                worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                package.Save();
                //// Указываем относительный путь для сохранения файла
                //string relativePath = @"Output\Departments.xlsx";
                //// Получаем полный путь в папке проекта
                //string fullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath);
                //// Убедимся, что директория существует
                //string directoryPath = Path.GetDirectoryName(fullPath); if (!Directory.Exists(directoryPath))
                //{
                //    Directory.CreateDirectory(directoryPath);
                //}
                //// Сохраняем пакет (файл)
                //FileInfo file = new FileInfo(fullPath);
               

                //await package.SaveAsAsync(file);
            }
            stream.Position = 0;
            var fileName = "Departments.xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            return File(stream, contentType, fileName);
        
        }

       
       
    }

}
