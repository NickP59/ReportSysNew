using ElectronNET.API;
using OfficeOpenXml;
using ReportSys.DAL;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
var builder = WebApplication.CreateBuilder(args);

builder.WebHost.UseElectron(args);



// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddDbContext<ReportSysContext>(options =>
{
    //options.UseNpgsql(builder.Configuration.GetConnectionString("Postgres"));
    options.UseSqlite("Filename=MyDatabase.db");
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}



//app.UseSession();

ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
using (var scope = app.Services.CreateScope())
{
    var serviceProvider = scope.ServiceProvider;


    var db = serviceProvider.GetRequiredService<ReportSysContext>();
    await db.Database.EnsureDeletedAsync();
    await db.Database.EnsureCreatedAsync();
    await ReportSysContextSeed.InitializeDb(db);


}




await Electron.WindowManager.CreateWindowAsync();

//app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

//app.UseAuthorization();

app.MapRazorPages();

app.Run();
