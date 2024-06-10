using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Text.Json;

public class ConfigureDbModel : PageModel
{
    private readonly IConfiguration _configuration;

    public ConfigureDbModel(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    [BindProperty]
    public DbSettings DbSettings { get; set; }

    public string Message { get; set; }

    public void OnGet()
    {
        // Загрузка текущих настроек (если они есть)
        DbSettings = _configuration.GetSection("ConnectionStrings").Get<DbSettings>();
    }

    public IActionResult OnPost()
    {
        if (!ModelState.IsValid)
        {
            return Page();
        }

        // Обновление файла appsettings.json
        var appSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
        var json = System.IO.File.ReadAllText(appSettingsPath);
        var jsonObj = JsonDocument.Parse(json);

        var connectionString = $"Server={DbSettings.Server};Database={DbSettings.Database};User Id={DbSettings.Username};Password={DbSettings.Password};";

        // Обновление строки подключения
        using (var doc = JsonDocument.Parse(json))
        {
            var root = doc.RootElement.Clone();
            var newConnectionString = new Dictionary<string, string>
            {
                ["DefaultConnection"] = connectionString
            };

            var jsonString = JsonSerializer.Serialize(new { ConnectionStrings = newConnectionString });
            System.IO.File.WriteAllText(appSettingsPath, jsonString);
        }

        Message = "Настройки подключения успешно сохранены.";

        return Page();
    }
}

public class DbSettings
{
    public string Server { get; set; }
    public string Database { get; set; }
    public string Username { get; set; }
    public string Password { get; set; }
}
