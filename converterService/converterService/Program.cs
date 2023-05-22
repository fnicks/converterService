using converterService.Services;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Win32;

await CreateHostBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddHostedService<ConverterService>();
    })
    .ConfigureWebHostDefaults(webBuilder =>
    {
        LoadEnvironmentVariablesFromFile();
        string port = Environment.GetEnvironmentVariable("PORT");
        OfficeExportsPdfA();

        webBuilder.UseUrls($"http://localhost:{port}") // <-----
            .ConfigureServices(services =>
            {
                services.AddControllers();
                services.AddEndpointsApiExplorer();
            })
            .Configure((hostContext, app) =>
            {
                //  app.UseHttpsRedirection();
                app.UseRouting();
                app.UseEndpoints(endpoints =>
                {
                    endpoints.MapControllers();
                });

            });
    })
    .Build()
    .RunAsync();

static IHostBuilder CreateHostBuilder(string[] args) =>

           Host.CreateDefaultBuilder(args)
               .UseWindowsService()
               .ConfigureServices((hostContext, services) =>
               {
                   services.AddHostedService<ConverterService>();
               });


static void LoadEnvironmentVariablesFromFile()
{
    string filePath = "./.env";
    if (File.Exists(filePath))
    {
        foreach (string line in File.ReadAllLines(filePath))
        {
            string[] parts = line.Split('=', 2);
            if (parts.Length == 2)
            {
                string key = parts[0].Trim();
                string value = parts[1].Trim();
                Environment.SetEnvironmentVariable(key, value);
            }
        }
    }
    else
    {
        Console.WriteLine($"Environment file not found: {filePath}");
    }
}


static void SetRegistryValue(string version)
{
    string subKey = $@"Software\Microsoft\Office\{version}\Common\FixedFormat";
    using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(subKey, true))
    {
        if (registryKey != null)
        {
            registryKey.SetValue("LastISO19005-1", 1, RegistryValueKind.DWord);
        }
        else
        {
            // Handle registry key not found error
            Console.WriteLine("Registry key not found.");
        }
    }
}

static void OfficeExportsPdfA()
{
    LoadEnvironmentVariablesFromFile();
    string excel = Environment.GetEnvironmentVariable("EXCEL_VERSION");
    SetRegistryValue(excel);
    string word = Environment.GetEnvironmentVariable("WORD_VERSION");
    SetRegistryValue(word);
    string powerPoint = Environment.GetEnvironmentVariable("POWER_POINT_VERSION");
    SetRegistryValue(powerPoint);
}