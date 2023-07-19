using converterService.Services;
using Microsoft.AspNetCore.Server.Kestrel.Core;

await CreateHostBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddHostedService<ConverterService>();
    })
    .ConfigureWebHostDefaults(webBuilder =>
    {
        LoadEnvironmentVariablesFromFile();
        string port = Environment.GetEnvironmentVariable("PORT");

        webBuilder.UseUrls($"http://+:{port}") // <-----
            .ConfigureServices(services =>
            {
                services.AddControllers();
                services.AddEndpointsApiExplorer();
                services.Configure<IISServerOptions>(options =>
                {
                    options.MaxRequestBodySize = 209715200; // 200 MB in bytes
                });
                services.Configure<KestrelServerOptions>(options =>
                {
                    options.Limits.MaxRequestBodySize = 209715200; // 200 MB in bytes
                });
            })
            .Configure((hostContext, app) =>
            {
                app.UseRouting();
                app.UseEndpoints(endpoints =>
                {
                    endpoints.MapControllers();
                });

            });
    })
    .UseWindowsService()
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
