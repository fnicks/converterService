using converterService.Services;
using Microsoft.AspNetCore.Hosting;

await CreateHostBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddHostedService<ConverterService>();
    })
    .ConfigureWebHostDefaults(webBuilder =>
    {
        LoadEnvironmentVariablesFromFile();
        string port = Environment.GetEnvironmentVariable("PORT");
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