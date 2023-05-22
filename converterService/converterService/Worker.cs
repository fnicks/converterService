namespace converterService.Services;

public class ConverterService : BackgroundService
{
    public ConverterService()
    {
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            await Task.Delay(TimeSpan.FromSeconds(5), stoppingToken);
        }
    }

    public override async Task StartAsync(CancellationToken cancellationToken)
    {
        Thread thread = new Thread(new ThreadStart(StartFunction));
        thread.Start();
        Console.WriteLine("start");
        await base.StartAsync(cancellationToken);
    }
    public override async Task StopAsync(CancellationToken cancellationToken)
    {
        Environment.Exit(0);
        await base.StopAsync(cancellationToken);
    }

    private void StartFunction()
    {
        Console.WriteLine("start");
    }
}