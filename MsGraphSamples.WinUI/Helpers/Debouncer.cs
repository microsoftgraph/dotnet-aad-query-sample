using Microsoft.UI.Xaml;

namespace MsGraphSamples.WinUI.Helpers;

public class Debouncer
{
    private readonly DispatcherTimer timer;
    private Action? action;

    public Debouncer(TimeSpan delay)
    {
        timer = new DispatcherTimer { Interval = delay };
        timer.Tick += Timer_Tick;
    }

    public void Debounce(Action action)
    {
        this.action = action;
        timer.Stop();
        timer.Start();
    }

    private void Timer_Tick(object? sender, object e)
    {
        timer.Stop();
        action?.Invoke();
    }
}