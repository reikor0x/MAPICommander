using System.Diagnostics;

internal sealed class ConsoleProgress : IDisposable
{
    private const int Width = 16;
    private readonly TextWriter? _output;
    private readonly string _label;
    private readonly bool _inline;
    private readonly long? _total;
    private readonly Stopwatch _refresh = Stopwatch.StartNew();
    private int _frame;
    private int _lastLength;

    public ConsoleProgress(TextWriter? output, string label, bool inline, long? total = null)
    {
        _output = output;
        _label = label;
        _inline = inline;
        _total = total;
        if (_output is not null) Render(0, complete: false);
    }

    public void Report(long processed)
    {
        if (_output is null || _refresh.ElapsedMilliseconds < (_inline ? 200 : 1000)) return;
        Render(processed, complete: false);
        _refresh.Restart();
    }

    public void Complete(long processed)
    {
        if (_output is not null) Render(processed, complete: true);
    }

    private void Render(long processed, bool complete)
    {
        string bar, detail;
        if (_total is long total)
        {
            var filled = complete ? Width : (int)(processed * Width / total);
            bar = new string('#', filled) + new string('-', Width - filled);
            detail = $"{processed}/{total} ({(complete ? 100 : processed * 100 / total)}%)";
        }
        else
        {
            var position = _frame++ % Width;
            bar = complete ? new string('#', Width) : new string('-', position) + '>' + new string('-', Width - position - 1);
            detail = $"{processed} corrispondenze";
        }

        var line = $"{_label} [{bar}] {detail}";
        if (_inline)
        {
            _output!.Write('\r' + line.PadRight(_lastLength));
            _lastLength = line.Length;
        }
        else
        {
            _output!.WriteLine(line);
        }
    }

    public void Dispose()
    {
        if (_output is not null && _inline) _output.WriteLine();
    }
}