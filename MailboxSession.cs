internal sealed class MailboxSession
{
    internal sealed record Target(byte[] StoreId, byte[]? FolderId);
    internal sealed record Row(long Received, byte[] Id);

    private readonly TaskCompletionSource<Control> _dispatcher = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly TaskCompletionSource _stopped = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private MapiRuntime? _runtime;
    private MapiObject? _session;
    private bool _online;

    public MailboxSession()
    {
        // Keep MAPI interfaces and their message pump on the same STA for the whole session.
        var thread = new Thread(Run) { Name = "MAPI session", IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    private void Run()
    {
        try
        {
            using var dispatcher = new Control();
            _ = dispatcher.Handle;
            _dispatcher.SetResult(dispatcher);
            Application.Run();
        }
        finally
        {
            Disconnect();
            _stopped.SetResult();
        }
    }

    public async Task StopAsync()
    {
        var dispatcher = await _dispatcher.Task;
        await dispatcher.InvokeAsync(Application.ExitThread);
        await _stopped.Task;
    }

    private async Task<T> InvokeAsync<T>(Func<T> action)
    {
        var dispatcher = await _dispatcher.Task;
        return await dispatcher.InvokeAsync(action);
    }

    public Task<List<Mapi.NamedEntry>> ConnectAsync(string? profile, bool online, nint owner, CancellationToken token) => InvokeAsync(() =>
    {
        Disconnect();
        _runtime = new MapiRuntime();
        try
        {
            _session = _runtime.Logon(profile, owner);
            _online = online;
            return Mapi.GetStores(_session, token);
        }
        catch
        {
            Disconnect();
            throw;
        }
    });

    public Task<List<Mapi.NamedEntry>> ChildrenAsync(Target target, CancellationToken token) => InvokeAsync(() =>
    {
        using var store = OpenStore(target, write: false);
        using var folder = Mapi.OpenFolder(store, target.FolderId ?? Mapi.GetMailRootId(store), write: false);
        return Mapi.GetChildren(folder, token);
    });

    public Task<List<Row>> PreviewAsync(Target target, long? from, long before, Action<long> progress, CancellationToken token) => InvokeAsync(() =>
    {
        using var store = OpenStore(target, write: false);
        using var folder = OpenTargetFolder(store, target, write: false);
        var rows = new List<Row>();
        Mapi.Scan(folder, from, before, (received, id) => rows.Add(new Row(received, id)), token, progress);
        return rows;
    });

    public Task<int> RemoveAsync(Target target, IReadOnlyList<Row> rows, bool hard, Action<long> progress, CancellationToken token) => InvokeAsync(() =>
    {
        token.ThrowIfCancellationRequested();
        using var store = OpenStore(target, write: true);
        using var source = OpenTargetFolder(store, target, write: true);
        using var destination = hard ? null : Mapi.OpenTrash(store, source);
        var batch = new byte[Math.Min(rows.Count, Mapi.MutationBatchSize)][];
        for (var offset = 0; offset < rows.Count;)
        {
            token.ThrowIfCancellationRequested();
            var length = Math.Min(batch.Length, rows.Count - offset);
            for (var i = 0; i < length; i++) batch[i] = rows[offset + i].Id;
            if (destination is null) Mapi.HardDelete(source, batch.AsSpan(0, length));
            else Mapi.Move(source, destination, batch.AsSpan(0, length));
            offset += length;
            progress(offset);
        }

        return rows.Count;
    });

    private MapiObject OpenStore(Target target, bool write) => Mapi.OpenStore(
        _session ?? throw new InvalidOperationException("Connect to an Outlook profile first."), target.StoreId, write, _online);

    private static MapiObject OpenTargetFolder(MapiObject store, Target target, bool write) =>
        Mapi.OpenNormalFolder(store, target.FolderId ?? throw new InvalidOperationException("Select a mail folder first."), write);

    private void Disconnect()
    {
        _session?.Dispose();
        _session = null;
        _runtime?.Dispose();
        _runtime = null;
    }
}