internal sealed class MailboxSession
{
    internal sealed record Target(byte[] StoreId, byte[]? FolderId, Target? Parent = null);
    internal sealed record Row(long Received, byte[] Id);

    private readonly TaskCompletionSource<Control> _dispatcher = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly TaskCompletionSource _stopped = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private MapiRuntime? _runtime;
    private MapiObject? _session;
    private readonly Dictionary<byte[], MapiObject> _browseStores = new(ReferenceEqualityComparer.Instance);
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

    public Task<List<Mapi.Profile>> ProfilesAsync(CancellationToken token) => InvokeAsync(() =>
    {
        using var runtime = _runtime is null ? new MapiRuntime() : null;
        return Mapi.GetProfiles(token);
    });

    public Task<List<Mapi.NamedEntry>> ConnectAsync(string profile, bool online, nint owner, CancellationToken token) => InvokeAsync(() =>
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
        var store = GetBrowseStore(target);
        using var folder = OpenTargetFolder(store, target, write: false, token);
        return Mapi.GetChildren(folder, token);
    });

    public Task<List<Row>> PreviewAsync(Target target, long? from, long before, Action<long> progress, CancellationToken token) => InvokeAsync(() =>
    {
        var store = GetBrowseStore(target);
        using var folder = OpenTargetFolder(store, target, write: false, token);
        Mapi.EnsureNormalFolder(folder);
        var rows = new List<Row>();
        Mapi.Scan(folder, from, before, (received, id) => rows.Add(new Row(received, id)), token, progress);
        return rows;
    });

    public Task<int> RemoveAsync(Target target, IReadOnlyList<Row> rows, bool hard, Action<long> progress, CancellationToken token) => InvokeAsync(() =>
    {
        token.ThrowIfCancellationRequested();
        if (_browseStores.Remove(target.StoreId, out var browseStore)) browseStore.Dispose();
        using var store = OpenStore(target, write: true);
        using var source = OpenTargetFolder(store, target, write: true, token);
        Mapi.EnsureNormalFolder(source);
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

    private MapiObject GetBrowseStore(Target target)
    {
        if (!_browseStores.TryGetValue(target.StoreId, out var store))
        {
            store = OpenStore(target, write: false);
            _browseStores.Add(target.StoreId, store);
        }
        return store;
    }

    private static MapiObject OpenTargetFolder(MapiObject store, Target target, bool write, CancellationToken token)
    {
        var path = new Stack<byte[]>();
        for (var node = target; node.FolderId is { } id; node = node.Parent ?? throw new InvalidOperationException("Folder has no parent in the current mailbox tree."))
            path.Push(id);
        token.ThrowIfCancellationRequested();
        var folder = Mapi.OpenMailRoot(store, write && path.Count == 0);
        try
        {
            while (path.TryPop(out var id))
            {
                token.ThrowIfCancellationRequested();
                var next = Mapi.OpenChildFolder(store, folder, id, write && path.Count == 0);
                folder.Dispose();
                folder = next;
            }
            return folder;
        }
        catch
        {
            folder.Dispose();
            throw;
        }
    }

    private void Disconnect()
    {
        foreach (var store in _browseStores.Values) store.Dispose();
        _browseStores.Clear();
        _session?.Dispose();
        _session = null;
        _runtime?.Dispose();
        _runtime = null;
    }
}