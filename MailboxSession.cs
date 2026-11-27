internal sealed class MailboxSession
{
    internal sealed record Target(byte[] StoreId, byte[]? FolderId, Target? Parent = null);
    internal sealed record Source(byte[] Id, string Path);
    internal sealed record Row(long Received, byte[] Id, uint? Size, string Subject, Source Source);
    internal sealed record Preview(Target Target, bool IncludesSize, IReadOnlyList<Row> Rows);
    internal enum SearchScope { Folder, Subtree, Mailbox }

    private readonly TaskCompletionSource<Control> _dispatcher = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly TaskCompletionSource _stopped = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private MapiRuntime? _runtime;
    private MapiObject? _session;
    private readonly Dictionary<byte[], MapiObject> _browseStores = new(ReferenceEqualityComparer.Instance);
    private Preview? _currentPreview;
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
        using var folder = OpenTargetFolder(store, target, token);
        return Mapi.GetChildren(folder, token);
    });

    public Task<Preview> PreviewAsync(Target target, string path, SearchScope scope, long? from, long? before, bool largest,
        Action<long> progress, Action<int?> percentage, CancellationToken token) => InvokeAsync(() =>
    {
        _currentPreview = null;
        if (scope == SearchScope.Mailbox ? target.FolderId is not null : target.FolderId is null)
            throw new InvalidOperationException("The selected target does not match the search scope.");
        var store = GetBrowseStore(target);
        using var root = OpenTargetFolder(store, target, token);
        var rootIsNormal = Mapi.IsNormalFolder(root);
        if (!rootIsNormal && scope != SearchScope.Mailbox) Mapi.EnsureNormalFolder(root);
        var rootSource = new Source(Mapi.GetFolderEntryId(root), path);
        var rows = new List<Row>();

        void ScanFolder(MapiObject folder, Source source, Action<long, uint?>? rowProgress)
        {
            var previous = rows.Count;
            Mapi.ScanPreview(folder, from, before, largest, (received, id, size, subject) => rows.Add(new Row(received, id, size, subject, source)),
                token, count => progress(previous + count), rowProgress);
        }

        if (scope != SearchScope.Folder)
        {
            var sources = DiscoverFolders(store, root, rootSource, rootIsNormal, token);
            for (var i = 0; i < sources.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                percentage((int)(i * 100L / sources.Count));
                var source = sources[i];
                if (ReferenceEquals(source, rootSource))
                {
                    ScanFolder(root, source, null);
                    continue;
                }
                using var folder = Mapi.OpenFolder(store, source.Id, write: false);
                Mapi.EnsureNormalFolder(folder);
                ScanFolder(folder, source, null);
            }
        }
        else
        {
            ScanFolder(root, rootSource, (examined, total) =>
                percentage(total is null ? null : total == 0 ? 0 : (int)Math.Min(99, examined * 100 / total.Value)));
        }

        token.ThrowIfCancellationRequested();
        if (largest) rows.Sort((left, right) => Nullable.Compare(right.Size, left.Size));
        token.ThrowIfCancellationRequested();
        return _currentPreview = new Preview(target, largest, rows);
    });

    private static List<Source> DiscoverFolders(MapiObject store, MapiObject root, Source rootSource, bool includeRoot, CancellationToken token)
    {
        var sources = new List<Source>();
        var seen = new HashSet<string>(StringComparer.Ordinal) { Convert.ToHexString(rootSource.Id) };
        if (includeRoot) sources.Add(rootSource);

        void VisitChildren(MapiObject parent, string parentPath)
        {
            foreach (var child in Mapi.GetChildren(parent, token))
            {
                token.ThrowIfCancellationRequested();
                using var folder = Mapi.OpenChildFolder(store, parent, child.Id, write: false);
                if (!Mapi.IsNormalFolder(folder)) continue;
                var id = Mapi.GetFolderEntryId(folder);
                if (!seen.Add(Convert.ToHexString(id))) continue;
                var childPath = $"{parentPath}\\{child.Name}";
                sources.Add(new Source(id, childPath));
                VisitChildren(folder, childPath);
            }
        }

        VisitChildren(root, rootSource.Path);
        return sources;
    }

    public Task<int> RemoveAsync(Preview preview, IReadOnlyList<int>? selectedIndices, bool hard, Action onFirstBatch,
        Action<long> progress, CancellationToken token) => InvokeAsync(() =>
    {
        token.ThrowIfCancellationRequested();
        if (!ReferenceEquals(preview, _currentPreview) || !_browseStores.TryGetValue(preview.Target.StoreId, out var store))
            throw new InvalidOperationException("The preview is no longer current. Run Preview again before removing messages.");
        _currentPreview = null;
        var groups = new Dictionary<Source, List<byte[]>>(ReferenceEqualityComparer.Instance);
        var count = selectedIndices?.Count ?? preview.Rows.Count;
        for (var i = 0; i < count; i++)
        {
            token.ThrowIfCancellationRequested();
            var row = preview.Rows[selectedIndices is null ? i : selectedIndices[i]];
            if (!groups.TryGetValue(row.Source, out var ids)) groups.Add(row.Source, ids = []);
            ids.Add(row.Id);
        }
        if (count == 0) return 0;

        var trashId = hard ? null : Mapi.GetTrashEntryId(store);
        if (trashId is not null)
        {
            foreach (var source in groups.Keys)
            {
                token.ThrowIfCancellationRequested();
                if (Mapi.SameEntry(store, source.Id, trashId))
                    throw new InvalidOperationException("The removal selection includes messages already in Deleted Items. Select only messages from other folders to move them.");
            }
        }

        using var destination = trashId is null ? null : Mapi.OpenFolder(store, trashId, write: true);
        var batch = new byte[Math.Min(count, Mapi.MutationBatchSize)][];
        var confirmed = 0;
        foreach (var (group, ids) in groups)
        {
            token.ThrowIfCancellationRequested();
            using var source = Mapi.OpenFolder(store, group.Id, write: true);
            Mapi.EnsureNormalFolder(source);
            for (var offset = 0; offset < ids.Count;)
            {
                var length = Math.Min(batch.Length, ids.Count - offset);
                ids.CopyTo(offset, batch, 0, length);
                token.ThrowIfCancellationRequested();
                var onSubmit = confirmed == 0 ? onFirstBatch : null;
                if (destination is null) Mapi.HardDelete(source, batch.AsSpan(0, length), onSubmit);
                else Mapi.Move(source, destination, batch.AsSpan(0, length), onSubmit);
                confirmed += length;
                offset += length;
                progress(confirmed);
            }
        }
        return confirmed;
    });

    private MapiObject GetBrowseStore(Target target)
    {
        if (!_browseStores.TryGetValue(target.StoreId, out var store))
        {
            var session = _session ?? throw new InvalidOperationException("Connect to an Outlook profile first.");
            store = Mapi.OpenStore(session, target.StoreId, Mapi.StoreAccess.BestAvailable, _online);
            _browseStores.Add(target.StoreId, store);
        }
        return store;
    }

    private static MapiObject OpenTargetFolder(MapiObject store, Target target, CancellationToken token)
    {
        var path = new Stack<byte[]>();
        for (var node = target; node.FolderId is { } id; node = node.Parent ?? throw new InvalidOperationException("Folder has no parent in the current mailbox tree."))
            path.Push(id);
        token.ThrowIfCancellationRequested();
        var folder = Mapi.OpenMailRoot(store);
        try
        {
            while (path.TryPop(out var id))
            {
                token.ThrowIfCancellationRequested();
                var next = Mapi.OpenChildFolder(store, folder, id, write: false);
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
        _currentPreview = null;
        foreach (var store in _browseStores.Values) store.Dispose();
        _browseStores.Clear();
        _session?.Dispose();
        _session = null;
        _runtime?.Dispose();
        _runtime = null;
    }
}