using System.Runtime.InteropServices;

internal sealed class MapiException(string operation, int result, Exception? inner = null)
    : Exception($"{operation}: MAPI HRESULT 0x{result:X8}.", inner)
{
    public int Result { get; } = result;
}

internal sealed unsafe class MapiObject : IDisposable
{
    public nint Pointer { get; private set; }

    public MapiObject(nint pointer) => Pointer = pointer != 0 ? pointer
        : throw new InvalidOperationException("MAPI did not return the requested interface.");

    public nint Method(int slot) => (*(nint**)Pointer)[slot];

    public void Dispose()
    {
        if (Pointer == 0) return;
        ((delegate* unmanaged[Stdcall]<nint, uint>)Method(2))(Pointer); // IUnknown::Release
        Pointer = 0;
    }
}

internal sealed unsafe class MapiRuntime : IDisposable
{
    public MapiRuntime() => Mapi.Check(Native.MAPIInitialize(0),
        "MAPIInitialize (check classic Outlook, the default MAPI client, and x86/x64 architecture)");

    public MapiObject Logon(string? profile, nint owner = 0)
    {
        const uint newSession = 0x2, extended = 0x20, noMail = 0x8000;
        const uint logonUi = 0x1, explicitProfile = 0x10;
        var flags = newSession | extended | noMail | Mapi.Unicode;
        if (profile is null) flags |= logonUi | explicitProfile;
        nint session = 0;
        fixed (char* name = profile)
        {
            var result = Native.MAPILogonEx((nuint)owner, name, null, flags, &session);
            if (result == unchecked((int)0x80040113)) throw new OperationCanceledException(); // MAPI_E_USER_CANCEL
            Mapi.Check(result, "MAPILogonEx (check the profile and access in classic Outlook)");
        }

        return new MapiObject(session);
    }

    public void Dispose() => Native.MAPIUninitialize();
}

internal static unsafe class Native
{
    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    [DllImport("mapi32.dll", ExactSpelling = true)]
    internal static extern int MAPIInitialize(nint init);

    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    [DllImport("mapi32.dll", ExactSpelling = true)]
    internal static extern void MAPIUninitialize();

    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    [DllImport("mapi32.dll", ExactSpelling = true)]
    internal static extern int MAPIAdminProfiles(uint flags, nint* admin);

    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    [DllImport("mapi32.dll", ExactSpelling = true)]
    internal static extern int MAPILogonEx(nuint ui, char* profile, char* password, uint flags, nint* session);

    [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
    [DllImport("mapi32.dll", ExactSpelling = true)]
    internal static extern uint MAPIFreeBuffer(void* buffer);
}

internal static unsafe class Mapi
{
    internal const uint Unicode = 0x80000000;
    internal const int MutationBatchSize = 256;
    internal const string NoStoresMessage = "No mailbox stores found. Verify that classic Outlook opens a mailbox with this profile, then try again.";
    private const int QueryBatchSize = 512;
    private const uint EntryId = 0x0FFF0102, DisplayName = 0x3001001F, DeliveryTime = 0x0E060040,
        MessageClass = 0x001A001F, IpmSubtree = 0x35E00102, IpmWastebasket = 0x35E30102, FolderType = 0x36010003,
        MessageSize = 0x0E080003, Subject = 0x0037001F;
    private const uint MapiModify = 1, MapiFolder = 3, FolderGeneric = 1;
    private const uint DisplayNameAnsi = 0x3001001E, DefaultProfile = 0x3D04000B;

    // Vtable slots follow Microsoft's MAPIX.h and MAPIDefS.h, including inherited methods.
    private const int MsgStoreFirstSlot = 3 + 11; // IUnknown + IMAPIProp; IMsgStore starts with Advise and Unadvise.
    private const int StoreCompareEntryIds = MsgStoreFirstSlot + 2, StoreOpenEntry = MsgStoreFirstSlot + 3;
    private const int ContainerOpenEntry = 16;
    private const int InvalidEntryId = unchecked((int)0x80040107);
    private const int GetProfileTable = 4, GetMsgStoresTable = 4, OpenMsgStore = 5, GetProps = 5,
        GetContentsTable = 14, GetHierarchyTable = 15, CopyMessages = 20, DeleteMessages = 21,
        SetColumns = 7, GetRowCount = 9, Restrict = 14, QueryRows = 19;

    internal record NamedEntry(string Name, byte[] Id);
    internal sealed record Profile(string Name, bool IsDefault);
    internal enum StoreAccess { ReadOnly, Write, BestAvailable }

    internal static void Check(int result, string operation)
    {
        if (result != 0) throw new MapiException(operation, result);
    }

    internal static List<Profile> GetProfiles(CancellationToken cancellation)
    {
        cancellation.ThrowIfCancellationRequested();
        nint adminPointer = 0;
        Check(Native.MAPIAdminProfiles(0, &adminPointer), "MAPIAdminProfiles");
        using var admin = new MapiObject(adminPointer);
        nint tablePointer = 0;
        var call = (delegate* unmanaged[Stdcall]<nint, uint, nint*, int>)admin.Method(GetProfileTable);
        // The profile table exposes ANSI display names, even in a Unicode Outlook build.
        Check(call(admin.Pointer, 0, &tablePointer), "IProfAdmin::GetProfileTable");
        using var table = new MapiObject(tablePointer);
        SelectColumns(table, [DisplayNameAnsi, DefaultProfile]);
        var profiles = new List<Profile>();
        ReadRows(table, row =>
        {
            var props = GetRowProperties(row, 2);
            var name = ReadAnsiDisplayName(props[0]);
            RequireTag(props[1], DefaultProfile, "PR_DEFAULT_PROFILE");
            profiles.Add(new Profile(name, props[1].Value.Boolean != 0));
        }, cancellation);
        return profiles;
    }

    internal static List<NamedEntry> GetStores(MapiObject session, CancellationToken cancellation)
    {
        nint table = 0;
        var call = (delegate* unmanaged[Stdcall]<nint, uint, nint*, int>)session.Method(GetMsgStoresTable);
        // Some Outlook providers reject MAPI_UNICODE on the table itself but still expose PR_DISPLAY_NAME_W.
        Check(call(session.Pointer, 0, &table), "IMAPISession::GetMsgStoresTable");
        using var handle = new MapiObject(table);
        return ReadNames(handle, cancellation, allowAnsi: true);
    }

    internal static MapiObject OpenStore(MapiObject session, byte[] entryId, StoreAccess access, bool online)
    {
        const uint noDialog = 0x1, mdbWrite = 0x4, bestAccess = 0x10, noMail = 0x80, mdbOnline = 0x100;
        nint store = 0;
        var accessFlags = access switch
        {
            StoreAccess.ReadOnly => 0u,
            StoreAccess.Write => mdbWrite,
            StoreAccess.BestAvailable => bestAccess,
            _ => throw new ArgumentOutOfRangeException(nameof(access))
        };
        var flags = noDialog | noMail | accessFlags | (online ? mdbOnline : 0);
        var call = (delegate* unmanaged[Stdcall]<nint, nuint, uint, byte*, void*, uint, nint*, int>)session.Method(OpenMsgStore);
        fixed (byte* id = entryId)
            Check(call(session.Pointer, 0, (uint)entryId.Length, id, null, flags, &store), "IMAPISession::OpenMsgStore");
        return new MapiObject(store);
    }

    internal static MapiObject ResolveFolder(MapiObject store, string? path, bool write, CancellationToken cancellation)
    {
        var segments = path?.Split('\\') ?? [];
        var folder = OpenMailRoot(store, write && segments.Length == 0);
        try
        {
            for (var i = 0; i < segments.Length; i++)
            {
                var child = FindUnique(GetChildren(folder, cancellation), segments[i], "Folder", "Run folders on the parent folder.");
                var next = OpenChildFolder(store, folder, child.Id, write && i == segments.Length - 1);
                folder.Dispose();
                folder = next;
            }

            if (segments.Length != 0) EnsureNormalFolder(folder);
            return folder;
        }
        catch
        {
            folder.Dispose();
            throw;
        }
    }

    internal static MapiObject OpenMailRoot(MapiObject store, bool write = false) =>
        OpenFolder(store, ReadBinaryProperty(store, IpmSubtree, "PR_IPM_SUBTREE_ENTRYID"), write);

    internal static byte[] GetFolderEntryId(MapiObject folder) => ReadBinaryProperty(folder, EntryId, "PR_ENTRYID (source folder)");

    internal static void EnsureNormalFolder(MapiObject folder)
    {
        if (!IsNormalFolder(folder))
            throw new InvalidOperationException("The path does not identify a normal folder. Search folders are excluded.");
    }

    internal static bool IsNormalFolder(MapiObject folder) =>
        ReadProperty(folder, FolderType, "PR_FOLDER_TYPE", p => p->Value.Number) == FolderGeneric;

    internal static List<NamedEntry> GetChildren(MapiObject folder, CancellationToken cancellation)
    {
        using var table = OpenFolderTable(folder, GetHierarchyTable, "IMAPIFolder::GetHierarchyTable");
        return ReadNames(table, cancellation);
    }

    internal static NamedEntry FindUnique(List<NamedEntry> entries, string name, string kind, string hint)
    {
        NamedEntry? match = null;
        foreach (var entry in entries)
        {
            if (!entry.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) continue;
            if (match is not null) throw new InvalidOperationException($"Ambiguous {kind}: {name}. No automatic selection was made.");
            match = entry;
        }

        return match ?? throw new InvalidOperationException($"{kind} not found: {name}. {hint}");
    }

    internal static string FormatReportRow(long received, byte[] id) => $"{DateTime.FromFileTimeUtc(received):O}\t{Convert.ToHexString(id)}";

    internal static long Scan(MapiObject folder, long? from, long before, Action<long, byte[]>? record,
        CancellationToken cancellation, Action<long>? reportProgress) =>
        ScanContents(folder, from, before, false, false, record is null ? null : (received, id, _, _) => record(received, id),
            cancellation, reportProgress, null);

    internal static long ScanPreview(MapiObject folder, long? from, long? before, bool includeSize, Action<long, byte[], uint?, string> record,
        CancellationToken cancellation, Action<long> reportProgress, Action<long, uint?>? reportRows) =>
        ScanContents(folder, from, before, includeSize, true, record, cancellation, reportProgress, reportRows);

    private static long ScanContents(MapiObject folder, long? from, long? before, bool includeSize, bool includeSubject,
        Action<long, byte[], uint?, string>? record,
        CancellationToken cancellation, Action<long>? reportProgress, Action<long, uint?>? reportRows)
    {
        cancellation.ThrowIfCancellationRequested();
        using var table = OpenFolderTable(folder, GetContentsTable, "IMAPIFolder::GetContentsTable");
        ReadOnlySpan<uint> columns = includeSubject
            ? includeSize ? [DeliveryTime, MessageClass, EntryId, Subject, MessageSize] : [DeliveryTime, MessageClass, EntryId, Subject]
            : record is not null ? [DeliveryTime, MessageClass, EntryId] : [DeliveryTime, MessageClass];
        var columnCount = (uint)columns.Length;
        SelectColumns(table, columns);
        if (before is { } end) RestrictDates(table, from, end);

        uint? totalRows = reportRows is null ? null : TryGetRowCount(table);
        long count = 0, examined = 0;
        reportRows?.Invoke(examined, totalRows);
        ReadRows(table, row =>
        {
            examined++;
            var properties = GetRowProperties(row, columnCount);
            var received = 0L;
            if (!IsMissing(properties[0], DeliveryTime))
            {
                RequireTag(properties[0], DeliveryTime, "PR_MESSAGE_DELIVERY_TIME");
                received = properties[0].Value.FileTime;
            }
            if (before is { } end && (received <= 0 || received >= end || received < from)) return;
            var type = ReadText(properties[1], MessageClass, "PR_MESSAGE_CLASS_W");
            if (!type.Equals("IPM.Note", StringComparison.OrdinalIgnoreCase) &&
                !type.StartsWith("IPM.Note.", StringComparison.OrdinalIgnoreCase)) return;

            uint? size = null;
            if (includeSize && !IsMissing(properties[4], MessageSize))
            {
                RequireTag(properties[4], MessageSize, "PR_MESSAGE_SIZE");
                size = unchecked((uint)properties[4].Value.Number);
            }
            var subject = includeSubject && !IsMissing(properties[3], Subject)
                ? ReadText(properties[3], Subject, "PR_SUBJECT_W").ToString() : "";
            record?.Invoke(received, ReadBinary(properties[2], EntryId, "PR_ENTRYID"), size, subject);
            count++;
        }, cancellation, reportProgress is null && reportRows is null ? null : () =>
        {
            reportProgress?.Invoke(count);
            reportRows?.Invoke(examined, totalRows);
        });
        return count;
    }

    private static void RestrictDates(MapiObject table, long? from, long before)
    {
        var upper = new Prop { Tag = DeliveryTime, Value = new PropValue { FileTime = before } };
        var lower = new Prop { Tag = DeliveryTime, Value = new PropValue { FileTime = from.GetValueOrDefault() } };
        const uint resAnd = 0, resProperty = 4, relopLessThan = 0, relopGreaterOrEqual = 3;
        var bounds = stackalloc Restriction[2];
        bounds[0] = new Restriction { Type = resProperty, Data = new() { Comparison = new() { Operator = relopLessThan, Tag = DeliveryTime, Property = &upper } } };
        bounds[1] = new Restriction { Type = resProperty, Data = new() { Comparison = new() { Operator = relopGreaterOrEqual, Tag = DeliveryTime, Property = &lower } } };
        var conjunction = new Restriction { Type = resAnd, Data = new() { And = new() { Count = 2, Children = bounds } } };
        var restrict = (delegate* unmanaged[Stdcall]<nint, Restriction*, uint, int>)table.Method(Restrict);
        Check(restrict(table.Pointer, from.HasValue ? &conjunction : bounds, 0), "IMAPITable::Restrict (PR_MESSAGE_DELIVERY_TIME)");
    }

    private static uint? TryGetRowCount(MapiObject table)
    {
        uint count = 0;
        var getCount = (delegate* unmanaged[Stdcall]<nint, uint, uint*, int>)table.Method(GetRowCount);
        return getCount(table.Pointer, 0, &count) == 0 ? count : null;
    }

    private static bool IsMissing(Prop property, uint tag) =>
        property.Tag == ((tag & 0xFFFF0000) | 0xA) && property.Value.Number == unchecked((int)0x8004010F);

    internal static MapiObject OpenTrash(MapiObject store, MapiObject source)
    {
        var trashId = GetTrashEntryId(store);
        if (SameEntry(store, GetFolderEntryId(source), trashId))
            throw new InvalidOperationException("The source is already Deleted Items. No messages will be moved.");
        return OpenFolder(store, trashId, true);
    }

    internal static byte[] GetTrashEntryId(MapiObject store) => ReadBinaryProperty(store, IpmWastebasket, "PR_IPM_WASTEBASKET_ENTRYID");

    internal static bool SameEntry(MapiObject store, byte[] first, byte[] second)
    {
        uint same = 0;
        var compare = (delegate* unmanaged[Stdcall]<nint, uint, byte*, uint, byte*, uint, uint*, int>)store.Method(StoreCompareEntryIds);
        fixed (byte* s = first, t = second)
            Check(compare(store.Pointer, (uint)first.Length, s, (uint)second.Length, t, 0, &same), "IMsgStore::CompareEntryIDs");
        return same != 0;
    }

    internal static void Move(MapiObject source, MapiObject destination, ReadOnlySpan<byte[]> ids, Action? onSubmit = null) =>
        ApplyToBatch(ids, list =>
        {
            const uint messageMove = 1;
            var move = (delegate* unmanaged[Stdcall]<nint, EntryList*, void*, nint, nuint, nint, uint, int>)source.Method(CopyMessages);
            onSubmit?.Invoke();
            Check(move(source.Pointer, list, null, destination.Pointer, 0, 0, messageMove), "IMAPIFolder::CopyMessages (MESSAGE_MOVE)");
        });

    internal static void HardDelete(MapiObject source, ReadOnlySpan<byte[]> ids, Action? onSubmit = null) =>
        ApplyToBatch(ids, list =>
        {
            const uint deleteHardDelete = 0x10;
            var delete = (delegate* unmanaged[Stdcall]<nint, EntryList*, nuint, nint, uint, int>)source.Method(DeleteMessages);
            onSubmit?.Invoke();
            Check(delete(source.Pointer, list, 0, 0, deleteHardDelete), "IMAPIFolder::DeleteMessages (DELETE_HARD_DELETE)");
        });

    private delegate void BatchAction(EntryList* entries);

    private static void ApplyToBatch(ReadOnlySpan<byte[]> ids, BatchAction apply)
    {
        var totalBytes = 0;
        foreach (var id in ids) totalBytes = checked(totalBytes + id.Length);
        var packedIds = new byte[totalBytes];
        var entries = stackalloc Binary[ids.Length];
        fixed (byte* storage = packedIds)
        {
            var offset = 0;
            for (var i = 0; i < ids.Length; i++)
            {
                var id = ids[i];
                id.AsSpan().CopyTo(packedIds.AsSpan(offset));
                entries[i] = new Binary { Count = (uint)id.Length, Data = storage + offset };
                offset += id.Length;
            }

            var list = new EntryList { Count = (uint)ids.Length, Entries = entries };
            apply(&list);
        }
    }

    internal static MapiObject OpenFolder(MapiObject store, byte[] entryId, bool write) =>
        OpenFolderEntry(store, StoreOpenEntry, "IMsgStore::OpenEntry", entryId, write);

    internal static MapiObject OpenChildFolder(MapiObject store, MapiObject parent, byte[] entryId, bool write)
    {
        try
        {
            return OpenFolder(store, entryId, write);
        }
        catch (MapiException e) when (e.Result == InvalidEntryId)
        {
            return OpenFolderEntry(parent, ContainerOpenEntry, "IMAPIContainer::OpenEntry (parent folder)", entryId, write, e);
        }
    }

    private static MapiObject OpenFolderEntry(MapiObject container, int slot, string operation, byte[] entryId, bool write,
        MapiException? previous = null)
    {
        uint type = 0;
        nint folder = 0;
        var call = (delegate* unmanaged[Stdcall]<nint, uint, byte*, void*, uint, uint*, nint*, int>)container.Method(slot);
        int resultCode;
        fixed (byte* id = entryId)
            resultCode = call(container.Pointer, (uint)entryId.Length, id, null, write ? MapiModify : 0, &type, &folder);
        if (resultCode != 0)
        {
            if (folder != 0) new MapiObject(folder).Dispose();
            if (previous is not null) operation += $"; store HRESULT 0x{previous.Result:X8}";
            throw new MapiException(operation, resultCode, previous);
        }
        var result = new MapiObject(folder);
        if (type == MapiFolder) return result;
        result.Dispose();
        throw new InvalidOperationException("The EntryID does not identify a MAPI folder.");
    }

    private static MapiObject OpenFolderTable(MapiObject folder, int method, string operation)
    {
        nint table = 0;
        var call = (delegate* unmanaged[Stdcall]<nint, uint, nint*, int>)folder.Method(method);
        Check(call(folder.Pointer, Unicode, &table), operation);
        return new MapiObject(table);
    }

    private static List<NamedEntry> ReadNames(MapiObject table, CancellationToken cancellation, bool allowAnsi = false)
    {
        SelectColumns(table, allowAnsi ? [DisplayName, EntryId, DisplayNameAnsi] : [DisplayName, EntryId]);
        var names = new List<NamedEntry>();
        ReadRows(table, row =>
        {
            var properties = GetRowProperties(row, allowAnsi ? 3u : 2u);
            const uint displayNameError = 0x3001000A;
            const int notFound = unchecked((int)0x8004010F), badCharWidth = unchecked((int)0x80040103);
            var name = allowAnsi && properties[0].Tag == displayNameError && properties[0].Value.Number is notFound or badCharWidth
                ? ReadAnsiDisplayName(properties[2])
                : ReadText(properties[0], DisplayName, "PR_DISPLAY_NAME_W").ToString();
            names.Add(new NamedEntry(name, ReadBinary(properties[1], EntryId, "PR_ENTRYID")));
        }, cancellation);
        return names;
    }

    private static void SelectColumns(MapiObject table, ReadOnlySpan<uint> columns)
    {
        var tags = stackalloc uint[columns.Length + 1];
        tags[0] = (uint)columns.Length;
        columns.CopyTo(new Span<uint>(tags + 1, columns.Length));
        var call = (delegate* unmanaged[Stdcall]<nint, uint*, uint, int>)table.Method(SetColumns);
        Check(call(table.Pointer, tags, 0), "IMAPITable::SetColumns");
    }

    private delegate void ReadRow(Row* row);

    private static void ReadRows(MapiObject table, ReadRow read, CancellationToken cancellation, Action? onBatch = null)
    {
        var query = (delegate* unmanaged[Stdcall]<nint, int, uint, RowSet**, int>)table.Method(QueryRows);
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            RowSet* batch = null;
            var result = query(table.Pointer, QueryBatchSize, 0, &batch);
            try
            {
                Check(result, "IMAPITable::QueryRows");
                if (batch == null) throw new InvalidOperationException("QueryRows did not return an SRowSet.");
                if (batch->Count == 0) return;
                var rows = &batch->First;
                for (var i = 0; i < batch->Count; i++) read(rows + i);
                onBatch?.Invoke();
            }
            finally
            {
                if (batch != null)
                {
                    // QueryRows allocates each row's property buffer separately from the row set.
                    var rows = &batch->First;
                    for (var i = 0; i < batch->Count; i++)
                        if (rows[i].Properties != null) Native.MAPIFreeBuffer(rows[i].Properties);
                    Native.MAPIFreeBuffer(batch);
                }
            }
        }
    }

    private static Prop* GetRowProperties(Row* row, uint expected) => row->Count == expected && row->Properties != null
        ? row->Properties : throw new InvalidOperationException("MAPI returned a row with unexpected columns.");

    private delegate T ConvertProp<T>(Prop* property);

    private static T ReadProperty<T>(MapiObject obj, uint tag, string label, ConvertProp<T> convert)
    {
        var request = stackalloc uint[] { 1, tag };
        Prop* props = null;
        uint count = 0;
        var call = (delegate* unmanaged[Stdcall]<nint, uint*, uint, uint*, Prop**, int>)obj.Method(GetProps);
        var result = call(obj.Pointer, request, Unicode, &count, &props);
        try
        {
            if (result < 0) Check(result, label);
            if (count != 1 || props == null) throw new InvalidOperationException($"{label}: property not returned.");
            RequireTag(*props, tag, label);
            Check(result, label);
            return convert(props);
        }
        finally
        {
            if (props != null) Native.MAPIFreeBuffer(props);
        }
    }

    private static byte[] ReadBinaryProperty(MapiObject obj, uint tag, string label) =>
        ReadProperty(obj, tag, label, p => ReadBinary(*p, tag, label));

    private static byte[] ReadBinary(Prop property, uint tag, string label)
    {
        RequireTag(property, tag, label);
        var binary = property.Value.Binary;
        if (binary.Count == 0 || binary.Data == null) throw new InvalidOperationException($"{label}: empty value.");
        return new ReadOnlySpan<byte>(binary.Data, checked((int)binary.Count)).ToArray();
    }

    private static ReadOnlySpan<char> ReadText(Prop property, uint tag, string label)
    {
        RequireTag(property, tag, label);
        if (property.Value.Text == null) throw new InvalidOperationException($"{label}: null value.");
        return MemoryMarshal.CreateReadOnlySpanFromNullTerminated(property.Value.Text);
    }

    private static string ReadAnsiDisplayName(Prop property)
    {
        RequireTag(property, DisplayNameAnsi, "PR_DISPLAY_NAME_A");
        var name = Marshal.PtrToStringAnsi(property.Value.Pointer);
        return !string.IsNullOrEmpty(name) ? name : throw new InvalidOperationException("PR_DISPLAY_NAME_A: empty value.");
    }

    private static void RequireTag(Prop property, uint tag, string label)
    {
        if (property.Tag == tag) return;
        if ((property.Tag & 0xFFFF) == 0xA) throw new MapiException(label, property.Value.Number); // PT_ERROR
        throw new InvalidOperationException($"{label}: unexpected tag 0x{property.Tag:X8}.");
    }

    [StructLayout(LayoutKind.Sequential)] private struct Binary { public uint Count; public byte* Data; }
    [StructLayout(LayoutKind.Sequential)] private struct EntryList { public uint Count; public Binary* Entries; }
    [StructLayout(LayoutKind.Sequential)] private struct Prop { public uint Tag, Padding; public PropValue Value; }
    [StructLayout(LayoutKind.Explicit)] private struct PropValue
    {
        [FieldOffset(0)] public Binary Binary;
        [FieldOffset(0)] public long FileTime;
        [FieldOffset(0)] public int Number;
        [FieldOffset(0)] public ushort Boolean;
        [FieldOffset(0)] public nint Pointer;
        [FieldOffset(0)] public char* Text;
    }

    [StructLayout(LayoutKind.Sequential)] private struct Row { public uint Padding, Count; public Prop* Properties; }
    [StructLayout(LayoutKind.Sequential)] private struct RowSet { public uint Count; public Row First; }
    [StructLayout(LayoutKind.Sequential)] private struct Restriction
    {
        public uint Type;
        public RestrictionUnion Data;
    }

    [StructLayout(LayoutKind.Explicit)] private struct RestrictionUnion
    {
        [FieldOffset(0)] public Comparison Comparison;
        [FieldOffset(0)] public AndRestriction And;
        [FieldOffset(0)] private Comment _comment; // Largest native union member determines stride/alignment.
    }

    [StructLayout(LayoutKind.Sequential)] private struct Comparison { public uint Operator, Tag; public Prop* Property; }
    [StructLayout(LayoutKind.Sequential)] private struct AndRestriction { public uint Count; public Restriction* Children; }
    [StructLayout(LayoutKind.Sequential)] private struct Comment { public uint Count; public Restriction* Child; public Prop* Properties; }
}