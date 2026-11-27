internal static unsafe partial class Mapi
{
    private const uint PstPath = 0x6700001F, PstConfigFlags = 0x67700003, ServiceUid = 0x3D0C0102;
    private static readonly Guid MessageInterface = new("00020307-0000-0000-C000-000000000046");

    internal static void WithNewPst(MapiRuntime runtime, string path, Action<MapiObject> write, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        nint pointer = 0;
        Check(Native.MAPIAdminProfiles(0, &pointer), "MAPIAdminProfiles (PST export)");
        using var admin = new MapiObject(pointer);
        var profile = $"MAPICommander-PST-{Guid.NewGuid():N}";
        var name = System.Text.Encoding.ASCII.GetBytes(profile + '\0');
        fixed (byte* profileName = name)
        {
            var create = (delegate* unmanaged[Stdcall]<nint, byte*, byte*, nuint, uint, int>)admin.Method(5);
            Check(create(admin.Pointer, profileName, null, 0, 0), "IProfAdmin::CreateProfile (temporary PST)");
            Exception? failure = null;
            try
            {
                ConfigurePst(admin, profileName, path, token);
                using var session = runtime.Logon(profile);
                var stores = GetStores(session, token);
                if (stores.Count != 1)
                    throw new InvalidOperationException("The temporary PST profile did not expose exactly one store.");
                using var store = OpenStore(session, stores[0].Id, StoreAccess.Write, online: false);
                using var root = OpenMailRoot(store, write: true);
                using var folder = CreateExportFolder(root);
                write(folder);
            }
            catch (Exception e) when (e is MapiException or InvalidOperationException or OperationCanceledException or IOException or UnauthorizedAccessException)
            {
                failure = e;
                throw;
            }
            finally
            {
                var delete = (delegate* unmanaged[Stdcall]<nint, byte*, uint, int>)admin.Method(6);
                var result = delete(admin.Pointer, profileName, 0);
                if (result != 0)
                    throw new MapiException($"IProfAdmin::DeleteProfile ({profile})", result, failure);
            }
        }
    }

    private static void ConfigurePst(MapiObject profileAdmin, byte* profileName, string path, CancellationToken token)
    {
        nint pointer = 0;
        var adminServices = (delegate* unmanaged[Stdcall]<nint, byte*, byte*, nuint, uint, nint*, int>)profileAdmin.Method(11);
        Check(adminServices(profileAdmin.Pointer, profileName, null, 0, 0, &pointer), "IProfAdmin::AdminServices (temporary PST)");
        using var services = new MapiObject(pointer);
        var create = (delegate* unmanaged[Stdcall]<nint, byte*, byte*, nuint, uint, int>)services.Method(5);
        fixed (byte* service = "MSUPST MS\0"u8, display = "MAPICommander Export\0"u8)
            Check(create(services.Pointer, service, display, 0, 0), "IMsgServiceAdmin::CreateMsgService (Unicode PST)");

        pointer = 0;
        var getTable = (delegate* unmanaged[Stdcall]<nint, uint, nint*, int>)services.Method(4);
        Check(getTable(services.Pointer, 0, &pointer), "IMsgServiceAdmin::GetMsgServiceTable (temporary PST)");
        byte[]? uid = null;
        using (var table = new MapiObject(pointer))
        {
            SelectColumns(table, [ServiceUid]);
            ReadRows(table, row =>
            {
                if (uid is not null) throw new InvalidOperationException("The temporary PST profile contains more than one service.");
                uid = ReadBinary(*GetRowProperties(row, 1), ServiceUid, "PR_SERVICE_UID");
            }, token);
        }
        if (uid is not { Length: 16 }) throw new InvalidOperationException("The Unicode PST service did not return a valid identifier.");

        var configure = (delegate* unmanaged[Stdcall]<nint, byte*, nuint, uint, uint, Prop*, int>)services.Method(9);
        fixed (byte* id = uid)
        fixed (char* file = path)
        {
            var props = stackalloc Prop[]
            {
                new() { Tag = PstPath, Value = new() { Text = file } },
                new() { Tag = PstConfigFlags, Value = new() { Number = unchecked((int)0x80000000) } }
            };
            Check(configure(services.Pointer, id, 0, 0, 2, props), "IMsgServiceAdmin::ConfigureMsgService (Unicode PST)");
        }
    }

    private static MapiObject CreateExportFolder(MapiObject root)
    {
        nint pointer = 0;
        var create = (delegate* unmanaged[Stdcall]<nint, uint, char*, char*, Guid*, uint, nint*, int>)root.Method(22);
        fixed (char* name = "MAPICommander Export")
            Check(create(root.Pointer, FolderGeneric, name, null, null, Unicode, &pointer), "IMAPIFolder::CreateFolder (PST export)");
        return new MapiObject(pointer);
    }

    internal static void CopyMessageToPst(MapiObject source, MapiObject destination, byte[] entryId)
    {
        const uint messageObjectType = 5;
        var messageInterface = MessageInterface;
        uint objectType = 0;
        nint pointer = 0;
        var open = (delegate* unmanaged[Stdcall]<nint, uint, byte*, Guid*, uint, uint*, nint*, int>)source.Method(ContainerOpenEntry);
        fixed (byte* id = entryId)
            Check(open(source.Pointer, (uint)entryId.Length, id, &messageInterface, 0, &objectType, &pointer), "IMAPIFolder::OpenEntry (source message)");
        using var original = new MapiObject(pointer);
        if (objectType != messageObjectType) throw new InvalidOperationException("A preview row did not identify a MAPI message.");

        pointer = 0;
        var create = (delegate* unmanaged[Stdcall]<nint, Guid*, uint, nint*, int>)destination.Method(19);
        Check(create(destination.Pointer, null, 0, &pointer), "IMAPIFolder::CreateMessage (PST export)");
        using var message = new MapiObject(pointer);
        uint* problems = null;
        // Exclude provider-specific identity and access metadata, not content or message timestamps.
        var exclude = stackalloc uint[] { 8, EntryId, 0x0E090102, 0x0FF60102, 0x0FF90102, 0x0FFA0102, 0x0FFB0102, 0x0FF40003, 0x0FF70003 };
        var copy = (delegate* unmanaged[Stdcall]<nint, uint, Guid*, uint*, nuint, nint, Guid*, nint, uint, uint**, int>)original.Method(10);
        try
        {
            Check(copy(original.Pointer, 0, null, exclude, 0, 0, &messageInterface, message.Pointer, 0, &problems), "IMessage::CopyTo (PST export)");
            if (problems != null && problems[0] != 0)
                throw new MapiException($"IMessage::CopyTo (property 0x{problems[2]:X8})", unchecked((int)problems[3]));
        }
        finally
        {
            if (problems != null) Native.MAPIFreeBuffer(problems);
        }
        var save = (delegate* unmanaged[Stdcall]<nint, uint, int>)message.Method(4);
        Check(save(message.Pointer, 0), "IMessage::SaveChanges (PST export)");
    }
}