using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

internal static class Program
{
    private const string Usage = """
        MAPICommander stores [--profile NAME]
        MAPICommander folders [--profile NAME] --store NAME [--folder "Folder\Subfolder"] [--online]
        MAPICommander scan [--profile NAME] --store NAME --folder "Folder\Subfolder"
            [--from DATE] (--before DATE | --to DATE) [--list] [--online]
            [--delete | --hard-delete] [--whatif | --yes]

        DATE: yyyy-MM-dd or d/M/yyyy (UTC), or yyyy-MM-ddTHH:mm:ssZ / yyyy-MM-ddTHH:mm:ss+01:00.
        --from is inclusive; --before is exclusive; --to includes the whole final day (date only).
        scan counts without making changes; --list reports the received time in UTC and EntryID as TSV.
        --whatif produces the TSV report without making changes, even with --delete or --hard-delete.
        --delete moves messages to Deleted Items; --hard-delete bypasses it. --yes skips confirmation.
        Omit --profile to open the classic Outlook profile chooser.
        """;

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 0 || args is ["--help" or "-h"])
        {
            Console.WriteLine(Usage);
            return 0;
        }

        using var cancellation = new CancellationTokenSource();
        ConsoleCancelEventHandler onCancel = (_, e) => { e.Cancel = true; cancellation.Cancel(); };
        Console.CancelKeyPress += onCancel;
        try
        {
            return Run(Options.Parse(args), cancellation.Token);
        }
        catch (ArgumentException e)
        {
            Console.Error.WriteLine($"{e.Message}\n\n{Usage}");
            return 2;
        }
        catch (OperationCanceledException)
        {
            Console.Error.WriteLine("Operazione annullata.");
            return 130;
        }
        catch (Exception e) when (e is MapiException or InvalidOperationException or IOException
            or DllNotFoundException or EntryPointNotFoundException or BadImageFormatException)
        {
            Console.Error.WriteLine($"Errore: {e.Message}");
            return 1;
        }
        finally
        {
            Console.CancelKeyPress -= onCancel;
        }
    }

    private static int Run(Options options, CancellationToken cancellation)
    {
        if (!OperatingSystem.IsWindows() || RuntimeInformation.ProcessArchitecture is not (Architecture.X64 or Architecture.X86))
            throw new InvalidOperationException("Serve Windows con Outlook classico e un processo x86/x64 della stessa architettura di Outlook.");

        if (options.Writes && !options.Yes && Console.IsInputRedirected)
            throw new InvalidOperationException("Con input reindirizzato, la cancellazione richiede anche --yes.");

        using var mapi = new MapiRuntime();
        using var session = mapi.Logon(options.Profile);
        var stores = Mapi.GetStores(session, cancellation);
        if (options.Command == "stores")
        {
            foreach (var item in stores) Console.WriteLine(item.Name);
            return 0;
        }

        var selected = Mapi.FindUnique(stores, options.Store!, "Store", "Usa il comando stores e verifica il profilo Outlook.");
        using var store = Mapi.OpenStore(session, selected.Id, options.Writes, options.Online);
        using var folder = Mapi.ResolveFolder(store, options.Folder, options.Writes, cancellation);
        if (options.Command == "folders")
        {
            foreach (var child in Mapi.GetChildren(folder, cancellation)) Console.WriteLine(child.Name);
            return 0;
        }

        var before = options.Before!.Value;
        var from = options.From is { } start ? $" | Da (incluso): {start:O}" : "";
        Console.Error.WriteLine($"Store: {selected.Name} | Cartella: {options.Folder}{from} | Prima di: {before:O}");
        if (options.WhatIf)
            Console.Error.WriteLine(options.Action == DeleteMode.Hard
                ? "WHATIF: anteprima hard delete. Nessuna modifica."
                : "WHATIF: anteprima spostamento in Posta eliminata. Nessuna modifica.");
        List<byte[]>? ids = options.Writes ? [] : null;
        var list = options.List || options.WhatIf;
        if (list) Console.WriteLine("received_utc\tentry_id_hex");
        var timer = Stopwatch.StartNew();
        var count = Mapi.Scan(folder, options.From?.ToFileTime(), before.ToFileTime(), list, ids, cancellation);
        Console.Error.WriteLine($"Corrispondenze: {count} ({timer.Elapsed.TotalSeconds:F2} s)");
        if (!list) Console.WriteLine(count);
        if (ids is null || ids.Count == 0) return 0;

        var hardDelete = options.Action == DeleteMode.Hard;
        using var trash = hardDelete ? null : Mapi.OpenTrash(store, folder);
        if (!options.Yes)
        {
            var confirmation = hardDelete ? "ELIMINA DEFINITIVAMENTE" : "SPOSTA";
            var action = hardDelete ? "Eliminare senza passare da Posta eliminata" : "Spostare in Posta eliminata";
            Console.Error.Write($"{action} {count} messaggi? Scrivi {confirmation} per confermare: ");
            if (Console.ReadLine() != confirmation) throw new OperationCanceledException();
        }

        var processed = 0;
        timer.Restart();
        try
        {
            while (processed < ids.Count)
            {
                cancellation.ThrowIfCancellationRequested();
                var batch = Math.Min(Mapi.MutationBatchSize, ids.Count - processed);
                var entries = CollectionsMarshal.AsSpan(ids).Slice(processed, batch);
                if (hardDelete) Mapi.HardDelete(folder, entries);
                else Mapi.Move(folder, trash!, entries);
                processed += batch;
            }
        }
        catch (MapiException)
        {
            Console.Error.WriteLine("Il blocco interrotto potrebbe essere stato applicato parzialmente. Riesegui scan prima di riprovare.");
            throw;
        }
        finally
        {
            var operation = hardDelete ? "Hard delete" : "Spostamenti";
            Console.Error.WriteLine($"{operation} confermati da MAPI: {processed}/{count} ({timer.Elapsed.TotalSeconds:F2} s)");
        }

        return 0;
    }

    private enum DeleteMode { None, Move, Hard }

    private sealed record Options(string Command, string? Profile, string? Store, string? Folder,
        DateTimeOffset? From, DateTimeOffset? Before, bool List, DeleteMode Action, bool WhatIf, bool Yes, bool Online)
    {
        private static readonly string[] CalendarFormats = ["yyyy-MM-dd", "d/M/yyyy", "dd/MM/yyyy"];
        public bool Writes => Action != DeleteMode.None && !WhatIf;

        public static Options Parse(string[] args)
        {
            var command = args[0];
            if (command is not ("stores" or "folders" or "scan"))
                throw new ArgumentException($"Comando sconosciuto: {command}");

            var values = new Dictionary<string, string?>(StringComparer.Ordinal);
            for (var i = 1; i < args.Length; i++)
            {
                var key = args[i];
                var takesValue = key is "--profile" or "--store" or "--folder" or "--from" or "--before" or "--to";
                var allowed = key == "--profile" || (command != "stores" && key is "--store" or "--folder" or "--online")
                    || (command == "scan" && key is "--from" or "--before" or "--to" or "--list" or "--delete" or "--hard-delete" or "--whatif" or "--yes");
                if (!allowed || !values.TryAdd(key, null))
                    throw new ArgumentException($"Opzione non valida o duplicata: {key}");
                if (!takesValue) continue;
                if (++i >= args.Length || string.IsNullOrWhiteSpace(args[i]) || args[i].StartsWith("--", StringComparison.Ordinal))
                    throw new ArgumentException($"Manca il valore di {key}.");
                values[key] = args[i];
            }

            var profile = values.GetValueOrDefault("--profile");
            var store = values.GetValueOrDefault("--store");
            var folder = values.GetValueOrDefault("--folder");
            var list = values.ContainsKey("--list");
            var delete = values.ContainsKey("--delete");
            var hardDelete = values.ContainsKey("--hard-delete");
            var whatIf = values.ContainsKey("--whatif");
            var yes = values.ContainsKey("--yes");
            if (command != "stores" && store is null) throw new ArgumentException("Specifica --store.");
            if (command == "scan" && folder is null) throw new ArgumentException("Specifica --folder.");
            if (profile?.Length > 64) throw new ArgumentException("Il profilo non puo superare 64 caratteri.");
            if (folder is not null && folder.Split('\\').Any(string.IsNullOrWhiteSpace))
                throw new ArgumentException("Il percorso deve essere relativo alla radice della posta e non avere segmenti vuoti.");
            if (delete && hardDelete) throw new ArgumentException("Scegli --delete oppure --hard-delete, non entrambi.");
            if (yes && whatIf) throw new ArgumentException("--whatif e --yes non sono combinabili.");
            if (yes && !delete && !hardDelete) throw new ArgumentException("--yes richiede --delete oppure --hard-delete.");

            DateTimeOffset? from = null;
            DateTimeOffset? before = null;
            if (command == "scan")
            {
                var beforeText = values.GetValueOrDefault("--before");
                var toText = values.GetValueOrDefault("--to");
                if ((beforeText is null) == (toText is null))
                    throw new ArgumentException("Specifica esattamente uno tra --before e --to.");

                if (values.GetValueOrDefault("--from") is { } fromText) from = ParseInstant(fromText, "--from");
                if (toText is not null)
                {
                    if (!TryParseDate(toText, out var to) || to == DateOnly.MaxValue || to.Year < 1601)
                        throw new ArgumentException("Specifica una data --to valida, senza orario, dal 1601-01-01 al 9999-12-30.");
                    before = AtUtcMidnight(to.AddDays(1));
                }
                else before = ParseInstant(beforeText!, "--before");

                if (from >= before) throw new ArgumentException("L'inizio dell'intervallo deve precedere il limite superiore esclusivo.");
            }

            var action = hardDelete ? DeleteMode.Hard : delete ? DeleteMode.Move : DeleteMode.None;
            return new Options(command, profile, store, folder, from, before, list, action, whatIf, yes, values.ContainsKey("--online"));
        }

        private static DateTimeOffset AtUtcMidnight(DateOnly date) => new(date.ToDateTime(TimeOnly.MinValue), TimeSpan.Zero);

        private static bool TryParseDate(string text, out DateOnly date) =>
            DateOnly.TryParseExact(text, CalendarFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);

        private static DateTimeOffset ParseInstant(string text, string option)
        {
            DateTimeOffset parsed;
            if (TryParseDate(text, out var date)) parsed = AtUtcMidnight(date);
            else if (!DateTimeOffset.TryParseExact(text, ["yyyy-MM-dd'T'HH:mm:ss'Z'", "yyyy-MM-dd'T'HH:mm:sszzz"],
                CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed))
                throw new ArgumentException($"Specifica una data {option} valida.");
            if (parsed <= DateTimeOffset.FromFileTime(0))
                throw new ArgumentException($"La data {option} deve essere successiva al 1601-01-01 UTC.");
            return parsed;
        }
    }
}
