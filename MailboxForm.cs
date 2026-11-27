internal sealed class MailboxForm : Form
{
    private readonly MailboxSession _session = new();
    private readonly TextBox _profile = new() { Width = 150, MaxLength = 64, PlaceholderText = "Choose at logon" };
    private readonly CheckBox _online = new() { Text = "Online (at logon)", AutoSize = true };
    private readonly Button _connect = new() { Text = "Log on", AutoSize = true };
    private readonly TreeView _folders = new() { Dock = DockStyle.Fill, HideSelection = false, AccessibleName = "Stores and folders" };
    private readonly DateTimePicker _from = CreateDatePicker(showCheckBox: true);
    private readonly DateTimePicker _to = CreateDatePicker(showCheckBox: false);
    private readonly Button _preview = new() { Text = "Preview", AutoSize = true };
    private readonly Button _export = new() { Text = "Export TSV", AutoSize = true };
    private readonly Button _move = new() { Text = "Move to Deleted Items", AutoSize = true };
    private readonly Button _hardDelete = new() { Text = "Hard delete...", AutoSize = true };
    private readonly Button _cancel = new() { Text = "Cancel operation", AutoSize = true };
    private readonly ListView _rows = new()
    {
        Dock = DockStyle.Fill, View = View.Details, VirtualMode = true, FullRowSelect = true, AccessibleName = "Preview results"
    };
    private readonly ToolStripStatusLabel _status = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly ToolStripProgressBar _progress = new() { Maximum = 1000, Visible = false };
    private readonly System.Windows.Forms.Timer _refresh = new() { Interval = 200 };
    private List<MailboxSession.Row> _report = [];
    private CancellationTokenSource? _operation;
    private long _processed;
    private bool _closing, _stopped;

    public MailboxForm()
    {
        Text = "MAPICommander - Exchange Mailbox";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(1030, 630);
        MinimumSize = new Size(820, 420);
        AutoScaleMode = AutoScaleMode.Dpi;
        _status.Text = "Log on to a classic Outlook profile to browse stores and folders.";
        _profile.AccessibleName = "Outlook profile (optional)";
        _from.AccessibleName = "Received from UTC (optional)";
        _to.AccessibleName = "Received through UTC";

        var split = new SplitContainer { Dock = DockStyle.Fill, Size = new Size(1000, 500), SplitterDistance = 265 };
        split.Panel1.Controls.Add(_folders);
        split.Panel2.Controls.Add(_rows);
        _rows.Columns.Add("Received (UTC)", 190);
        _rows.Columns.Add("EntryID (hex)", 500);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(6) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(Bar(LabelFor("Outlook profile:"), _profile, _online, _connect), 0, 0);
        layout.Controls.Add(Bar(LabelFor("Received from (UTC):"), _from, LabelFor("Through (inclusive):"), _to, _preview), 0, 1);
        layout.Controls.Add(split, 0, 2);
        layout.Controls.Add(Bar(_export, _move, _hardDelete, _cancel), 0, 3);
        layout.Controls.Add(new StatusStrip { Dock = DockStyle.Fill, SizingGrip = false, Items = { _status, _progress } }, 0, 4);
        Controls.Add(layout);

        _connect.Click += async (_, _) => await ConnectAsync();
        _folders.BeforeExpand += async (_, e) =>
        {
            if (e.Node is not { Tag: MailboxSession.Target target } node || node.Nodes.Count != 1 || node.Nodes[0].Tag is not null)
                return;
            e.Cancel = true;
            await ExecuteAsync("Reading folders", async token =>
            {
                var children = await _session.ChildrenAsync(target, token);
                node.Nodes.Clear();
                foreach (var child in children)
                    node.Nodes.Add(CreateNode(child.Name, new MailboxSession.Target(target.StoreId, child.Id)));
                node.Expand();
                _status.Text = $"{children.Count:N0} subfolders in {node.Text}.";
            });
        };
        _folders.AfterSelect += (_, _) => ClearReport();
        _from.ValueChanged += (_, _) => ClearReport();
        _to.ValueChanged += (_, _) => ClearReport();
        _preview.Click += async (_, _) => await PreviewAsync();
        _export.Click += async (_, _) => await ExportAsync();
        _move.Click += async (_, _) => await RemoveAsync(hard: false);
        _hardDelete.Click += async (_, _) => await RemoveAsync(hard: true);
        _cancel.Click += (_, _) => RequestCancellation();
        _rows.RetrieveVirtualItem += (_, e) =>
        {
            if ((uint)e.ItemIndex >= (uint)_report.Count)
            {
                e.Item = new ListViewItem();
                return;
            }

            var row = _report[e.ItemIndex];
            e.Item = new ListViewItem([DateTime.FromFileTimeUtc(row.Received).ToString("u"), Convert.ToHexString(row.Id)]);
        };
        UpdateCommands();
    }

    private static DateTimePicker CreateDatePicker(bool showCheckBox) => new()
    {
        Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", ShowCheckBox = showCheckBox,
        Width = 135, Value = DateTime.UtcNow.Date.AddDays(-1), Checked = !showCheckBox
    };

    private static Label LabelFor(string text) => new() { Text = text, AutoSize = true, Margin = new Padding(3, 8, 3, 0) };

    private static FlowLayoutPanel Bar(params Control[] controls)
    {
        var bar = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
        bar.Controls.AddRange(controls);
        return bar;
    }

    private static TreeNode CreateNode(string name, MailboxSession.Target target)
    {
        var node = new TreeNode(name) { Tag = target };
        node.Nodes.Add("Loading...");
        return node;
    }

    private MailboxSession.Target? SelectedFolder => _folders.SelectedNode?.Tag is MailboxSession.Target { FolderId: not null } target ? target : null;

    private Task ConnectAsync()
    {
        _folders.Nodes.Clear();
        ClearReport();
        var profile = string.IsNullOrWhiteSpace(_profile.Text) ? null : _profile.Text;
        var online = _online.Checked;
        var owner = Handle;
        return ExecuteAsync("Logging on", async token =>
        {
            var stores = await _session.ConnectAsync(profile, online, owner, token);
            foreach (var store in stores)
                _folders.Nodes.Add(CreateNode(store.Name, new MailboxSession.Target(store.Id, FolderId: null)));
            _status.Text = stores.Count == 0 ? "No stores found in this profile." : "Expand a store and select a mail folder.";
        });
    }

    private Task PreviewAsync()
    {
        if (SelectedFolder is not { } target) return Task.CompletedTask;
        var start = _from.Checked ? _from.Value.Date : (DateTime?)null;
        var end = _to.Value.Date;
        if (start > end)
        {
            MessageBox.Show(this, "The start date cannot be after the final date.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return Task.CompletedTask;
        }

        ClearReport();
        return ExecuteAsync("Scanning", async token =>
        {
            _report = await _session.PreviewAsync(target,
                start is { } first ? DateTime.SpecifyKind(first, DateTimeKind.Utc).ToFileTimeUtc() : null,
                DateTime.SpecifyKind(end.AddDays(1), DateTimeKind.Utc).ToFileTimeUtc(), ReportProgress, token);
            _rows.VirtualListSize = _report.Count;
            _status.Text = $"{_report.Count:N0} matching messages. No changes made.";
        }, showCount: true);
    }

    private Task ExportAsync()
    {
        using var dialog = new SaveFileDialog { Filter = "Tab-separated report (*.tsv)|*.tsv", FileName = "preview.tsv" };
        if (dialog.ShowDialog(this) != DialogResult.OK) return Task.CompletedTask;
        var path = dialog.FileName;
        var report = _report;
        return ExecuteAsync("Exporting", async token =>
        {
            await Task.Run(() =>
            {
                using var writer = new StreamWriter(path);
                writer.WriteLine("received_utc\tentry_id_hex");
                foreach (var row in report)
                {
                    token.ThrowIfCancellationRequested();
                    writer.WriteLine(Mapi.FormatReportRow(row.Received, row.Id));
                }
            }, token);
            _status.Text = $"Exported {report.Count:N0} messages.";
        }, interruptedHint: " The output file may be incomplete.");
    }

    private Task RemoveAsync(bool hard)
    {
        if (SelectedFolder is not { } target || _report.Count == 0) return Task.CompletedTask;
        var report = _report;
        var prompt = hard
            ? $"Hard-delete {report.Count:N0} previewed messages from '{_folders.SelectedNode!.FullPath}'?\n\nThis bypasses Deleted Items and recovery may not be possible."
            : $"Move {report.Count:N0} previewed messages from '{_folders.SelectedNode!.FullPath}' to Deleted Items?";
        if (MessageBox.Show(this, prompt, hard ? "Confirm hard delete" : "Confirm move", MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
            return Task.CompletedTask;

        ClearReport();
        return ExecuteAsync(hard ? "Hard deleting" : "Moving", async token =>
        {
            var count = await _session.RemoveAsync(target, report, hard, ReportProgress, token);
            _status.Text = $"{count:N0} messages {(hard ? "hard-deleted" : "moved to Deleted Items")}.";
        }, total: report.Count, interruptedHint: " Some messages may already have been removed. Preview again before retrying.");
    }

    private async Task ExecuteAsync(string label, Func<CancellationToken, Task> work, int? total = null,
        bool showCount = false, string interruptedHint = "")
    {
        if (_operation is not null || _closing) return;
        using var cancellation = new CancellationTokenSource();
        _operation = cancellation;
        ReportProgress(0);
        UpdateCommands();
        _progress.Style = total is null ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
        _progress.Value = 0;
        _progress.Visible = true;
        void RefreshProgress(object? sender, EventArgs args)
        {
            if (cancellation.IsCancellationRequested) return;
            var processed = Interlocked.Read(ref _processed);
            if (total is int count)
            {
                _progress.Value = (int)(processed * _progress.Maximum / count);
                _status.Text = $"{label}: {processed:N0} / {count:N0} confirmed";
            }
            else _status.Text = showCount ? $"{label}: {processed:N0} matches" : $"{label}...";
        }

        RefreshProgress(null, EventArgs.Empty);
        _refresh.Tick += RefreshProgress;
        _refresh.Start();
        try
        {
            await work(cancellation.Token);
        }
        catch (OperationCanceledException)
        {
            _status.Text = $"{label} cancelled.{interruptedHint}";
        }
        catch (Exception e) when (e is MapiException or InvalidOperationException or IOException or UnauthorizedAccessException
            or DllNotFoundException or EntryPointNotFoundException or BadImageFormatException)
        {
            _refresh.Stop();
            _status.Text = $"{label} failed.{interruptedHint}";
            MessageBox.Show(this, e.Message + interruptedHint, $"{label} failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _refresh.Stop();
            _refresh.Tick -= RefreshProgress;
            _progress.Visible = false;
            _operation = null;
            UpdateCommands();
        }
    }

    private void ReportProgress(long count) => Interlocked.Exchange(ref _processed, count);

    private void RequestCancellation()
    {
        _operation?.Cancel();
        _cancel.Enabled = false;
        _status.Text = "Cancelling after the current operation returns...";
    }

    private void ClearReport()
    {
        if (_report.Count != 0) _status.Text = "Preview cleared. Run Preview for the current selection.";
        _rows.VirtualListSize = 0;
        _report = [];
        UpdateCommands();
    }

    private void UpdateCommands()
    {
        var idle = _operation is null && !_closing;
        _profile.Enabled = _online.Enabled = _connect.Enabled = _folders.Enabled = _from.Enabled = _to.Enabled = idle;
        _preview.Enabled = idle && SelectedFolder is not null;
        _export.Enabled = _move.Enabled = _hardDelete.Enabled = _preview.Enabled && _report.Count > 0;
        _cancel.Enabled = _operation is { IsCancellationRequested: false };
    }

    protected override async void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        if (e.Cancel || _stopped) return;
        e.Cancel = true;
        if (_operation is not null)
        {
            RequestCancellation();
            return;
        }

        if (_closing) return;
        _closing = true;
        UpdateCommands();
        await _session.StopAsync();
        _stopped = true;
        Close();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _refresh.Dispose();
        base.Dispose(disposing);
    }
}