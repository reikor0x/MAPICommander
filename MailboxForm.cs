using System.Reflection;
using System.Security;

internal sealed class MailboxForm : Form
{
    private readonly MailboxSession _session = new();
    private readonly GuiLocalization _localization = new();
    private readonly ToolStripMenuItem _options = new();
    private readonly ToolStripMenuItem _language = new();
    private readonly ToolStripLabel _version = new() { Alignment = ToolStripItemAlignment.Right };
    private readonly Icon _appIcon;
    private readonly Bitmap _versionIcon;
    private readonly Label _profileLabel = LabelFor();
    private readonly Label _fromLabel = LabelFor();
    private readonly Label _toLabel = LabelFor();
    private readonly Label _scopeBanner = new() { Dock = DockStyle.Fill, AutoEllipsis = true, Padding = new Padding(12, 0, 12, 0), TextAlign = ContentAlignment.MiddleLeft };
    private readonly Font _headingFont;
    private TreeNode? _operationNode;
    private (string Text, object?[] Values) _statusMessage;
    private readonly ComboBox _profile = new() { Width = 225, DropDownStyle = ComboBoxStyle.DropDownList, DisplayMember = nameof(Mapi.Profile.Name) };
    private readonly Button _reloadProfiles = new() { Text = "Refresh profiles", AutoSize = true };
    private readonly CheckBox _online = new() { AutoSize = true };
    private readonly Button _connect = new() { Text = "Log on", AutoSize = true };
    private readonly TreeView _folders = new()
    {
        Dock = DockStyle.Fill, HideSelection = false, BorderStyle = BorderStyle.None, DrawMode = TreeViewDrawMode.OwnerDrawText,
        ShowLines = false, ItemHeight = 24, AccessibleName = "Stores and folders"
    };
    private readonly DateTimePicker _from = CreateDatePicker(showCheckBox: true);
    private readonly DateTimePicker _to = CreateDatePicker(showCheckBox: false);
    private readonly CheckBox _limitDates = new() { Text = "Filter by date", AutoSize = true, Checked = true };
    private readonly CheckBox _wholeMailbox = new() { Text = "Entire mailbox", AutoSize = true };
    private readonly CheckBox _includeSubfolders = new() { AutoSize = true };
    private readonly CheckBox _largest = new() { Text = "Largest first", AutoSize = true };
    private readonly Button _preview = new() { Text = "Preview", AutoSize = true };
    private readonly Button _export = new() { Text = "Export TSV", AutoSize = true };
    private readonly Button _move = new() { Text = "Move to Deleted Items", AutoSize = true };
    private readonly Button _hardDelete = new() { Text = "Hard delete...", AutoSize = true };
    private readonly Button _cancel = new() { Text = "Cancel operation", AutoSize = true };
    private readonly ListView _rows = new()
    {
        Dock = DockStyle.Fill, View = View.Details, VirtualMode = true, FullRowSelect = true,
        MultiSelect = true, HideSelection = false, BorderStyle = BorderStyle.None, AccessibleName = "Preview results"
    };
    private readonly Label _status = new() { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly PercentProgressBar _progress = new() { MinimumSize = new Size(0, 44), Dock = DockStyle.Fill, Visible = false };
    private readonly System.Windows.Forms.Timer _refresh = new() { Interval = 200 };
    private MailboxSession.Preview? _report;
    private CancellationTokenSource? _operation;
    private long _processed;
    private int _searchPercentage = -1;
    private bool _closing, _stopped;

    public MailboxForm()
    {
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(1030, 630);
        MinimumSize = new Size(820, 420);
        AutoScaleMode = AutoScaleMode.Dpi;
        using var iconStream = typeof(MailboxForm).Assembly.GetManifestResourceStream("MAPICommander.AppIcon")
            ?? throw new InvalidOperationException("The application icon resource is missing.");
        _appIcon = new Icon(iconStream);
        Icon = _appIcon;
        using var menuIcon = new Icon(_appIcon, SystemInformation.SmallIconSize);
        _versionIcon = menuIcon.ToBitmap();
        _version.Image = _versionIcon;
        _headingFont = new Font(Font, FontStyle.Bold);
        _scopeBanner.Font = _headingFont;
        _scopeBanner.Height = 36;
        _preview.Font = _headingFont;
        foreach (var button in new[] { _reloadProfiles, _connect, _preview, _export, _move, _hardDelete, _cancel })
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = SystemColors.ControlDark;
            button.Padding = new Padding(8, 2, 8, 2);
        }
        string? preferenceError = null;
        try { _localization.LoadPreference(); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or SecurityException) { preferenceError = e.Message; }
        var menu = new MenuStrip { Dock = DockStyle.Fill, Items = { _options, _version } };
        _options.DropDownItems.Add(_language);
        foreach (var language in GuiLocalization.Languages)
        {
            var item = new ToolStripMenuItem(language.Name) { Tag = language.Code };
            item.Click += (_, _) => ChangeLanguage(language.Code);
            _language.DropDownItems.Add(item);
        }
        MainMenuStrip = menu;

        var split = new SplitContainer { Dock = DockStyle.Fill, Size = new Size(1000, 500), SplitterDistance = 265, SplitterWidth = 8 };
        split.Panel1.BackColor = split.Panel2.BackColor = SystemColors.Window;
        split.Panel1.Padding = split.Panel2.Padding = new Padding(8);
        split.Panel1.Controls.Add(_folders);
        split.Panel2.Controls.Add(_rows);
        foreach (var width in new[] { 280, 155, 0, 205, 500 }) _rows.Columns.Add("", width);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 9, Padding = new Padding(12, 6, 12, 8) };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.Controls.Add(menu, 0, 0);
        layout.Controls.Add(Bar(_profileLabel, _profile, _reloadProfiles, _online, _connect), 0, 1);
        layout.Controls.Add(Bar(_limitDates, _fromLabel, _from, _toLabel, _to), 0, 2);
        layout.Controls.Add(Bar(_wholeMailbox, _includeSubfolders, _largest, _preview), 0, 3);
        layout.Controls.Add(_scopeBanner, 0, 4);
        layout.Controls.Add(split, 0, 5);
        layout.Controls.Add(Bar(_export, _move, _hardDelete, _cancel), 0, 6);
        _status.Height = 26;
        layout.Controls.Add(_status, 0, 7);
        layout.Controls.Add(_progress, 0, 8);
        Controls.Add(layout);

        SetStatus("Select a classic Outlook profile to browse stores and folders.");
        ApplyLanguage();
        Shown += async (_, _) =>
        {
            if (preferenceError is not null) ShowPreferencesError(preferenceError);
            await LoadProfilesAsync();
        };
        _reloadProfiles.Click += async (_, _) => await LoadProfilesAsync();
        _profile.SelectedIndexChanged += (_, _) => UpdateCommands();
        _profile.SelectionChangeCommitted += (_, _) =>
        {
            ResetBrowser();
            SetStatus("Click Log on to browse the selected Outlook profile.");
        };
        _connect.Click += async (_, _) => await ConnectAsync();
        _folders.DrawNode += DrawFolder;
        _folders.BeforeSelect += (_, e) => e.Cancel = _operation is not null || _closing;
        _folders.BeforeCollapse += (_, e) => e.Cancel = _operation is not null || _closing;
        _folders.BeforeExpand += async (_, e) =>
        {
            if (_operation is not null || _closing)
            {
                e.Cancel = true;
                return;
            }
            if (e.Node is not { Tag: MailboxSession.Target target } node || node.Nodes.Count != 1 || node.Nodes[0].Tag is not null)
                return;
            e.Cancel = true;
            var loaded = false;
            await ExecuteAsync("Reading folders", async token =>
            {
                var children = await _session.ChildrenAsync(target, token);
                node.Nodes.Clear();
                foreach (var child in children)
                    node.Nodes.Add(CreateNode(child.Name, new MailboxSession.Target(target.StoreId, child.Id, target)));
                loaded = true;
                SetStatus("Subfolders in {1}: {0:N0}.", children.Count, node.Text);
            }, activeNode: node);
            if (loaded) node.Expand();
        };
        _folders.AfterSelect += (_, _) => ClearReport();
        _limitDates.CheckedChanged += (_, _) => ClearReport();
        _wholeMailbox.CheckedChanged += (_, _) => ClearReport();
        _includeSubfolders.CheckedChanged += (_, _) => ClearReport();
        _largest.CheckedChanged += (_, _) => ClearReport();
        _from.ValueChanged += (_, _) => ClearReport();
        _to.ValueChanged += (_, _) => ClearReport();
        _preview.Click += async (_, _) => await PreviewAsync();
        _export.Click += async (_, _) => await ExportAsync();
        _move.Click += async (_, _) => await RemoveAsync(hard: false);
        _hardDelete.Click += async (_, _) => await RemoveAsync(hard: true);
        _cancel.Click += (_, _) => RequestCancellation();
        _rows.RetrieveVirtualItem += (_, e) =>
        {
            if (_report is not { Rows: var rows } || (uint)e.ItemIndex >= (uint)rows.Count)
            {
                e.Item = new ListViewItem();
                return;
            }

            var row = rows[e.ItemIndex];
            e.Item = new ListViewItem([row.Subject, row.Received > 0 ? DateTime.FromFileTimeUtc(row.Received).ToString("u") : "",
                row.Size is { } size ? T("{0:N0} B", size) : "", row.Source.Path, Convert.ToHexString(row.Id)]);
        };
        UpdateCommands();
    }

    private static DateTimePicker CreateDatePicker(bool showCheckBox) => new()
    {
        Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", ShowCheckBox = showCheckBox,
        Width = 135, Value = DateTime.UtcNow.Date.AddDays(-1), Checked = !showCheckBox
    };

    private string T(string text, params object?[] values) => _localization.Text(text, values);

    private void SetStatus(string text, params object?[] values)
    {
        _statusMessage = (text, values);
        _status.Text = T(text, values);
    }

    private void ChangeLanguage(string code)
    {
        _localization.Select(code);
        ApplyLanguage();
        try { _localization.SavePreference(); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or SecurityException) { ShowPreferencesError(e.Message); }
    }

    private void ShowPreferencesError(string detail) => MessageBox.Show(this,
        T("The language preference could not be read or saved.") + "\n\n" + detail, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);

    private void ApplyLanguage()
    {
        Text = T("MAPICommander - Exchange Mailbox");
        _options.Text = T("Options");
        _language.Text = T("Language");
        foreach (ToolStripMenuItem item in _language.DropDownItems) item.Checked = Equals(item.Tag, _localization.LanguageCode);
        var assembly = typeof(MailboxForm).Assembly;
        var version = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0]
            ?? assembly.GetName().Version?.ToString() ?? "?";
        _version.Text = T("Version {0}", version);
        _profileLabel.Text = T("Outlook profile:");
        _fromLabel.Text = T("Received from (UTC):");
        _toLabel.Text = T("Through (inclusive):");
        _reloadProfiles.Text = T("Refresh profiles");
        _online.Text = T("Online mode no cache");
        _connect.Text = T("Log on");
        _limitDates.Text = T("Filter by date");
        _wholeMailbox.Text = T("Entire mailbox");
        _includeSubfolders.Text = T("Include subfolders");
        _largest.Text = T("Largest first");
        _preview.Text = T("Preview");
        _export.Text = T("Export TSV");
        _move.Text = T("Move to Deleted Items");
        _hardDelete.Text = T("Hard delete...");
        _cancel.Text = T("Cancel operation");
        _profile.AccessibleName = T("Outlook profile");
        _from.AccessibleName = T("Received from UTC (optional)");
        _to.AccessibleName = T("Received through UTC");
        _folders.AccessibleName = T("Stores and folders");
        _rows.AccessibleName = T("Preview results");
        _progress.AccessibleName = T("Operation progress");
        _progress.SetWorkingText(T("Working..."));
        _rows.Columns[0].Text = T("Subject");
        _rows.Columns[1].Text = T("Received (UTC)");
        _rows.Columns[2].Text = T("Size");
        _rows.Columns[3].Text = T("Folder");
        _rows.Columns[4].Text = T("EntryID (hex)");
        _status.Text = T(_statusMessage.Text, _statusMessage.Values);
        UpdateScopeBanner();
        _rows.Invalidate();
    }

    private static Label LabelFor() => new() { AutoSize = true, Margin = new Padding(3, 8, 3, 0) };

    private static FlowLayoutPanel Bar(params Control[] controls)
    {
        var bar = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        bar.Controls.AddRange(controls);
        return bar;
    }

    private static TreeNode CreateNode(string name, MailboxSession.Target target)
    {
        var node = new TreeNode(name) { Tag = target };
        node.Nodes.Add("");
        return node;
    }

    private TreeNode? SearchNode
    {
        get
        {
            var node = _folders.SelectedNode;
            if (!_wholeMailbox.Checked) return node?.Tag is MailboxSession.Target { FolderId: not null } ? node : null;
            while (node?.Parent is { } parent) node = parent;
            return node?.Tag is MailboxSession.Target { FolderId: null } ? node : null;
        }
    }

    private MailboxSession.SearchScope SearchScope => _wholeMailbox.Checked ? MailboxSession.SearchScope.Mailbox
        : _includeSubfolders.Checked ? MailboxSession.SearchScope.Subtree : MailboxSession.SearchScope.Folder;

    private string DescribeScope(TreeNode node) => T(node.Tag is MailboxSession.Target { FolderId: null } ? "Mailbox: {0}"
        : SearchScope == MailboxSession.SearchScope.Subtree ? "Folder and subfolders: {0}" : "Folder: {0}", node.FullPath);

    private void UpdateScopeBanner()
    {
        var node = _operationNode ?? SearchNode ?? _folders.SelectedNode;
        var active = node?.Tag is MailboxSession.Target;
        _scopeBanner.BackColor = active ? SystemColors.Highlight : SystemColors.ControlLight;
        _scopeBanner.ForeColor = active ? SystemColors.HighlightText : SystemColors.ControlText;
        _scopeBanner.Text = active ? DescribeScope(node!) : T("No folder selected");
        _scopeBanner.AccessibleName = _scopeBanner.Text;
        _folders.Invalidate();
    }

    private void DrawFolder(object? sender, DrawTreeNodeEventArgs e)
    {
        if (e.Node is not { } node)
        {
            e.DrawDefault = true;
            return;
        }

        var highlighted = ReferenceEquals(node, _operationNode ?? _folders.SelectedNode);
        if (!highlighted && _operation is null)
        {
            e.DrawDefault = true;
            return;
        }

        var bounds = e.Bounds;
        bounds.Width = Math.Max(0, _folders.ClientSize.Width - bounds.Left);
        var font = highlighted ? _headingFont : _folders.Font;
        e.Graphics.FillRectangle(highlighted ? SystemBrushes.Highlight : SystemBrushes.Window, bounds);
        TextRenderer.DrawText(e.Graphics, node.Text, font, bounds, highlighted ? SystemColors.HighlightText : SystemColors.GrayText,
            TextFormatFlags.NoPrefix | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
        if (highlighted && _folders.Focused && _operation is null)
            ControlPaint.DrawFocusRectangle(e.Graphics, bounds, SystemColors.HighlightText, SystemColors.Highlight);
    }

    private Task LoadProfilesAsync() => ExecuteAsync("Reading Outlook profiles", async token =>
    {
        var previous = (_profile.SelectedItem as Mapi.Profile)?.Name;
        ResetBrowser();
        _profile.Items.Clear();
        var profiles = await _session.ProfilesAsync(token);
        profiles.Sort((left, right) => StringComparer.CurrentCultureIgnoreCase.Compare(left.Name, right.Name));
        _profile.Items.AddRange(profiles.ToArray());
        var selected = profiles.FindIndex(profile => profile.Name.Equals(previous, StringComparison.Ordinal));
        if (selected < 0) selected = profiles.FindIndex(profile => profile.IsDefault);
        _profile.SelectedIndex = selected < 0 && profiles.Count == 1 ? 0 : selected;
        SetStatus(profiles.Count == 0
            ? "No Outlook profiles found. Open classic Outlook as this Windows user, then select Refresh profiles."
            : "Select an Outlook profile and click Log on.");
    });

    private Task ConnectAsync()
    {
        if (_profile.SelectedItem is not Mapi.Profile profile) return Task.CompletedTask;
        ResetBrowser();
        var online = _online.Checked;
        var owner = Handle;
        return ExecuteAsync("Logging on", async token =>
        {
            var stores = await _session.ConnectAsync(profile.Name, online, owner, token);
            foreach (var store in stores)
                _folders.Nodes.Add(CreateNode(store.Name, new MailboxSession.Target(store.Id, FolderId: null)));
            SetStatus(stores.Count == 0 ? Mapi.NoStoresMessage : "Expand a store and select a mail folder.");
        });
    }

    private Task PreviewAsync()
    {
        if (SearchNode is not { Tag: MailboxSession.Target target } node) return Task.CompletedTask;
        var scope = SearchScope;
        var largest = _largest.Checked;
        var limitDates = _limitDates.Checked;
        var start = limitDates && _from.Checked ? _from.Value.Date : (DateTime?)null;
        var end = _to.Value.Date;
        if (start > end)
        {
            MessageBox.Show(this, T("The start date cannot be after the final date."), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return Task.CompletedTask;
        }
        if (scope == MailboxSession.SearchScope.Mailbox)
        {
            var search = new TaskDialogButton(T("Search"));
            var cancel = new TaskDialogButton(T("Cancel"));
            var warning = new TaskDialogPage
            {
                Caption = T("Entire mailbox search"),
                Heading = T("Search all mail folders in '{0}'?", node.Text),
                Text = T("This can take a long time, especially for large or online mailboxes. No messages will be changed."),
                Icon = TaskDialogIcon.Warning, Buttons = { search, cancel }, DefaultButton = cancel, AllowCancel = true
            };
            if (TaskDialog.ShowDialog(this, warning) != search) return Task.CompletedTask;
        }

        ClearReport();
        return ExecuteAsync("Searching", async token =>
        {
            _report = await _session.PreviewAsync(target, node.FullPath, scope,
                start is { } first ? DateTime.SpecifyKind(first, DateTimeKind.Utc).ToFileTimeUtc() : null,
                limitDates ? DateTime.SpecifyKind(end.AddDays(1), DateTimeKind.Utc).ToFileTimeUtc() : null,
                largest, ReportProgress, value => Volatile.Write(ref _searchPercentage, value ?? -1), token);
            _rows.Columns[2].Width = largest ? 115 : 0;
            _rows.VirtualListSize = _report.Rows.Count;
            SetStatus("Search ended: {0:N0} matching messages. No changes made.", _report.Rows.Count);
        }, showCount: true, activeNode: node);
    }

    private Task ExportAsync()
    {
        if (_report is not { Rows: var report, IncludesSize: var includeSize }) return Task.CompletedTask;
        using var dialog = new SaveFileDialog { Filter = T("Tab-separated report (*.tsv)") + "|*.tsv", FileName = "preview.tsv" };
        if (dialog.ShowDialog(this) != DialogResult.OK) return Task.CompletedTask;
        var path = dialog.FileName;
        return ExecuteAsync("Exporting", async token =>
        {
            await Task.Run(() =>
            {
                using var writer = new StreamWriter(path);
                writer.WriteLine(includeSize ? "received_utc\tentry_id_hex\tfolder\tsize_bytes" : "received_utc\tentry_id_hex\tfolder");
                foreach (var row in report)
                {
                    token.ThrowIfCancellationRequested();
                    var identity = row.Received > 0 ? Mapi.FormatReportRow(row.Received, row.Id) : $"\t{Convert.ToHexString(row.Id)}";
                    var folder = row.Source.Path.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
                    writer.WriteLine(includeSize ? FormattableString.Invariant($"{identity}\t{folder}\t{row.Size}") : $"{identity}\t{folder}");
                }
            }, token);
            SetStatus("Exported {0:N0} messages.", report.Count);
        }, interruptedHint: () => T("The output file may be incomplete."), activeNode: SearchNode);
    }

    private Task RemoveAsync(bool hard)
    {
        if (_report is not { Rows.Count: > 0 } report || SearchNode is not { } node || !ReferenceEquals(node.Tag, report.Target)) return Task.CompletedTask;
        var selected = _rows.SelectedIndices.Cast<int>().ToArray();
        var allButton = new TaskDialogButton(T("All previewed messages ({0:N0})", report.Rows.Count));
        var selectedButton = new TaskDialogButton(T("Only highlighted messages ({0:N0})", selected.Length)) { Enabled = selected.Length > 0 };
        var cancelButton = new TaskDialogButton(T("Cancel"));
        var page = new TaskDialogPage
        {
            Caption = T(hard ? "Confirm hard delete" : "Confirm move to Deleted Items"),
            Heading = T("Which previewed messages do you want to remove?"),
            Text = DescribeScope(node) + "\n\n" + T(hard ? "Hard delete bypasses Deleted Items. Recovery may not be possible." : "The chosen messages will be moved to Deleted Items."),
            Icon = TaskDialogIcon.Warning,
            Buttons = { selectedButton, allButton, cancelButton },
            DefaultButton = cancelButton,
            AllowCancel = true
        };
        var choice = TaskDialog.ShowDialog(this, page);
        if (choice != allButton && (choice != selectedButton || selected.Length == 0)) return Task.CompletedTask;
        var indices = choice == selectedButton ? selected : null;
        var total = indices?.Length ?? report.Rows.Count;

        ClearReport();
        var batchStarted = 0;
        return ExecuteAsync(hard ? "Hard deleting" : "Moving", async token =>
        {
            var count = await _session.RemoveAsync(report, indices, hard, () => Volatile.Write(ref batchStarted, 1), ReportProgress, token);
            SetStatus(hard ? "Permanently deleted: {0:N0}." : "Moved to Deleted Items: {0:N0}.", count);
        }, total: total, interruptedHint: () => Volatile.Read(ref batchStarted) == 0
            ? T("No removal batch was submitted. Run Preview again before retrying.")
            : T("Messages confirmed: {0:N0}. An unconfirmed batch may have changed the mailbox. Run Preview again before retrying.", Interlocked.Read(ref _processed)), activeNode: node);
    }

    private async Task ExecuteAsync(string label, Func<CancellationToken, Task> work, int? total = null,
        bool showCount = false, Func<string>? interruptedHint = null, TreeNode? activeNode = null)
    {
        if (_operation is not null || _closing) return;
        var operation = T(label);
        using var cancellation = new CancellationTokenSource();
        _operation = cancellation;
        _operationNode = activeNode;
        ReportProgress(0);
        Volatile.Write(ref _searchPercentage, -1);
        UpdateCommands();
        _progress.Visible = true;
        void RefreshProgress(object? sender, EventArgs args)
        {
            if (cancellation.IsCancellationRequested) return;
            var processed = Interlocked.Read(ref _processed);
            if (total is int count)
            {
                _progress.SetProgress((int)(processed * 100 / count));
                SetStatus("{0}: {1:N0} / {2:N0} confirmed", operation, processed, count);
            }
            else
            {
                var percentage = Volatile.Read(ref _searchPercentage);
                _progress.SetProgress(percentage >= 0 ? percentage : null);
                if (showCount) SetStatus("{0}: {1:N0} matches", operation, processed);
                else SetStatus("{0}...", operation);
            }
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
            SetStatus("{0}: cancelled. {1}", operation, interruptedHint?.Invoke());
        }
        catch (Exception e) when (e is MapiException or InvalidOperationException or IOException or UnauthorizedAccessException
            or DllNotFoundException or EntryPointNotFoundException or BadImageFormatException)
        {
            _refresh.Stop();
            var hint = interruptedHint?.Invoke();
            SetStatus("{0}: failed. {1}", operation, hint);
            MessageBox.Show(this, T(e.Message) + (hint is null ? "" : "\n\n" + hint), T("{0}: failed", operation), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _refresh.Stop();
            _refresh.Tick -= RefreshProgress;
            _progress.Visible = false;
            _operationNode = null;
            _operation = null;
            UpdateCommands();
        }
    }

    private void ReportProgress(long count) => Interlocked.Exchange(ref _processed, count);

    private void RequestCancellation()
    {
        _operation?.Cancel();
        _cancel.Enabled = false;
        SetStatus("Cancelling after the current operation returns...");
    }

    private void ClearReport()
    {
        if (_report is { Rows.Count: > 0 }) SetStatus("Preview cleared. Run Preview for the current selection.");
        _rows.VirtualListSize = 0;
        _report = null;
        UpdateCommands();
    }

    private void ResetBrowser()
    {
        _folders.Nodes.Clear();
        ClearReport();
    }

    private void UpdateCommands()
    {
        var idle = _operation is null && !_closing;
        _options.Enabled = idle;
        _reloadProfiles.Enabled = _online.Enabled = _rows.Enabled = idle;
        _limitDates.Enabled = _wholeMailbox.Enabled = _largest.Enabled = idle;
        _includeSubfolders.Enabled = idle && !_wholeMailbox.Checked && SearchNode is not null;
        _from.Enabled = _to.Enabled = idle && _limitDates.Checked;
        _profile.Enabled = idle && _profile.Items.Count > 0;
        _connect.Enabled = idle && _profile.SelectedItem is Mapi.Profile;
        _preview.Enabled = idle && SearchNode is not null;
        _export.Enabled = _move.Enabled = _hardDelete.Enabled = _preview.Enabled && _report is { Rows.Count: > 0 };
        _cancel.Enabled = _operation is { IsCancellationRequested: false };
        _preview.BackColor = _preview.Enabled ? SystemColors.Highlight : SystemColors.Control;
        _preview.ForeColor = _preview.Enabled ? SystemColors.HighlightText : SystemColors.GrayText;
        UpdateScopeBanner();
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
        base.Dispose(disposing);
        if (disposing)
        {
            _refresh.Dispose();
            _headingFont.Dispose();
            _versionIcon.Dispose();
            _appIcon.Dispose();
        }
    }

    private sealed class PercentProgressBar : Control
    {
        private int? _percentage;
        private int _step;
        private string _workingText = "Working...";

        public PercentProgressBar()
        {
            DoubleBuffered = true;
            AccessibleRole = AccessibleRole.ProgressBar;
        }

        public void SetProgress(int? percentage)
        {
            _percentage = percentage;
            if (percentage is null) _step = (_step + 1) % 12;
            Invalidate();
        }

        public void SetWorkingText(string text)
        {
            _workingText = text;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var track = ClientRectangle;
            e.Graphics.FillRectangle(SystemBrushes.ControlDark, track);
            track.Inflate(-1, -1);
            e.Graphics.FillRectangle(SystemBrushes.Window, track);
            var fill = track;
            if (_percentage is int percent) fill.Width = (int)(fill.Width * percent / 100L);
            else
            {
                fill.Width /= 4;
                fill.X += (track.Width - fill.Width) * _step / 11;
            }
            e.Graphics.FillRectangle(SystemBrushes.Highlight, fill);
            var text = _percentage is int value ? $"{value}%" : _workingText;
            using var progressFont = new Font(Font.FontFamily, Font.Size * 1.2f, FontStyle.Bold, Font.Unit);
            const TextFormatFlags flags = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                TextFormatFlags.SingleLine | TextFormatFlags.PreserveGraphicsClipping;
            TextRenderer.DrawText(e.Graphics, text, progressFont, track, SystemColors.WindowText, flags);
            var state = e.Graphics.Save();
            e.Graphics.SetClip(fill, System.Drawing.Drawing2D.CombineMode.Intersect);
            TextRenderer.DrawText(e.Graphics, text, progressFont, track, SystemColors.HighlightText, flags);
            e.Graphics.Restore(state);
        }
    }
}