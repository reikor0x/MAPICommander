using System.Globalization;
using Microsoft.Win32;

internal sealed class GuiLocalization
{
    internal sealed record Language(string Code, string Name, string Culture);

    internal static IReadOnlyList<Language> Languages { get; } = Array.AsReadOnly<Language>(
        [new("en", "English", "en-US"), new("it", "Italiano", "it-IT"), new("fr", "Français", "fr-FR"),
         new("de", "Deutsch", "de-DE"), new("es", "Español", "es-ES")]);

    private const string SettingsPath = @"Software\MAPICommander";
    private const string LanguageSetting = "GuiLanguage";
    private static readonly Lazy<Dictionary<string, string[]>> Catalog = new(LoadCatalog);
    private int _languageIndex;
    private CultureInfo _culture = CultureInfo.GetCultureInfo(Languages[0].Culture);

    public string LanguageCode => Languages[_languageIndex].Code;

    public void LoadPreference()
    {
        using var key = Registry.CurrentUser.OpenSubKey(SettingsPath);
        if (key?.GetValue(LanguageSetting) is string code) Select(code);
    }

    public void SavePreference()
    {
        using var key = Registry.CurrentUser.CreateSubKey(SettingsPath);
        key.SetValue(LanguageSetting, LanguageCode, RegistryValueKind.String);
    }

    public void Select(string code)
    {
        for (var i = 0; i < Languages.Count; i++)
        {
            if (Languages[i].Code != code) continue;
            _languageIndex = i;
            _culture = CultureInfo.GetCultureInfo(Languages[i].Culture);
            return;
        }
        _languageIndex = 0;
        _culture = CultureInfo.GetCultureInfo(Languages[0].Culture);
    }

    public string Text(string english, params object?[] values)
    {
        var text = _languageIndex != 0 && Catalog.Value.TryGetValue(english, out var variants)
            ? variants[_languageIndex] : english;
        return values.Length == 0 ? text : string.Format(_culture, text, values);
    }

    private static Dictionary<string, string[]> LoadCatalog()
    {
        using var stream = typeof(GuiLocalization).Assembly.GetManifestResourceStream("MAPICommander.GuiTranslations")
            ?? throw new InvalidOperationException("The GUI translation catalog is missing.");
        using var reader = new StreamReader(stream);
        var header = reader.ReadLine();
        if (header != string.Join('|', Languages.Select(language => language.Code)))
            throw new InvalidDataException("The GUI translation catalog has an invalid language header.");
        var catalog = new Dictionary<string, string[]>(StringComparer.Ordinal);
        while (reader.ReadLine() is { } line)
        {
            var columns = line.Split('|');
            if (columns.Length != Languages.Count || columns.Any(string.IsNullOrWhiteSpace) || !catalog.TryAdd(columns[0], columns))
                throw new InvalidDataException("The GUI translation catalog has an invalid or duplicate entry.");
        }
        return catalog;
    }
}