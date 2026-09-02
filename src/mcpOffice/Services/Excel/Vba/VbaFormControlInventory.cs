using System.Text.RegularExpressions;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

/// <summary>
/// Infers a UserForm's controls from its code-behind: handler names (`cmdOK_Click`), `Me.ctrl`
/// references, Hungarian-prefixed bare references and `As MSForms.X` declarations. The binary
/// `.frx` designer part is not read (VBA-015 cheap version). Pure.
/// </summary>
internal static partial class VbaFormControlInventory
{
    [GeneratedRegex(@"^\s*(?:Public|Private|Friend)?\s*(?:Static\s+)?Sub\s+(?<ctrl>[A-Za-z_]\w*?)_(?<evt>Click|DblClick|Change|KeyPress|KeyDown|KeyUp|AfterUpdate|BeforeUpdate|Enter|Exit|MouseDown|MouseUp|MouseMove|GotFocus|LostFocus|DropButtonClick|SpinUp|SpinDown|Scroll|Initialize|Activate|Deactivate|Terminate|QueryClose|Layout|Resize|BeforeDragOver|BeforeDropOrPaste|Error|AddControl|RemoveControl|Zoom)\s*\(", RegexOptions.IgnoreCase)]
    private static partial Regex HandlerRegex();

    [GeneratedRegex(@"\bMe\.(?<ctrl>[A-Za-z_]\w*)(?:\.(?<prop>[A-Za-z_]\w*))?", RegexOptions.IgnoreCase)]
    private static partial Regex MeRefRegex();

    [GeneratedRegex(@"(?<![\w.])(?<ctrl>[A-Za-z_]\w*)\.(?<prop>[A-Za-z_]\w*)\b", RegexOptions.IgnoreCase)]
    private static partial Regex BareRefRegex();

    [GeneratedRegex(@"\b(?:Dim|Private|Public|Friend)\s+(?:WithEvents\s+)?(?<name>[A-Za-z_]\w*)\s+As\s+(?:MSForms\.)?(?<type>TextBox|CommandButton|ListBox|ComboBox|CheckBox|OptionButton|Label|Frame|Image|SpinButton|ScrollBar|MultiPage|TabStrip|ToggleButton|Control)\b", RegexOptions.IgnoreCase)]
    private static partial Regex DeclarationRegex();

    [GeneratedRegex(@"\b(?:Dim|Private|Public|Friend|Static|Const)\s+(?<name>[A-Za-z_]\w*)\b", RegexOptions.IgnoreCase)]
    private static partial Regex AnyDeclarationRegex();

    private static readonly StringComparer Ci = StringComparer.OrdinalIgnoreCase;

    // Hungarian prefixes seen in the corpus and the usual MSForms conventions.
    private static readonly (string Prefix, string Type)[] Prefixes =
    [
        ("txt", "TextBox"), ("txb", "TextBox"), ("cmd", "CommandButton"), ("btn", "CommandButton"),
        ("lst", "ListBox"), ("lbx", "ListBox"), ("cbo", "ComboBox"), ("cmb", "ComboBox"), ("ddl", "ComboBox"),
        ("chk", "CheckBox"), ("opt", "OptionButton"), ("rb", "OptionButton"), ("lbl", "Label"),
        ("fra", "Frame"), ("frm", "Frame"), ("img", "Image"), ("spn", "SpinButton"), ("scr", "ScrollBar"),
        ("mpg", "MultiPage"), ("tab", "TabStrip"), ("tgl", "ToggleButton"),
    ];

    private static readonly Dictionary<string, string> EventTypeHints = new(Ci)
    {
        ["DropButtonClick"] = "ComboBox",
        ["SpinUp"] = "SpinButton",
        ["SpinDown"] = "SpinButton",
        ["Scroll"] = "ScrollBar",
        ["Click"] = "CommandButton",
        ["Change"] = "TextBox",
        ["AfterUpdate"] = "TextBox",
        ["BeforeUpdate"] = "TextBox",
        ["KeyPress"] = "TextBox",
    };

    private static readonly Dictionary<string, string> MemberTypeHints = new(Ci)
    {
        ["AddItem"] = "ListBox", ["RemoveItem"] = "ListBox", ["ListIndex"] = "ListBox", ["ListCount"] = "ListBox",
        ["List"] = "ListBox", ["RowSource"] = "ListBox", ["Selected"] = "ListBox", ["MultiSelect"] = "ListBox",
        ["Caption"] = "Label", ["PasswordChar"] = "TextBox", ["MaxLength"] = "TextBox", ["Text"] = "TextBox",
        ["Picture"] = "Image", ["Pages"] = "MultiPage", ["Tabs"] = "TabStrip",
    };

    private sealed class Info
    {
        public string Name = "";
        public string? DeclaredType;
        public readonly SortedSet<string> Events = new(Ci);
        public readonly SortedSet<string> Properties = new(Ci);
        public bool SeenViaMeOrHandler;
    }

    public static ExcelVbaUserForm Analyze(ExcelVbaModule form)
    {
        var lines = VbaLineCleaner.Clean(form.Code ?? "");
        var controls = new Dictionary<string, Info>(Ci);
        var formEvents = new SortedSet<string>(Ci);
        var declaredNonControls = new HashSet<string>(Ci);
        int handlerCount = 0;

        Info Get(string name)
        {
            if (!controls.TryGetValue(name, out var info)) controls[name] = info = new Info { Name = name };
            return info;
        }

        foreach (var line in lines)
        {
            var text = line.Text;

            var decl = DeclarationRegex().Match(text);
            if (decl.Success)
            {
                var i = Get(decl.Groups["name"].Value);
                i.DeclaredType = Normalize(decl.Groups["type"].Value);
                i.SeenViaMeOrHandler = true;
                continue;
            }
            var anyDecl = AnyDeclarationRegex().Match(text);
            if (anyDecl.Success && !controls.ContainsKey(anyDecl.Groups["name"].Value))
                declaredNonControls.Add(anyDecl.Groups["name"].Value);

            var handler = HandlerRegex().Match(text);
            if (handler.Success)
            {
                handlerCount++;
                var ctrl = handler.Groups["ctrl"].Value;
                var evt = handler.Groups["evt"].Value;
                if (Ci.Equals(ctrl, "UserForm") || Ci.Equals(ctrl, form.Name))
                {
                    formEvents.Add(evt);
                }
                else
                {
                    var i = Get(ctrl);
                    i.Events.Add(evt);
                    i.SeenViaMeOrHandler = true;
                }
                continue;
            }

            foreach (Match m in MeRefRegex().Matches(text))
            {
                var name = m.Groups["ctrl"].Value;
                if (IsFormMember(name)) continue;
                var i = Get(name);
                i.SeenViaMeOrHandler = true;
                if (m.Groups["prop"].Success) i.Properties.Add(m.Groups["prop"].Value);
            }

            foreach (Match m in BareRefRegex().Matches(text))
            {
                var name = m.Groups["ctrl"].Value;
                if (Ci.Equals(name, "Me")) continue;
                if (declaredNonControls.Contains(name)) continue;
                var known = controls.ContainsKey(name);
                if (!known && PrefixType(name) is null) continue;   // an unknown bare name only counts with a Hungarian prefix
                Get(name).Properties.Add(m.Groups["prop"].Value);
            }
        }

        var result = controls.Values
            .Where(i => i.SeenViaMeOrHandler || PrefixType(i.Name) is not null)
            .OrderBy(i => i.Name, Ci)
            .Select(i =>
            {
                var (type, confidence) = InferType(i);
                return new ExcelVbaFormControl(i.Name, type, confidence, i.Events.ToList(), i.Properties.ToList());
            })
            .ToList();

        return new ExcelVbaUserForm(form.Name, result, formEvents.ToList(), handlerCount);
    }

    private static (string Type, string Confidence) InferType(Info i)
    {
        if (i.DeclaredType is not null) return (i.DeclaredType, "declared");
        if (PrefixType(i.Name) is { } byPrefix) return (byPrefix, "prefix");
        foreach (var evt in i.Events)
            if (EventTypeHints.TryGetValue(evt, out var byEvent))
            {
                // A ListBox/ComboBox also raises Click/Change; members decide when they disagree.
                var byMember = i.Properties.Select(p => MemberTypeHints.GetValueOrDefault(p)).FirstOrDefault(t => t is not null);
                return (byMember ?? byEvent, byMember is not null ? "member" : "event");
            }
        foreach (var p in i.Properties)
            if (MemberTypeHints.TryGetValue(p, out var t)) return (t, "member");
        return ("Control", "none");
    }

    private static string? PrefixType(string name)
    {
        foreach (var (prefix, type) in Prefixes)
            if (name.Length > prefix.Length && name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                && (char.IsUpper(name[prefix.Length]) || name[prefix.Length] == '_' || char.IsDigit(name[prefix.Length])))
                return type;
        return null;
    }

    private static string Normalize(string type) =>
        type.ToLowerInvariant() switch
        {
            "textbox" => "TextBox", "commandbutton" => "CommandButton", "listbox" => "ListBox", "combobox" => "ComboBox",
            "checkbox" => "CheckBox", "optionbutton" => "OptionButton", "label" => "Label", "frame" => "Frame",
            "image" => "Image", "spinbutton" => "SpinButton", "scrollbar" => "ScrollBar", "multipage" => "MultiPage",
            "tabstrip" => "TabStrip", "togglebutton" => "ToggleButton", _ => "Control",
        };

    // Members of the form itself, not controls on it.
    private static bool IsFormMember(string name) =>
        name is "Show" or "Hide" or "Caption" or "Controls" or "Name" or "Tag" or "Left" or "Top" or "Width" or "Height"
            or "Repaint" or "StartUpPosition" or "BackColor" or "ForeColor" or "Font" or "ActiveControl" or "Enabled"
            or "Visible" or "Zoom" or "ScrollTop" or "ScrollLeft" or "Move" or "PrintForm" or "Unload"
        || name.StartsWith("Show", StringComparison.OrdinalIgnoreCase) && name.Length == 4;
}
