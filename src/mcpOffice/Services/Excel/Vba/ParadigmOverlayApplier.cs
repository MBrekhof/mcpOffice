using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class ParadigmOverlayApplier
{
    public static readonly IReadOnlyList<string> SupportedParadigms = new[]
    {
        "classLibrary", "workerService", "webApi", "console"
    };

    public static CSharpSuggestion Apply(
        string module,
        string procedureName,
        string? scope,
        ProcedureAxes axes,
        string paradigm)
    {
        var className = StripModulePrefix(module);
        var methodName = ToPascalCase(procedureName);
        var isPublic = !string.Equals(scope, "Private", StringComparison.OrdinalIgnoreCase);

        return paradigm switch
        {
            "classLibrary" => ApplyClassLibrary(className, methodName, isPublic, axes),
            "workerService" => ApplyWorkerService(className, methodName, isPublic, axes),
            "webApi" => ApplyWebApi(className, methodName, isPublic, axes),
            "console" => ApplyConsole(className, methodName, isPublic, axes),
            _ => throw new ArgumentException($"Unsupported paradigm '{paradigm}'", nameof(paradigm))
        };
    }

    private static CSharpSuggestion ApplyClassLibrary(string c, string m, bool pub, ProcedureAxes axes)
    {
        if (axes.Trigger == "eventHandler")
            return new("requiresManualReview", c, m, null, pub,
                new[] { "event_handler_no_pure_classlib_target" });

        bool depsEmpty = axes.Dependencies.Count == 0;
        bool excelOnly = axes.Dependencies.Count == 1 && axes.Dependencies[0] == "excelObjectModel";
        bool hasDbOrNet = axes.Dependencies.Any(d => d == "database" || d == "network");

        if (axes.Purity == "pure" && axes.Shape == "leaf")
            return new("staticMethod", c, m, "static", pub, Array.Empty<string>());

        if ((axes.Purity == "pure" || axes.Purity == "readsState") && depsEmpty)
            return new("staticMethod", c, m, "static", pub, Array.Empty<string>());

        if (axes.Purity == "sideEffectful" && hasDbOrNet)
            return new("instanceMethod", c, m, "scoped", pub,
                new[] { "requires_external_dependency_injection" });

        if (axes.Purity == "writesState" && excelOnly)
            return new("instanceMethod", c, m, "scoped", pub,
                new[] { "depends_on_excel_object_model" });

        // Catch-all: instance method, conservative.
        return new("instanceMethod", c, m, "scoped", pub, Array.Empty<string>());
    }

    private static CSharpSuggestion ApplyWorkerService(string c, string m, bool pub, ProcedureAxes axes)
    {
        bool isBackgroundEntry =
            (axes.Trigger == "macroEntryPoint" &&
             (axes.Purity == "writesState" || axes.Purity == "sideEffectful")) ||
            (axes.Trigger == "eventHandler" && IsBackgroundEventName(m));

        var blockers = new List<string>();
        if (axes.Dependencies.Contains("excelObjectModel"))
            blockers.Add("depends_on_excel_object_model");

        if (isBackgroundEntry)
            return new("backgroundService", c, m, "singleton", pub, blockers);

        return new("instanceMethod", c, m, "scoped", pub, blockers);
    }

    private static bool IsBackgroundEventName(string method) =>
        string.Equals(method, "WorkbookOpen", StringComparison.Ordinal) ||
        string.Equals(method, "AutoOpen", StringComparison.Ordinal) ||
        method.Contains("OnTime", StringComparison.Ordinal);

    private static CSharpSuggestion ApplyWebApi(string c, string m, bool pub, ProcedureAxes axes)
    {
        var blockers = new List<string>();
        if (axes.Dependencies.Contains("excelObjectModel"))
            blockers.Add("depends_on_excel_object_model");

        if (axes.Trigger == "macroEntryPoint" && pub)
            return new("apiAction", c, m, "scoped", pub, blockers);

        return new("instanceMethod", c, m, "scoped", pub, blockers);
    }

    private static CSharpSuggestion ApplyConsole(string c, string m, bool pub, ProcedureAxes axes)
    {
        if (axes.Trigger == "macroEntryPoint")
            return new("consoleEntryPoint", c, m, null, pub, Array.Empty<string>());

        bool isStaticCandidate = axes.Purity == "pure" || axes.Purity == "readsState";
        return isStaticCandidate
            ? new("staticMethod", c, m, "static", pub, Array.Empty<string>())
            : new("instanceMethod", c, m, "scoped", pub, Array.Empty<string>());
    }

    private static string StripModulePrefix(string moduleName)
    {
        // Hungarian module prefixes seen in the wild: mod/mdl/bas for standard modules
        // (Air.xlsm uses mdl), cls for class modules, frm for UserForms.
        foreach (var prefix in new[] { "mod", "mdl", "bas", "cls", "frm" })
        {
            if (moduleName.Length > prefix.Length &&
                moduleName.StartsWith(prefix, StringComparison.Ordinal) &&
                char.IsUpper(moduleName[prefix.Length]))
            {
                return moduleName[prefix.Length..];
            }
        }
        return ToPascalCase(moduleName);
    }

    private static string ToPascalCase(string identifier)
    {
        if (string.IsNullOrEmpty(identifier)) return identifier;
        var parts = identifier.Split('_', StringSplitOptions.RemoveEmptyEntries);
        return string.Concat(parts.Select(p =>
            char.ToUpperInvariant(p[0]) + (p.Length > 1 ? p[1..] : "")));
    }
}
