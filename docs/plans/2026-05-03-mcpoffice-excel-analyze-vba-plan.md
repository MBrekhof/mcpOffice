# excel_analyze_vba Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Ship `excel_analyze_vba` — an MCP tool that layers structural analysis (procedures, event handlers, call graph, object-model references, dependencies) on top of `excel_extract_vba`'s raw VBA source. Single tool, tiered output via opt-in toggles, ~24th tool overall.

**Architecture:** New `VbaSourceAnalyzer` composed of four single-purpose components — `VbaLineCleaner` (comment/string/continuation handling), `VbaProcedureScanner` (boundaries + signatures), `VbaCallGraphBuilder` (callee resolution), `VbaReferenceCollector` (object-model + dependency dispatch). Tool method calls existing `VbaProjectReader.Read` then runs the analyzer. No new external dependencies.

**Tech Stack:** .NET 9 · existing OpenMcdf/cp1252 stack from `excel_extract_vba` · xUnit · no third-party VBA parsing libs.

**Reference design:** `docs/plans/2026-05-03-mcpoffice-excel-analyze-vba-design.md` — single source of truth for tool surface, output DTO, parsing strategy, error model, enum vocabularies. Read it before starting.

---

## Conventions used in this plan

- All paths are relative to `C:\Projects\mcpOffice\`.
- Each task is one TDD micro-cycle: write failing test → run → implement → run → commit.
- "Run tests" means: `dotnet test --nologo --filter "FullyQualifiedName~<TestClassName>"` from repo root for the targeted suite, then `dotnet test --nologo` once per task to confirm the whole suite stays green.
- Conventional Commits: `feat:`, `test:`, `chore:`, `docs:`.
- After every task: `dotnet build` is 0 warnings / 0 errors AND every test passes. If either fails, stop and fix before moving on (per superpowers:verification-before-completion).
- All new files live under `src/mcpOffice/Services/Excel/Vba/` (production) and `tests/mcpOffice.Tests/Excel/Vba/` (unit).

---

# Phase 1 — DTOs

### Task 1: Add public analysis DTOs

**Files:**
- Create: `src/mcpOffice/Models/ExcelVbaAnalysis.cs`
- Create: `src/mcpOffice/Models/ExcelVbaAnalysisSummary.cs`
- Create: `src/mcpOffice/Models/ExcelVbaModuleAnalysis.cs`
- Create: `src/mcpOffice/Models/ExcelVbaProcedure.cs`
- Create: `src/mcpOffice/Models/ExcelVbaParameter.cs`
- Create: `src/mcpOffice/Models/ExcelVbaCallEdge.cs`
- Create: `src/mcpOffice/Models/ExcelVbaSiteRef.cs`
- Create: `src/mcpOffice/Models/ExcelVbaReferences.cs`
- Create: `src/mcpOffice/Models/ExcelVbaObjectModelRef.cs`
- Create: `src/mcpOffice/Models/ExcelVbaDependency.cs`

**Step 1: Write the records (one per file, file-scoped namespace `McpOffice.Models`)**

```csharp
// ExcelVbaAnalysis.cs
namespace McpOffice.Models;
public sealed record ExcelVbaAnalysis(
    bool HasVbaProject,
    ExcelVbaAnalysisSummary Summary,
    IReadOnlyList<ExcelVbaModuleAnalysis>? Modules,
    IReadOnlyList<ExcelVbaCallEdge>? CallGraph,
    ExcelVbaReferences? References);
```

```csharp
// ExcelVbaAnalysisSummary.cs
namespace McpOffice.Models;
public sealed record ExcelVbaAnalysisSummary(
    int ModuleCount,
    int ParsedModuleCount,
    int UnparsedModuleCount,
    int ProcedureCount,
    int EventHandlerCount,
    int CallEdgeCount,
    int ObjectModelReferenceCount,
    int DependencyCount);
```

```csharp
// ExcelVbaModuleAnalysis.cs
namespace McpOffice.Models;
public sealed record ExcelVbaModuleAnalysis(
    string Name,
    string Kind,                                  // "standardModule" | "classModule" | "documentModule" | "userForm"
    bool Parsed,
    string? Reason,                               // populated when Parsed = false
    IReadOnlyList<ExcelVbaProcedure> Procedures);
```

```csharp
// ExcelVbaProcedure.cs
namespace McpOffice.Models;
public sealed record ExcelVbaProcedure(
    string Name,
    string FullyQualifiedName,                    // "<module>.<name>"
    string Kind,                                  // "Sub" | "Function" | "PropertyGet" | "PropertyLet" | "PropertySet"
    string? Scope,                                // "Public" | "Private" | "Friend" | null
    IReadOnlyList<ExcelVbaParameter> Parameters,
    string? ReturnType,
    int LineStart,                                // 1-based inclusive
    int LineEnd,                                  // 1-based inclusive
    bool IsEventHandler,
    string? EventTarget);
```

```csharp
// ExcelVbaParameter.cs
namespace McpOffice.Models;
public sealed record ExcelVbaParameter(
    string Name,
    string? Type,
    bool ByRef,
    bool Optional,
    string? DefaultValue);
```

```csharp
// ExcelVbaCallEdge.cs
namespace McpOffice.Models;
public sealed record ExcelVbaCallEdge(
    string From,                                  // procedure FullyQualifiedName
    string To,                                    // FQN if resolved, bare name or "<dynamic>" otherwise
    bool Resolved,
    ExcelVbaSiteRef Site);
```

```csharp
// ExcelVbaSiteRef.cs
namespace McpOffice.Models;
public sealed record ExcelVbaSiteRef(
    string Module,
    string Procedure,
    int Line);
```

```csharp
// ExcelVbaReferences.cs
namespace McpOffice.Models;
public sealed record ExcelVbaReferences(
    IReadOnlyList<ExcelVbaObjectModelRef> ObjectModel,
    IReadOnlyList<ExcelVbaDependency> Dependencies);
```

```csharp
// ExcelVbaObjectModelRef.cs
namespace McpOffice.Models;
public sealed record ExcelVbaObjectModelRef(
    string Module,
    string Procedure,
    int Line,
    string Api,                                   // "Worksheets" | "Sheets" | "Range" | "Cells" | "ActiveSheet" | ...
    string? Literal);                             // first string-literal arg when present
```

```csharp
// ExcelVbaDependency.cs
namespace McpOffice.Models;
public sealed record ExcelVbaDependency(
    string Module,
    string Procedure,
    int Line,
    string Kind,                                  // "file" | "database" | "network" | "automation" | "shell"
    string? Target,                               // ProgID / path / URL when literal
    string? Operation);                           // best-effort verb, e.g. "Open", "Kill"
```

**Step 2: Build to verify they compile**

Run: `dotnet build --nologo`
Expected: 0 warnings / 0 errors.

**Step 3: Commit**

```bash
git add src/mcpOffice/Models/ExcelVba*.cs
git commit -m "feat: DTOs for excel_analyze_vba (ExcelVbaAnalysis et al.)"
```

---

# Phase 2 — Line cleaner

### Task 2: `VbaLineCleaner` — strip comments, sentinel strings, fold continuations

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaLineCleaner.cs`
- Create: `src/mcpOffice/Services/Excel/Vba/CleanedLine.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaLineCleanerTests.cs`

**Step 1: Write the failing tests**

```csharp
// tests/mcpOffice.Tests/Excel/Vba/VbaLineCleanerTests.cs
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaLineCleanerTests
{
    [Fact]
    public void Strips_apostrophe_comment()
    {
        var lines = VbaLineCleaner.Clean("x = 1 ' set x");
        Assert.Single(lines);
        Assert.Equal("x = 1", lines[0].Text.TrimEnd());
        Assert.Equal(1, lines[0].LineNumber);
    }

    [Fact]
    public void Apostrophe_inside_string_is_not_a_comment()
    {
        var lines = VbaLineCleaner.Clean("s = \"isn't a comment\"");
        Assert.Single(lines);
        Assert.Contains("<STR>", lines[0].Text);
        Assert.DoesNotContain("isn't", lines[0].Text);
    }

    [Fact]
    public void Doubled_quote_escape_inside_string()
    {
        var lines = VbaLineCleaner.Clean("s = \"he said \"\"hi\"\"\"");
        Assert.Single(lines);
        Assert.Contains("<STR>", lines[0].Text);
        // No stray apostrophe / unbalanced quote in cleaned text:
        Assert.DoesNotContain("he said", lines[0].Text);
    }

    [Fact]
    public void Rem_statement_is_treated_as_comment()
    {
        var lines = VbaLineCleaner.Clean("Rem this is a comment");
        Assert.Single(lines);
        Assert.Equal("", lines[0].Text.Trim());
    }

    [Fact]
    public void Folds_underscore_continuation()
    {
        var src = "Sub Foo(x As Long, _\r\n            y As Long)";
        var lines = VbaLineCleaner.Clean(src);
        Assert.Single(lines);
        Assert.Contains("Sub Foo(x As Long,", lines[0].Text);
        Assert.Contains("y As Long)", lines[0].Text);
        Assert.Equal(1, lines[0].LineNumber); // start line preserved
    }

    [Fact]
    public void Preserves_originalText_for_string_literal_capture()
    {
        var lines = VbaLineCleaner.Clean("Set ws = Worksheets(\"Data\")");
        Assert.Single(lines);
        Assert.Contains("\"Data\"", lines[0].OriginalText);
        Assert.Contains("<STR>", lines[0].Text);
    }
}
```

**Step 2: Run — fail (no class yet)**

Run: `dotnet test --nologo --filter "FullyQualifiedName~VbaLineCleanerTests"`
Expected: build error / FAIL: type does not exist.

**Step 3: Implement**

```csharp
// src/mcpOffice/Services/Excel/Vba/CleanedLine.cs
namespace McpOffice.Services.Excel.Vba;

internal sealed record CleanedLine(int LineNumber, string Text, string OriginalText);
```

```csharp
// src/mcpOffice/Services/Excel/Vba/VbaLineCleaner.cs
using System.Text;

namespace McpOffice.Services.Excel.Vba;

internal static class VbaLineCleaner
{
    private const string StringSentinel = "<STR>";

    public static IReadOnlyList<CleanedLine> Clean(string source)
    {
        if (string.IsNullOrEmpty(source)) return [];
        var rawLines = source.Replace("\r\n", "\n").Split('\n');
        var result = new List<CleanedLine>(rawLines.Length);

        var pending = new StringBuilder();
        var pendingOriginal = new StringBuilder();
        int? pendingStart = null;

        for (int i = 0; i < rawLines.Length; i++)
        {
            var raw = rawLines[i];
            var cleaned = CleanSingleLine(raw);

            if (pending.Length == 0)
            {
                pendingStart = i + 1;
            }

            // _ continuation = trailing whitespace + underscore (after cleaning may not end with _ if comment stripped)
            var endsWithContinuation = EndsWithContinuation(raw);

            if (endsWithContinuation)
            {
                // Strip the trailing _ from the cleaned text and the original
                pending.Append(StripTrailingContinuation(cleaned));
                pending.Append(' ');
                pendingOriginal.Append(StripTrailingContinuation(raw));
                pendingOriginal.Append(' ');
                continue;
            }

            pending.Append(cleaned);
            pendingOriginal.Append(raw);

            result.Add(new CleanedLine(pendingStart ?? (i + 1), pending.ToString(), pendingOriginal.ToString()));
            pending.Clear();
            pendingOriginal.Clear();
            pendingStart = null;
        }

        // Any dangling pending (file ended on a continuation) — flush as-is.
        if (pending.Length > 0)
        {
            result.Add(new CleanedLine(pendingStart ?? rawLines.Length, pending.ToString(), pendingOriginal.ToString()));
        }

        return result;
    }

    private static string CleanSingleLine(string raw)
    {
        // Handle the leading "Rem " comment-statement form (case-insensitive, must be first non-whitespace token).
        var trimmed = raw.TrimStart();
        if (trimmed.Length >= 4 &&
            (trimmed.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase) ||
             trimmed.Equals("Rem", StringComparison.OrdinalIgnoreCase)))
        {
            return new string(' ', raw.Length - trimmed.Length); // preserve leading whitespace
        }

        var sb = new StringBuilder(raw.Length);
        bool inString = false;
        for (int i = 0; i < raw.Length; i++)
        {
            char c = raw[i];

            if (inString)
            {
                if (c == '"')
                {
                    // Doubled "" inside string = escaped quote, stay in string.
                    if (i + 1 < raw.Length && raw[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    inString = false;
                    sb.Append('"').Append(StringSentinel).Append('"');
                    sb.Length -= StringSentinel.Length + 2; // we'll append properly below
                    sb.Append('"');
                }
                continue;
            }

            if (c == '"')
            {
                inString = true;
                sb.Append('"').Append(StringSentinel).Append('"');
                // Skip past the entire string in raw to avoid double processing.
                int j = i + 1;
                while (j < raw.Length)
                {
                    if (raw[j] == '"')
                    {
                        if (j + 1 < raw.Length && raw[j + 1] == '"') { j += 2; continue; }
                        break;
                    }
                    j++;
                }
                i = j; // i++ at loop end takes us past closing quote
                inString = false;
                continue;
            }

            if (c == '\'') return sb.ToString(); // comment to end of line

            sb.Append(c);
        }
        return sb.ToString();
    }

    private static bool EndsWithContinuation(string raw)
    {
        // VBA continuation: <whitespace>_  at end of line, outside any string.
        // Cheap check: trim trailing whitespace, last char is '_' AND char before is whitespace.
        var trimmed = raw.TrimEnd();
        if (trimmed.Length < 2) return false;
        if (trimmed[^1] != '_') return false;
        return char.IsWhiteSpace(trimmed[^2]);
    }

    private static string StripTrailingContinuation(string s)
    {
        var trimmed = s.TrimEnd();
        // Drop the final '_' character.
        return trimmed[..^1];
    }
}
```

**Step 4: Run — pass**

Run: `dotnet test --nologo --filter "FullyQualifiedName~VbaLineCleanerTests"`
Expected: all 6 tests PASS.

**Step 5: Run the full suite to catch any regression**

Run: `dotnet test --nologo`
Expected: still green (existing 86 passed / 1 skipped + 6 new = 92 / 1 skipped).

**Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaLineCleaner.cs src/mcpOffice/Services/Excel/Vba/CleanedLine.cs tests/mcpOffice.Tests/Excel/Vba/VbaLineCleanerTests.cs
git commit -m "feat: VbaLineCleaner — comment/string/continuation handling"
```

---

# Phase 3 — Procedure scanner

### Task 3: `VbaProcedureScanner` — boundaries, signatures, parameters, event detection

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaProcedureScanner.cs`
- Create: `src/mcpOffice/Services/Excel/Vba/ScannedProcedure.cs` (internal — carries the cleaned-line range so later passes can scan only the body)
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaProcedureScannerTests.cs`

**Step 1: Write failing tests**

```csharp
// tests/mcpOffice.Tests/Excel/Vba/VbaProcedureScannerTests.cs
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaProcedureScannerTests
{
    private static IReadOnlyList<ScannedProcedure> Scan(string moduleKind, string moduleName, string source)
    {
        var lines = VbaLineCleaner.Clean(source);
        return VbaProcedureScanner.Scan(moduleKind, moduleName, lines);
    }

    [Fact]
    public void Detects_simple_sub()
    {
        var procs = Scan("standardModule", "Module1",
            "Public Sub DoIt()\nEnd Sub");
        Assert.Single(procs);
        Assert.Equal("DoIt", procs[0].Procedure.Name);
        Assert.Equal("Module1.DoIt", procs[0].Procedure.FullyQualifiedName);
        Assert.Equal("Sub", procs[0].Procedure.Kind);
        Assert.Equal("Public", procs[0].Procedure.Scope);
    }

    [Fact]
    public void Detects_function_with_return_type()
    {
        var procs = Scan("standardModule", "Module1",
            "Function Add(a As Long, b As Long) As Long\nAdd = a + b\nEnd Function");
        var p = procs.Single().Procedure;
        Assert.Equal("Function", p.Kind);
        Assert.Equal("Long", p.ReturnType);
        Assert.Equal(2, p.Parameters.Count);
        Assert.Equal("a", p.Parameters[0].Name);
        Assert.Equal("Long", p.Parameters[0].Type);
    }

    [Fact]
    public void Detects_property_get()
    {
        var procs = Scan("classModule", "MyClass",
            "Public Property Get Name() As String\nEnd Property");
        Assert.Equal("PropertyGet", procs.Single().Procedure.Kind);
    }

    [Fact]
    public void Parses_optional_byval_with_default()
    {
        var procs = Scan("standardModule", "M",
            "Sub F(Optional ByVal x As String = \"d\")\nEnd Sub");
        var p = procs.Single().Procedure.Parameters.Single();
        Assert.True(p.Optional);
        Assert.False(p.ByRef);
        Assert.Equal("x", p.Name);
        Assert.Equal("String", p.Type);
        Assert.NotNull(p.DefaultValue);
    }

    [Fact]
    public void Detects_event_handler_in_document_module()
    {
        var procs = Scan("documentModule", "ThisWorkbook",
            "Private Sub Workbook_Open()\nEnd Sub");
        var p = procs.Single().Procedure;
        Assert.True(p.IsEventHandler);
        Assert.Equal("Workbook", p.EventTarget);
    }

    [Fact]
    public void Standard_module_with_underscore_name_is_not_event_handler()
    {
        var procs = Scan("standardModule", "Utils",
            "Sub Foo_Bar()\nEnd Sub");
        Assert.False(procs.Single().Procedure.IsEventHandler);
    }

    [Fact]
    public void Records_line_range()
    {
        var procs = Scan("standardModule", "M",
            "Sub A()\nx = 1\nEnd Sub");
        var p = procs.Single().Procedure;
        Assert.Equal(1, p.LineStart);
        Assert.Equal(3, p.LineEnd);
    }

    [Fact]
    public void Defaults_scope_to_null_when_unspecified()
    {
        var procs = Scan("standardModule", "M", "Sub A()\nEnd Sub");
        Assert.Null(procs.Single().Procedure.Scope);
    }

    [Fact]
    public void Multiple_procedures()
    {
        var procs = Scan("standardModule", "M",
            "Sub A()\nEnd Sub\n\nSub B()\nEnd Sub");
        Assert.Equal(2, procs.Count);
        Assert.Equal("A", procs[0].Procedure.Name);
        Assert.Equal("B", procs[1].Procedure.Name);
    }
}
```

**Step 2: Run — fail**

**Step 3: Implement**

```csharp
// src/mcpOffice/Services/Excel/Vba/ScannedProcedure.cs
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal sealed record ScannedProcedure(
    ExcelVbaProcedure Procedure,
    int CleanedLineStartIndex,    // index into the CleanedLine list (inclusive, after the Sub/Function header)
    int CleanedLineEndIndex);     // index into the CleanedLine list (inclusive, line before End Sub/Function)
```

```csharp
// src/mcpOffice/Services/Excel/Vba/VbaProcedureScanner.cs
using System.Text.RegularExpressions;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static partial class VbaProcedureScanner
{
    [GeneratedRegex(@"^\s*(?<scope>Public|Private|Friend)?\s*(Static\s+)?(?<kind>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+(?<name>\w+)\s*\((?<params>[^)]*)\)(\s+As\s+(?<ret>\w+))?", RegexOptions.IgnoreCase)]
    private static partial Regex ProcOpenRegex();

    [GeneratedRegex(@"^\s*End\s+(Sub|Function|Property)\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex ProcCloseRegex();

    public static IReadOnlyList<ScannedProcedure> Scan(string moduleKind, string moduleName, IReadOnlyList<CleanedLine> lines)
    {
        var procs = new List<ScannedProcedure>();

        for (int i = 0; i < lines.Count; i++)
        {
            var open = ProcOpenRegex().Match(lines[i].Text);
            if (!open.Success) continue;

            int startLine = lines[i].LineNumber;
            int bodyStartIdx = i + 1;
            int closeIdx = -1;
            for (int j = i + 1; j < lines.Count; j++)
            {
                if (ProcCloseRegex().IsMatch(lines[j].Text)) { closeIdx = j; break; }
            }
            int endLine = closeIdx >= 0 ? lines[closeIdx].LineNumber : lines[^1].LineNumber;

            var name = open.Groups["name"].Value;
            var kindRaw = open.Groups["kind"].Value;
            var kind = NormalizeKind(kindRaw);
            var scope = open.Groups["scope"].Success ? open.Groups["scope"].Value : null;
            var returnType = open.Groups["ret"].Success ? open.Groups["ret"].Value : null;
            var paramList = ParseParameters(open.Groups["params"].Value);

            var (isEvent, target) = ClassifyEventHandler(moduleKind, name);

            var proc = new ExcelVbaProcedure(
                Name: name,
                FullyQualifiedName: $"{moduleName}.{name}",
                Kind: kind,
                Scope: scope,
                Parameters: paramList,
                ReturnType: returnType,
                LineStart: startLine,
                LineEnd: endLine,
                IsEventHandler: isEvent,
                EventTarget: target);

            procs.Add(new ScannedProcedure(proc, bodyStartIdx, closeIdx >= 0 ? closeIdx - 1 : lines.Count - 1));
            i = closeIdx >= 0 ? closeIdx : lines.Count;
        }

        return procs;
    }

    private static string NormalizeKind(string raw)
    {
        var collapsed = Regex.Replace(raw, @"\s+", "");
        return collapsed switch
        {
            "Sub" => "Sub",
            "Function" => "Function",
            "PropertyGet" => "PropertyGet",
            "PropertyLet" => "PropertyLet",
            "PropertySet" => "PropertySet",
            _ => collapsed
        };
    }

    private static IReadOnlyList<ExcelVbaParameter> ParseParameters(string paramText)
    {
        if (string.IsNullOrWhiteSpace(paramText)) return [];
        var parts = SplitOnTopLevelCommas(paramText);
        var list = new List<ExcelVbaParameter>(parts.Count);
        foreach (var part in parts)
        {
            var p = part.Trim();
            if (p.Length == 0) continue;

            bool optional = false, byRef = true; // VBA default is ByRef
            // Strip leading modifiers
            var modPattern = @"^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?";
            var m = Regex.Match(p, modPattern, RegexOptions.IgnoreCase);
            if (m.Success)
            {
                if (m.Value.IndexOf("Optional", StringComparison.OrdinalIgnoreCase) >= 0) optional = true;
                if (m.Value.IndexOf("ByVal", StringComparison.OrdinalIgnoreCase) >= 0) byRef = false;
                p = p[m.Length..];
            }

            string name;
            string? type = null;
            string? defaultValue = null;

            // Split on " = " for default value
            var eq = FindTopLevelEquals(p);
            if (eq >= 0)
            {
                defaultValue = p[(eq + 1)..].Trim();
                p = p[..eq].Trim();
            }

            var asIdx = Regex.Match(p, @"\s+As\s+", RegexOptions.IgnoreCase);
            if (asIdx.Success)
            {
                name = p[..asIdx.Index].Trim();
                type = p[(asIdx.Index + asIdx.Length)..].Trim();
            }
            else
            {
                name = p.Trim();
            }

            list.Add(new ExcelVbaParameter(name, type, byRef, optional, defaultValue));
        }
        return list;
    }

    private static List<string> SplitOnTopLevelCommas(string s)
    {
        var parts = new List<string>();
        int depth = 0;
        int start = 0;
        for (int i = 0; i < s.Length; i++)
        {
            var c = s[i];
            if (c == '(' || c == '[') depth++;
            else if (c == ')' || c == ']') depth--;
            else if (c == ',' && depth == 0)
            {
                parts.Add(s[start..i]);
                start = i + 1;
            }
        }
        parts.Add(s[start..]);
        return parts;
    }

    private static int FindTopLevelEquals(string s)
    {
        int depth = 0;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '(' || s[i] == '[') depth++;
            else if (s[i] == ')' || s[i] == ']') depth--;
            else if (s[i] == '=' && depth == 0) return i;
        }
        return -1;
    }

    private static (bool IsEvent, string? Target) ClassifyEventHandler(string moduleKind, string name)
    {
        if (moduleKind == "standardModule") return (false, null);
        var idx = name.IndexOf('_');
        if (idx <= 0 || idx == name.Length - 1) return (false, null);
        return (true, name[..idx]);
    }
}
```

**Step 4: Run — pass**

Run: `dotnet test --nologo --filter "FullyQualifiedName~VbaProcedureScannerTests"`
Expected: all 9 tests PASS.

**Step 5: Full suite green**

Run: `dotnet test --nologo`

**Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaProcedureScanner.cs src/mcpOffice/Services/Excel/Vba/ScannedProcedure.cs tests/mcpOffice.Tests/Excel/Vba/VbaProcedureScannerTests.cs
git commit -m "feat: VbaProcedureScanner — boundaries, signatures, event handler detection"
```

---

# Phase 4 — Call graph

### Task 4: `VbaCallGraphBuilder` — direct calls + Application.Run, resolved against the procedure index

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaCallGraphBuilder.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaCallGraphBuilderTests.cs`

**Step 1: Failing tests**

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaCallGraphBuilderTests
{
    private static IReadOnlyList<ExcelVbaCallEdge> Build(params (string moduleName, string moduleKind, string source)[] modules)
    {
        var scanned = new List<(string ModuleName, IReadOnlyList<CleanedLine> Lines, IReadOnlyList<ScannedProcedure> Procs)>();
        foreach (var (n, k, s) in modules)
        {
            var lines = VbaLineCleaner.Clean(s);
            var procs = VbaProcedureScanner.Scan(k, n, lines);
            scanned.Add((n, lines, procs));
        }
        return VbaCallGraphBuilder.Build(scanned);
    }

    [Fact]
    public void Resolves_direct_call_within_module()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nB\nEnd Sub\nSub B()\nEnd Sub"));
        var edge = Assert.Single(edges);
        Assert.Equal("M.A", edge.From);
        Assert.Equal("M.B", edge.To);
        Assert.True(edge.Resolved);
    }

    [Fact]
    public void Resolves_call_keyword_form()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nCall B\nEnd Sub\nSub B()\nEnd Sub"));
        Assert.Single(edges);
        Assert.True(edges[0].Resolved);
    }

    [Fact]
    public void Resolves_cross_module_call()
    {
        var edges = Build(
            ("Caller", "standardModule", "Sub A()\nDoLog\nEnd Sub"),
            ("Utils", "standardModule", "Sub DoLog()\nEnd Sub"));
        var edge = Assert.Single(edges);
        Assert.Equal("Caller.A", edge.From);
        Assert.Equal("Utils.DoLog", edge.To);
        Assert.True(edge.Resolved);
    }

    [Fact]
    public void Captures_application_run_as_dynamic_unresolved()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nApplication.Run \"OtherWb.xlsm!Foo\"\nEnd Sub"));
        var edge = Assert.Single(edges);
        Assert.False(edge.Resolved);
        Assert.Equal("OtherWb.xlsm!Foo", edge.To);
    }

    [Fact]
    public void Skips_vba_keywords_and_string_sentinels()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nIf x Then\nDim y As Long\nEnd If\nEnd Sub"));
        Assert.Empty(edges);
    }

    [Fact]
    public void Records_call_site_line()
    {
        var edges = Build(("M", "standardModule",
            "Sub A()\nx = 1\nB\nEnd Sub\nSub B()\nEnd Sub"));
        Assert.Equal(3, edges[0].Site.Line);
    }
}
```

**Step 2: Implement**

```csharp
// src/mcpOffice/Services/Excel/Vba/VbaCallGraphBuilder.cs
using System.Text.RegularExpressions;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static partial class VbaCallGraphBuilder
{
    // Reserved words / control-flow tokens that look like a procedure name but aren't a call.
    private static readonly HashSet<string> Keywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "If", "Then", "Else", "ElseIf", "End", "Sub", "Function", "Property",
        "Dim", "Set", "Const", "Public", "Private", "Friend", "Static",
        "For", "Next", "While", "Wend", "Do", "Loop", "Until", "Each", "In",
        "Select", "Case", "With", "Exit", "GoTo", "On", "Error", "Resume",
        "True", "False", "Nothing", "Null", "Empty", "And", "Or", "Not", "Xor", "Mod",
        "Is", "Like", "Optional", "ByVal", "ByRef", "ParamArray", "As",
        "Return", "Stop", "Rem", "Type", "Enum", "Declare", "Lib", "Alias",
        "Me", "Application"  // Application.Run is handled separately
    };

    [GeneratedRegex(@"^\s*(Call\s+)?(?<name>[A-Za-z_]\w*)\s*(\(|$)", RegexOptions.IgnoreCase)]
    private static partial Regex DirectCallRegex();

    [GeneratedRegex(@"Application\s*\.\s*Run\s+""(?<target>[^""]+)""", RegexOptions.IgnoreCase)]
    private static partial Regex AppRunRegex();

    [GeneratedRegex(@"^\s*[A-Za-z_]\w*\s*=", RegexOptions.IgnoreCase)]
    private static partial Regex AssignmentRegex();

    public static IReadOnlyList<ExcelVbaCallEdge> Build(
        IReadOnlyList<(string ModuleName, IReadOnlyList<CleanedLine> Lines, IReadOnlyList<ScannedProcedure> Procs)> modules)
    {
        // Build the procedure index: bare-name + FQN.
        var byBareName = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        var allFqn = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (_, _, procs) in modules)
        {
            foreach (var sp in procs)
            {
                allFqn.Add(sp.Procedure.FullyQualifiedName);
                if (!byBareName.TryGetValue(sp.Procedure.Name, out var list))
                {
                    list = [];
                    byBareName[sp.Procedure.Name] = list;
                }
                list.Add(sp.Procedure.FullyQualifiedName);
            }
        }

        var edges = new List<ExcelVbaCallEdge>();
        foreach (var (moduleName, lines, procs) in modules)
        {
            foreach (var sp in procs)
            {
                for (int i = sp.CleanedLineStartIndex; i <= sp.CleanedLineEndIndex && i < lines.Count; i++)
                {
                    var line = lines[i];

                    // Application.Run "Foo"
                    foreach (Match m in AppRunRegex().Matches(line.Text))
                    {
                        edges.Add(new ExcelVbaCallEdge(
                            From: sp.Procedure.FullyQualifiedName,
                            To: m.Groups["target"].Value,
                            Resolved: false,
                            Site: new ExcelVbaSiteRef(moduleName, sp.Procedure.Name, line.LineNumber)));
                    }

                    // Skip lines that are obviously assignments
                    if (AssignmentRegex().IsMatch(line.Text)) continue;

                    var dc = DirectCallRegex().Match(line.Text);
                    if (!dc.Success) continue;

                    var name = dc.Groups["name"].Value;
                    if (Keywords.Contains(name)) continue;
                    // Skip self-name (the procedure header line was already excluded by [start, end] body window)
                    if (string.Equals(name, sp.Procedure.Name, StringComparison.OrdinalIgnoreCase)) continue;

                    string to;
                    bool resolved;
                    if (byBareName.TryGetValue(name, out var fqns))
                    {
                        // Prefer same-module match; otherwise first match
                        var sameMod = fqns.FirstOrDefault(f => f.StartsWith(moduleName + ".", StringComparison.OrdinalIgnoreCase));
                        to = sameMod ?? fqns[0];
                        resolved = true;
                    }
                    else
                    {
                        to = name;
                        resolved = false;
                    }

                    edges.Add(new ExcelVbaCallEdge(
                        From: sp.Procedure.FullyQualifiedName,
                        To: to,
                        Resolved: resolved,
                        Site: new ExcelVbaSiteRef(moduleName, sp.Procedure.Name, line.LineNumber)));
                }
            }
        }
        return edges;
    }
}
```

**Step 3-5: Run failing → implement → run passing → full suite green**

**Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaCallGraphBuilder.cs tests/mcpOffice.Tests/Excel/Vba/VbaCallGraphBuilderTests.cs
git commit -m "feat: VbaCallGraphBuilder — direct + Application.Run edges with FQN resolution"
```

---

# Phase 5 — Reference collector

### Task 5: `VbaReferenceCollector` — object-model APIs + dependency dispatch

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaReferenceCollector.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaReferenceCollectorTests.cs`

**Step 1: Failing tests** (covers Worksheets/Range literal capture, file Open, ADODB → database, MSXML → network, CreateObject fallback to automation, Shell)

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaReferenceCollectorTests
{
    private static (IReadOnlyList<ExcelVbaObjectModelRef> Om, IReadOnlyList<ExcelVbaDependency> Deps) Collect(
        string moduleName, string moduleKind, string source)
    {
        var lines = VbaLineCleaner.Clean(source);
        var procs = VbaProcedureScanner.Scan(moduleKind, moduleName, lines);
        var om = new List<ExcelVbaObjectModelRef>();
        var deps = new List<ExcelVbaDependency>();
        VbaReferenceCollector.Collect(moduleName, lines, procs, om, deps);
        return (om, deps);
    }

    [Fact]
    public void Captures_worksheets_with_literal()
    {
        var (om, _) = Collect("M", "standardModule",
            "Sub A()\nSet ws = Worksheets(\"Data\")\nEnd Sub");
        var r = Assert.Single(om);
        Assert.Equal("Worksheets", r.Api);
        Assert.Equal("Data", r.Literal);
    }

    [Fact]
    public void Captures_range()
    {
        var (om, _) = Collect("M", "standardModule",
            "Sub A()\nRange(\"A1:B10\").Value = 0\nEnd Sub");
        Assert.Equal("Range", om.Single().Api);
        Assert.Equal("A1:B10", om.Single().Literal);
    }

    [Fact]
    public void File_open_classified_as_file()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nOpen \"C:\\f.txt\" For Input As #1\nEnd Sub");
        var d = Assert.Single(deps);
        Assert.Equal("file", d.Kind);
    }

    [Fact]
    public void Adodb_classified_as_database()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nSet c = CreateObject(\"ADODB.Connection\")\nEnd Sub");
        var d = Assert.Single(deps);
        Assert.Equal("database", d.Kind);
        Assert.Equal("ADODB.Connection", d.Target);
    }

    [Fact]
    public void Msxml_classified_as_network()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nSet h = CreateObject(\"MSXML2.XMLHTTP\")\nEnd Sub");
        Assert.Equal("network", deps.Single().Kind);
    }

    [Fact]
    public void Outlook_falls_back_to_automation()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nSet o = CreateObject(\"Outlook.Application\")\nEnd Sub");
        var d = Assert.Single(deps);
        Assert.Equal("automation", d.Kind);
        Assert.Equal("Outlook.Application", d.Target);
    }

    [Fact]
    public void Shell_classified_as_shell()
    {
        var (_, deps) = Collect("M", "standardModule",
            "Sub A()\nShell(\"notepad.exe\")\nEnd Sub");
        Assert.Equal("shell", deps.Single().Kind);
    }
}
```

**Step 2: Implement**

```csharp
// src/mcpOffice/Services/Excel/Vba/VbaReferenceCollector.cs
using System.Text.RegularExpressions;
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static partial class VbaReferenceCollector
{
    private static readonly string[] ObjectModelApis =
    [
        "Worksheets", "Sheets", "Range", "Cells",
        "ActiveSheet", "ActiveWorkbook", "ThisWorkbook",
        "Application", "Selection", "Names"
    ];

    [GeneratedRegex(@"\b(?<api>Worksheets|Sheets|Range|Cells|ActiveSheet|ActiveWorkbook|ThisWorkbook|Application|Selection|Names)\b\s*\(\s*""(?<lit>[^""]*)""", RegexOptions.IgnoreCase)]
    private static partial Regex OmWithLiteralRegex();

    [GeneratedRegex(@"\b(?<api>Worksheets|Sheets|Range|Cells|ActiveSheet|ActiveWorkbook|ThisWorkbook|Application|Selection|Names)\b", RegexOptions.IgnoreCase)]
    private static partial Regex OmAnyRegex();

    [GeneratedRegex(@"^\s*Open\s+", RegexOptions.IgnoreCase)]
    private static partial Regex FileOpenRegex();

    [GeneratedRegex(@"\b(Kill|MkDir|RmDir|ChDir|Dir|FileSystemObject|Workbooks\.Open|Workbooks\.OpenText|Scripting\.FileSystemObject)\b", RegexOptions.IgnoreCase)]
    private static partial Regex FileApiRegex();

    [GeneratedRegex(@"(?:CreateObject|GetObject)\s*\(\s*""(?<progid>[^""]+)""", RegexOptions.IgnoreCase)]
    private static partial Regex CreateGetObjectRegex();

    [GeneratedRegex(@"\b(ADODB\.|DAO\.|OpenDatabase)\b", RegexOptions.IgnoreCase)]
    private static partial Regex DatabaseApiRegex();

    [GeneratedRegex(@"\b(MSXML2\.XMLHTTP|WinHttp\.WinHttpRequest|URLDownloadToFile|InternetExplorer\.Application)\b", RegexOptions.IgnoreCase)]
    private static partial Regex NetworkApiRegex();

    [GeneratedRegex(@"^\s*Shell\s*\(", RegexOptions.IgnoreCase)]
    private static partial Regex ShellRegex();

    public static void Collect(
        string moduleName,
        IReadOnlyList<CleanedLine> lines,
        IReadOnlyList<ScannedProcedure> procs,
        List<ExcelVbaObjectModelRef> objectModelOut,
        List<ExcelVbaDependency> dependenciesOut)
    {
        foreach (var sp in procs)
        {
            for (int i = sp.CleanedLineStartIndex; i <= sp.CleanedLineEndIndex && i < lines.Count; i++)
            {
                var line = lines[i];

                CollectObjectModel(moduleName, sp.Procedure.Name, line, objectModelOut);
                CollectDependencies(moduleName, sp.Procedure.Name, line, dependenciesOut);
            }
        }
    }

    private static void CollectObjectModel(string module, string proc, CleanedLine line, List<ExcelVbaObjectModelRef> sink)
    {
        var seenAt = new HashSet<int>();

        // First pass: APIs invoked with a string-literal arg — capture from OriginalText.
        foreach (Match m in OmWithLiteralRegex().Matches(line.OriginalText))
        {
            sink.Add(new ExcelVbaObjectModelRef(module, proc, line.LineNumber, m.Groups["api"].Value, m.Groups["lit"].Value));
            seenAt.Add(m.Index);
        }

        // Second pass: bare references (no literal arg). De-dupe by start position to avoid double-counting.
        foreach (Match m in OmAnyRegex().Matches(line.OriginalText))
        {
            if (seenAt.Contains(m.Index)) continue;
            sink.Add(new ExcelVbaObjectModelRef(module, proc, line.LineNumber, m.Groups["api"].Value, null));
        }
    }

    private static void CollectDependencies(string module, string proc, CleanedLine line, List<ExcelVbaDependency> sink)
    {
        // Shell builtin (statement form, leading)
        if (ShellRegex().IsMatch(line.Text))
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "shell", null, "Shell"));
            return;
        }

        // File: Open ... For ...
        if (FileOpenRegex().IsMatch(line.Text))
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "file", null, "Open"));
            return;
        }

        var fileM = FileApiRegex().Match(line.Text);
        if (fileM.Success)
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "file", null, fileM.Groups[1].Value));
            return;
        }

        // CreateObject / GetObject — dispatch by ProgID.
        var co = CreateGetObjectRegex().Match(line.OriginalText);
        if (co.Success)
        {
            var progId = co.Groups["progid"].Value;
            string kind = ClassifyProgId(progId);
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, kind, progId, "CreateObject"));
            return;
        }

        if (DatabaseApiRegex().IsMatch(line.Text))
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "database", null, null));
            return;
        }
        if (NetworkApiRegex().IsMatch(line.Text))
        {
            sink.Add(new ExcelVbaDependency(module, proc, line.LineNumber, "network", null, null));
            return;
        }
    }

    private static string ClassifyProgId(string progId)
    {
        if (progId.StartsWith("ADODB.", StringComparison.OrdinalIgnoreCase) ||
            progId.StartsWith("DAO.", StringComparison.OrdinalIgnoreCase)) return "database";
        if (progId.StartsWith("MSXML2.", StringComparison.OrdinalIgnoreCase) ||
            progId.StartsWith("WinHttp.", StringComparison.OrdinalIgnoreCase) ||
            progId.Equals("InternetExplorer.Application", StringComparison.OrdinalIgnoreCase)) return "network";
        if (progId.Contains("FileSystemObject", StringComparison.OrdinalIgnoreCase)) return "file";
        return "automation";
    }
}
```

**Step 3-5: Run → fail → implement → pass → full suite**

**Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaReferenceCollector.cs tests/mcpOffice.Tests/Excel/Vba/VbaReferenceCollectorTests.cs
git commit -m "feat: VbaReferenceCollector — object-model APIs + dependency dispatch"
```

---

# Phase 6 — Orchestrator

### Task 6: `VbaSourceAnalyzer` — compose cleaner / scanner / call graph / refs into the public DTO

**Files:**
- Create: `src/mcpOffice/Services/Excel/Vba/VbaSourceAnalyzer.cs`
- Create: `tests/mcpOffice.Tests/Excel/Vba/VbaSourceAnalyzerTests.cs`

**Step 1: Failing tests**

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaSourceAnalyzerTests
{
    private static ExcelVbaProject Project(params (string Name, string Kind, string Code)[] modules) =>
        new(true, modules.Select(m => new ExcelVbaModule(m.Name, m.Kind, m.Code.Split('\n').Length, m.Code)).ToList());

    [Fact]
    public void HasVbaProject_false_returns_zeroed_summary()
    {
        var result = VbaSourceAnalyzer.Analyze(
            new ExcelVbaProject(false, []), includeProcedures: true, includeCallGraph: true, includeReferences: true);
        Assert.False(result.HasVbaProject);
        Assert.Equal(0, result.Summary.ModuleCount);
        Assert.Null(result.Modules);  // not present when no project
        Assert.Null(result.CallGraph);
        Assert.Null(result.References);
    }

    [Fact]
    public void Summary_counts_match_collections()
    {
        var p = Project(("Util", "standardModule", "Sub Log()\nEnd Sub\nSub Warn()\nLog\nEnd Sub"));
        var r = VbaSourceAnalyzer.Analyze(p, true, true, true);
        Assert.Equal(1, r.Summary.ModuleCount);
        Assert.Equal(1, r.Summary.ParsedModuleCount);
        Assert.Equal(0, r.Summary.UnparsedModuleCount);
        Assert.Equal(2, r.Summary.ProcedureCount);
        Assert.Equal(1, r.Summary.CallEdgeCount);
    }

    [Fact]
    public void Toggles_omit_collections()
    {
        var p = Project(("M", "standardModule", "Sub A()\nEnd Sub"));
        var r = VbaSourceAnalyzer.Analyze(p, includeProcedures: false, includeCallGraph: false, includeReferences: false);
        Assert.Null(r.Modules);
        Assert.Null(r.CallGraph);
        Assert.Null(r.References);
        // Summary still populated — analysis runs internally to compute counts.
        Assert.Equal(1, r.Summary.ProcedureCount);
    }

    [Fact]
    public void Module_too_large_marked_unparsed()
    {
        var bigSource = string.Join("\n", Enumerable.Repeat("x = 1", 5001));
        var p = Project(("Big", "standardModule", "Sub A()\n" + bigSource + "\nEnd Sub"));
        var r = VbaSourceAnalyzer.Analyze(p, true, false, false);
        Assert.Single(r.Modules!);
        Assert.False(r.Modules![0].Parsed);
        Assert.Equal("module_too_large", r.Modules[0].Reason);
        Assert.Equal(1, r.Summary.UnparsedModuleCount);
    }

    [Fact]
    public void Event_handler_count_in_summary()
    {
        var p = Project(("ThisWorkbook", "documentModule",
            "Private Sub Workbook_Open()\nEnd Sub"));
        var r = VbaSourceAnalyzer.Analyze(p, true, false, false);
        Assert.Equal(1, r.Summary.EventHandlerCount);
    }
}
```

**Step 2: Implement**

```csharp
// src/mcpOffice/Services/Excel/Vba/VbaSourceAnalyzer.cs
using McpOffice.Models;

namespace McpOffice.Services.Excel.Vba;

internal static class VbaSourceAnalyzer
{
    private const int MaxLinesPerModule = 5000;

    public static ExcelVbaAnalysis Analyze(
        ExcelVbaProject project,
        bool includeProcedures,
        bool includeCallGraph,
        bool includeReferences)
    {
        if (!project.HasVbaProject)
        {
            return new ExcelVbaAnalysis(
                HasVbaProject: false,
                Summary: new ExcelVbaAnalysisSummary(0, 0, 0, 0, 0, 0, 0, 0),
                Modules: null,
                CallGraph: null,
                References: null);
        }

        var perModule = new List<(string ModuleName, string ModuleKind, IReadOnlyList<CleanedLine> Lines, IReadOnlyList<ScannedProcedure> Procs, bool Parsed, string? Reason)>();
        var moduleAnalyses = new List<ExcelVbaModuleAnalysis>(project.Modules.Count);

        foreach (var m in project.Modules)
        {
            if (string.IsNullOrEmpty(m.Code))
            {
                moduleAnalyses.Add(new ExcelVbaModuleAnalysis(m.Name, m.Kind, false, "empty_source", []));
                perModule.Add((m.Name, m.Kind, [], [], false, "empty_source"));
                continue;
            }

            var cleaned = VbaLineCleaner.Clean(m.Code);
            if (cleaned.Count > MaxLinesPerModule)
            {
                moduleAnalyses.Add(new ExcelVbaModuleAnalysis(m.Name, m.Kind, false, "module_too_large", []));
                perModule.Add((m.Name, m.Kind, cleaned, [], false, "module_too_large"));
                continue;
            }

            var procs = VbaProcedureScanner.Scan(m.Kind, m.Name, cleaned);
            moduleAnalyses.Add(new ExcelVbaModuleAnalysis(
                m.Name, m.Kind, true, null, procs.Select(sp => sp.Procedure).ToList()));
            perModule.Add((m.Name, m.Kind, cleaned, procs, true, null));
        }

        // Call graph + references always built (cheap relative to extraction); we just decide whether to expose.
        var callModules = perModule.Where(p => p.Parsed)
            .Select(p => (p.ModuleName, p.Lines, p.Procs)).ToList();
        var edges = VbaCallGraphBuilder.Build(callModules);

        var omRefs = new List<ExcelVbaObjectModelRef>();
        var deps = new List<ExcelVbaDependency>();
        foreach (var (moduleName, _, lines, procs, parsed, _) in perModule)
        {
            if (!parsed) continue;
            VbaReferenceCollector.Collect(moduleName, lines, procs, omRefs, deps);
        }

        var procedureCount = moduleAnalyses.Sum(m => m.Procedures.Count);
        var eventHandlerCount = moduleAnalyses.Sum(m => m.Procedures.Count(p => p.IsEventHandler));
        var summary = new ExcelVbaAnalysisSummary(
            ModuleCount: project.Modules.Count,
            ParsedModuleCount: moduleAnalyses.Count(m => m.Parsed),
            UnparsedModuleCount: moduleAnalyses.Count(m => !m.Parsed),
            ProcedureCount: procedureCount,
            EventHandlerCount: eventHandlerCount,
            CallEdgeCount: edges.Count,
            ObjectModelReferenceCount: omRefs.Count,
            DependencyCount: deps.Count);

        return new ExcelVbaAnalysis(
            HasVbaProject: true,
            Summary: summary,
            Modules: includeProcedures ? moduleAnalyses : null,
            CallGraph: includeCallGraph ? edges : null,
            References: includeReferences ? new ExcelVbaReferences(omRefs, deps) : null);
    }
}
```

**Step 3-5: Run failing → implement → pass → full suite**

**Step 6: Commit**

```bash
git add src/mcpOffice/Services/Excel/Vba/VbaSourceAnalyzer.cs tests/mcpOffice.Tests/Excel/Vba/VbaSourceAnalyzerTests.cs
git commit -m "feat: VbaSourceAnalyzer orchestrator — composes cleaner/scanner/callgraph/refs"
```

---

# Phase 7 — Service + tool wiring

### Task 7: Add `AnalyzeVba` to `IExcelWorkbookService` + `ExcelWorkbookService`

**Files:**
- Modify: `src/mcpOffice/Services/Excel/IExcelWorkbookService.cs`
- Modify: `src/mcpOffice/Services/Excel/ExcelWorkbookService.cs`

**Step 1: Extend the interface**

```csharp
// add to IExcelWorkbookService
ExcelVbaAnalysis AnalyzeVba(
    string path,
    bool includeProcedures,
    bool includeCallGraph,
    bool includeReferences);
```

**Step 2: Implement on the service**

```csharp
// add to ExcelWorkbookService — reuse the existing VbaProjectReader instance creation pattern from ExtractVba
public ExcelVbaAnalysis AnalyzeVba(
    string path,
    bool includeProcedures,
    bool includeCallGraph,
    bool includeReferences)
{
    PathGuard.RequireExists(path);
    var reader = new VbaProjectReader();
    var project = reader.Read(path);  // throws vba_project_locked / vba_parse_error as needed
    return VbaSourceAnalyzer.Analyze(project, includeProcedures, includeCallGraph, includeReferences);
}
```

(If `ExtractVba` constructs `VbaProjectReader` differently, mirror that pattern exactly.)

**Step 3: Build to confirm it compiles**

Run: `dotnet build --nologo`
Expected: 0 warnings / 0 errors. The MCP tool method doesn't exist yet — no integration test will fail at this point.

**Step 4: Commit**

```bash
git add src/mcpOffice/Services/Excel/IExcelWorkbookService.cs src/mcpOffice/Services/Excel/ExcelWorkbookService.cs
git commit -m "feat: IExcelWorkbookService.AnalyzeVba"
```

---

### Task 8: Add the `excel_analyze_vba` MCP tool method

**Files:**
- Modify: `src/mcpOffice/Tools/ExcelTools.cs` — append after `ExcelGetStructure`

**Step 1: Add the method**

```csharp
[McpServerTool(Name = "excel_analyze_vba")]
[Description("Layers structural analysis on top of excel_extract_vba's source: procedures with signatures, event handlers, call graph, Excel object-model references (Worksheets/Range/Cells/...), and external dependencies (file/database/network/automation/shell). Tiered output via toggles. Returns hasVbaProject=false (with zeroed summary) for workbooks without a VBA project.")]
public static object ExcelAnalyzeVba(
    [Description("Absolute path to the .xlsm/.xlsb workbook")] string path,
    [Description("Include the per-module procedure list. Default true.")] bool includeProcedures = true,
    [Description("Include the call graph edges. Default false (medium cost).")] bool includeCallGraph = false,
    [Description("Include object-model and dependency references. Default false (heaviest output).")] bool includeReferences = false)
    => Service.AnalyzeVba(path, includeProcedures, includeCallGraph, includeReferences);
```

**Step 2: Build**

Run: `dotnet build --nologo`
Expected: 0 warnings / 0 errors.

**Step 3: Commit**

```bash
git add src/mcpOffice/Tools/ExcelTools.cs
git commit -m "feat: excel_analyze_vba MCP tool"
```

---

### Task 9: Update `ToolSurfaceTests` to include the new tool

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs:8-33` — add `"excel_analyze_vba"` to the expected array (alphabetic position: between `"excel_"` entries — first).

**Step 1: Add it**

```csharp
string[] expected =
[
    "excel_analyze_vba",
    "excel_extract_vba",
    // ... rest unchanged
];
```

**Step 2: Run the tool surface test**

Run: `dotnet test tests/mcpOffice.Tests.Integration --nologo --filter "FullyQualifiedName~ToolSurfaceTests"`
Expected: PASS — tool count is now 24.

**Step 3: Commit**

```bash
git add tests/mcpOffice.Tests.Integration/ToolSurfaceTests.cs
git commit -m "test: include excel_analyze_vba in tool surface assertion"
```

---

# Phase 8 — Integration coverage

### Task 10: Add stdio integration test against the synthetic fixture

**Files:**
- Modify: `tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs` — append a new test method.

**Step 1: Add the test**

```csharp
[Fact]
public async Task Analyze_vba_via_stdio_returns_summary()
{
    var fixture = ResolveFixturePath("sample-with-macros.xlsm");
    if (!File.Exists(fixture)) return;  // synthetic fixture optional; same skip pattern as Extract_vba_via_stdio

    await using var harness = await ServerHarness.StartAsync();
    var result = await harness.Client.CallToolAsync(
        "excel_analyze_vba",
        new Dictionary<string, object?>
        {
            ["path"] = fixture,
            ["includeProcedures"] = true,
            ["includeCallGraph"] = true,
            ["includeReferences"] = true
        });

    var text = result.Content.OfType<ModelContextProtocol.Protocol.TextContentBlock>().Single().Text;
    Assert.Contains("\"hasVbaProject\":true", text);
    Assert.Contains("\"summary\":", text);
    Assert.Contains("\"modules\":", text);
}
```

**Step 2-3: Run + commit**

Run: `dotnet test tests/mcpOffice.Tests.Integration --nologo --filter "FullyQualifiedName~Analyze_vba"`
Expected: PASS.

```bash
git add tests/mcpOffice.Tests.Integration/ExcelWorkflowTests.cs
git commit -m "test: stdio integration coverage for excel_analyze_vba"
```

---

### Task 11: Add real-world benchmark against `Air.xlsm`

**Files:**
- Create: `tests/mcpOffice.Tests/Excel/Vba/AirSampleAnalysisTests.cs`

**Step 1: Add the gated test**

```csharp
using McpOffice.Models;
using McpOffice.Services.Excel;

namespace McpOffice.Tests.Excel.Vba;

public class AirSampleAnalysisTests
{
    private const string SamplePath = @"C:\Projects\mcpOffice-samples\Air.xlsm";

    [Fact]
    public void Analyzes_real_air_workbook_without_exceptions()
    {
        if (!File.Exists(SamplePath)) return;  // gracefully no-op on machines without the sample

        var svc = new ExcelWorkbookService();
        var analysis = svc.AnalyzeVba(SamplePath,
            includeProcedures: true, includeCallGraph: true, includeReferences: true);

        Assert.True(analysis.HasVbaProject);
        Assert.NotNull(analysis.Modules);
        Assert.NotNull(analysis.CallGraph);
        Assert.NotNull(analysis.References);

        // Every module is either parsed or carries a reason — never both null.
        foreach (var m in analysis.Modules!)
            Assert.True(m.Parsed || m.Reason is not null);

        // Plausible floors for 107 modules of real macro code.
        Assert.True(analysis.Summary.ProcedureCount > 50,
            $"expected > 50 procedures, got {analysis.Summary.ProcedureCount}");
        Assert.NotEmpty(analysis.CallGraph!);
        Assert.Contains(analysis.References!.ObjectModel, r => r.Api == "Worksheets");
        Assert.Contains(analysis.References!.ObjectModel, r => r.Api == "Range");
    }
}
```

**Step 2: Run**

Run: `dotnet test tests/mcpOffice.Tests --nologo --filter "FullyQualifiedName~AirSampleAnalysisTests"`
Expected: PASS on machines that have `Air.xlsm`; silent no-op elsewhere.

**Step 3: Commit**

```bash
git add tests/mcpOffice.Tests/Excel/Vba/AirSampleAnalysisTests.cs
git commit -m "test: real-world benchmark against Air.xlsm (gated on file existence)"
```

---

# Phase 9 — Final verification

### Task 12: Full suite green + docs update

**Step 1: Build + test**

```bash
dotnet build --nologo
dotnet test --nologo
```

Expected: 0 warnings / 0 errors. Tests: pre-existing 86 + ~30 new unit + 1 new integration ≈ 117 passed / 1 skipped (locked-VBA fixture still). 11 integration tests (10 + 1 new).

**Step 2: Update `TODO.md`**

- Mark `excel_analyze_vba` as DONE.
- Move "Spike file `tests/mcpOffice.Tests/Spikes/VbaExtractionSpike.cs` is historical reference; consider removing once `excel_analyze_vba` lands" into a now-actionable item.

**Step 3: Update `SESSION_HANDOFF.md`**

- Record the new tool (24 total: 1 Ping + 15 Word + 8 Excel).
- Note any surprises observed against `Air.xlsm` (procedure count, edge count, dependency counts) — these become the v2 (`excel_analyze_vba` conversion hints) starting evidence.

**Step 4: Commit + push**

```bash
git add TODO.md SESSION_HANDOFF.md
git commit -m "docs: handoff after excel_analyze_vba lands"
git push origin main
```

---

## What this plan deliberately does NOT do

- No conversion hints (v2 work, designed against actual `Air.xlsm` evidence).
- No moduleName / procedureName filters on the tool — whole-workbook only for v1.
- No async / parallel module analysis. Sequential is fine for the expected payload sizes.
- No userForm layout introspection. Event handlers on form controls still parse via name pattern.
- No timeout guard. The 5,000-line per-module cap is the only pathological-input safeguard.
- No spike-file deletion as part of this plan — it's a clean-up follow-up after the new tool is in agent hands.

## Risks called out

1. **Regex false matches on call graph.** The keyword list is conservative; if real `Air.xlsm` produces noise (e.g., type names matching a procedure name), the `resolved` flag and the keyword list are the right places to tighten. Don't expand into a real parser without clear evidence.
2. **`Application.Run "X!Y"` won't be cross-resolved.** Captured as unresolved; agent can decide. This is documented in the design.
3. **The 5,000-line cap is a guess.** If a real legitimate module exceeds it, raise the cap rather than splitting its parsing — but keep the sentinel so a runaway never hangs the analyzer.
