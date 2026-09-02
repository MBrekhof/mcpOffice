using McpOffice.Services.Excel.Vba;

namespace McpOffice.Tests.Excel.Vba;

public class VbaProcedureHasherTests
{
    private static IReadOnlyList<string> Body(string source)
    {
        var lines = VbaLineCleaner.Clean(source);
        var proc = Assert.Single(VbaProcedureScanner.Scan("standardModule", "M", lines));
        return VbaProcedureHasher.Normalize(lines, proc.CleanedLineStartIndex, proc.CleanedLineEndIndex);
    }

    [Fact]
    public void Normalize_drops_comments_blank_lines_case_and_extra_whitespace()
    {
        var body = Body("Sub A()\n    Dim   x As Long   ' note\n\n    X = 1\nEnd Sub");
        Assert.Equal(["dim x as long", "x = 1"], body);
    }

    [Fact]
    public void Hash_ignores_name_and_formatting_but_not_logic_or_literals()
    {
        var a = VbaProcedureHasher.Hash(Body("Sub Alpha()\n  x = 1\n  s = \"tab\"\nEnd Sub"));
        var b = VbaProcedureHasher.Hash(Body("Public Sub Beta()\nX = 1\n\nS   = \"tab\"\nEnd Sub"));
        var c = VbaProcedureHasher.Hash(Body("Sub Gamma()\n  x = 1\n  s = \"TAB\"\nEnd Sub"));   // literal case-folded too
        var d = VbaProcedureHasher.Hash(Body("Sub Delta()\n  x = 1\n  s = \"space\"\nEnd Sub"));
        Assert.Equal(a, b);
        Assert.Equal(a, c);
        Assert.NotEqual(a, d);
    }

    [Fact]
    public void Similarity_is_line_multiset_overlap()
    {
        string[] a = ["x = 1", "y = 2", "z = 3", "w = 4"];
        string[] b = ["x = 1", "y = 2", "z = 3", "w = 5"];
        Assert.Equal(0.75, VbaProcedureHasher.Similarity(a, b), 3);
        Assert.Equal(1.0, VbaProcedureHasher.Similarity(a, a));
        Assert.Equal(0.0, VbaProcedureHasher.Similarity(a, ["q = 9"]));
        Assert.Equal(1.0, VbaProcedureHasher.Similarity([], []));
    }
}
