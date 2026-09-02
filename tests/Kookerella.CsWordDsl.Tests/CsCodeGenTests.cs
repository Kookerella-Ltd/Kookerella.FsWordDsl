using System.Diagnostics;
using Xunit;
using static Kookerella.CsWordDsl.Tests.TestHelpers;

namespace Kookerella.CsWordDsl.Tests;

/// <summary>
/// Actually executes a <see cref="CsCodeGen"/>-generated file via <c>dotnet run --file</c>
/// (.NET's "file-based apps" feature) and reloads the result - not just generating the
/// source and trusting it compiles. Mirrors the F# test suite's own <c>Category=Slow</c>
/// group, which runs every generated <c>script.fsx</c> via <c>dotnet fsi</c> for the same
/// reason: generating source that merely *looks* plausible isn't the same guarantee as
/// generating source that actually runs and reproduces the file.
/// </summary>
public class CsCodeGenTests
{
    private static string WrapperCsprojPath()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "Kookerella.FsWordDsl.slnx")))
            dir = dir.Parent;

        if (dir is null)
            throw new InvalidOperationException($"Could not locate the repo root from {AppContext.BaseDirectory}");

        return Path.Combine(dir.FullName, "src", "Kookerella.CsWordDsl", "Kookerella.CsWordDsl.csproj");
    }

    [Fact]
    public void Generated_script_reproduces_the_document()
    {
        var doc = Document.Create(
            Section.Of([
                Block.Paragraph([new Inline.Run("Quarterly Report")], styleId: "Title"),
                Block.Paragraph([
                    new Inline.Run("This report covers "),
                    new Inline.Run("Q1 2026", new RunStyle { Bold = true }),
                    Inline.HyperlinkText("full dataset", new HyperlinkTarget.ExternalUrl("https://example.com/data"))
                ]),
                Block.Table(
                    [TableRow.Of([TableCell.Of([Block.Paragraph([new Inline.Run("A")])])], height: 20.0, repeatAsHeader: true)],
                    [100.0])
            ]));

        var outputFileName = TempDocxPath();
        var scriptPath = Path.Combine(Path.GetTempPath(), $"CsWordDslCodeGenTest_{Guid.NewGuid():N}.cs");
        var script = CsCodeGen.Generate([$"#:project {WrapperCsprojPath()}"], outputFileName, doc);
        File.WriteAllText(scriptPath, script);

        try
        {
            var psi = new ProcessStartInfo("dotnet", $"run --file \"{scriptPath}\"")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            };

            using var process = Process.Start(psi)!;
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit(120_000);

            Assert.True(process.ExitCode == 0, $"dotnet run --file failed:\n{stdout}\n{stderr}");
            Assert.True(File.Exists(outputFileName), $"Generated script did not produce {outputFileName}");

            var loaded = DocumentIO.Load(outputFileName);
            var body = loaded.Sections[0].Body;
            Assert.Equal("Quarterly Report", ((Inline.Run)((Block.ParagraphBlock)body[0]).Para.Inlines[0]).Text);
            Assert.True(((Inline.Run)((Block.ParagraphBlock)body[1]).Para.Inlines[1]).Style?.Bold);
            Assert.Equal("A", ((Inline.Run)((Block.ParagraphBlock)((Block.TableBlock)body[2]).Entry.Rows[0].Cells[0].Content[0]).Para.Inlines[0]).Text);
        }
        finally
        {
            File.Delete(scriptPath);
            if (File.Exists(outputFileName)) File.Delete(outputFileName);
        }
    }
}
