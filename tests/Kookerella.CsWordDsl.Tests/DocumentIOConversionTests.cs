using System.Diagnostics;
using Xunit;
using Fs = Kookerella.FsWordDsl;
using static Kookerella.CsWordDsl.Tests.TestHelpers;

namespace Kookerella.CsWordDsl.Tests;

/// <summary>Covers the <see cref="DocumentIO"/> members beyond <c>Save</c>/<c>Load</c> -
/// <c>ToXml</c>/<c>FromXml</c>, <c>ToJson</c>/<c>FromJson</c>, and
/// <c>GenerateFSharpScript</c> - the C# wrapper's own entry points to the F# core's
/// <c>Xml.fs</c>/<c>Json.fs</c>/<c>Document.generateScript</c>, added so a C# caller never
/// needs to reference <c>Kookerella.FsWordDsl</c> directly to reach them (see
/// <see cref="DocumentIO"/>'s own doc comment).</summary>
public class DocumentIOConversionTests
{
    private static Document SampleDocument() =>
        Document.Create(
            Section.Of([
                Block.Paragraph([new Inline.Run("Quarterly Report")], styleId: "Title"),
                Block.Paragraph([
                    new Inline.Run("Plain, "),
                    new Inline.Run("bold.", new RunStyle { Bold = true })
                ])
            ]));

    [Fact]
    public void Xml_round_trips()
    {
        var doc = SampleDocument();

        var xml = DocumentIO.ToXml(doc);
        Assert.Contains("Quarterly Report", xml);

        var loaded = DocumentIO.FromXml(xml);
        var body = loaded.Sections[0].Body;
        Assert.Equal("Title", ((Block.ParagraphBlock)body[0]).Para.StyleId);
        Assert.Equal("Quarterly Report", ((Inline.Run)((Block.ParagraphBlock)body[0]).Para.Inlines[0]).Text);
        Assert.True(((Inline.Run)((Block.ParagraphBlock)body[1]).Para.Inlines[1]).Style?.Bold);
    }

    [Fact]
    public void Json_round_trips()
    {
        var doc = SampleDocument();

        var json = DocumentIO.ToJson(doc);
        Assert.Contains("Quarterly Report", json);

        var loaded = DocumentIO.FromJson(json);
        var body = loaded.Sections[0].Body;
        Assert.Equal("Title", ((Block.ParagraphBlock)body[0]).Para.StyleId);
        Assert.Equal("Quarterly Report", ((Inline.Run)((Block.ParagraphBlock)body[0]).Para.Inlines[0]).Text);
        Assert.True(((Inline.Run)((Block.ParagraphBlock)body[1]).Para.Inlines[1]).Style?.Bold);
    }

    /// <summary>Actually executes the generated script via <c>dotnet fsi</c> rather than
    /// just generating source and trusting it compiles - same discipline
    /// <see cref="CsCodeGenTests"/> applies to <c>CsCodeGen.Generate</c>, and the F# test
    /// suite's own <c>Category=Slow</c> group applies to every committed example script.
    /// </summary>
    [Fact]
    public void Generated_fsharp_script_reproduces_the_document()
    {
        var doc = SampleDocument();
        var outputFileName = TempDocxPath();
        var scriptPath = Path.Combine(Path.GetTempPath(), $"CsWordDslFsScriptTest_{Guid.NewGuid():N}.fsx");

        var referenceLines = new[]
        {
            $"#r \"{typeof(Fs.Model.Document).Assembly.Location.Replace("\\", "\\\\")}\"",
            $"#r \"{typeof(DocumentFormat.OpenXml.Wordprocessing.Paragraph).Assembly.Location.Replace("\\", "\\\\")}\""
        };

        var script = DocumentIO.GenerateFSharpScript(referenceLines, outputFileName, doc);
        File.WriteAllText(scriptPath, script);

        try
        {
            var psi = new ProcessStartInfo("dotnet", $"fsi \"{scriptPath}\"")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            };

            using var process = Process.Start(psi)!;
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit(120_000);

            Assert.True(process.ExitCode == 0, $"dotnet fsi failed:\n{stdout}\n{stderr}");
            Assert.True(File.Exists(outputFileName), $"Generated script did not produce {outputFileName}");

            var loaded = DocumentIO.Load(outputFileName);
            var body = loaded.Sections[0].Body;
            Assert.Equal("Quarterly Report", ((Inline.Run)((Block.ParagraphBlock)body[0]).Para.Inlines[0]).Text);
            Assert.True(((Inline.Run)((Block.ParagraphBlock)body[1]).Para.Inlines[1]).Style?.Bold);
        }
        finally
        {
            File.Delete(scriptPath);
            if (File.Exists(outputFileName)) File.Delete(outputFileName);
        }
    }
}
