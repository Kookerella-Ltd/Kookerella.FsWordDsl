using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Kookerella.CsWordDsl.Tests;

/// <summary>Shared fixtures/assertions used by both <c>DocumentTests</c> (behavioral
/// round-trip tests against throwaway temp files) and <c>ExampleTests</c> (which reload the
/// F# test suite's own checked-in <c>Examples/*/output.docx</c> fixtures).</summary>
internal static class TestHelpers
{
    public static string TempDocxPath() =>
        Path.Combine(Path.GetTempPath(), $"CsWordDslTest_{Guid.NewGuid():N}.docx");

    /// <summary>A minimal, valid 1x1 transparent PNG - real enough for a real embedded
    /// image part, tiny enough to inline here. Same fixture the F# suite uses.</summary>
    public static byte[] OnePixelPng() =>
        Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    public static void AssertSchemaValid(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        // Office2010, not the parameterless (Office2007) default - matches the F# test
        // suite's own reasoning: `w:tblLook`'s named boolean attributes, which real modern
        // Word itself writes, aren't valid under the older Office2007 transitional schema.
        var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2010);
        var errors = validator.Validate(document).ToList();

        Assert.True(errors.Count == 0, string.Join("\n", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }

    /// <summary>Where the F# test suite's own <c>Examples/</c> fixtures land in this
    /// project's own output directory - see this project's .csproj for how they're linked
    /// in rather than copy-pasted.</summary>
    public static string ExamplePath(string scenarioName, string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "Examples", scenarioName, fileName);
}
