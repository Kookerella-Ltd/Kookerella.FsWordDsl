namespace Kookerella.CsWordDsl;

/// <summary><see cref="Width"/> is in points; <see langword="null"/> uses Word's own
/// default weight.</summary>
public sealed record BorderSide(BorderLineStyle Style, double? Width = null, Color? Color = null);
