namespace Kookerella.CsWordDsl;

/// <summary>WordprocessingML's <c>w:highlight</c> only accepts this fixed, enumerated
/// palette - unlike <see cref="Color"/>, arbitrary RGB is not valid here.</summary>
public enum HighlightColor
{
    Yellow,
    Green,
    Cyan,
    Magenta,
    Blue,
    Red,
    DarkBlue,
    DarkCyan,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    DarkGray,
    LightGray,
    Black
}
