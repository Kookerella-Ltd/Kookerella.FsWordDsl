namespace Kookerella.CsWordDsl;

public enum PageOrientation
{
    Portrait,
    Landscape
}

/// <summary>How a section begins relative to the previous one - meaningless, and not
/// written, for the very first section.</summary>
public enum SectionBreakType
{
    NextPage,
    Continuous,
    EvenPage,
    OddPage
}
