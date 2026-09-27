using Xunit;
using OfficeAi.Shared;

public class OfficeLanguageTests
{
    [Fact]
    public void ResolveUiLanguage_HebrewLcid_ReturnsHe()
    {
        Assert.Equal("he", OfficeLanguage.ResolveUiLanguage(1037));
    }

    [Fact]
    public void ResolveUiLanguage_EnglishUsLcid_ReturnsEn()
    {
        Assert.Equal("en", OfficeLanguage.ResolveUiLanguage(1033));
    }

    [Fact]
    public void ResolveUiLanguage_UnsupportedLanguage_ReturnsEn()
    {
        // French (1036) - not one of this panel's two supported UI
        // languages, so it degrades to the safe default rather than
        // throwing or returning something the UI can't render.
        Assert.Equal("en", OfficeLanguage.ResolveUiLanguage(1036));
    }

    [Fact]
    public void ResolveUiLanguage_ZeroOrNegativeLcid_ReturnsEn()
    {
        Assert.Equal("en", OfficeLanguage.ResolveUiLanguage(0));
        Assert.Equal("en", OfficeLanguage.ResolveUiLanguage(-1));
    }

    [Fact]
    public void ResolveBrandName_HebrewLcid_ReturnsTransliteration()
    {
        Assert.Equal("אופן דוקס", OfficeLanguage.ResolveBrandName(1037));
    }

    [Fact]
    public void ResolveBrandName_NonHebrewLcid_ReturnsLatinName()
    {
        Assert.Equal("OpenDocs", OfficeLanguage.ResolveBrandName(1033));
    }

    [Fact]
    public void ResolveBrandName_FailureSentinel_ReturnsLatinName()
    {
        // 0 is what the guarded caller (ThisAddIn.GetOfficeUiLanguageId) is
        // expected to return when the LanguageSettings COM call throws -
        // must degrade to the safe default, not an unrenderable value.
        Assert.Equal("OpenDocs", OfficeLanguage.ResolveBrandName(0));
    }
}
