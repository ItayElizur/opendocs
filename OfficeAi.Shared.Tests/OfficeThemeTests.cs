using Xunit;
using OfficeAi.Shared;

public class OfficeThemeTests
{
    private const string OfficeThemeKey = @"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common";
    private const string OfficeThemeValue = "UI Theme";
    private const string PersonalizeKey = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
    private const string PersonalizeValue = "AppsUseLightTheme";

    // ---- fake registry: fixed responses per (keyName, valueName), also
    // pins the exact two registry paths this method is expected to read ----
    private static System.Func<string, string, object, object> Fake(object officeThemeValue, object personalizeValue = null)
    {
        return (keyName, valueName, defaultValue) =>
        {
            if (keyName == OfficeThemeKey && valueName == OfficeThemeValue) return officeThemeValue;
            if (keyName == PersonalizeKey && valueName == PersonalizeValue) return personalizeValue;
            Assert.Fail("Unexpected registry read: " + keyName + " / " + valueName);
            return null;
        };
    }

    [Fact]
    public void ReadEffectiveTheme_MissingOfficeKey_ReturnsLight()
    {
        Assert.Equal("light", OfficeTheme.ReadEffectiveTheme(Fake(null)));
    }

    [Fact]
    public void ReadEffectiveTheme_DarkGray_ReturnsDark()
    {
        Assert.Equal("dark", OfficeTheme.ReadEffectiveTheme(Fake(3)));
    }

    [Fact]
    public void ReadEffectiveTheme_Black_ReturnsDark()
    {
        Assert.Equal("dark", OfficeTheme.ReadEffectiveTheme(Fake(4)));
    }

    [Fact]
    public void ReadEffectiveTheme_White_ReturnsLight()
    {
        Assert.Equal("light", OfficeTheme.ReadEffectiveTheme(Fake(5)));
    }

    [Fact]
    public void ReadEffectiveTheme_Colorful_ReturnsLight()
    {
        Assert.Equal("light", OfficeTheme.ReadEffectiveTheme(Fake(7)));
    }

    [Fact]
    public void ReadEffectiveTheme_UseSystemSetting_WithAppsUseLightThemeZero_ReturnsDark()
    {
        Assert.Equal("dark", OfficeTheme.ReadEffectiveTheme(Fake(6, 0)));
    }

    [Fact]
    public void ReadEffectiveTheme_UseSystemSetting_WithAppsUseLightThemeOne_ReturnsLight()
    {
        Assert.Equal("light", OfficeTheme.ReadEffectiveTheme(Fake(6, 1)));
    }

    [Fact]
    public void ReadEffectiveTheme_UseSystemSetting_PersonalizeKeyMissing_ReturnsLight()
    {
        Assert.Equal("light", OfficeTheme.ReadEffectiveTheme(Fake(6, null)));
    }

    [Fact]
    public void ReadEffectiveTheme_UnknownNumericValue_ReturnsLight()
    {
        Assert.Equal("light", OfficeTheme.ReadEffectiveTheme(Fake(99)));
    }

    [Fact]
    public void ReadEffectiveTheme_RegistryThrows_ReturnsLight()
    {
        System.Func<string, string, object, object> throwing = (k, v, d) => throw new System.UnauthorizedAccessException();
        Assert.Equal("light", OfficeTheme.ReadEffectiveTheme(throwing));
    }
}
