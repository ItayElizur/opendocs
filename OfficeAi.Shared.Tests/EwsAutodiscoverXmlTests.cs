using Xunit;
using OfficeAi.Shared;

public class EwsAutodiscoverXmlTests
{
    private const string Ns = "http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a";

    private static string Wrap(string protocols) =>
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
        "<Autodiscover xmlns=\"http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006\">" +
        "<Response xmlns=\"" + Ns + "\"><User><DisplayName>T</DisplayName></User>" +
        "<Account><AccountType>email</AccountType><Action>settings</Action>" +
        protocols +
        "</Account></Response></Autodiscover>";

    [Fact]
    public void ExchAndExpr_PrefersExch()
    {
        string xml = Wrap(
            "<Protocol><Type>EXPR</Type><ASUrl>https://ext.corp.com/EWS/Exchange.asmx</ASUrl></Protocol>" +
            "<Protocol><Type>EXCH</Type><ASUrl>https://int.corp.local/EWS/Exchange.asmx</ASUrl></Protocol>");
        Assert.Equal("https://int.corp.local/EWS/Exchange.asmx", EwsAutodiscoverXml.ParseEwsUrl(xml));
    }

    [Fact]
    public void ExprOnly_ReturnsExpr()
    {
        string xml = Wrap("<Protocol><Type>EXPR</Type><ASUrl>https://ext.corp.com/EWS/Exchange.asmx</ASUrl></Protocol>");
        Assert.Equal("https://ext.corp.com/EWS/Exchange.asmx", EwsAutodiscoverXml.ParseEwsUrl(xml));
    }

    [Fact]
    public void EwsUrlElement_PreferredOverAsUrl()
    {
        string xml = Wrap(
            "<Protocol><Type>EXCH</Type>" +
            "<EwsUrl>https://int.corp.local/EWS/Exchange.asmx</EwsUrl>" +
            "<ASUrl>https://avail.corp.local/EWS/Exchange.asmx</ASUrl></Protocol>");
        Assert.Equal("https://int.corp.local/EWS/Exchange.asmx", EwsAutodiscoverXml.ParseEwsUrl(xml));
    }

    [Fact]
    public void UnknownProtocolType_StillUsedAsLastResort()
    {
        string xml = Wrap("<Protocol><Type>WEB</Type><EwsUrl>https://web.corp.local/EWS/Exchange.asmx</EwsUrl></Protocol>");
        Assert.Equal("https://web.corp.local/EWS/Exchange.asmx", EwsAutodiscoverXml.ParseEwsUrl(xml));
    }

    [Fact]
    public void Malformed_ReturnsNull()
    {
        Assert.Null(EwsAutodiscoverXml.ParseEwsUrl("<Autodiscover><Response><oops"));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("   ")]
    public void EmptyOrWhitespace_ReturnsNull(string input)
    {
        Assert.Null(EwsAutodiscoverXml.ParseEwsUrl(input));
    }

    [Fact]
    public void ErrorResponse_NoProtocol_ReturnsNull()
    {
        string xml =
            "<Autodiscover xmlns=\"http://schemas.microsoft.com/exchange/autodiscover/responseschema/2006\">" +
            "<Response><Error Time=\"x\" Id=\"1\"><ErrorCode>600</ErrorCode><Message>Invalid Request</Message></Error></Response>" +
            "</Autodiscover>";
        Assert.Null(EwsAutodiscoverXml.ParseEwsUrl(xml));
    }

    [Fact]
    public void ProtocolWithoutUrl_IsSkipped()
    {
        string xml = Wrap(
            "<Protocol><Type>EXCH</Type><Server>int.corp.local</Server></Protocol>" +
            "<Protocol><Type>EXPR</Type><ASUrl>https://ext.corp.com/EWS/Exchange.asmx</ASUrl></Protocol>");
        Assert.Equal("https://ext.corp.com/EWS/Exchange.asmx", EwsAutodiscoverXml.ParseEwsUrl(xml));
    }

    [Fact]
    public void RealisticNamespacedSample_ReturnsExchAsUrl()
    {
        string xml = Wrap(
            "<Protocol><Type>EXCH</Type><Server>CONTOSO-EX01.corp.local</Server>" +
            "<ServerDN>/o=Contoso/ou=Exchange/cn=Servers/cn=EX01</ServerDN>" +
            "<ASUrl>https://mail.corp.local/EWS/Exchange.asmx</ASUrl>" +
            "<OOFUrl>https://mail.corp.local/EWS/Exchange.asmx</OOFUrl></Protocol>" +
            "<Protocol><Type>EXPR</Type><Server>mail.contoso.com</Server>" +
            "<ASUrl>https://mail.contoso.com/EWS/Exchange.asmx</ASUrl></Protocol>" +
            "<Protocol><Type>WEB</Type></Protocol>");
        Assert.Equal("https://mail.corp.local/EWS/Exchange.asmx", EwsAutodiscoverXml.ParseEwsUrl(xml));
    }
}
