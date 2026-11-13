using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class InputValidatorTests
{
    [Theory]
    [InlineData("PC-01")]
    [InlineData("DOMINIO-80600")]
    [InlineData("srv-sccm.dominio.local")]
    [InlineData("10.0.0.15")]
    [InlineData("localhost")]
    public void IsValidHostName_AcceptsLegitimateNames(string host)
        => Assert.True(InputValidator.IsValidHostName(host));

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("pc01 /admin")]
    [InlineData("pc01&calc")]
    [InlineData("pc01;calc")]
    [InlineData("pc01|calc")]
    [InlineData("pc01\"")]
    [InlineData("pc01'")]
    [InlineData("pc01\n")]
    [InlineData("pc01\r\ncalc")]
    [InlineData("..\\..\\windows")]
    [InlineData("\\\\evil\\share")]
    [InlineData("pc$(whoami)")]
    [InlineData("*")]
    [InlineData("-pc01")]
    [InlineData("pc01-")]
    [InlineData("a..b")]
    public void IsValidHostName_RejectsInjectionAndMalformed(string? host)
        => Assert.False(InputValidator.IsValidHostName(host));

    [Fact]
    public void IsValidHostName_RejectsOverlongValues()
        => Assert.False(InputValidator.IsValidHostName(new string('a', 300)));

    [Theory]
    [InlineData("usuario.sobrenome")]
    [InlineData("maria silva")]
    [InlineData("00:1A:2B:3C:4D:5E")]
    [InlineData("DOMINIO-01")]
    public void IsValidSearchTerm_AcceptsLegitimateTerms(string term)
        => Assert.True(InputValidator.IsValidSearchTerm(term));

    [Theory]
    [InlineData("*")]
    [InlineData("*)(objectClass=*")]
    [InlineData("x' OR '1'='1")]
    [InlineData("a'; Remove-Item C:\\ -Recurse #")]
    [InlineData("a)(cn=*")]
    [InlineData("a\\5c")]
    [InlineData("a%")]
    [InlineData("")]
    [InlineData(null)]
    public void IsValidSearchTerm_RejectsLdapWqlAndShellMetacharacters(string? term)
        => Assert.False(InputValidator.IsValidSearchTerm(term));

    [Theory]
    [InlineData("usuario.teste", true)]
    [InlineData("DOMINIO\\usuario.teste", true)]
    [InlineData("usuario\" /delete & calc", false)]
    [InlineData("usuario&calc", false)]
    [InlineData("..\\admin", false)]
    [InlineData("", false)]
    public void IsValidUserName_Whitelist(string user, bool expected)
        => Assert.Equal(expected, InputValidator.IsValidUserName(user));

    [Theory]
    [InlineData("Ubuntu-24.04", true)]
    [InlineData("Debian", true)]
    [InlineData("Ubuntu; calc", false)]
    [InlineData("Ubuntu && calc", false)]
    [InlineData("-d evil", false)]
    [InlineData("", false)]
    public void IsValidWslDistro_Whitelist(string distro, bool expected)
        => Assert.Equal(expected, InputValidator.IsValidWslDistro(distro));

    [Theory]
    [InlineData("/home/usuario/projeto", true)]
    [InlineData("/srv/app-1_2.x", true)]
    [InlineData("/home/u/$(id)", false)]
    [InlineData("/home/u/`id`", false)]
    [InlineData("/home/u'; rm -rf / #", false)]
    [InlineData("/home/u\"; calc", false)]
    [InlineData("/home/u/../../etc", false)]
    [InlineData("relativo/dir", false)]
    [InlineData("/home/u with space", false)]
    [InlineData("", false)]
    public void IsValidWslPath_Whitelist(string path, bool expected)
        => Assert.Equal(expected, InputValidator.IsValidWslPath(path));

    [Theory]
    [InlineData("ABC", true)]
    [InlineData("S01", true)]
    [InlineData("AB", false)]
    [InlineData("ABCD", false)]
    [InlineData("A'B", false)]
    public void IsValidSiteCode_Whitelist(string code, bool expected)
        => Assert.Equal(expected, InputValidator.IsValidSiteCode(code));

    [Theory]
    [InlineData("https://exemplo.local/", true)]
    [InlineData("http://localhost:8501/", true)]
    [InlineData("file:///C:/Windows/System32/calc.exe", false)]
    [InlineData("javascript:alert(1)", false)]
    [InlineData("C:\\Windows\\System32\\calc.exe", false)]
    [InlineData("https://user:pass@exemplo.local/", false)]
    [InlineData("https://exemplo.local/\" -new-window", false)]
    [InlineData("ftp://exemplo.local/", false)]
    [InlineData("", false)]
    public void IsSafeHttpUrl_OnlyHttpAndHttps(string url, bool expected)
        => Assert.Equal(expected, InputValidator.IsSafeHttpUrl(url));

    [Theory]
    [InlineData("C:\\Program Files\\Mozilla Firefox\\firefox.exe", true)]
    [InlineData("\\\\servidor\\share\\tool.exe", false)]
    [InlineData("tool.exe", false)]
    [InlineData("C:\\tool.bat", false)]
    [InlineData("C:\\tool.exe\" & calc", false)]
    [InlineData("", false)]
    public void IsSafeLocalExecutablePath_Whitelist(string path, bool expected)
        => Assert.Equal(expected, InputValidator.IsSafeLocalExecutablePath(path));

    [Theory]
    [InlineData("FOO", true)]
    [InlineData("URL_OTRS", true)]
    [InlineData("1ABC", false)]
    [InlineData("A=B", false)]
    [InlineData("A\nB", false)]
    [InlineData("", false)]
    public void IsValidEnvKey_Whitelist(string key, bool expected)
        => Assert.Equal(expected, InputValidator.IsValidEnvKey(key));

    [Fact]
    public void SanitizeEnvValue_RemovesLineBreaks_PreventingKeyInjection()
    {
        string sanitized = InputValidator.SanitizeEnvValue("valor\r\nFIREFOX_PATH=C:\\malware.exe");
        Assert.DoesNotContain('\n', sanitized);
        Assert.DoesNotContain('\r', sanitized);
        Assert.Equal("valorFIREFOX_PATH=C:\\malware.exe", sanitized);
    }

    [Fact]
    public void SanitizeLogText_RemovesLineBreaks_PreventingLogForging()
    {
        string sanitized = InputValidator.SanitizeLogText("ok\n[2026-01-01 00:00:00] [INFO] forjado");
        Assert.DoesNotContain('\n', sanitized);
    }

    [Theory]
    [InlineData("*", @"\2a")]
    [InlineData("(", @"\28")]
    [InlineData(")", @"\29")]
    [InlineData("\\", @"\5c")]
    public void EscapeLdap_EscapesRfc4515SpecialCharacters(string input, string expected)
        => Assert.Equal(expected, InputValidator.EscapeLdap(input));

    [Fact]
    public void EscapeWql_EscapesQuoteAndBackslash()
        => Assert.Equal(@"a\'b\\c", InputValidator.EscapeWql(@"a'b\c"));

    [Fact]
    public void EscapePowerShellSingleQuoted_DoublesQuotes()
        => Assert.Equal("a''b", InputValidator.EscapePowerShellSingleQuoted("a'b"));
}

