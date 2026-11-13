using AmbientesHub.Models;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class AdUserDetailsTests
{
    [Theory]
    [InlineData("CN=Grupo Suporte,OU=Grupos,DC=dominio,DC=local", "Grupo Suporte")]
    [InlineData(@"CN=Sobrenome\, Nome,OU=Usuarios,DC=dominio,DC=local", "Sobrenome, Nome")]
    [InlineData("usuario.sobrenome", "usuario.sobrenome")]
    [InlineData("", "")]
    public void ExtractCnFromDn_ExtractsCleanCommonName(string dn, string expectedCn)
    {
        Assert.Equal(expectedCn, AdUserDetails.ExtractCnFromDn(dn));
    }

    [Theory]
    [InlineData("CN=Usuario Teste,OU=TI,OU=Usuarios,DC=dominio,DC=local", "OU=TI,OU=Usuarios,DC=dominio,DC=local")]
    [InlineData("", "Não informado")]
    public void ExtractOuFromDn_ExtractsOrganizationalUnitPath(string dn, string expectedOu)
    {
        Assert.Equal(expectedOu, AdUserDetails.ExtractOuFromDn(dn));
    }

    [Fact]
    public void StatusBadgeText_ReflectsAccountFlagsCorrectly()
    {
        var active = new AdUserDetails { IsEnabled = true, IsLockedOut = false };
        var disabled = new AdUserDetails { IsEnabled = false, IsLockedOut = true };
        var locked = new AdUserDetails { IsEnabled = true, IsLockedOut = true };
        var mustChangePwd = new AdUserDetails { IsEnabled = true, IsLockedOut = false, MustChangePasswordNextLogon = true };

        Assert.Contains("CONTA ATIVA", active.StatusBadgeText);
        Assert.Contains("DESATIVADA", disabled.StatusBadgeText);
        Assert.Contains("BLOQUEADA", locked.StatusBadgeText);
        Assert.Contains("TROCAR SENHA", mustChangePwd.StatusBadgeText);
    }

    [Fact]
    public void FormatFullDossier_IncludesAllEssentialSections()
    {
        var user = new AdUserDetails
        {
            SamAccountName = "usuario.sobrenome",
            UserPrincipalName = "usuario.sobrenome@dominio.local",
            DisplayName = "Usuario Sobrenome Teste",
            Mail = "usuario.sobrenome@dominio.local",
            Title = "Analista de Sistemas",
            Department = "Tecnologia da Informacao",
            Office = "Sala 101",
            Company = "Organizacao",
            Groups = ["G_SUPORTE_TI", "G_VPN_ACESSO"]
        };

        string dossier = user.FormatFullDossier();

        Assert.Contains("usuario.sobrenome", dossier);
        Assert.Contains("Usuario Sobrenome Teste", dossier);
        Assert.Contains("Analista de Sistemas", dossier);
        Assert.Contains("Tecnologia da Informacao", dossier);
        Assert.Contains("G_SUPORTE_TI, G_VPN_ACESSO", dossier);
    }

    [Theory]
    [InlineData(@"DOMINIO\usuario.teste", "usuario.teste")]
    [InlineData("  usuario.sobrenome  ", "usuario.sobrenome")]
    [InlineData("", "")]
    public void NormalizeAdUserQuery_StripsDomainPrefixAndWhitespace(string raw, string expected)
    {
        Assert.Equal(expected, AdSccmService.NormalizeAdUserQuery(raw));
    }

    [Fact]
    public void BuildAdUserLdapFilter_BuildsMultiAttributeFilterAndEscapesSpecialChars()
    {
        string filterSingle = AdSccmService.BuildAdUserLdapFilter("usuario.sobrenome");
        Assert.Contains("(sAMAccountName=*usuario.sobrenome*)", filterSingle);
        Assert.Contains("(mail=*usuario.sobrenome*)", filterSingle);
        Assert.Contains("(department=*usuario.sobrenome*)", filterSingle);

        string filterMultiWord = AdSccmService.BuildAdUserLdapFilter("Usuario Teste");
        Assert.Contains("*Usuario*Teste*", filterMultiWord);
    }

    [Fact]
    public void ParseUserAccountFlags_DecodesBitmaskCorrectly()
    {
        // 512 = NORMAL_ACCOUNT (Ativa)
        var normal = AdSccmService.ParseUserAccountFlags(512, 0, 0, 133500000000000000);
        Assert.True(normal.isEnabled);
        Assert.False(normal.isLockedOut);
        Assert.False(normal.passwordNeverExpires);
        Assert.False(normal.mustChangePasswordNextLogon);

        // 514 = NORMAL_ACCOUNT (512) + ACCOUNTDISABLE (2)
        var disabled = AdSccmService.ParseUserAccountFlags(514, 0, 0, 133500000000000000);
        Assert.False(disabled.isEnabled);

        // 66048 = NORMAL_ACCOUNT (512) + DONT_EXPIRE_PASSWD (65536)
        var neverExpires = AdSccmService.ParseUserAccountFlags(66048, 0, 0, 0);
        Assert.True(neverExpires.isEnabled);
        Assert.True(neverExpires.passwordNeverExpires);
        Assert.False(neverExpires.mustChangePasswordNextLogon);

        // pwdLastSet == 0 sem DONT_EXPIRE_PASSWD e lockoutTime > 0
        var lockedAndExpiredPwd = AdSccmService.ParseUserAccountFlags(512, 0, 133500000000000000, 0);
        Assert.True(lockedAndExpiredPwd.isLockedOut);
        Assert.True(lockedAndExpiredPwd.mustChangePasswordNextLogon);
    }

    [Fact]
    public void FormatAdFileTime_HandlesNeverExpiresAndValidDates()
    {
        Assert.Equal("Nunca expira", AdSccmService.FormatAdFileTime(0, isAccountExpires: true));
        Assert.Equal("Nunca expira", AdSccmService.FormatAdFileTime(long.MaxValue, isAccountExpires: true));
        Assert.Equal(string.Empty, AdSccmService.FormatAdFileTime(0, isAccountExpires: false));

        var sampleDate = new DateTime(2026, 5, 15, 14, 30, 0, DateTimeKind.Local);
        string formatted = AdSccmService.FormatAdFileTime(sampleDate.ToFileTime(), isAccountExpires: false);
        Assert.Equal("15/05/2026 14:30", formatted);
    }
}

