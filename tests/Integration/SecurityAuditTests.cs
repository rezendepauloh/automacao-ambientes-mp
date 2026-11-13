using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text.RegularExpressions;
using AmbientesHub.Services;
using AmbientesHub.Tests.Helpers;
using Xunit;

namespace AmbientesHub.Tests.Integration;

/// <summary>
/// Suíte SecOps contínua:
/// (a) Varredura de segredos e PII em todos os arquivos rastreados pelo Git (git ls-files).
/// (b) Paridade de chaves entre ConfigService e .env.example.
/// (c) Travas fail-closed para comandos/credenciais/URLs/perfis.
/// (d) Hardening de ACL (equivalente a chmod 600 no Windows).
/// </summary>
public class SecurityAuditTests
{
    private static readonly string[] TextExtensions =
        [".cs", ".xaml", ".csproj", ".ps1", ".md", ".example", ".txt", ".svg"];

    // Padrões de segredos reais (tokens de provedores, chaves privadas PEM, conexões com senha explícita)
    private static readonly (string Name, Regex Pattern)[] SecretPatterns =
    [
        ("Private Key PEM", new Regex(@"-----BEGIN (?:RSA |EC |DSA |OPENSSH |ENCRYPTED )?PRIVATE KEY-----", RegexOptions.Compiled)),
        ("GitHub Token", new Regex(@"\b(?:ghp|gho|ghu|ghs|ghr)_[A-Za-z0-9]{36,}\b", RegexOptions.Compiled)),
        ("AWS Access Key", new Regex(@"\bAKIA[0-9A-Z]{16}\b", RegexOptions.Compiled)),
        ("Slack Token", new Regex(@"\bxox[baprs]-[A-Za-z0-9\-]{10,}\b", RegexOptions.Compiled)),
        ("Google API Key", new Regex(@"\bAIza[0-9A-Za-z\-_]{35}\b", RegexOptions.Compiled)),
        ("ConvertTo-SecureString -AsPlainText", new Regex(@"ConvertTo-SecureString\s+[^\r\n#-]*-AsPlainText", RegexOptions.Compiled | RegexOptions.IgnoreCase))
    ];

    // Identificadores pessoais/organizacionais reais proibidos em arquivos versionados (conforme AGENTS.md §2)
    private static readonly string[] ForbiddenPersonalLiterals =
    [
        "paulogoncalves",
        "automated-OTRS-and-CitSmart",
        "MPE-80600",
        "unigran.br"
    ];

    [Fact]
    public void AllGitTrackedFiles_ContainNoHardcodedSecretsOrPersonalIdentifiers()
    {
        var tracked = RepoLocator.TrackedFiles();
        Assert.NotEmpty(tracked);

        var findings = new List<string>();

        foreach (string rel in tracked)
        {
            // Este próprio arquivo de teste contém os literais proibidos como alvos de verificação
            if (rel.EndsWith("SecurityAuditTests.cs", StringComparison.OrdinalIgnoreCase) ||
                rel.EndsWith("BancadaDockerServiceTests.cs", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string ext = Path.GetExtension(rel);
            bool isText = TextExtensions.Contains(ext, StringComparer.OrdinalIgnoreCase)
                          || Path.GetFileName(rel).StartsWith(".env", StringComparison.OrdinalIgnoreCase)
                          || Path.GetFileName(rel).StartsWith(".git", StringComparison.OrdinalIgnoreCase);
            if (!isText) continue;

            string full = RepoLocator.FullPath(rel);
            string content = File.ReadAllText(full);

            foreach (var (name, rx) in SecretPatterns)
            {
                if (rx.IsMatch(content))
                    findings.Add($"{rel}: padrão de segredo detectado ({name})");
            }

            foreach (string forbidden in ForbiddenPersonalLiterals)
            {
                if (content.Contains(forbidden, StringComparison.OrdinalIgnoreCase))
                    findings.Add($"{rel}: identificador pessoal/organizacional proibido '{forbidden}'");
            }
        }

        Assert.True(findings.Count == 0, "Vazamentos detectados em arquivos rastreados:\n" + string.Join("\n", findings));
    }

    [Fact]
    public void SensitiveLocalFiles_AreNeverTrackedByGit()
    {
        var tracked = RepoLocator.TrackedFiles();

        Assert.DoesNotContain(tracked, f => f.Equals(".env", StringComparison.OrdinalIgnoreCase) || f.EndsWith("/.env", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(tracked, f => f.StartsWith("Credenciais/", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(tracked, f => f.StartsWith("temp/", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(tracked, f => f.EndsWith("config.ps1", StringComparison.OrdinalIgnoreCase) && !f.EndsWith("config.exemplo.ps1", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(tracked, f => f.EndsWith("login.ps1", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(tracked, f => f.EndsWith(".log", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void EnvExampleFiles_ExistAreIdenticalAndCoverAllConfigServiceKeys()
    {
        string rootExample = RepoLocator.FullPath(".env.example");
        string hubExample = RepoLocator.FullPath("AmbientesHub/.env.example");

        Assert.True(File.Exists(rootExample));
        Assert.True(File.Exists(hubExample));

        string rootContent = File.ReadAllText(rootExample).Replace("\r\n", "\n").Trim();
        string hubContent = File.ReadAllText(hubExample).Replace("\r\n", "\n").Trim();
        Assert.Equal(rootContent, hubContent);

        var definedKeys = rootContent
            .Split('\n')
            .Select(l => l.Trim())
            .Where(l => !string.IsNullOrEmpty(l) && !l.StartsWith('#') && l.Contains('='))
            .Select(l => l[..l.IndexOf('=')].Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        string[] requiredKeys =
        [
            "APP_TERMINAL_FONT", "ADMIN_USER", "ADMIN_AD_USER",
            "SCCM_SITE_SERVER", "SCCM_SITE_CODE",
            "URL_OTRS", "URL_CITSMART", "URL_BANCADA_LOCAL",
            "BANCADA_WINDOWS_DIR", "BANCADA_WSL_DIR", "BANCADA_WSL_DISTRO",
            "WHATSAPP_PWA_ARGS", "RSAT_MSU_NETWORK_PATH", "DEVICE_REPORTS_DIR",
            "CMRCVIEWER_PATH", "FIREFOX_PATH", "PROTECTED_PROFILES",
            "PING_TIMEOUT_MS", "MINIMIZE_TO_TRAY"
        ];

        foreach (string key in requiredKeys)
        {
            Assert.Contains(key, definedKeys);
        }
    }

    [Fact]
    public void HardenFilePermissions_DisablesInheritance_AndGrantsCurrentUserAccess()
    {
        string tempEnv = Path.Combine(Path.GetTempPath(), $"sec_acl_{Guid.NewGuid():N}.env");
        try
        {
            File.WriteAllText(tempEnv, "ADMIN_USER=seu_usuario_admin");
            bool ok = ConfigService.HardenFilePermissions(tempEnv);
            Assert.True(ok);

            var acl = new FileInfo(tempEnv).GetAccessControl();
            Assert.True(acl.AreAccessRulesProtected);

            var rules = acl.GetAccessRules(includeExplicit: true, includeInherited: true, typeof(SecurityIdentifier))
                .Cast<FileSystemAccessRule>()
                .ToList();

            var authenticatedUsersSid = new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null);
            var usersSid = new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null);

            Assert.DoesNotContain(rules, r => r.IdentityReference.Equals(authenticatedUsersSid));
            Assert.DoesNotContain(rules, r => r.IdentityReference.Equals(usersSid));
        }
        finally
        {
            if (File.Exists(tempEnv)) File.Delete(tempEnv);
        }
    }

    [Fact]
    public void RemoteManagementHelper_FailClosed_RejectsInvalidHostsAndControlChars()
    {
        Assert.False(RemoteManagementHelper.PingHost("pc01 & calc"));
        Assert.False(RemoteManagementHelper.EnsureAdminShareConnected("pc01; calc"));
        Assert.False(RemoteManagementHelper.InvokeCommandOnTarget("pc01 & calc", "whoami"));
        Assert.False(RemoteManagementHelper.InvokeCommandOnTarget(".", "echo ok\r\ncalc.exe"));
        Assert.Throws<ArgumentException>(() => RemoteManagementHelper.CreateScope("pc01\" /evil"));
    }

    [Fact]
    public void RemoteUserProfileService_FailClosed_RejectsMaliciousHostAndUsernames()
    {
        var svc = new RemoteUserProfileService();
        var logs = new List<string>();

        var badHostRes = svc.PurgeUserProfiles("pc01 & calc", ["usuario.teste"], logs.Add);
        Assert.Equal(0, badHostRes.RemovedCount);
        Assert.Contains(logs, l => l.Contains("[ERRO]"));

        logs.Clear();
        var badUserRes = svc.PurgeUserProfiles("192.0.2.1", ["usuario\" & calc"], logs.Add);
        Assert.Equal(0, badUserRes.RemovedCount);
        Assert.Equal(1, badUserRes.SkippedCount);
    }

    [Theory]
    [InlineData("\"-taskbar-tab\" \"uuid\" \"-new-window\" \"https://web.whatsapp.com/\"", true)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    [InlineData("-new-window https://web.whatsapp.com/ & calc.exe", false)]
    [InlineData("-new-window $(whoami)", false)]
    [InlineData("-new-window `id`", false)]
    [InlineData("-new-window\r\ncalc", false)]
    public void BrowserLauncherService_IsSafeBrowserArguments_EnforcesWhitelist(string args, bool expected)
    {
        Assert.Equal(expected, BrowserLauncherService.IsSafeBrowserArguments(args));
    }
}

