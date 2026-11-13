using System.Text.RegularExpressions;

namespace AmbientesHub.Services;

/// <summary>
/// Validação e sanitização centralizadas (whitelists estritas) para qualquer dado que
/// possa chegar a processos externos, filtros LDAP/WQL, scripts PowerShell, UNC ou .env.
/// Política: fail-closed — entrada fora da whitelist é rejeitada, nunca "corrigida".
/// </summary>
public static class InputValidator
{
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromMilliseconds(250);

    // Hostname NetBIOS/DNS: rótulos alfanuméricos com hífen, até 5 pontos, máx. 253 caracteres.
    private static readonly Regex HostNameRegex = new(
        @"^[A-Za-z0-9]([A-Za-z0-9\-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9\-]{0,61}[A-Za-z0-9])?){0,5}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RegexTimeout);

    // Termo de busca livre (máquina, usuário, MAC): sem aspas, curingas, parênteses, barras ou metacaracteres.
    private static readonly Regex SearchTermRegex = new(
        @"^[\p{L}\p{N}][\p{L}\p{N} ._\-@:]{0,63}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RegexTimeout);

    // Conta de usuário local/AD (sAMAccountName simplificado, com DOMINIO\ opcional).
    private static readonly Regex UserNameRegex = new(
        @"^([A-Za-z0-9._\-]{1,32}\\)?[A-Za-z0-9._\-]{1,64}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RegexTimeout);

    private static readonly Regex WslDistroRegex = new(
        @"^[A-Za-z0-9][A-Za-z0-9._\-]{0,63}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RegexTimeout);

    // Caminho POSIX absoluto, sem metacaracteres de shell, sem ".." .
    private static readonly Regex WslPathRegex = new(
        @"^/[A-Za-z0-9._\-/]{0,255}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RegexTimeout);

    private static readonly Regex SiteCodeRegex = new(
        @"^[A-Za-z0-9]{3}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RegexTimeout);

    private static readonly Regex EnvKeyRegex = new(
        @"^[A-Za-z_][A-Za-z0-9_]{0,63}$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, RegexTimeout);

    public static bool IsValidHostName(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 253) return false;
        return SafeIsMatch(HostNameRegex, value);
    }

    public static bool IsValidSearchTerm(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        return SafeIsMatch(SearchTermRegex, value);
    }

    public static bool IsValidUserName(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        return SafeIsMatch(UserNameRegex, value);
    }

    public static bool IsValidWslDistro(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        return SafeIsMatch(WslDistroRegex, value);
    }

    public static bool IsValidWslPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        if (value.Contains("..", StringComparison.Ordinal)) return false;
        return SafeIsMatch(WslPathRegex, value);
    }

    public static bool IsValidSiteCode(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        return SafeIsMatch(SiteCodeRegex, value);
    }

    public static bool IsValidEnvKey(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        return SafeIsMatch(EnvKeyRegex, value);
    }

    /// <summary>Aceita apenas URLs absolutas http/https sem credenciais embutidas.</summary>
    public static bool IsSafeHttpUrl(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        if (value.IndexOfAny(['"', '\r', '\n', '\0']) >= 0) return false;
        if (!Uri.TryCreate(value, UriKind.Absolute, out var uri)) return false;
        if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps) return false;
        return string.IsNullOrEmpty(uri.UserInfo);
    }

    /// <summary>Aceita apenas executáveis com caminho absoluto local (sem UNC, sem aspas/CRLF).</summary>
    public static bool IsSafeLocalExecutablePath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        if (value.IndexOfAny(['"', '\r', '\n', '\0', '<', '>', '|', '*', '?']) >= 0) return false;
        if (value.StartsWith(@"\\", StringComparison.Ordinal)) return false;
        if (!Path.IsPathFullyQualified(value)) return false;
        return value.EndsWith(".exe", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Escapa literal WQL (barra invertida e aspa simples).</summary>
    public static string EscapeWql(string value)
        => (value ?? string.Empty).Replace(@"\", @"\\").Replace("'", @"\'");

    /// <summary>Escapa literal PowerShell entre aspas simples.</summary>
    public static string EscapePowerShellSingleQuoted(string value)
        => (value ?? string.Empty).Replace("'", "''");

    /// <summary>Escapa valor para filtro LDAP (RFC 4515).</summary>
    public static string EscapeLdap(string value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        var sb = new System.Text.StringBuilder(value.Length);
        foreach (char c in value)
        {
            switch (c)
            {
                case '\\': sb.Append(@"\5c"); break;
                case '*': sb.Append(@"\2a"); break;
                case '(': sb.Append(@"\28"); break;
                case ')': sb.Append(@"\29"); break;
                case '\0': sb.Append(@"\00"); break;
                default: sb.Append(c); break;
            }
        }
        return sb.ToString();
    }

    /// <summary>Remove CR/LF/NUL de valores gravados em .env (impede injeção de chaves).</summary>
    public static string SanitizeEnvValue(string? value)
        => (value ?? string.Empty).Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\0", string.Empty).Trim();

    /// <summary>Remove caracteres de controle de mensagens de log (anti log-injection).</summary>
    public static string SanitizeLogText(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return value.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
    }

    private static bool SafeIsMatch(Regex regex, string value)
    {
        // '$' aceita "\n" final; exigir que o match cubra a string inteira elimina esse bypass.
        try
        {
            var m = regex.Match(value);
            return m.Success && m.Index == 0 && m.Length == value.Length;
        }
        catch (RegexMatchTimeoutException) { return false; }
    }
}
