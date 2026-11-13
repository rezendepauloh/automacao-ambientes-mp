using System.Text;

namespace AmbientesHub.Models;

public class AdUserDetails
{
    public string SamAccountName { get; set; } = string.Empty;
    public string UserPrincipalName { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string CommonName { get; set; } = string.Empty;
    public string GivenName { get; set; } = string.Empty;
    public string Surname { get; set; } = string.Empty;

    public string Mail { get; set; } = string.Empty;
    public string TelephoneNumber { get; set; } = string.Empty;
    public string Mobile { get; set; } = string.Empty;
    public string Department { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Company { get; set; } = string.Empty;
    public string Office { get; set; } = string.Empty;
    public string EmployeeId { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string ManagerDn { get; set; } = string.Empty;

    public string DistinguishedName { get; set; } = string.Empty;
    public string HomeDirectory { get; set; } = string.Empty;
    public string HomeDrive { get; set; } = string.Empty;
    public string ProfilePath { get; set; } = string.Empty;
    public string ScriptPath { get; set; } = string.Empty;

    public string WhenCreated { get; set; } = string.Empty;
    public string WhenChanged { get; set; } = string.Empty;
    public string LastLogon { get; set; } = string.Empty;
    public string PasswordLastSet { get; set; } = string.Empty;
    public string AccountExpires { get; set; } = string.Empty;

    public bool IsEnabled { get; set; } = true;
    public bool IsLockedOut { get; set; } = false;
    public bool PasswordNeverExpires { get; set; } = false;
    public bool MustChangePasswordNextLogon { get; set; } = false;
    public int UserAccountControl { get; set; }
    public int BadPasswordCount { get; set; }
    public int LogonCount { get; set; }

    public List<string> Groups { get; set; } = [];

    public string HeaderNameDisplay => !string.IsNullOrWhiteSpace(DisplayName)
        ? DisplayName
        : (!string.IsNullOrWhiteSpace(CommonName) ? CommonName : SamAccountName);

    public string SamAndUpnDisplay => !string.IsNullOrWhiteSpace(UserPrincipalName)
        ? $"{SamAccountName} • {UserPrincipalName}"
        : SamAccountName;

    public string StatusBadgeText
    {
        get
        {
            if (!IsEnabled) return "🔴 CONTA DESATIVADA";
            if (IsLockedOut) return "🔒 CONTA BLOQUEADA";
            if (MustChangePasswordNextLogon) return "🟡 TROCAR SENHA NO LOGON";
            return "🟢 CONTA ATIVA";
        }
    }

    public string StatusBadgeBackground
    {
        get
        {
            if (!IsEnabled) return "#7f1d1d";
            if (IsLockedOut) return "#78350f";
            if (MustChangePasswordNextLogon) return "#713f12";
            return "#064e3b";
        }
    }

    public string StatusBadgeForeground
    {
        get
        {
            if (!IsEnabled) return "#f87171";
            if (IsLockedOut) return "#fbbf24";
            if (MustChangePasswordNextLogon) return "#fde047";
            return "#34d399";
        }
    }

    public string MailDisplay => !string.IsNullOrWhiteSpace(Mail) ? Mail : "Não informado";

    public string TitleDisplay => !string.IsNullOrWhiteSpace(Title) ? Title : "Não informado";

    public string DepartmentAndOfficeDisplay
    {
        get
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(Department)) parts.Add(Department.Trim());
            if (!string.IsNullOrWhiteSpace(Office) && !string.Equals(Office.Trim(), Department.Trim(), StringComparison.OrdinalIgnoreCase))
                parts.Add(Office.Trim());
            if (!string.IsNullOrWhiteSpace(Company)) parts.Add(Company.Trim());

            return parts.Count > 0 ? string.Join(" • ", parts) : "Não informado";
        }
    }

    public string PhonesDisplay
    {
        get
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(TelephoneNumber)) parts.Add($"Tel: {TelephoneNumber.Trim()}");
            if (!string.IsNullOrWhiteSpace(Mobile)) parts.Add($"Cel: {Mobile.Trim()}");
            if (!string.IsNullOrWhiteSpace(EmployeeId)) parts.Add($"Matrícula: {EmployeeId.Trim()}");

            return parts.Count > 0 ? string.Join(" | ", parts) : "Não informado";
        }
    }

    public string DescriptionDisplay => !string.IsNullOrWhiteSpace(Description) ? Description : "Sem descrição";

    public string ManagerDisplay => !string.IsNullOrWhiteSpace(ManagerDn) ? ExtractCnFromDn(ManagerDn) : "Não informado";

    public string OrganizationalUnitDisplay => ExtractOuFromDn(DistinguishedName);

    public string PasswordPolicyDisplay
    {
        get
        {
            var flags = new List<string>();
            if (PasswordNeverExpires) flags.Add("Senha nunca expira");
            if (MustChangePasswordNextLogon) flags.Add("Trocar no próximo logon");
            if (IsLockedOut) flags.Add("Bloqueada por tentativas");
            if (!IsEnabled) flags.Add("Conta desabilitada");
            if (flags.Count == 0) flags.Add("Padrão do domínio");
            return $"{string.Join(" • ", flags)} (UAC: {UserAccountControl})";
        }
    }

    public string LastLogonDisplay => !string.IsNullOrWhiteSpace(LastLogon) ? LastLogon : "Nunca / Não registrado";

    public string PasswordLastSetDisplay => !string.IsNullOrWhiteSpace(PasswordLastSet) ? PasswordLastSet : "Nunca definida";

    public string AccountExpiresDisplay => !string.IsNullOrWhiteSpace(AccountExpires) ? AccountExpires : "Nunca expira";

    public string AuditCountersDisplay => $"Erros de senha: {BadPasswordCount}  •  Total de logons: {LogonCount}";

    public string CreatedAndChangedDisplay
    {
        get
        {
            string created = !string.IsNullOrWhiteSpace(WhenCreated) ? WhenCreated : "N/D";
            string changed = !string.IsNullOrWhiteSpace(WhenChanged) ? WhenChanged : "N/D";
            return $"Criado: {created}  |  Modificado: {changed}";
        }
    }

    public string ProfileAndPathsDisplay
    {
        get
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(ScriptPath)) parts.Add($"Script: {ScriptPath}");
            if (!string.IsNullOrWhiteSpace(HomeDirectory))
                parts.Add(!string.IsNullOrWhiteSpace(HomeDrive) ? $"Home ({HomeDrive}): {HomeDirectory}" : $"Home: {HomeDirectory}");
            if (!string.IsNullOrWhiteSpace(ProfilePath)) parts.Add($"Perfil: {ProfilePath}");

            return parts.Count > 0 ? string.Join(" | ", parts) : "Padrão (sem script/diretório base)";
        }
    }

    public string GroupsHeaderDisplay => $"👥 GRUPOS DE SEGURANÇA & DISTRIBUIÇÃO ({Groups.Count})";

    public string GroupsDisplay => Groups.Count > 0
        ? string.Join("  •  ", Groups)
        : "Nenhum grupo adicional listado em memberOf";

    public static string ExtractCnFromDn(string? dn)
    {
        if (string.IsNullOrWhiteSpace(dn))
            return string.Empty;

        string trimmed = dn.Trim();
        if (trimmed.StartsWith("CN=", StringComparison.OrdinalIgnoreCase))
        {
            int commaIdx = FindUnescapedComma(trimmed, 3);
            string cn = commaIdx > 3 ? trimmed[3..commaIdx] : trimmed[3..];
            return cn.Replace(@"\,", ",");
        }

        return trimmed;
    }

    public static string ExtractOuFromDn(string? dn)
    {
        if (string.IsNullOrWhiteSpace(dn))
            return "Não informado";

        string trimmed = dn.Trim();
        if (trimmed.StartsWith("CN=", StringComparison.OrdinalIgnoreCase))
        {
            int commaIdx = FindUnescapedComma(trimmed, 3);
            if (commaIdx > 0 && commaIdx + 1 < trimmed.Length)
            {
                return trimmed[(commaIdx + 1)..].Trim();
            }
        }

        return trimmed;
    }

    private static int FindUnescapedComma(string text, int startIndex)
    {
        for (int i = startIndex; i < text.Length; i++)
        {
            if (text[i] == ',' && (i == 0 || text[i - 1] != '\\'))
                return i;
        }
        return -1;
    }

    public string FormatFullDossier()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"=== FICHA DO USUÁRIO ACTIVE DIRECTORY: {HeaderNameDisplay} ===");
        sb.AppendLine($"Login (sAMAccountName) : {SamAccountName}");
        sb.AppendLine($"UPN                    : {UserPrincipalName}");
        sb.AppendLine($"Status da Conta        : {StatusBadgeText}");
        sb.AppendLine($"E-mail                 : {MailDisplay}");
        sb.AppendLine($"Cargo / Função         : {TitleDisplay}");
        sb.AppendLine($"Lotação / Setor        : {DepartmentAndOfficeDisplay}");
        sb.AppendLine($"Contatos / Matrícula   : {PhonesDisplay}");
        sb.AppendLine($"Descrição              : {DescriptionDisplay}");
        sb.AppendLine($"Gestor                 : {ManagerDisplay}");
        sb.AppendLine($"Política de Senha      : {PasswordPolicyDisplay}");
        sb.AppendLine($"Último Logon           : {LastLogonDisplay}");
        sb.AppendLine($"Última Troca de Senha  : {PasswordLastSetDisplay}");
        sb.AppendLine($"Expiração da Conta     : {AccountExpiresDisplay}");
        sb.AppendLine($"Contadores de Auditoria: {AuditCountersDisplay}");
        sb.AppendLine($"Datas do Objeto        : {CreatedAndChangedDisplay}");
        sb.AppendLine($"Caminhos / Script      : {ProfileAndPathsDisplay}");
        sb.AppendLine($"Unidade Organizacional : {OrganizationalUnitDisplay}");
        sb.AppendLine($"Distinguished Name (DN): {DistinguishedName}");
        sb.AppendLine($"Grupos ({Groups.Count})           : {(Groups.Count > 0 ? string.Join(", ", Groups) : "Nenhum")}");
        return sb.ToString();
    }
}

