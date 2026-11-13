using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace AmbientesHub.Services;

public static class CredentialService
{
    public const string DefaultAdminCredPath = @"C:\AmbientesVirtuais\Credenciais\cred_admin.xml";
    private static readonly XNamespace Ns = "http://schemas.microsoft.com/powershell/2004/04";

    public static (bool Success, string Message, string? Username, string? Password) ReadPlaintextCredential(string xmlPath = DefaultAdminCredPath)
    {
        if (!File.Exists(xmlPath))
        {
            return (false, $"Arquivo de credencial não encontrado: {xmlPath}", null, null);
        }

        try
        {
            // Anti XXE/DoS: arquivo limitado a 64 KB, DTD proibido e resolver de entidades externas desabilitado.
            if (new FileInfo(xmlPath).Length > 64 * 1024)
            {
                return (false, "Arquivo de credencial excede o tamanho máximo permitido (64 KB).", null, null);
            }

            var readerSettings = new System.Xml.XmlReaderSettings
            {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 64 * 1024
            };
            XDocument doc;
            using (var xmlReader = System.Xml.XmlReader.Create(xmlPath, readerSettings))
            {
                doc = XDocument.Load(xmlReader);
            }
            var userEl = doc.Descendants(Ns + "S").FirstOrDefault(e => (string?)e.Attribute("N") == "UserName");
            var passEl = doc.Descendants(Ns + "SS").FirstOrDefault(e => (string?)e.Attribute("N") == "Password");

            if (userEl == null || passEl == null)
            {
                return (false, "Formato Clixml inválido: elementos de usuário ou senha ausentes.", null, null);
            }

            string username = userEl.Value.Trim();
            string hexPassword = passEl.Value.Trim();

            byte[] encryptedBytes = Convert.FromHexString(hexPassword);
            byte[] decryptedBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
            string plainPassword = Encoding.Unicode.GetString(decryptedBytes);

            return (true, "Credencial DPAPI descriptografada com sucesso.", username, plainPassword);
        }
        catch (CryptographicException)
        {
            return (false, "Falha na descriptografia DPAPI: a credencial pertence a outro usuário ou máquina.", null, null);
        }
        catch (Exception ex)
        {
            return (false, $"Erro ao verificar cofre: {ex.Message}", null, null);
        }
    }

    public static (bool Success, string Message, string? Username) VerifyCredential(string xmlPath)
    {
        var res = ReadPlaintextCredential(xmlPath);
        if (!res.Success)
        {
            return (false, res.Message, null);
        }

        return (true, "Credencial DPAPI ativa e pronta para uso nas automações!", res.Username);
    }

    public static (bool Success, string Message) SaveCredential(string xmlPath, string username, string plainPassword)
    {
        if (string.IsNullOrWhiteSpace(username))
        {
            return (false, "O nome de usuário não pode estar em branco.");
        }

        if (string.IsNullOrEmpty(plainPassword))
        {
            return (false, "A senha não pode estar em branco.");
        }

        try
        {
            string? dir = Path.GetDirectoryName(xmlPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            byte[] plainBytes = Encoding.Unicode.GetBytes(plainPassword);
            byte[] encryptedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
            string hexEncrypted = Convert.ToHexString(encryptedBytes).ToLowerInvariant();

            string cleanUser = System.Security.SecurityElement.Escape(username.Trim());

            string clixml =
$@"<Objs Version=""1.1.0.1"" xmlns=""http://schemas.microsoft.com/powershell/2004/04"">
  <Obj RefId=""0"">
    <TN RefId=""0"">
      <T>System.Management.Automation.PSCredential</T>
      <T>System.Object</T>
    </TN>
    <ToString>System.Management.Automation.PSCredential</ToString>
    <Props>
      <S N=""UserName"">{cleanUser}</S>
      <SS N=""Password"">{hexEncrypted}</SS>
    </Props>
  </Obj>
</Objs>";

            File.WriteAllText(xmlPath, clixml, new UTF8Encoding(false));
            ConfigService.HardenFilePermissions(xmlPath);

            // Teste imediato de validação
            var test = VerifyCredential(xmlPath);
            if (!test.Success)
            {
                return (false, $"O arquivo foi gravado, mas a validação falhou: {test.Message}");
            }

            return (true, $"Cofre seguro gravado e validado em: {xmlPath}");
        }
        catch (Exception ex)
        {
            return (false, $"Erro ao gravar credencial no cofre: {ex.Message}");
        }
    }
}

