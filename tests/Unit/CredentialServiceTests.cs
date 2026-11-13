using System.IO;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class CredentialServiceTests
{
    [Fact]
    public void VerifyCredential_ReturnsFalse_WhenFileDoesNotExist()
    {
        string nonExistentPath = Path.Combine(Path.GetTempPath(), $"inexistente_{Guid.NewGuid():N}.xml");
        var (success, message, username) = CredentialService.VerifyCredential(nonExistentPath);

        Assert.False(success);
        Assert.Contains("não encontrado", message);
        Assert.Null(username);
    }

    [Theory]
    [InlineData("", "SenhaForte123!")]
    [InlineData("   ", "SenhaForte123!")]
    [InlineData("usuario.teste", "")]
    public void SaveCredential_FailsClosed_WhenUsernameOrPasswordIsEmpty(string user, string pass)
    {
        string tempXml = Path.Combine(Path.GetTempPath(), $"cred_invalid_{Guid.NewGuid():N}.xml");
        var (success, _) = CredentialService.SaveCredential(tempXml, user, pass);

        Assert.False(success);
        Assert.False(File.Exists(tempXml));
    }

    [Fact]
    public void ReadPlaintextCredential_RejectsXmlWithDtd_PreventingXxe()
    {
        string tempXml = Path.Combine(Path.GetTempPath(), $"cred_xxe_{Guid.NewGuid():N}.xml");
        const string xxePayload = """
            <?xml version="1.0" encoding="utf-8"?>
            <!DOCTYPE foo [ <!ENTITY xxe SYSTEM "file:///C:/Windows/win.ini"> ]>
            <Objs xmlns="http://schemas.microsoft.com/powershell/2004/04">
              <S N="UserName">&xxe;</S>
              <SS N="Password">00</SS>
            </Objs>
            """;

        try
        {
            File.WriteAllText(tempXml, xxePayload);
            var (success, _, username, password) = CredentialService.ReadPlaintextCredential(tempXml);

            Assert.False(success);
            Assert.Null(username);
            Assert.Null(password);
        }
        finally
        {
            if (File.Exists(tempXml)) File.Delete(tempXml);
        }
    }

    [Fact]
    public void ReadPlaintextCredential_RejectsOversizedFile_PreventingXmlDoS()
    {
        string tempXml = Path.Combine(Path.GetTempPath(), $"cred_big_{Guid.NewGuid():N}.xml");
        try
        {
            File.WriteAllText(tempXml, new string('A', 70 * 1024));
            var (success, msg, _, _) = CredentialService.ReadPlaintextCredential(tempXml);

            Assert.False(success);
            Assert.Contains("tamanho máximo", msg);
        }
        finally
        {
            if (File.Exists(tempXml)) File.Delete(tempXml);
        }
    }

    [Fact]
    public void SaveAndVerifyCredential_Roundtrip_WorksWithCurrentDpapiScope()
    {
        string tempXml = Path.Combine(Path.GetTempPath(), $"cred_test_{Guid.NewGuid():N}.xml");
        string testUser = "teste_unitario";
        string testPass = "SenhaForte123!";

        try
        {
            var saveResult = CredentialService.SaveCredential(tempXml, testUser, testPass);
            Assert.True(saveResult.Success);
            Assert.True(File.Exists(tempXml));

            // Senha em texto plano nunca deve constar no XML gravado
            string rawXml = File.ReadAllText(tempXml);
            Assert.DoesNotContain(testPass, rawXml);

            var verifyResult = CredentialService.VerifyCredential(tempXml);
            Assert.True(verifyResult.Success);
            Assert.Equal(testUser, verifyResult.Username);
        }
        finally
        {
            if (File.Exists(tempXml))
            {
                try { File.Delete(tempXml); } catch { }
            }
        }
    }
}
