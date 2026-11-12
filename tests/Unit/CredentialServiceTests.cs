using System.IO;
using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class CredentialServiceTests
{
    [Fact]
    public void VerifyCredential_ReturnsFalse_WhenFileDoesNotExist()
    {
        string nonExistentPath = @"C:\AmbientesVirtuais\Credenciais\inexistente_12345.xml";
        var (success, message, username) = CredentialService.VerifyCredential(nonExistentPath);

        Assert.False(success);
        Assert.Contains("não encontrado", message);
        Assert.Null(username);
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

