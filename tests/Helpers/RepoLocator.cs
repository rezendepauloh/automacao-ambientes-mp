using System.Diagnostics;
using System.IO;

namespace AmbientesHub.Tests.Helpers;

/// <summary>Localiza a raiz do repositório e enumera os arquivos rastreados pelo Git (git ls-files).</summary>
public static class RepoLocator
{
    public static string Root { get; } = FindRoot();

    private static string FindRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")) &&
                Directory.Exists(Path.Combine(dir.FullName, "AmbientesHub")))
            {
                return dir.FullName;
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Raiz do repositório (AGENTS.md + AmbientesHub/) não localizada a partir de " + AppContext.BaseDirectory);
    }

    /// <summary>Caminhos relativos (separador '/') de todos os arquivos rastreados.</summary>
    public static IReadOnlyList<string> TrackedFiles()
    {
        var viaGit = TryGitLsFiles();
        if (viaGit.Count > 0) return viaGit;

        // Fallback (sem git no PATH): enumera o disco ignorando diretórios não versionados.
        string[] skipDirs = ["/.git/", "/bin/", "/obj/", "/temp/", "/Credenciais/", "/logs/", "/Atalhos/", "/TestResults/", "/.vs/"];
        return Directory.EnumerateFiles(Root, "*", SearchOption.AllDirectories)
            .Select(f => Path.GetRelativePath(Root, f).Replace('\\', '/'))
            .Where(rel => !skipDirs.Any(s => ("/" + rel).Contains(s, StringComparison.OrdinalIgnoreCase)))
            .Where(rel => !rel.EndsWith(".log", StringComparison.OrdinalIgnoreCase))
            .Where(rel => !rel.Equals(".env", StringComparison.OrdinalIgnoreCase) && !rel.EndsWith("/.env", StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static List<string> TryGitLsFiles()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "git",
                WorkingDirectory = Root,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8
            };
            psi.ArgumentList.Add("-c");
            psi.ArgumentList.Add("core.quotepath=off");
            psi.ArgumentList.Add("ls-files");
            psi.ArgumentList.Add("-z");

            using var p = Process.Start(psi);
            if (p == null) return [];
            string output = p.StandardOutput.ReadToEnd();
            if (!p.WaitForExit(20000) || p.ExitCode != 0) return [];

            return output.Split('\0', StringSplitOptions.RemoveEmptyEntries)
                .Where(rel => File.Exists(Path.Combine(Root, rel)))
                .ToList();
        }
        catch
        {
            return [];
        }
    }

    public static string FullPath(string relative) => Path.Combine(Root, relative.Replace('/', Path.DirectorySeparatorChar));
}

