using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

/// <summary>
/// Serviço nativo para verificação de instalação e provisionamento automatizado
/// de ferramentas administrativas (RSAT AD MMC, SCCM Console e pacotes via winget).
/// </summary>
public class SoftwareProvisioningService
{
    public const string DefaultSccmTargetDir = @"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole";

    private static readonly string[] SccmExecutableCandidates =
    [
        @"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe",
        @"C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe",
        @"C:\Program Files (x86)\ConfigMgr Console\bin\Microsoft.ConfigurationManagement.exe",
        @"C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\Microsoft.ConfigurationManagement.exe"
    ];

    public static List<SoftwarePackageItem> CreateDefaultCatalog()
    {
        return
        [
            new SoftwarePackageItem
            {
                Id = "rsat-ad-console",
                Name = "Active Directory Users and Computers (MMC / dsa.msc)",
                Category = "Administração AD • RSAT",
                Description = "Instala o RSAT Active Directory nativo via DISM (WindowsCapability) com fallback automático para o pacote .msu corporativo e cria atalho no Menu Iniciar.",
                Method = SoftwareInstallMethod.RsatAdConsole,
                IconEmoji = "🏛️",
                RequiresAdmin = true
            },
            new SoftwarePackageItem
            {
                Id = "sccm-console",
                Name = "Configuration Manager Console (SCCM / MECM)",
                Category = "Gestão Corporativa • SCCM",
                Description = "Localiza o instalador oficial consolesetup.exe no servidor do site SCCM configurado e realiza a instalação local do console e módulo PowerShell.",
                Method = SoftwareInstallMethod.SccmConsole,
                IconEmoji = "💻",
                RequiresAdmin = true
            },
            new SoftwarePackageItem
            {
                Id = "mozilla-firefox",
                Name = "Mozilla Firefox",
                Category = "Navegador • Web",
                Description = "Navegador padrão do AmbientesHub para abertura de chamados, portais corporativos e isolamento de perfis.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "Mozilla.Firefox",
                IconEmoji = "🦊",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "windows-terminal",
                Name = "Windows Terminal",
                Category = "Desenvolvimento • Terminal",
                Description = "Terminal moderno do Windows com abas, suporte a PowerShell 7, WSL e temas com Nerd Fonts.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "Microsoft.WindowsTerminal",
                IconEmoji = "💻",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "powershell-7",
                Name = "PowerShell 7 (pwsh)",
                Category = "Automação • Shell",
                Description = "Versão moderna e multiplataforma do PowerShell baseada no .NET Core para automações de infraestrutura.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "Microsoft.PowerShell",
                IconEmoji = "⚡",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "microsoft-powertoys",
                Name = "Microsoft PowerToys",
                Category = "Produtividade • Sistema",
                Description = "Conjunto de utilitários avançados do Windows (PowerToys Run, FancyZones, Text Extractor, Always on Top).",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "Microsoft.PowerToys",
                IconEmoji = "🛠️",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "microsoft-wsl",
                Name = "Windows Subsystem for Linux (WSL)",
                Category = "Desenvolvimento • Linux",
                Description = "Kernel Linux integrado ao Windows para rodar Docker, servidores locais e bancadas de desenvolvimento.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "Microsoft.WSL",
                IconEmoji = "🐧",
                RequiresAdmin = true
            },
            new SoftwarePackageItem
            {
                Id = "voidtools-everything",
                Name = "Everything (voidtools)",
                Category = "Busca • Utilitário",
                Description = "Localizador instantâneo de arquivos e pastas para o Windows indexado em tempo real via NTFS MFT.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "voidtools.Everything",
                IconEmoji = "🔍",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "everything-toolbar",
                Name = "Everything Toolbar",
                Category = "Busca • Barra de Tarefas",
                Description = "Integração do motor de pesquisa Everything diretamente na barra de tarefas do Windows 10/11.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "srwi.everythingtoolbar.launcher",
                IconEmoji = "🔎",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "copyq-clipboard",
                Name = "CopyQ Clipboard Manager",
                Category = "Produtividade • Área de Transferência",
                Description = "Gerenciador avançado da área de transferência com histórico pesquisável, abas e comandos personalizados.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "hluk.CopyQ",
                IconEmoji = "📋",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "espanso-text-expander",
                Name = "Espanso (Text Expander)",
                Category = "Produtividade • Automação",
                Description = "Expansor de texto veloz para atalhos de preenchimento de chamados, e-mails e respostas padrão.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "Espanso.Espanso",
                IconEmoji = "🚀",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "kando-menu",
                Name = "Kando (Pie Menu)",
                Category = "Navegação • Menus Radiais",
                Description = "Menu radial flutuante e ágil acionado por atalho de teclado para disparar apps e rotinas no Windows.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "KandoMenu.Kando",
                IconEmoji = "🥧",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "altsnap-window-manager",
                Name = "AltSnap",
                Category = "Janelas • Produtividade",
                Description = "Permite arrastar e redimensionar qualquer janela no Windows simplesmente segurando a tecla Alt.",
                Method = SoftwareInstallMethod.Winget,
                WingetId = "AltSnap.AltSnap",
                IconEmoji = "🪟",
                RequiresAdmin = false
            },
            new SoftwarePackageItem
            {
                Id = "artcraft-launcher",
                Name = "Artcraft Launcher",
                Category = "Utilitários • Launcher",
                Description = "Gerenciador e inicializador Artcraft para estações e ferramentas customizadas (via instalador .msi oficial do GitHub).",
                Method = SoftwareInstallMethod.DirectMsi,
                DownloadUrl = "https://github.com/storytold/craft-launcher/releases/download/v0.19.2/artcraft-launcher-0.19.2-windows-x64.msi",
                LocalInstallPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\artcraft-launcher\Artcraft Launcher.exe"),
                UninstallMatchName = "Artcraft Launcher",
                IconEmoji = "🎨",
                RequiresAdmin = false
            }
        ];
    }

    public static string? FindAdConsoleMscPath()
    {
        string sys32 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"System32\dsa.msc");
        if (File.Exists(sys32)) return sys32;

        string sysNative = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"Sysnative\dsa.msc");
        if (File.Exists(sysNative)) return sysNative;

        return null;
    }

    public static string? FindSccmConsoleExecutablePath()
    {
        foreach (string path in SccmExecutableCandidates)
        {
            if (File.Exists(path)) return path;
        }

        string? psd1 = AdSccmService.FindSccmModulePath();
        if (!string.IsNullOrEmpty(psd1))
        {
            string? binDir = Path.GetDirectoryName(psd1);
            if (!string.IsNullOrEmpty(binDir))
            {
                string exeNearModule = Path.Combine(binDir, "Microsoft.ConfigurationManagement.exe");
                if (File.Exists(exeNearModule)) return exeNearModule;
                return psd1;
            }
        }

        return null;
    }

    public static string BuildWingetInstallArguments(string wingetId, bool silent = true)
    {
        if (string.IsNullOrWhiteSpace(wingetId))
            throw new ArgumentException("ID do pacote winget não informado.", nameof(wingetId));

        string cleanId = wingetId.Trim();
        var args = new StringBuilder($"install --id \"{cleanId}\" --exact --source winget --accept-package-agreements --accept-source-agreements");
        
        // Se for AltSnap, força escopo de máquina (--scope machine) para instalar para todos os usuários do sistema
        if (cleanId.Equals("AltSnap.AltSnap", StringComparison.OrdinalIgnoreCase))
        {
            args.Append(" --scope machine");
        }

        if (silent)
            args.Append(" --silent");

        return args.ToString();
    }

    public static string BuildSccmSetupArguments(string siteServer, string targetDir = DefaultSccmTargetDir, bool silent = true)
    {
        if (string.IsNullOrWhiteSpace(siteServer))
            throw new ArgumentException("Servidor SCCM não informado.", nameof(siteServer));

        string core = $"TargetDir=\"{targetDir}\" DefaultSiteServerName={siteServer.Trim()}";
        return silent ? $"/q {core}" : core;
    }

    public static (string? capabilityName, bool isInstalled) ParseRsatCapabilityFromDism(string dismOutput)
    {
        if (string.IsNullOrWhiteSpace(dismOutput))
            return (null, false);

        string? currentCap = null;
        foreach (string rawLine in dismOutput.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
        {
            string line = rawLine.Trim();
            if (line.StartsWith("Capability Identity", StringComparison.OrdinalIgnoreCase) ||
                line.StartsWith("Identidade da Capacidade", StringComparison.OrdinalIgnoreCase))
            {
                int colon = line.IndexOf(':');
                if (colon >= 0)
                {
                    string candidate = line[(colon + 1)..].Trim();
                    currentCap = candidate.Contains("Rsat.ActiveDirectory.DS-LDS.Tools", StringComparison.OrdinalIgnoreCase)
                        ? candidate
                        : null;
                }
            }
            else if (currentCap != null &&
                     (line.StartsWith("State", StringComparison.OrdinalIgnoreCase) ||
                      line.StartsWith("Estado", StringComparison.OrdinalIgnoreCase)))
            {
                int colon = line.IndexOf(':');
                string stateVal = colon >= 0 ? line[(colon + 1)..].Trim() : string.Empty;
                bool installed = stateVal.Equals("Installed", StringComparison.OrdinalIgnoreCase) ||
                                 stateVal.Equals("Instalado", StringComparison.OrdinalIgnoreCase);
                return (currentCap, installed);
            }
        }

        return (null, false);
    }

    public async Task RefreshStatusAsync(IEnumerable<SoftwarePackageItem> catalog)
    {
        var items = catalog.ToList();
        await Task.Run(() =>
        {
            var updateIds = GetWingetUpgradablePackageIds();
            foreach (var item in items)
            {
                RefreshSingleItem(item, updateIds);
            }
        });
    }

    public void RefreshSingleItem(SoftwarePackageItem item, HashSet<string>? knownUpdateIds = null)
    {
        switch (item.Method)
        {
            case SoftwareInstallMethod.RsatAdConsole:
            {
                string? msc = FindAdConsoleMscPath();
                item.IsInstalled = msc != null;
                item.HasUpdate = false;
                item.StatusDetail = msc != null
                    ? $"Detectado em: {msc}"
                    : "Não localizado no System32 (dsa.msc ausente).";
                break;
            }
            case SoftwareInstallMethod.SccmConsole:
            {
                string? exe = FindSccmConsoleExecutablePath();
                item.IsInstalled = exe != null;
                item.HasUpdate = false;
                item.StatusDetail = exe != null
                    ? $"Detectado em: {exe}"
                    : "Console AdminConsole não localizado na estação.";
                break;
            }
            case SoftwareInstallMethod.Winget:
            {
                string wId = item.WingetId ?? string.Empty;
                bool installed = !string.IsNullOrWhiteSpace(wId) && CheckWingetInstalled(wId);

                // Fallbacks locais para pacotes instalados por usuário Admin ou executáveis diretos (ex: AltSnap, EverythingToolbar)
                string? detectedCustomPath = null;
                if (!installed && item.Id == "altsnap-window-manager")
                {
                    detectedCustomPath = FindAltSnapInstalledPath();
                    if (detectedCustomPath != null)
                    {
                        installed = true;
                    }
                }
                else if (!installed && item.Id == "everything-toolbar")
                {
                    detectedCustomPath = FindEverythingToolbarInstalledPath();
                    if (detectedCustomPath != null)
                    {
                        installed = true;
                    }
                }

                item.IsInstalled = installed;

                bool hasUpdate = false;
                if (installed)
                {
                    if (knownUpdateIds != null)
                    {
                        hasUpdate = knownUpdateIds.Contains(wId);
                    }
                    else
                    {
                        hasUpdate = CheckSingleWingetUpgradeAvailable(wId);
                    }
                }
                item.HasUpdate = hasUpdate;

                if (!installed)
                {
                    item.StatusDetail = $"Pacote {wId} não encontrado localmente.";
                }
                else if (detectedCustomPath != null)
                {
                    item.StatusDetail = $"Detectado em: {detectedCustomPath}";
                }
                else if (hasUpdate)
                {
                    item.StatusDetail = $"Nova versão disponível para atualização via winget!";
                }
                else
                {
                    item.StatusDetail = $"Instalado e atualizado na versão mais recente.";
                }
                break;
            }
            case SoftwareInstallMethod.DirectMsi:
            {
                bool installed = CheckDirectAppInstalled(item);
                item.IsInstalled = installed;
                item.HasUpdate = false;
                item.StatusDetail = installed
                    ? "Instalado no sistema (Artcraft Launcher)."
                    : "Não instalado (disponível via .msi oficial do GitHub).";
                break;
            }
        }
    }

    public static string? FindAltSnapInstalledPath()
    {
        // 1. Usuário atual
        string currentUserRoaming = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            @"AltSnap\AltSnap.exe");
        if (File.Exists(currentUserRoaming)) return currentUserRoaming;

        // 2. Program Files
        string pf64 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"AltSnap\AltSnap.exe");
        if (File.Exists(pf64)) return pf64;
        string pf32 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"AltSnap\AltSnap.exe");
        if (File.Exists(pf32)) return pf32;

        // 3. Perfis em C:\Users (ex: paulo_admin ou outros perfis com privilégio elevado)
        string usersRoot = Path.GetPathRoot(Environment.SystemDirectory) ?? @"C:\";
        string usersDir = Path.Combine(usersRoot, "Users");
        if (Directory.Exists(usersDir))
        {
            try
            {
                foreach (var userFolder in Directory.GetDirectories(usersDir))
                {
                    string candidate = Path.Combine(userFolder, @"AppData\Roaming\AltSnap\AltSnap.exe");
                    if (File.Exists(candidate)) return candidate;
                }
            }
            catch { }
        }

        return null;
    }

    public static string? FindEverythingToolbarInstalledPath()
    {
        // 1. Program Files / Program Files (x86)
        string pf64 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"EverythingToolbar\EverythingToolbar.Launcher.exe");
        if (File.Exists(pf64)) return pf64;
        string pf32 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"EverythingToolbar\EverythingToolbar.Launcher.exe");
        if (File.Exists(pf32)) return pf32;

        // Variações com EverythingToolbar.exe
        string pf64Alt = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"EverythingToolbar\EverythingToolbar.exe");
        if (File.Exists(pf64Alt)) return pf64Alt;
        string pf32Alt = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"EverythingToolbar\EverythingToolbar.exe");
        if (File.Exists(pf32Alt)) return pf32Alt;

        // 2. LocalAppData / AppData dos usuários
        string localApp = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\EverythingToolbar\EverythingToolbar.Launcher.exe");
        if (File.Exists(localApp)) return localApp;

        string usersRoot = Path.GetPathRoot(Environment.SystemDirectory) ?? @"C:\";
        string usersDir = Path.Combine(usersRoot, "Users");
        if (Directory.Exists(usersDir))
        {
            try
            {
                foreach (var userFolder in Directory.GetDirectories(usersDir))
                {
                    string cand1 = Path.Combine(userFolder, @"AppData\Local\Programs\EverythingToolbar\EverythingToolbar.Launcher.exe");
                    if (File.Exists(cand1)) return cand1;
                    string cand2 = Path.Combine(userFolder, @"AppData\Roaming\EverythingToolbar\EverythingToolbar.Launcher.exe");
                    if (File.Exists(cand2)) return cand2;
                }
            }
            catch { }
        }

        // 3. Registro de Desinstalação do Windows
        if (IsAppInUninstallRegistry("EverythingToolbar") || IsAppInUninstallRegistry("Everything Toolbar"))
        {
            return "Registro do Windows (EverythingToolbar)";
        }

        return null;
    }

    private static bool CheckDirectAppInstalled(SoftwarePackageItem item)
    {
        if (!string.IsNullOrWhiteSpace(item.LocalInstallPath) && File.Exists(item.LocalInstallPath))
            return true;

        // Procura também na pasta padrão de programas
        string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        string progFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        string[] possiblePaths =
        [
            Path.Combine(programFiles, @"Artcraft Launcher\Artcraft Launcher.exe"),
            Path.Combine(progFilesX86, @"Artcraft Launcher\Artcraft Launcher.exe"),
            Path.Combine(localAppData, @"Programs\artcraft-launcher\Artcraft Launcher.exe"),
            Path.Combine(localAppData, @"Programs\Artcraft Launcher\Artcraft Launcher.exe")
        ];

        if (possiblePaths.Any(File.Exists))
            return true;

        // Procura no registro de desinstalação (HKCU e HKLM Uninstall)
        return IsAppInUninstallRegistry(item.UninstallMatchName ?? item.Name);
    }

    private static bool IsAppInUninstallRegistry(string displayNamePart)
    {
        try
        {
            string[] uninstallKeys =
            [
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            ];

            foreach (var keyPath in uninstallKeys)
            {
                using var hklmKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(keyPath);
                if (hklmKey != null)
                {
                    foreach (var subKeyName in hklmKey.GetSubKeyNames())
                    {
                        using var subKey = hklmKey.OpenSubKey(subKeyName);
                        string? dispName = subKey?.GetValue("DisplayName")?.ToString();
                        if (!string.IsNullOrEmpty(dispName) && dispName.Contains(displayNamePart, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }

                using var hkcuKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyPath);
                if (hkcuKey != null)
                {
                    foreach (var subKeyName in hkcuKey.GetSubKeyNames())
                    {
                        using var subKey = hkcuKey.OpenSubKey(subKeyName);
                        string? dispName = subKey?.GetValue("DisplayName")?.ToString();
                        if (!string.IsNullOrEmpty(dispName) && dispName.Contains(displayNamePart, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }
        }
        catch { }
        return false;
    }

    public static HashSet<string> GetWingetUpgradablePackageIds()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var res = RunProcess("winget.exe", "upgrade --source winget", 25000, logOutput: false);
            if (!string.IsNullOrWhiteSpace(res.stdout))
            {
                foreach (string line in res.stdout.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
                {
                    string trimmed = line.Trim();
                    foreach (var pkg in CreateDefaultCatalog().Where(p => !string.IsNullOrEmpty(p.WingetId)))
                    {
                        if (trimmed.Contains(pkg.WingetId!, StringComparison.OrdinalIgnoreCase))
                        {
                            set.Add(pkg.WingetId!);
                        }
                    }
                }
            }
        }
        catch { }
        return set;
    }

    public static bool CheckSingleWingetUpgradeAvailable(string wingetId)
    {
        try
        {
            var res = RunProcess("winget.exe", $"upgrade --id \"{wingetId.Trim()}\" --exact --source winget", 20000, logOutput: false);
            return res.exitCode == 0 && res.stdout.Contains(wingetId.Trim(), StringComparison.OrdinalIgnoreCase) && !res.stdout.Contains("No applicable update found", StringComparison.OrdinalIgnoreCase) && !res.stdout.Contains("Nenhuma atualização aplicável encontrada", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    public async Task<bool> UninstallSoftwareAsync(SoftwarePackageItem item, bool silent = true, Action<string>? onLog = null)
    {
        return await Task.Run(() => UninstallSoftwareCore(item, silent, onLog));
    }

    private bool UninstallSoftwareCore(SoftwarePackageItem item, bool silent, Action<string>? onLog)
    {
        void Log(string msg)
        {
            onLog?.Invoke(msg);
            AppLoggerService.Info(msg, nameof(SoftwareProvisioningService));
        }

        Log("==========================================================");
        Log($"=== DESINSTALAÇÃO DE PACOTE: {item.Name} ===");
        Log("==========================================================");

        bool ok = false;
        switch (item.Method)
        {
            case SoftwareInstallMethod.Winget:
            {
                string args = $"uninstall --id \"{item.WingetId}\" --exact --source winget";
                if (silent) args += " --silent";
                Log($"-> Executando: winget {args}");
                var res = RunProcess("winget.exe", args, 300000, Log);
                ok = res.exitCode == 0;
                Log(ok
                    ? $"[SUCESSO] {item.Name} desinstalado com sucesso!"
                    : $"[AVISO] winget uninstall retornou código {res.exitCode}.");
                break;
            }
            case SoftwareInstallMethod.RsatAdConsole:
            {
                if (!SystemPerformanceService.IsUserAdministrator())
                {
                    Log("[ERRO] Desinstalar RSAT AD requer privilégios de Administrador.");
                    return false;
                }
                Log("-> Removendo capacidade RSAT Active Directory via DISM...");
                var caps = RunProcess("dism.exe", "/Online /Get-Capabilities", 30000, logOutput: false);
                var (capName, _) = ParseRsatCapabilityFromDism(caps.stdout);
                if (!string.IsNullOrWhiteSpace(capName))
                {
                    var rem = RunProcess("dism.exe", $"/Online /Remove-Capability /CapabilityName:{capName}", 180000, Log);
                    ok = rem.exitCode == 0;
                }
                else
                {
                    Log("  -> Capacidade RSAT não encontrada no DISM.");
                }
                break;
            }
            case SoftwareInstallMethod.SccmConsole:
            {
                Log("-> Desinstalando Configuration Manager Console via consolesetup /uninstall...");
                string siteServer = ConfigService.SccmSiteServer.Trim();
                string siteCode = ConfigService.SccmSiteCode.Trim();

                if (string.IsNullOrWhiteSpace(siteServer) || string.IsNullOrWhiteSpace(siteCode))
                {
                    Log("[ERRO] SCCM_SITE_SERVER ou SCCM_SITE_CODE não configurados no .env.");
                    ok = false;
                    break;
                }

                string installer = $@"\\{siteServer}\SMS_{siteCode}\bin\i386\consolesetup.exe";
                if (File.Exists(installer))
                {
                    var res = RunProcess(installer, "/uninstall /q", 180000, Log);
                    ok = res.exitCode == 0;
                }
                else
                {
                    Log("[ERRO] Instalador do SCCM inacessível para desinstalação.");
                }
                break;
            }
            case SoftwareInstallMethod.DirectMsi:
            {
                ok = UninstallMsiDirectCore(item, silent, Log);
                break;
            }
        }

        RefreshSingleItem(item);
        return ok;
    }

    public async Task<bool> InstallSoftwareAsync(SoftwarePackageItem item, bool silent = true, Action<string>? onLog = null)
    {
        return await Task.Run(() => InstallSoftwareCore(item, silent, onLog));
    }

    private bool InstallSoftwareCore(SoftwarePackageItem item, bool silent, Action<string>? onLog)
    {
        void Log(string msg)
        {
            onLog?.Invoke(msg);
            AppLoggerService.Info(msg, nameof(SoftwareProvisioningService));
        }

        bool ok = item.Method switch
        {
            SoftwareInstallMethod.RsatAdConsole => InstallAdConsoleCore(silent, Log),
            SoftwareInstallMethod.SccmConsole => InstallSccmConsoleCore(silent, Log),
            SoftwareInstallMethod.Winget => InstallWingetCore(item.WingetId, silent, Log),
            SoftwareInstallMethod.DirectMsi => InstallMsiDirectCore(item, silent, Log),
            _ => false
        };

        RefreshSingleItem(item);
        return ok || item.IsInstalled;
    }

    private bool InstallAdConsoleCore(bool silent, Action<string> log)
    {
        log("==========================================================");
        log("=== INSTALAÇÃO DO ACTIVE DIRECTORY MMC (RSAT AD) ===");
        log("==========================================================");

        if (FindAdConsoleMscPath() != null)
        {
            log("[OK] dsa.msc já está instalado neste computador.");
            EnsureAdShortcut(log);
            return true;
        }

        if (!SystemPerformanceService.IsUserAdministrator())
        {
            log("[ERRO] Instalar o RSAT / pacote .msu exige privilégios de Administrador.");
            log("[i] Dica: clique em 'Reiniciar como Admin' no canto superior direito desta aba.");
            return false;
        }

        bool installed = false;

        // Método 1: DISM WindowsCapability (Nativo Windows 10/11)
        log("[1/3] Verificando Recurso Opcional Nativo (Rsat.ActiveDirectory.DS-LDS.Tools) via DISM...");
        try
        {
            var caps = RunProcess("dism.exe", "/Online /Get-Capabilities", 45000, logOutput: false);
            var (capName, alreadyInstalled) = ParseRsatCapabilityFromDism(caps.stdout);

            if (alreadyInstalled)
            {
                log($"  -> [OK] Capacidade {capName} já marcada como instalada.");
                installed = true;
            }
            else if (!string.IsNullOrWhiteSpace(capName))
            {
                log($"  -> Instalando capacidade nativa '{capName}' (isso pode levar alguns minutos)...");
                var addRes = RunProcess("dism.exe", $"/Online /Add-Capability /CapabilityName:{capName}", 300000, log);
                if (addRes.exitCode == 0 || addRes.exitCode == 3010)
                {
                    log("  -> [OK] RSAT Active Directory instalado via WindowsCapability!");
                    installed = true;
                }
                else
                {
                    log($"  -> [AVISO] DISM retornou código {addRes.exitCode}. Tentando fallback via pacote .msu corporativo...");
                }
            }
            else
            {
                log("  -> [AVISO] Capacidade RSAT AD não listada online pelo DISM. Recorrendo ao pacote .msu...");
            }
        }
        catch (Exception ex)
        {
            log($"  -> [AVISO] Falha ao consultar DISM Capabilities: {ex.Message}");
        }

        // Método 2: Fallback pacote .msu na rede corporativa
        if (!installed && FindAdConsoleMscPath() == null)
        {
            string networkMsu = ConfigService.RsatMsuNetworkPath;
            log("[2/3] Verificando pacote .msu no compartilhamento de rede...");
            log($"      Caminho: {networkMsu}");

            if (File.Exists(networkMsu))
            {
                string tempMsu = Path.Combine(Path.GetTempPath(), "WindowsTH-RSAT-AD.msu");
                try
                {
                    log($"  -> Copiando pacote para {tempMsu}...");
                    File.Copy(networkMsu, tempMsu, overwrite: true);

                    string wusaArgs = silent
                        ? $"\"{tempMsu}\" /quiet /norestart"
                        : $"\"{tempMsu}\" /norestart";

                    log($"  -> Executando wusa.exe {(silent ? "(modo silencioso)" : "(modo interativo)")}...");
                    var wusaRes = RunProcess("wusa.exe", wusaArgs, 300000, log);
                    if (wusaRes.exitCode == 0 || wusaRes.exitCode == 3010)
                    {
                        log(wusaRes.exitCode == 3010
                            ? "  -> [OK] Pacote MSU instalado! Reinicialização pendente para concluir ativação."
                            : "  -> [OK] Pacote MSU instalado com sucesso!");
                        installed = true;
                    }
                    else
                    {
                        log($"  -> [ERRO] wusa.exe finalizou com código {wusaRes.exitCode}.");
                    }
                }
                catch (Exception ex)
                {
                    log($"  -> [ERRO] Falha ao instalar pacote MSU: {ex.Message}");
                }
                finally
                {
                    try { if (File.Exists(tempMsu)) File.Delete(tempMsu); } catch { }
                }
            }
            else
            {
                log("[ERRO] Pacote .msu não localizado na rede. Verifique a conexão corporativa/VPN.");
            }
        }

        // Método 3: Atalho no Menu Iniciar
        log("[3/3] Configurando atalho no Menu Iniciar...");
        EnsureAdShortcut(log);

        bool finalCheck = FindAdConsoleMscPath() != null || installed;
        log(finalCheck
            ? "[SUCESSO] Provisionamento do Active Directory MMC concluído! (Win+R -> dsa.msc)"
            : "[AVISO] Não foi possível confirmar dsa.msc após a execução.");
        return finalCheck;
    }

    private static void EnsureAdShortcut(Action<string> log)
    {
        string? dsaMsc = FindAdConsoleMscPath();
        if (dsaMsc == null)
        {
            log("  -> [AVISO] dsa.msc ainda não encontrado no System32 (pode exigir reinicialização se instalado via MSU).");
            return;
        }

        object? wshShell = null;
        object? shortcut = null;
        try
        {
            string progData = Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms);
            string adminTools = Path.Combine(progData, "Administrative Tools");
            string targetFolder = Directory.Exists(adminTools) ? adminTools : progData;
            string shortcutPath = Path.Combine(targetFolder, "Active Directory Users and Computers.lnk");

            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                log("  -> [AVISO] WScript.Shell indisponível para criar atalho.");
                return;
            }

            wshShell = Activator.CreateInstance(shellType);
            if (wshShell == null) return;

            dynamic dynShell = wshShell;
            shortcut = dynShell.CreateShortcut(shortcutPath);
            dynamic dynShortcut = shortcut;
            dynShortcut.TargetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "mmc.exe");
            dynShortcut.Arguments = $"\"{dsaMsc}\"";
            dynShortcut.Description = "Active Directory Users and Computers (MMC)";
            dynShortcut.IconLocation = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "dsadmin.dll") + ",0";
            dynShortcut.Save();

            log($"  -> [OK] Atalho criado em: {shortcutPath}");
        }
        catch (Exception ex)
        {
            log($"  -> [AVISO] Não foi possível criar atalho: {ex.Message}");
        }
        finally
        {
            if (shortcut != null && Marshal.IsComObject(shortcut))
                Marshal.ReleaseComObject(shortcut);
            if (wshShell != null && Marshal.IsComObject(wshShell))
                Marshal.ReleaseComObject(wshShell);
        }
    }

    private bool InstallSccmConsoleCore(bool silent, Action<string> log)
    {
        log("==========================================================");
        log("=== INSTALAÇÃO DO CONFIGURATION MANAGER CONSOLE (SCCM) ===");
        log("==========================================================");

        if (FindSccmConsoleExecutablePath() != null)
        {
            log("[OK] Configuration Manager Console já está instalado neste computador.");
            return true;
        }

        string siteServer = ConfigService.SccmSiteServer.Trim();
        string siteCode = ConfigService.SccmSiteCode.Trim();

        if (string.IsNullOrWhiteSpace(siteServer) || string.IsNullOrWhiteSpace(siteCode))
        {
            log("[ERRO] Servidor ou Site Code do SCCM não configurados no arquivo .env (SCCM_SITE_SERVER / SCCM_SITE_CODE).");
            return false;
        }

        string installerPath = $@"\\{siteServer}\SMS_{siteCode}\bin\i386\consolesetup.exe";
        log($"[1/2] Verificando instalador corporativo em: {installerPath}");

        if (!File.Exists(installerPath))
        {
            log("[ERRO] Não foi possível acessar consolesetup.exe no compartilhamento de rede!");
            log("[i] Verifique se o computador está na rede corporativa / VPN ou ajuste SCCM_SITE_SERVER em Configurações.");
            return false;
        }

        log("  -> [OK] Instalador localizado! Copiando/executando...");
        string args = BuildSccmSetupArguments(siteServer, DefaultSccmTargetDir, silent);
        log($"[2/2] Executando consolesetup.exe {(silent ? "(modo silencioso /q)" : "(modo interativo)")}...");

        var res = RunProcess(installerPath, args, 600000, log);
        if (res.exitCode == 0)
        {
            log("[SUCESSO] Configuration Manager Console instalado com sucesso!");
            return true;
        }

        log($"[AVISO] consolesetup.exe encerrou com código {res.exitCode}. Consulte C:\\Program Files (x86)\\Microsoft Configuration Manager\\AdminConsole\\AdminConsole.log");
        return false;
    }

    private bool InstallWingetCore(string? wingetId, bool silent, Action<string> log)
    {
        if (string.IsNullOrWhiteSpace(wingetId))
        {
            log("[ERRO] Identificador do pacote winget não configurado.");
            return false;
        }

        log("==========================================================");
        log($"=== INSTALAÇÃO VIA WINGET (SOURCE: winget): {wingetId} ===");
        log("==========================================================");

        string args = BuildWingetInstallArguments(wingetId, silent);
        log($"-> Executando: winget {args}");

        var res = RunProcess("winget.exe", args, 600000, log);
        if (res.exitCode == 0)
        {
            log($"[SUCESSO] Pacote '{wingetId}' instalado via winget!");
            return true;
        }

        log($"[AVISO] winget retornou código {res.exitCode}.");
        return false;
    }

    private static bool InstallMsiDirectCore(SoftwarePackageItem item, bool silent, Action<string> log)
    {
        if (string.IsNullOrWhiteSpace(item.DownloadUrl))
        {
            log("[ERRO] URL de download do instalador .msi não informada.");
            return false;
        }

        log("==========================================================");
        log($"=== INSTALAÇÃO DIRETA VIA MSI: {item.Name} ===");
        log("==========================================================");

        string tempDir = Path.GetTempPath();
        string msiFileName = Path.GetFileName(new Uri(item.DownloadUrl).LocalPath);
        if (string.IsNullOrWhiteSpace(msiFileName)) msiFileName = $"{item.Id}.msi";
        string localMsiPath = Path.Combine(tempDir, msiFileName);

        try
        {
            log($"-> Baixando instalador oficial de: {item.DownloadUrl}...");
            using (var httpClient = new System.Net.Http.HttpClient())
            {
                httpClient.Timeout = TimeSpan.FromMinutes(5);
                var data = httpClient.GetByteArrayAsync(item.DownloadUrl).GetAwaiter().GetResult();
                File.WriteAllBytes(localMsiPath, data);
            }
            log($"  -> Download concluído: {localMsiPath} ({new FileInfo(localMsiPath).Length / (1024 * 1024.0):F2} MB)");

            string msiArgs = silent ? $"/i \"{localMsiPath}\" /qn /norestart" : $"/i \"{localMsiPath}\" /norestart";
            log($"-> Executando msiexec.exe {(silent ? "(modo silencioso /qn)" : "(modo interativo)")}...");

            var res = RunProcess("msiexec.exe", msiArgs, 300000, log);
            if (res.exitCode == 0 || res.exitCode == 3010)
            {
                log($"[SUCESSO] {item.Name} instalado com sucesso via MSI!");
                return true;
            }

            log($"[AVISO] msiexec.exe finalizou com código {res.exitCode}.");
            return false;
        }
        catch (Exception ex)
        {
            log($"[ERRO] Falha ao baixar ou executar instalador MSI: {ex.Message}");
            return false;
        }
        finally
        {
            try { if (File.Exists(localMsiPath)) File.Delete(localMsiPath); } catch { }
        }
    }

    private static bool UninstallMsiDirectCore(SoftwarePackageItem item, bool silent, Action<string> log)
    {
        log("-> Localizando identificador do produto MSI no registro para desinstalação...");
        string matchName = item.UninstallMatchName ?? item.Name;
        string? productCode = FindMsiProductCodeByName(matchName);

        if (string.IsNullOrWhiteSpace(productCode))
        {
            log($"[AVISO] Código do produto para '{matchName}' não encontrado no Registro de desinstalação.");
            return false;
        }

        log($"  -> ProductCode localizado: {productCode}");
        string msiArgs = silent ? $"/x {productCode} /qn /norestart" : $"/x {productCode} /norestart";
        log($"-> Executando msiexec.exe {msiArgs}...");

        var res = RunProcess("msiexec.exe", msiArgs, 300000, log);
        bool ok = res.exitCode == 0 || res.exitCode == 3010;
        log(ok
            ? $"[SUCESSO] {item.Name} desinstalado com sucesso!"
            : $"[AVISO] msiexec retornou código {res.exitCode}.");
        return ok;
    }

    private static string? FindMsiProductCodeByName(string displayNamePart)
    {
        try
        {
            string[] uninstallKeys =
            [
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            ];

            foreach (var keyPath in uninstallKeys)
            {
                using var hklmKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(keyPath);
                if (hklmKey != null)
                {
                    foreach (var subKeyName in hklmKey.GetSubKeyNames())
                    {
                        using var subKey = hklmKey.OpenSubKey(subKeyName);
                        string? dispName = subKey?.GetValue("DisplayName")?.ToString();
                        if (!string.IsNullOrEmpty(dispName) && dispName.Contains(displayNamePart, StringComparison.OrdinalIgnoreCase))
                        {
                            return subKeyName.StartsWith("{") && subKeyName.EndsWith("}") ? subKeyName : null;
                        }
                    }
                }

                using var hkcuKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyPath);
                if (hkcuKey != null)
                {
                    foreach (var subKeyName in hkcuKey.GetSubKeyNames())
                    {
                        using var subKey = hkcuKey.OpenSubKey(subKeyName);
                        string? dispName = subKey?.GetValue("DisplayName")?.ToString();
                        if (!string.IsNullOrEmpty(dispName) && dispName.Contains(displayNamePart, StringComparison.OrdinalIgnoreCase))
                        {
                            return subKeyName.StartsWith("{") && subKeyName.EndsWith("}") ? subKeyName : null;
                        }
                    }
                }
            }
        }
        catch { }
        return null;
    }

    private static bool CheckWingetInstalled(string wingetId)
    {
        var res = RunProcess("winget.exe", $"list --id \"{wingetId.Trim()}\" --exact --source winget", 15000, logOutput: false);
        return res.exitCode == 0 && res.stdout.Contains(wingetId.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static (int exitCode, string stdout, string stderr) RunProcess(
        string fileName,
        string arguments,
        int timeoutMs,
        Action<string>? onOutput = null,
        bool logOutput = true)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using var proc = new Process { StartInfo = psi };
            var stdoutSb = new StringBuilder();
            var stderrSb = new StringBuilder();

            proc.OutputDataReceived += (_, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                {
                    stdoutSb.AppendLine(e.Data);
                    if (logOutput) onOutput?.Invoke($"  {e.Data.Trim()}");
                }
            };
            proc.ErrorDataReceived += (_, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                {
                    stderrSb.AppendLine(e.Data);
                    if (logOutput) onOutput?.Invoke($"  [AVISO] {e.Data.Trim()}");
                }
            };

            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();

            if (!proc.WaitForExit(timeoutMs))
            {
                try { proc.Kill(entireProcessTree: true); } catch { }
                return (-1, stdoutSb.ToString(), "Tempo limite excedido.");
            }

            proc.WaitForExit();
            return (proc.ExitCode, stdoutSb.ToString(), stderrSb.ToString());
        }
        catch (Exception ex)
        {
            if (logOutput) onOutput?.Invoke($"  [ERRO] Falha ao iniciar {fileName}: {ex.Message}");
            return (-1, string.Empty, ex.Message);
        }
    }
}

