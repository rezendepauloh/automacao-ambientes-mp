using System.IO;
using System.Management;
using System.Text.RegularExpressions;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

public class DeviceAnalyzerService
{
    public static string DefaultOutputFolder => ConfigService.DeviceReportsDir;

    public DeviceAnalysisResult AnalyzeDevice(
        string computerName,
        bool skipMajorData = false,
        bool autoOpenHtmlInFirefox = true,
        string? outputFolder = null,
        Action<string>? onLog = null)
    {
        string target = RemoteManagementHelper.NormalizeTargetHost(computerName);
        string outDir = string.IsNullOrWhiteSpace(outputFolder) ? DefaultOutputFolder : outputFolder.Trim();
        var result = new DeviceAnalysisResult { OutputFolder = outDir };

        void Log(string msg) => onLog?.Invoke(msg);

        Log($"=== INICIANDO ANÁLISE COMPLETA DO DISPOSITIVO: {target} ===");

        if (!RemoteManagementHelper.PingHost(target))
        {
            result.Message = $"Máquina '{target}' não respondeu ao ping (ICMP).";
            Log($"[ERRO] {result.Message}");
            return result;
        }

        try
        {
            Directory.CreateDirectory(outDir);
            // Inventário é estritamente somente-leitura: não habilita WinRM nem altera Firewall/TrustedHosts no host remoto.
            RemoteManagementHelper.EnsureAdminShareConnected(target);

            var cimScope = RemoteManagementHelper.CreateScope(target, @"root\cimv2");
            var regScope = RemoteManagementHelper.CreateScope(target, @"root\default");

            var data = new DeviceReportData
            {
                ComputerName = target,
                SkipMajorData = skipMajorData
            };

            // 1. Hardware Principal, BIOS e Usuário Logado
            Log("[1/12] Coletando Hardware Principal, BIOS e Usuário Logado...");
            CollectHardwareBiosAndUser(cimScope, target, data);

            // 2. Sistema Operacional
            Log("[2/12] Coletando Sistema Operacional e Build (HnH)...");
            CollectOperatingSystem(cimScope, regScope, data);

            // 3. Processador, Placa-Mãe e Chassi
            Log("[3/12] Coletando Processador, Placa-Mãe e Chassi...");
            CollectCpuMotherboardAndChassis(cimScope, data);

            // 4. Memória RAM (JEDEC)
            Log("[4/12] Coletando Módulos de Memória RAM (JEDEC)...");
            CollectMemoryModules(cimScope, data);

            // 5. Monitores (EDID via root\wmi) e Placa de Vídeo
            Log("[5/12] Coletando Monitores (EDID) e Controlador de Vídeo...");
            CollectMonitorsAndVideo(cimScope, target, data);

            // 6. Discos Físicos (SSD/HDD) e Lógicos
            Log("[6/12] Coletando Discos Físicos (SSD/HDD) e Partições Lógicas...");
            CollectDisks(cimScope, target, data);

            // 7. Rede e Dispositivos de Som
            Log("[7/12] Coletando Adaptadores de Rede e Dispositivos de Áudio...");
            CollectNetworkAndSound(cimScope, data);

            // 8. Impressoras Locais e Mapeadas
            Log("[8/12] Coletando Impressoras Locais e Mapeadas por Usuário...");
            CollectPrinters(cimScope, regScope, target, data);

            if (!skipMajorData)
            {
                // 9. Perfis de Usuários no Disco
                Log("[9/12] Coletando Perfis de Usuários e tamanho em disco...");
                var profileSvc = new RemoteUserProfileService();
                data.UserProfiles = profileSvc.QueryUserProfiles(target, onLog: null, calculateDiskSize: true);

                // 10. Programas Instalados (Registro 64/32-bit)
                Log("[10/12] Coletando Programas Instalados via Registro...");
                CollectInstalledPrograms(regScope, data);

                // 11. Serviços em Execução e Processos Ativos
                Log("[11/12] Coletando Serviços em Execução e Processos Ativos...");
                CollectServicesAndProcesses(cimScope, data);

                // 12. Drivers PnP Assinados
                Log("[12/12] Coletando Drivers PnP Instalados...");
                CollectDrivers(cimScope, data);
            }
            else
            {
                Log("[9-12] Coleta estendida (Perfis/Programas/Serviços/Processos/Drivers) pulada conforme opção rápida.");
            }

            result.Data = data;

            // Nome seguro do relatório
            string rawName = $"{data.HardwareInfo.Computador}_{data.OS.NomeMaquina}";
            string safeName = Regex.Replace(rawName, @"[\\/:\*\?""<>|\s]+", "_").Trim('_');
            if (string.IsNullOrWhiteSpace(safeName)) safeName = target;
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            string htmlPath = Path.Combine(outDir, $"Report_{safeName}_{timestamp}.html");
            string pdfPath = Path.Combine(outDir, $"Report_{safeName}_{timestamp}.pdf");
            string excelPath = Path.Combine(outDir, $"Report_{safeName}_{timestamp}.xlsx");

            // A) Gerar HTML
            Log("-> Gerando relatório interativo HTML...");
            string htmlContent = DeviceReportExportService.BuildHtmlReport(data);
            File.WriteAllText(htmlPath, htmlContent, new System.Text.UTF8Encoding(false));
            result.HtmlPath = htmlPath;
            Log($"   [OK] HTML salvo em: {htmlPath}");

            // B) Abrir no Firefox automaticamente
            if (autoOpenHtmlInFirefox)
            {
                try
                {
                    Log("-> Abrindo relatório HTML no Mozilla Firefox...");
                    var browser = new BrowserLauncherService();
                    browser.OpenUrlsInFirefox([htmlPath]);
                }
                catch (Exception exBrowser)
                {
                    Log($"   [AVISO] Não foi possível abrir o Firefox automaticamente: {exBrowser.Message}");
                }
            }

            // C) Gerar PDF via Edge/Chrome Headless
            Log("-> Gerando relatório PDF (Headless Chromium)...");
            if (DeviceReportExportService.GeneratePdfReport(htmlPath, pdfPath))
            {
                result.PdfPath = pdfPath;
                Log($"   [OK] PDF salvo em: {pdfPath}");
            }
            else
            {
                Log("   [AVISO] Navegador Edge/Chrome não encontrado para conversão PDF headless.");
            }

            // D) Gerar Excel (.xlsx OpenXML nativo em aba única 'Relatorio')
            Log("-> Gerando planilha Excel (.xlsx OpenXML nativo)...");
            DeviceReportExportService.GenerateExcelReport(data, excelPath);
            result.ExcelPath = excelPath;
            Log($"   [OK] Excel salvo em: {excelPath}");

            result.Success = true;
            result.Message = $"Análise de '{target}' concluída com sucesso!";
            Log("==================================================================");
            Log($"[SUCESSO] Relatórios HTML, PDF e Excel gerados em: {outDir}");
        }
        catch (Exception ex)
        {
            result.Message = $"Falha ao analisar '{target}': {ex.Message}";
            Log($"[ERRO] {result.Message}");
        }

        return result;
    }

    private static void CollectHardwareBiosAndUser(ManagementScope cimScope, string target, DeviceReportData data)
    {
        string serial = string.Empty;
        try
        {
            using var biosSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Manufacturer, SMBIOSBIOSVersion, SerialNumber, ReleaseDate FROM Win32_BIOS"));
            foreach (ManagementObject b in biosSearcher.Get())
            {
                using (b)
                {
                    serial = b["SerialNumber"]?.ToString()?.Trim() ?? string.Empty;
                    string relDate = string.Empty;
                    if (b["ReleaseDate"] != null)
                    {
                        try { relDate = ManagementDateTimeConverter.ToDateTime(b["ReleaseDate"].ToString()!).ToString("dd/MM/yyyy"); }
                        catch { relDate = b["ReleaseDate"]?.ToString() ?? string.Empty; }
                    }

                    data.BIOS = new BiosSummaryItem
                    {
                        Fabricante = HardwareMappings.GetFriendlyManufacturer(b["Manufacturer"]?.ToString()),
                        VersaoBios = b["SMBIOSBIOSVersion"]?.ToString() ?? string.Empty,
                        NumeroSerial = serial,
                        DataLancamento = relDate
                    };
                }
            }
        }
        catch { }

        var activePorts = GetActiveVideoPorts(target);
        string activePortsStr = activePorts.Count > 0 ? string.Join(", ", activePorts) : "Nenhuma detectada";

        try
        {
            using var csSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Model, Manufacturer, TotalPhysicalMemory, NumberOfLogicalProcessors, NumberOfProcessors, UserName, Domain FROM Win32_ComputerSystem"));
            foreach (ManagementObject cs in csSearcher.Get())
            {
                using (cs)
                {
                    string rawModel = cs["Model"]?.ToString()?.Trim() ?? string.Empty;
                    string friendlyComp = rawModel;
                    string videoInputs = "Não mapeado";

                    if (HardwareMappings.ModelMap.TryGetValue(rawModel, out var mapped))
                    {
                        friendlyComp = mapped.FriendlyName;
                        videoInputs = mapped.VideoPorts;
                    }

                    double ramGb = cs["TotalPhysicalMemory"] != null
                        ? Math.Round(Convert.ToDouble(cs["TotalPhysicalMemory"]) / (1024.0 * 1024.0 * 1024.0), 2)
                        : 0;

                    data.HardwareInfo = new HardwareSummaryItem
                    {
                        Computador = string.IsNullOrWhiteSpace(friendlyComp) ? target : friendlyComp,
                        Modelo = rawModel,
                        Fabricante = HardwareMappings.GetFriendlyManufacturer(cs["Manufacturer"]?.ToString()),
                        MemoriaFisicaGb = $"{ramGb} GB",
                        ProcessadoresLogicos = cs["NumberOfLogicalProcessors"] != null ? Convert.ToInt32(cs["NumberOfLogicalProcessors"]) : 0,
                        ProcessadoresFisicos = cs["NumberOfProcessors"] != null ? Convert.ToInt32(cs["NumberOfProcessors"]) : 0,
                        SerialNumber = serial,
                        PortasVideoAtivas = activePortsStr,
                        EntradasVideo = videoInputs
                    };

                    string loggedRaw = cs["UserName"]?.ToString() ?? string.Empty;
                    data.CurrentUserDomain = cs["Domain"]?.ToString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(loggedRaw))
                    {
                        string simpleUser = loggedRaw.Contains('\\') ? loggedRaw.Split('\\').Last() : loggedRaw;
                        data.CurrentUser = simpleUser;
                        var adInfo = RemoteManagementHelper.LookupAdUser(simpleUser);
                        data.CurrentUserDisplayName = adInfo.DisplayName;
                    }
                }
            }
        }
        catch { }
    }

    private static List<string> GetActiveVideoPorts(string target)
    {
        var list = new List<string>();
        try
        {
            var wmiScope = RemoteManagementHelper.CreateScope(target, @"root\wmi", TimeSpan.FromSeconds(10));
            var inputTypeByInst = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            using (var paramSearcher = new ManagementObjectSearcher(wmiScope, new ObjectQuery("SELECT InstanceName, VideoInputType FROM WmiMonitorBasicDisplayParams")))
            {
                foreach (ManagementObject p in paramSearcher.Get())
                {
                    using (p)
                    {
                        string inst = p["InstanceName"]?.ToString() ?? string.Empty;
                        int vit = p["VideoInputType"] != null ? Convert.ToInt32(p["VideoInputType"]) : -1;
                        if (!string.IsNullOrEmpty(inst)) inputTypeByInst[inst] = vit;
                    }
                }
            }

            using (var connSearcher = new ManagementObjectSearcher(wmiScope, new ObjectQuery("SELECT InstanceName, Active, VideoOutputTechnology FROM WmiMonitorConnectionParams")))
            {
                foreach (ManagementObject c in connSearcher.Get())
                {
                    using (c)
                    {
                        bool active = c["Active"] != null && Convert.ToBoolean(c["Active"]);
                        if (!active) continue;

                        string inst = c["InstanceName"]?.ToString() ?? string.Empty;
                        int tech = c["VideoOutputTechnology"] != null ? Convert.ToInt32(c["VideoOutputTechnology"]) : -1;

                        if (inputTypeByInst.TryGetValue(inst, out int vit) && vit == 0)
                        {
                            list.Add("VGA");
                        }
                        else
                        {
                            list.Add(HardwareMappings.ConvertVideoOutputTech(tech));
                        }
                    }
                }
            }
        }
        catch { }

        return list.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
    }

    private static void CollectOperatingSystem(ManagementScope cimScope, ManagementScope regScope, DeviceReportData data)
    {
        try
        {
            string? winRelease = RemoteManagementHelper.GetStdRegStringValue(
                regScope,
                RemoteManagementHelper.HKEY_LOCAL_MACHINE,
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
                "DisplayVersion")
                ?? RemoteManagementHelper.GetStdRegStringValue(
                    regScope,
                    RemoteManagementHelper.HKEY_LOCAL_MACHINE,
                    @"SOFTWARE\Microsoft\Windows NT\CurrentVersion",
                    "ReleaseId");

            using var osSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Caption, Version, BuildNumber, CSName, LastBootUpTime FROM Win32_OperatingSystem"));
            foreach (ManagementObject os in osSearcher.Get())
            {
                using (os)
                {
                    string version = os["Version"]?.ToString() ?? string.Empty;
                    string lastBoot = string.Empty;
                    if (os["LastBootUpTime"] != null)
                    {
                        try { lastBoot = ManagementDateTimeConverter.ToDateTime(os["LastBootUpTime"].ToString()!).ToString("dd/MM/yyyy HH:mm:ss"); }
                        catch { lastBoot = os["LastBootUpTime"]?.ToString() ?? string.Empty; }
                    }

                    data.LastBootTime = lastBoot;
                    data.OS = new OsSummaryItem
                    {
                        SO = os["Caption"]?.ToString() ?? string.Empty,
                        VersaoWmi = version,
                        VersaoWindowsHnH = string.IsNullOrWhiteSpace(winRelease) ? version : winRelease,
                        NumeroBuild = os["BuildNumber"]?.ToString() ?? string.Empty,
                        NomeMaquina = os["CSName"]?.ToString() ?? data.ComputerName,
                        UltimoBoot = lastBoot
                    };
                }
            }
        }
        catch { }
    }

    private static void CollectCpuMotherboardAndChassis(ManagementScope cimScope, DeviceReportData data)
    {
        try
        {
            using var cpuSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, CurrentVoltage, DataWidth, L2CacheSize, L3CacheSize FROM Win32_Processor"));
            foreach (ManagementObject cpu in cpuSearcher.Get())
            {
                using (cpu)
                {
                    data.Processor = new ProcessorSummaryItem
                    {
                        Nome = cpu["Name"]?.ToString()?.Trim() ?? string.Empty,
                        NumeroCores = cpu["NumberOfCores"] != null ? Convert.ToInt32(cpu["NumberOfCores"]) : 0,
                        ProcessadoresLogicos = cpu["NumberOfLogicalProcessors"] != null ? Convert.ToInt32(cpu["NumberOfLogicalProcessors"]) : 0,
                        ClockMaximo = cpu["MaxClockSpeed"] != null ? $"{cpu["MaxClockSpeed"]} MHz" : "N/D",
                        Voltagem = cpu["CurrentVoltage"] != null ? $"{cpu["CurrentVoltage"]} V" : "N/D",
                        L1Cache = cpu["DataWidth"] != null ? $"{cpu["DataWidth"]} KB" : "N/D",
                        L2Cache = cpu["L2CacheSize"] != null ? $"{cpu["L2CacheSize"]} KB" : "N/D",
                        L3Cache = cpu["L3CacheSize"] != null ? $"{cpu["L3CacheSize"]} KB" : "N/D"
                    };
                    break;
                }
            }
        }
        catch { }

        try
        {
            using var mbSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Manufacturer, Product, Version FROM Win32_BaseBoard"));
            foreach (ManagementObject mb in mbSearcher.Get())
            {
                using (mb)
                {
                    data.Motherboard = new MotherboardSummaryItem
                    {
                        Fabricante = HardwareMappings.GetFriendlyManufacturer(mb["Manufacturer"]?.ToString()),
                        Produto = mb["Product"]?.ToString() ?? string.Empty,
                        Versao = mb["Version"]?.ToString() ?? string.Empty
                    };
                    break;
                }
            }
        }
        catch { }

        try
        {
            using var chSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT ChassisTypes, Manufacturer FROM Win32_SystemEnclosure"));
            foreach (ManagementObject ch in chSearcher.Get())
            {
                using (ch)
                {
                    var typesList = new List<string>();
                    if (ch["ChassisTypes"] is Array arr)
                    {
                        foreach (var item in arr)
                        {
                            int code = Convert.ToInt32(item);
                            typesList.Add(HardwareMappings.ChassisTypeMap.TryGetValue(code, out string? name) ? name : $"Unknown({code})");
                        }
                    }

                    data.Chassis = new ChassisSummaryItem
                    {
                        TipoChassi = string.Join(", ", typesList),
                        Fabricante = HardwareMappings.GetFriendlyManufacturer(ch["Manufacturer"]?.ToString())
                    };
                    break;
                }
            }
        }
        catch { }
    }

    private static void CollectMemoryModules(ManagementScope cimScope, DeviceReportData data)
    {
        try
        {
            using var memSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Manufacturer, Capacity, Speed, PartNumber, SerialNumber FROM Win32_PhysicalMemory"));
            foreach (ManagementObject m in memSearcher.Get())
            {
                using (m)
                {
                    double capGb = m["Capacity"] != null
                        ? Math.Round(Convert.ToDouble(m["Capacity"]) / (1024.0 * 1024.0 * 1024.0), 2)
                        : 0;

                    data.MemoryModules.Add(new RamModuleItem
                    {
                        Fabricante = HardwareMappings.FormatRamManufacturer(m["Manufacturer"]?.ToString()),
                        CapacidadeGb = $"{capGb} GB",
                        VelocidadeMhz = m["Speed"] != null ? $"{m["Speed"]} MHz" : "N/D",
                        PartNumber = m["PartNumber"]?.ToString()?.Trim() ?? "N/D",
                        NumeroSerie = m["SerialNumber"]?.ToString()?.Trim() ?? "N/D"
                    });
                }
            }
        }
        catch { }
    }

    private static void CollectMonitorsAndVideo(ManagementScope cimScope, string target, DeviceReportData data)
    {
        try
        {
            var wmiScope = RemoteManagementHelper.CreateScope(target, @"root\wmi", TimeSpan.FromSeconds(10));

            var connMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            using (var cSearcher = new ManagementObjectSearcher(wmiScope, new ObjectQuery("SELECT InstanceName, VideoOutputTechnology FROM WmiMonitorConnectionParams")))
            {
                foreach (ManagementObject c in cSearcher.Get())
                {
                    using (c)
                    {
                        string inst = c["InstanceName"]?.ToString() ?? string.Empty;
                        int tech = c["VideoOutputTechnology"] != null ? Convert.ToInt32(c["VideoOutputTechnology"]) : -1;
                        if (!string.IsNullOrEmpty(inst)) connMap[inst] = tech;
                    }
                }
            }

            var paramMap = new Dictionary<string, (int H, int V, int InputType)>(StringComparer.OrdinalIgnoreCase);
            using (var pSearcher = new ManagementObjectSearcher(wmiScope, new ObjectQuery("SELECT InstanceName, MaxHorizontalImageSize, MaxVerticalImageSize, VideoInputType FROM WmiMonitorBasicDisplayParams")))
            {
                foreach (ManagementObject p in pSearcher.Get())
                {
                    using (p)
                    {
                        string inst = p["InstanceName"]?.ToString() ?? string.Empty;
                        int h = p["MaxHorizontalImageSize"] != null ? Convert.ToInt32(p["MaxHorizontalImageSize"]) : 0;
                        int v = p["MaxVerticalImageSize"] != null ? Convert.ToInt32(p["MaxVerticalImageSize"]) : 0;
                        int vit = p["VideoInputType"] != null ? Convert.ToInt32(p["VideoInputType"]) : -1;
                        if (!string.IsNullOrEmpty(inst)) paramMap[inst] = (h, v, vit);
                    }
                }
            }

            using (var mSearcher = new ManagementObjectSearcher(wmiScope, new ObjectQuery("SELECT InstanceName, ManufacturerName, SerialNumberID, UserFriendlyName, Active FROM WmiMonitorID")))
            {
                foreach (ManagementObject m in mSearcher.Get())
                {
                    using (m)
                    {
                        string inst = m["InstanceName"]?.ToString() ?? string.Empty;
                        string rawName = inst.Contains('\\') ? inst.Split('\\')[1] : inst;
                        string mfrCode = DecodeUInt16CharArray(m["ManufacturerName"]);
                        string serial = DecodeUInt16CharArray(m["SerialNumberID"]);
                        string friendly = DecodeUInt16CharArray(m["UserFriendlyName"]);
                        bool active = m["Active"] != null && Convert.ToBoolean(m["Active"]);

                        bool hasParams = paramMap.TryGetValue(inst, out var p);
                        bool hasConn = connMap.TryGetValue(inst, out int techCode);

                        string activePort = hasParams && p.InputType == 0
                            ? "VGA"
                            : (hasConn ? HardwareMappings.ConvertVideoOutputTech(techCode) : "N/A");

                        var item = new MonitorInfoItem
                        {
                            Nome = rawName,
                            Polegada = hasParams && p.H > 0 && p.V > 0 ? Math.Round((p.H / 2.54 * (p.V / 2.54)) / 100.0, 1).ToString() : "N/A",
                            Resolucao = hasParams ? $"{p.H}x{p.V}" : "N/A",
                            Tamanho = hasParams ? $"{p.H}mm x {p.V}mm" : "N/A",
                            Ativo = active ? "Sim" : "Não",
                            SerialNumber = serial,
                            NomeAmigavel = friendly,
                            Fabricante = HardwareMappings.ConvertMonitorManufacturer(mfrCode),
                            EntradasVideo = "N/A",
                            PortaAtiva = activePort
                        };

                        if (HardwareMappings.MonitorMap.TryGetValue(rawName, out var cfg))
                        {
                            item.Nome = cfg.Nome;
                            item.Polegada = cfg.Polegada;
                            item.Resolucao = cfg.Resolucao;
                            item.Tamanho = cfg.Tamanho;
                            item.EntradasVideo = cfg.VideoPorts;
                        }

                        data.Monitors.Add(item);
                    }
                }
            }
        }
        catch { }

        try
        {
            using var gpuSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Name, VideoProcessor, AdapterCompatibility, AdapterRAM, DeviceID, CurrentHorizontalResolution, CurrentVerticalResolution, CurrentNumberOfColors, DriverVersion, DriverDate FROM Win32_VideoController"));
            foreach (ManagementObject v in gpuSearcher.Get())
            {
                using (v)
                {
                    double ramGb = v["AdapterRAM"] != null
                        ? Math.Round(Convert.ToDouble(v["AdapterRAM"]) / (1024.0 * 1024.0 * 1024.0), 2)
                        : 0;
                    string drvDate = string.Empty;
                    if (v["DriverDate"] != null)
                    {
                        try { drvDate = ManagementDateTimeConverter.ToDateTime(v["DriverDate"].ToString()!).ToString("dd/MM/yyyy HH:mm:ss"); }
                        catch { drvDate = v["DriverDate"]?.ToString() ?? string.Empty; }
                    }

                    data.VideoControllers.Add(new VideoControllerItem
                    {
                        Nome = v["Name"]?.ToString() ?? string.Empty,
                        ProcessadorVideo = v["VideoProcessor"]?.ToString() ?? string.Empty,
                        AdaptadorCompatibilidade = v["AdapterCompatibility"]?.ToString() ?? string.Empty,
                        CapacidadeGb = $"{ramGb} GB",
                        DeviceId = v["DeviceID"]?.ToString() ?? string.Empty,
                        ResolucaoHorizontal = v["CurrentHorizontalResolution"]?.ToString() ?? string.Empty,
                        ResolucaoVertical = v["CurrentVerticalResolution"]?.ToString() ?? string.Empty,
                        NumeroCores = v["CurrentNumberOfColors"]?.ToString() ?? string.Empty,
                        VersaoDriver = v["DriverVersion"]?.ToString() ?? string.Empty,
                        DataDriverInstalado = drvDate
                    });
                }
            }
        }
        catch { }
    }

    private static void CollectDisks(ManagementScope cimScope, string target, DeviceReportData data)
    {
        var mediaTypesByDriveNum = new Dictionary<int, string>();
        try
        {
            var storageScope = RemoteManagementHelper.CreateScope(target, @"root\Microsoft\Windows\Storage", TimeSpan.FromSeconds(10));
            using var pdSearcher = new ManagementObjectSearcher(storageScope, new ObjectQuery("SELECT DeviceId, MediaType FROM MSFT_PhysicalDisk"));
            foreach (ManagementObject pd in pdSearcher.Get())
            {
                using (pd)
                {
                    if (int.TryParse(pd["DeviceId"]?.ToString(), out int devId))
                    {
                        ushort mt = pd["MediaType"] != null ? Convert.ToUInt16(pd["MediaType"]) : (ushort)0;
                        string mtStr = mt switch
                        {
                            3 => "HDD",
                            4 => "SSD",
                            5 => "SCM",
                            _ => "Unspecified"
                        };
                        mediaTypesByDriveNum[devId] = mtStr;
                    }
                }
            }
        }
        catch { }

        try
        {
            using var ddSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT DeviceID, Model, InterfaceType, SerialNumber, Size FROM Win32_DiskDrive"));
            foreach (ManagementObject d in ddSearcher.Get())
            {
                using (d)
                {
                    string devId = d["DeviceID"]?.ToString() ?? string.Empty;
                    string mediaType = "Unspecified";
                    var m = Regex.Match(devId, @"PHYSICALDRIVE(\d+)$", RegexOptions.IgnoreCase);
                    if (m.Success && int.TryParse(m.Groups[1].Value, out int num) && mediaTypesByDriveNum.TryGetValue(num, out string? mt))
                    {
                        mediaType = mt;
                    }

                    double sizeGb = d["Size"] != null
                        ? Math.Round(Convert.ToDouble(d["Size"]) / (1024.0 * 1024.0 * 1024.0), 2)
                        : 0;

                    var parsed = HardwareMappings.GetParsedDiskInfo(d["Model"]?.ToString());

                    data.DiskDrives.Add(new PhysicalDiskItem
                    {
                        Fabricante = parsed.Fabricante,
                        Modelo = parsed.Modelo,
                        TipoInterface = d["InterfaceType"]?.ToString() ?? "N/D",
                        TamanhoGb = sizeGb > 0 ? $"{sizeGb} GB" : "N/D",
                        TipoMidia = mediaType,
                        NumeroSerie = d["SerialNumber"]?.ToString()?.Trim() ?? "N/D"
                    });
                }
            }
        }
        catch { }

        try
        {
            using var ldSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT DeviceID, FreeSpace, Size, VolumeName, FileSystem FROM Win32_LogicalDisk WHERE Size > 0"));
            foreach (ManagementObject ld in ldSearcher.Get())
            {
                using (ld)
                {
                    double freeGb = ld["FreeSpace"] != null ? Math.Round(Convert.ToDouble(ld["FreeSpace"]) / (1024.0 * 1024.0 * 1024.0), 2) : 0;
                    double totalGb = ld["Size"] != null ? Math.Round(Convert.ToDouble(ld["Size"]) / (1024.0 * 1024.0 * 1024.0), 2) : 0;
                    double usedGb = Math.Round(Math.Max(0, totalGb - freeGb), 2);
                    string volName = ld["VolumeName"]?.ToString() ?? string.Empty;

                    data.LogicalDisks.Add(new LogicalDiskItem
                    {
                        DeviceId = ld["DeviceID"]?.ToString() ?? string.Empty,
                        EspacoLivre = $"{freeGb} GB",
                        EspacoUsado = $"{usedGb} GB",
                        EspacoTotal = $"{totalGb} GB",
                        NomeVolume = string.IsNullOrWhiteSpace(volName) ? "Windows" : volName,
                        SistemaArquivos = ld["FileSystem"]?.ToString() ?? "NTFS"
                    });
                }
            }
        }
        catch { }
    }

    private static void CollectNetworkAndSound(ManagementScope cimScope, DeviceReportData data)
    {
        try
        {
            using var netSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Description, MACAddress, IPAddress, DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True"));
            foreach (ManagementObject n in netSearcher.Get())
            {
                using (n)
                {
                    string ips = n["IPAddress"] is string[] ipArr ? string.Join(", ", ipArr) : string.Empty;
                    string gws = n["DefaultIPGateway"] is string[] gwArr ? string.Join(", ", gwArr) : string.Empty;

                    data.NetworkConfigs.Add(new NetworkConfigItem
                    {
                        NomePlaca = n["Description"]?.ToString() ?? string.Empty,
                        MacAddress = n["MACAddress"]?.ToString() ?? string.Empty,
                        EnderecosIp = ips,
                        Gateway = gws
                    });
                }
            }
        }
        catch { }

        try
        {
            using var sndSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Name, Manufacturer, Status, StatusInfo FROM Win32_SoundDevice"));
            foreach (ManagementObject s in sndSearcher.Get())
            {
                using (s)
                {
                    int si = s["StatusInfo"] != null ? Convert.ToInt32(s["StatusInfo"]) : 0;
                    string statInfo = si switch
                    {
                        1 => "Outro",
                        2 => "Desconhecido",
                        3 => "Habilitado",
                        4 => "Desabilitado",
                        5 => "Não aplicável",
                        _ => si.ToString()
                    };

                    data.SoundDevices.Add(new SoundDeviceItem
                    {
                        Nome = s["Name"]?.ToString() ?? string.Empty,
                        Fabricante = HardwareMappings.GetFriendlyManufacturer(s["Manufacturer"]?.ToString()),
                        Status = s["Status"]?.ToString() ?? string.Empty,
                        InformacaoStatus = statInfo
                    });
                }
            }
        }
        catch { }
    }

    private static void CollectPrinters(ManagementScope cimScope, ManagementScope regScope, string target, DeviceReportData data)
    {
        try
        {
            using var pSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Name, PrinterState, PrinterStatus, ShareName, PortName FROM Win32_Printer"));
            foreach (ManagementObject p in pSearcher.Get())
            {
                using (p)
                {
                    int status = p["PrinterStatus"] != null ? Convert.ToInt32(p["PrinterStatus"]) : 0;
                    int state = p["PrinterState"] != null ? Convert.ToInt32(p["PrinterState"]) : 0;

                    string statusStr = status switch
                    {
                        1 => "Outro Status",
                        2 => "Desconhecido",
                        3 => "Parada",
                        4 => "Imprimindo",
                        5 => "Está em aquecimento",
                        6 => "Impressão interrompida",
                        7 => "Offline",
                        _ => status.ToString()
                    };

                    string stateStr = state switch
                    {
                        0 => "Parada",
                        1 => "Pausada",
                        2 => "Erro",
                        3 => "Pendente de deletar",
                        4 => "Atolamento de papel",
                        5 => "Sem papel",
                        6 => "Alimentação manual",
                        7 => "Problema de papel",
                        8 => "Offline",
                        9 => "Entrada/Saída ativa",
                        10 => "Ocupada",
                        11 => "Imprimindo",
                        12 => "",
                        _ => state.ToString()
                    };

                    data.PrintersLocal.Add(new LocalPrinterItem
                    {
                        Nome = p["Name"]?.ToString() ?? string.Empty,
                        EstadoImpressao = stateStr,
                        StatusImpressao = statusStr,
                        NomeCompartilhamento = p["ShareName"]?.ToString() ?? string.Empty,
                        Porta = p["PortName"]?.ToString() ?? string.Empty
                    });
                }
            }
        }
        catch { }

        // Impressoras mapeadas por usuário (HKEY_USERS\<SID>\Printers\Connections)
        try
        {
            var sids = RemoteManagementHelper.EnumStdRegSubKeys(regScope, RemoteManagementHelper.HKEY_USERS, "");
            if (sids != null)
            {
                foreach (string sid in sids)
                {
                    if (!Regex.IsMatch(sid, @"^S-1-5-21-\d+-\d+-\d+-\d+$")) continue;

                    var conns = RemoteManagementHelper.EnumStdRegSubKeys(regScope, RemoteManagementHelper.HKEY_USERS, $@"{sid}\Printers\Connections");
                    if (conns != null)
                    {
                        foreach (string entry in conns)
                        {
                            var parts = entry.Split(',', StringSplitOptions.RemoveEmptyEntries);
                            data.PrintersMapped.Add(new MappedPrinterItem
                            {
                                ComputerName = target,
                                UserSid = sid,
                                UserName = sid,
                                Type = "Connected",
                                Server = parts.Length > 0 ? parts[0] : string.Empty,
                                PrinterName = parts.Length > 1 ? parts[1] : entry
                            });
                        }
                    }
                }
            }
        }
        catch { }
    }

    private static void CollectInstalledPrograms(ManagementScope regScope, DeviceReportData data)
    {
        string[] uninstallRoots =
        [
            @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ];

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (string root in uninstallRoots)
        {
            var subKeys = RemoteManagementHelper.EnumStdRegSubKeys(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, root);
            if (subKeys == null) continue;

            foreach (string key in subKeys)
            {
                string fullKey = $@"{root}\{key}";
                string? displayName = RemoteManagementHelper.GetStdRegStringValue(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, fullKey, "DisplayName");
                if (string.IsNullOrWhiteSpace(displayName)) continue;

                string cleanName = displayName.Trim();
                if (!seen.Add(cleanName)) continue;

                string? version = RemoteManagementHelper.GetStdRegStringValue(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, fullKey, "DisplayVersion");
                string? publisher = RemoteManagementHelper.GetStdRegStringValue(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, fullKey, "Publisher");
                string? installDate = RemoteManagementHelper.GetStdRegStringValue(regScope, RemoteManagementHelper.HKEY_LOCAL_MACHINE, fullKey, "InstallDate");

                data.InstalledPrograms.Add(new InstalledProgramItem
                {
                    NomePrograma = cleanName,
                    Versao = version?.Trim() ?? string.Empty,
                    Fabricante = publisher?.Trim() ?? string.Empty,
                    DataInstalacao = HardwareMappings.FormatInstallDate(installDate)
                });
            }
        }

        data.InstalledPrograms = data.InstalledPrograms.OrderBy(p => p.NomePrograma).ToList();
    }

    private static void CollectServicesAndProcesses(ManagementScope cimScope, DeviceReportData data)
    {
        try
        {
            using var svcSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT Name, DisplayName, State, StartMode, StartName, PathName FROM Win32_Service WHERE State='Running'"));
            foreach (ManagementObject s in svcSearcher.Get())
            {
                using (s)
                {
                    string startMode = (s["StartMode"]?.ToString() ?? "") switch
                    {
                        "Auto" => "Automático",
                        "Manual" => "Manual",
                        "Disabled" => "Desativado",
                        "System" => "Sistema",
                        "Boot" => "Boot",
                        var m => m
                    };

                    data.RunningServices.Add(new RunningServiceItem
                    {
                        NomeTecnico = s["Name"]?.ToString() ?? string.Empty,
                        NomeExibicao = s["DisplayName"]?.ToString() ?? string.Empty,
                        Status = "Em execução",
                        TipoInicializacao = startMode,
                        ContaLogon = s["StartName"]?.ToString() ?? string.Empty,
                        CaminhoExecutavel = s["PathName"]?.ToString() ?? string.Empty
                    });
                }
            }
        }
        catch { }

        try
        {
            using var procSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT ProcessId, Name, WorkingSetSize, ThreadCount, HandleCount, ExecutablePath, CommandLine FROM Win32_Process"));
            foreach (ManagementObject p in procSearcher.Get())
            {
                using (p)
                {
                    double memMb = p["WorkingSetSize"] != null
                        ? Math.Round(Convert.ToDouble(p["WorkingSetSize"]) / (1024.0 * 1024.0), 2)
                        : 0;

                    data.ActiveProcesses.Add(new ActiveProcessItem
                    {
                        Pid = p["ProcessId"] != null ? Convert.ToUInt32(p["ProcessId"]) : 0,
                        NomeExecutavel = p["Name"]?.ToString() ?? string.Empty,
                        Usuario = "Sistema / Ativo",
                        UsoMemoriaMb = memMb,
                        Threads = p["ThreadCount"] != null ? Convert.ToUInt32(p["ThreadCount"]) : 0,
                        Handles = p["HandleCount"] != null ? Convert.ToUInt32(p["HandleCount"]) : 0,
                        CaminhoArquivo = p["ExecutablePath"]?.ToString() ?? string.Empty,
                        LinhaComando = RedactSensitiveCommandLine(p["CommandLine"]?.ToString() ?? string.Empty)
                    });
                }
            }
        }
        catch { }
    }

    private static readonly Regex SensitiveCmdArgRegex = new(
        @"(?i)(?:(?:/|--|-)(?:p|pass|password|pwd|secret|token|apikey|api[_-]?key|key|authorization|bearer|cred|credential)\s*(?:=|:|\s+)\s*)(?:""[^""]*""|'[^']*'|\S+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(250));

    /// <summary>Mascaramento de senhas, tokens e chaves passados em linha de comando de processos coletados no inventário.</summary>
    public static string RedactSensitiveCommandLine(string? commandLine)
    {
        if (string.IsNullOrEmpty(commandLine)) return string.Empty;
        try
        {
            return SensitiveCmdArgRegex.Replace(commandLine, m =>
            {
                int sep = m.Value.IndexOfAny(['=', ':', ' ']);
                return sep > 0 ? m.Value[..(sep + 1)] + "***REDACTED***" : "***REDACTED***";
            });
        }
        catch (RegexMatchTimeoutException)
        {
            return "[REDACTED-TIMEOUT]";
        }
    }

    private static void CollectDrivers(ManagementScope cimScope, DeviceReportData data)
    {
        try
        {
            using var drvSearcher = new ManagementObjectSearcher(cimScope, new ObjectQuery("SELECT DeviceName, Manufacturer, DriverVersion, DriverDate, InfName, IsSigned FROM Win32_PnPSignedDriver WHERE DeviceName IS NOT NULL"));
            foreach (ManagementObject d in drvSearcher.Get())
            {
                using (d)
                {
                    string drvDate = "N/D";
                    if (d["DriverDate"] != null)
                    {
                        try { drvDate = ManagementDateTimeConverter.ToDateTime(d["DriverDate"].ToString()!).ToString("dd/MM/yyyy"); }
                        catch { drvDate = d["DriverDate"]?.ToString() ?? "N/D"; }
                    }

                    bool signed = d["IsSigned"] != null && Convert.ToBoolean(d["IsSigned"]);

                    data.InstalledDrivers.Add(new InstalledDriverItem
                    {
                        NomeDispositivo = d["DeviceName"]?.ToString() ?? string.Empty,
                        Fabricante = d["Manufacturer"]?.ToString() ?? string.Empty,
                        VersaoDriver = d["DriverVersion"]?.ToString() ?? string.Empty,
                        DataDriver = drvDate,
                        ArquivoInf = d["InfName"]?.ToString() ?? string.Empty,
                        Assinado = signed ? "Sim" : "Não"
                    });
                }
            }
        }
        catch { }
    }

    private static string DecodeUInt16CharArray(object? raw)
    {
        if (raw is ushort[] u16Arr)
        {
            return new string(u16Arr.Where(c => c != 0).Select(c => (char)c).ToArray()).Trim();
        }
        if (raw is Array arr)
        {
            var chars = new List<char>();
            foreach (var item in arr)
            {
                int val = Convert.ToInt32(item);
                if (val != 0) chars.Add((char)val);
            }
            return new string(chars.ToArray()).Trim();
        }
        return raw?.ToString()?.Trim() ?? string.Empty;
    }
}

