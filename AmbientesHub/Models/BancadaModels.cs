using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace AmbientesHub.Models;

public class HardwareSummaryItem
{
    public string Computador { get; set; } = string.Empty;
    public string Modelo { get; set; } = string.Empty;
    public string Fabricante { get; set; } = string.Empty;
    public string MemoriaFisicaGb { get; set; } = string.Empty;
    public int ProcessadoresLogicos { get; set; }
    public int ProcessadoresFisicos { get; set; }
    public string SerialNumber { get; set; } = string.Empty;
    public string PortasVideoAtivas { get; set; } = "Nenhuma detectada";
    public string EntradasVideo { get; set; } = "Não mapeado";
}

public class OsSummaryItem
{
    public string SO { get; set; } = string.Empty;
    public string VersaoWmi { get; set; } = string.Empty;
    public string VersaoWindowsHnH { get; set; } = string.Empty;
    public string NumeroBuild { get; set; } = string.Empty;
    public string NomeMaquina { get; set; } = string.Empty;
    public string UltimoBoot { get; set; } = string.Empty;
}

public class BiosSummaryItem
{
    public string Fabricante { get; set; } = string.Empty;
    public string VersaoBios { get; set; } = string.Empty;
    public string NumeroSerial { get; set; } = string.Empty;
    public string DataLancamento { get; set; } = string.Empty;
}

public class ProcessorSummaryItem
{
    public string Nome { get; set; } = string.Empty;
    public int NumeroCores { get; set; }
    public int ProcessadoresLogicos { get; set; }
    public string ClockMaximo { get; set; } = "N/D";
    public string Voltagem { get; set; } = "N/D";
    public string L1Cache { get; set; } = "N/D";
    public string L2Cache { get; set; } = "N/D";
    public string L3Cache { get; set; } = "N/D";
}

public class MotherboardSummaryItem
{
    public string Fabricante { get; set; } = string.Empty;
    public string Produto { get; set; } = string.Empty;
    public string Versao { get; set; } = string.Empty;
}

public class ChassisSummaryItem
{
    public string TipoChassi { get; set; } = string.Empty;
    public string Fabricante { get; set; } = string.Empty;
}

public class RamModuleItem
{
    public string Fabricante { get; set; } = "N/D";
    public string CapacidadeGb { get; set; } = "0 GB";
    public string VelocidadeMhz { get; set; } = "N/D";
    public string PartNumber { get; set; } = "N/D";
    public string NumeroSerie { get; set; } = "N/D";
}

public class PhysicalDiskItem
{
    public string Fabricante { get; set; } = "N/D";
    public string Modelo { get; set; } = "N/D";
    public string TipoInterface { get; set; } = "N/D";
    public string TamanhoGb { get; set; } = "N/D";
    public string TipoMidia { get; set; } = "Unspecified";
    public string NumeroSerie { get; set; } = "N/D";
}

public class LogicalDiskItem
{
    public string DeviceId { get; set; } = string.Empty;
    public string EspacoLivre { get; set; } = "0 GB";
    public string EspacoUsado { get; set; } = "0 GB";
    public string EspacoTotal { get; set; } = "0 GB";
    public string NomeVolume { get; set; } = "Windows";
    public string SistemaArquivos { get; set; } = "NTFS";
}

public class NetworkConfigItem
{
    public string NomePlaca { get; set; } = string.Empty;
    public string MacAddress { get; set; } = string.Empty;
    public string EnderecosIp { get; set; } = string.Empty;
    public string Gateway { get; set; } = string.Empty;
}

public class MonitorInfoItem
{
    public string Nome { get; set; } = string.Empty;
    public string Polegada { get; set; } = "N/A";
    public string Resolucao { get; set; } = "N/A";
    public string Tamanho { get; set; } = "N/A";
    public string Ativo { get; set; } = "Sim";
    public string SerialNumber { get; set; } = string.Empty;
    public string NomeAmigavel { get; set; } = string.Empty;
    public string Fabricante { get; set; } = string.Empty;
    public string EntradasVideo { get; set; } = "N/A";
    public string PortaAtiva { get; set; } = "N/A";
}

public class VideoControllerItem
{
    public string Nome { get; set; } = string.Empty;
    public string ProcessadorVideo { get; set; } = string.Empty;
    public string AdaptadorCompatibilidade { get; set; } = string.Empty;
    public string CapacidadeGb { get; set; } = "0 GB";
    public string DeviceId { get; set; } = string.Empty;
    public string ResolucaoHorizontal { get; set; } = string.Empty;
    public string ResolucaoVertical { get; set; } = string.Empty;
    public string NumeroCores { get; set; } = string.Empty;
    public string VersaoDriver { get; set; } = string.Empty;
    public string DataDriverInstalado { get; set; } = string.Empty;
}

public class SoundDeviceItem
{
    public string Nome { get; set; } = string.Empty;
    public string Fabricante { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string InformacaoStatus { get; set; } = string.Empty;
}

public class LocalPrinterItem
{
    public string Nome { get; set; } = string.Empty;
    public string EstadoImpressao { get; set; } = string.Empty;
    public string StatusImpressao { get; set; } = string.Empty;
    public string NomeCompartilhamento { get; set; } = string.Empty;
    public string Porta { get; set; } = string.Empty;
}

public class MappedPrinterItem
{
    public string ComputerName { get; set; } = string.Empty;
    public string UserSid { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;
    public string Type { get; set; } = "Connected";
    public string Server { get; set; } = string.Empty;
    public string PrinterName { get; set; } = string.Empty;
}

public class InstalledDriverItem
{
    public string NomeDispositivo { get; set; } = string.Empty;
    public string Fabricante { get; set; } = string.Empty;
    public string VersaoDriver { get; set; } = string.Empty;
    public string DataDriver { get; set; } = "N/D";
    public string ArquivoInf { get; set; } = string.Empty;
    public string Assinado { get; set; } = "Não";
}

public class InstalledProgramItem
{
    public string NomePrograma { get; set; } = string.Empty;
    public string Versao { get; set; } = string.Empty;
    public string Fabricante { get; set; } = string.Empty;
    public string DataInstalacao { get; set; } = string.Empty;
}

public class RunningServiceItem
{
    public string NomeTecnico { get; set; } = string.Empty;
    public string NomeExibicao { get; set; } = string.Empty;
    public string Status { get; set; } = "Em execução";
    public string TipoInicializacao { get; set; } = "Automático";
    public string ContaLogon { get; set; } = string.Empty;
    public string CaminhoExecutavel { get; set; } = string.Empty;
}

public class ActiveProcessItem
{
    public uint Pid { get; set; }
    public string NomeExecutavel { get; set; } = string.Empty;
    public string Usuario { get; set; } = "Sistema";
    public double UsoMemoriaMb { get; set; }
    public uint Threads { get; set; }
    public uint Handles { get; set; }
    public string CaminhoArquivo { get; set; } = string.Empty;
    public string LinhaComando { get; set; } = string.Empty;
}

public class RemoteUserProfileItem : INotifyPropertyChanged
{
    private bool _isSelected;

    public string Nome { get; set; } = string.Empty;
    public string Usuario { get; set; } = string.Empty;
    public string Descricao { get; set; } = string.Empty;
    public string Escritorio { get; set; } = string.Empty;
    public string CaminhoPerfil { get; set; } = string.Empty;
    public string Sid { get; set; } = string.Empty;
    public string UltimoUso { get; set; } = "N/D";
    public DateTime? UltimoUsoDate { get; set; }
    public bool Loaded { get; set; }
    public string StatusSessao => Loaded ? "Logado (Ativo)" : "Não logado";
    public bool Special { get; set; }
    public double EspacoDiscoGb { get; set; }
    public string EspacoDiscoTexto => $"{EspacoDiscoGb:N2} GB";

    public bool CanBeRemoved => !Loaded && !Special;

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;
            _isSelected = CanBeRemoved && value;
            OnPropertyChanged();
        }
    }

    public Visibility ActiveBadgeVisibility => Loaded ? Visibility.Visible : Visibility.Collapsed;
    public Visibility RemovableBadgeVisibility => CanBeRemoved ? Visibility.Visible : Visibility.Collapsed;

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public class DeviceReportData
{
    public string ComputerName { get; set; } = string.Empty;
    public HardwareSummaryItem HardwareInfo { get; set; } = new();
    public string CurrentUser { get; set; } = string.Empty;
    public string CurrentUserDisplayName { get; set; } = string.Empty;
    public string CurrentUserDomain { get; set; } = string.Empty;
    public OsSummaryItem OS { get; set; } = new();
    public BiosSummaryItem BIOS { get; set; } = new();
    public ProcessorSummaryItem Processor { get; set; } = new();
    public MotherboardSummaryItem Motherboard { get; set; } = new();
    public ChassisSummaryItem Chassis { get; set; } = new();
    public List<RamModuleItem> MemoryModules { get; set; } = [];
    public List<PhysicalDiskItem> DiskDrives { get; set; } = [];
    public List<LogicalDiskItem> LogicalDisks { get; set; } = [];
    public List<NetworkConfigItem> NetworkConfigs { get; set; } = [];
    public List<MonitorInfoItem> Monitors { get; set; } = [];
    public List<VideoControllerItem> VideoControllers { get; set; } = [];
    public List<SoundDeviceItem> SoundDevices { get; set; } = [];
    public List<LocalPrinterItem> PrintersLocal { get; set; } = [];
    public List<MappedPrinterItem> PrintersMapped { get; set; } = [];
    public List<InstalledDriverItem> InstalledDrivers { get; set; } = [];
    public List<InstalledProgramItem> InstalledPrograms { get; set; } = [];
    public List<RunningServiceItem> RunningServices { get; set; } = [];
    public List<ActiveProcessItem> ActiveProcesses { get; set; } = [];
    public List<RemoteUserProfileItem> UserProfiles { get; set; } = [];
    public string LastBootTime { get; set; } = string.Empty;
    public bool SkipMajorData { get; set; }
}

public class DeviceAnalysisResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public DeviceReportData? Data { get; set; }
    public string? HtmlPath { get; set; }
    public string? PdfPath { get; set; }
    public string? ExcelPath { get; set; }
    public string OutputFolder { get; set; } = string.Empty;
}

public class RemoteMaintenanceOptions
{
    public bool CleanTempFiles { get; set; } = true;
    public bool CleanPrefetch { get; set; } = true;
    public bool CleanRecycleBin { get; set; } = true;
    public bool CleanEventLogs { get; set; } = true;
    public bool CleanWindowsUpdateCache { get; set; } = true;
    public bool FlushDnsCache { get; set; } = true;
    public bool CleanWinSxS { get; set; } = true;
    public bool CleanThumbnailCache { get; set; } = true;
    public bool CleanStoreCache { get; set; } = true;
    public bool CleanCbsLogs { get; set; } = true;
    public bool CleanWindowsOld { get; set; } = false;
    public bool CleanDeliveryOptimization { get; set; } = true;
    public bool CleanCrashDumps { get; set; } = true;
    public bool RunRepairVolume { get; set; } = false;
    public bool RunDefragIfHdd { get; set; } = false;
    public bool RunSfcIfCorrupted { get; set; } = false;
}

public class RemoteCleanupInspectionItem
{
    public string Category { get; set; } = string.Empty;
    public string TargetPath { get; set; } = string.Empty;
    public double SizeMb { get; set; }
    public int FileCount { get; set; }
    public bool Exists { get; set; }
}

public class RemoteMaintenanceResult
{
    public bool Success { get; set; }
    public string ComputerName { get; set; } = string.Empty;
    public double FreeSpaceBeforeGb { get; set; }
    public double FreeSpaceAfterGb { get; set; }
    public double TotalSpaceGb { get; set; }
    public double FreedGb => Math.Round(Math.Max(0, FreeSpaceAfterGb - FreeSpaceBeforeGb), 2);
    public double EstimatedCleanableGb { get; set; }
    public List<RemoteCleanupInspectionItem> InspectionItems { get; set; } = [];
}

public class RemoteProfilePurgeResult
{
    public bool Success { get; set; }
    public string ComputerName { get; set; } = string.Empty;
    public int RemovedCount { get; set; }
    public int SkippedCount { get; set; }
    public int FailedCount { get; set; }
    public double FreedEstimateGb { get; set; }
}

