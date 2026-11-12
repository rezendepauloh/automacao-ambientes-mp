using System.Globalization;
using System.Text.RegularExpressions;

namespace AmbientesHub.Services;

public static class HardwareMappings
{
    public static readonly Dictionary<string, (string FriendlyName, string VideoPorts)> ModelMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["3209N4P"] = ("Lenovo ThinkCentre M92p", "1x VGA / 1x DisplayPort"),
            ["10A9000WBP"] = ("Lenovo ThinkCentre M93p", "1x VGA / 1x DisplayPort"),
            ["11DUSD3R00"] = ("Lenovo ThinkCentre M70q", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["12TES8R800"] = ("Lenovo ThinkCentre M70q v2", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["HP EliteDesk 800 G5 Desktop Mini"] = ("HP EliteDesk 800 G5 Desktop Mini", "1x VGA / 2x DisplayPort"),
            ["HP EliteDesk 800 G3 DM 35W (Brazil)"] = ("HP EliteDesk 800 G3 DM 35W", "1x VGA / 2x DisplayPort"),
            ["HP ProDesk 600 G1 SFF"] = ("HP ProDesk 600 G1 SFF", "1x VGA / 2x DisplayPort"),
            ["HP Z440 Workstation"] = ("HP Z440 Workstation", "1x VGA / Placa de vídeo com 4x DisplayPort"),
            ["HP Pro SFF 400 G9 Desktop PC"] = ("HP Pro SFF 400 G9 Desktop PC", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["OptiPlex 7050"] = ("Dell OptiPlex 7050", "1x HDMI / 2x DisplayPort"),
            ["OptiPlex 3020"] = ("Dell OptiPlex 3020", "1x VGA / 1x DisplayPort"),
            ["20W1S6CB00"] = ("Lenovo T14 Gen 2", "1x HDMI"),
            ["13E0S00400"] = ("Lenovo E14", "1x HDMI")
        };

    public static readonly Dictionary<string, (string Nome, string Polegada, string Resolucao, string Tamanho, string VideoPorts)> MonitorMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["DELA0D7"] = ("Dell P2217H", "21.7", "1920x1080", "480x270mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["DELA113"] = ("Dell P2217H", "21.7", "1920x1080", "480x270mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["DELA0D9"] = ("Dell P2217H", "21.7", "1920x1080", "480x270mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["DELA0D8"] = ("Dell P2217H", "21.7", "1920x1080", "480x270mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["DEL40F4"] = ("Dell P2317H", "23.1", "1920x1080", "510x290mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["DEL40F3"] = ("Dell P2317H", "23.1", "1920x1080", "510x290mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["DELA026"] = ("Dell P2317H", "24.2", "1920x1080", "510x290mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["DELD03A"] = ("Dell E1914H", "18.5", "1366x768", "410x230mm", "1x VGA"),
            ["DELA1D4"] = ("Dell C2423H", "24.0", "1920x1080", "510x290mm", "1x HDMI / 2x DisplayPort"),
            ["DELA1D5"] = ("Dell C2423H", "24.0", "1920x1080", "510x290mm", "1x HDMI / 2x DisplayPort"),
            ["AOC2401"] = ("AOC 24P1U", "24", "1920x1080", "530x300mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["AOC2470"] = ("AOC M2470PWH", "23.4", "1920x1080", "520x290mm", "1x VGA / 2x HDMI"),
            ["GSM4ED3"] = ("LG E2011", "20.3", "1600x900", "450x250mm", "1x VGA / 1x DVI"),
            ["GSM4ED4"] = ("LG E2011", "20.3", "1600x900", "450x250mm", "1x VGA / 1x DVI"),
            ["GSM4C1E"] = ("LG E2011", "20.3", "1600x900", "450x250mm", "1x VGA / 1x DVI"),
            ["HWP3139"] = ("HP 20", "20.3", "1920x1080", "450x250mm", "1x VGA / 1x DVI"),
            ["HWP289A"] = ("HP L190hb", "19.1", "1440x900", "410x260mm", "1x VGA / 1x DVI"),
            ["HWP3279"] = ("HP H P E3", "23.1", "1920x1080", "510x290mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["HWP2990"] = ("HP W2072a", "19.9", "1600x900", "440x250mm", "1x VGA / 1x DVI"),
            ["GSM0001"] = ("LG TV", "41.9", "1920x1080", "930x520mm", ""),
            ["SAM0C39"] = ("Samsung", "31.5", "1920x1080", "700x390mm", ""),
            ["LEN40A1"] = ("Lenovo", "13.9", "1600x900", "310x170mm", ""),
            ["LEN61F9"] = ("Lenovo ThinkVision S24e", "24.0", "1920x1080", "450x250mm", "1x VGA / 1x HDMI"),
            ["DELD0D8"] = ("Dell P2419H", "23.8", "1920x1080", "530x300mm", "1x VGA / 1x DisplayPort / 1x HDMI"),
            ["DELD0DA"] = ("Dell P2419H", "23.8", "1920x1080", "530x300mm", "1x VGA / 1x DisplayPort / 1x HDMI")
        };

    public static readonly Dictionary<string, string> MonitorManufacturerMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["DEL"] = "Dell",
            ["HWP"] = "HP",
            ["GSM"] = "LG",
            ["SAM"] = "Samsung",
            ["AOC"] = "AOC",
            ["LEN"] = "Lenovo",
            ["ACR"] = "Acer",
            ["ASU"] = "Asus",
            ["BEN"] = "BenQ",
            ["SNY"] = "Sony",
            ["PHL"] = "Philips",
            ["NEC"] = "NEC",
            ["VSC"] = "ViewSonic"
        };

    public static readonly Dictionary<string, string> GeneralManufacturerMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["LENOVO"] = "Lenovo",
            ["DELL"] = "Dell",
            ["DELL INC."] = "Dell",
            ["HP"] = "HP",
            ["HEWLETT-PACKARD"] = "HP",
            ["INTEL"] = "Intel",
            ["REALTEK"] = "Realtek",
            ["NVIDIA"] = "Nvidia",
            ["MICROSOFT"] = "Microsoft"
        };

    public static readonly Dictionary<int, string> JedecMap = new()
    {
        [0x4B] = "Kingston",
        [0x73] = "Samsung",
        [0x0198] = "Kingston",
        [0x802C] = "Micron",
        [0x80AD] = "Hynix",
        [0x015B] = "Elpida",
        [0x43] = "Corsair",
        [0x4D] = "Micron",
        [0xAD] = "Hynix",
        [0x83] = "Infinion",
        [0x85] = "Nanya",
        [0x02FE] = "Elpida",
        [0x0194] = "Smart Modular",
        [0x029E] = "ADATA",
        [0x04CB] = "ADATA",
        [0x014F] = "Transcend",
        [0x02C4] = "Crucial",
        [0x059B] = "Crucial",
        [0x070B] = "Crucial",
        [0x04CD] = "G.Skill",
        [0x0325] = "Patriot",
        [0x0A] = "Union Memory",
        [0x01A4] = "Union Memory"
    };

    public static readonly Dictionary<int, string> VideoTechMap = new()
    {
        [-1] = "Outro",
        [0] = "VGA",
        [1] = "S-Video",
        [2] = "Video Composto",
        [3] = "Video Componente",
        [4] = "DVI",
        [5] = "HDMI",
        [6] = "LVDS (Tela Interna)",
        [8] = "D-JPN",
        [9] = "SDI",
        [10] = "DisplayPort",
        [11] = "DisplayPort (Embedded)",
        [12] = "UDI",
        [13] = "UDI (Embedded)",
        [14] = "SDTV",
        [15] = "Miracast",
        [16] = "Indirect Wired"
    };

    public static readonly Dictionary<int, string> ChassisTypeMap = new()
    {
        [1] = "Other",
        [2] = "Unknown",
        [3] = "Desktop",
        [4] = "Low Profile Desktop",
        [5] = "Pizza Box",
        [6] = "Mini Tower",
        [7] = "Tower",
        [8] = "Portable",
        [9] = "Laptop",
        [10] = "Notebook",
        [11] = "Hand Held",
        [12] = "Docking Station",
        [13] = "All in One",
        [14] = "Sub Notebook",
        [15] = "Space-Saving",
        [16] = "Lunch Box",
        [17] = "Main System Chassis",
        [18] = "Expansion Chassis",
        [19] = "SubChassis",
        [20] = "Bus Expansion Chassis",
        [21] = "Peripheral Chassis",
        [22] = "RAID Chassis",
        [23] = "Rack Mount Chassis",
        [24] = "Sealed-Case PC",
        [25] = "Multi-System Chassis",
        [26] = "CompactPCI",
        [27] = "AdvancedTCA",
        [28] = "Blade",
        [29] = "Blade Enclosure",
        [30] = "Tablet",
        [31] = "Convertible",
        [32] = "Detachable",
        [33] = "IoT Gateway",
        [34] = "Embedded PC",
        [35] = "Mini PC",
        [36] = "Stick PC"
    };

    public static readonly Dictionary<string, (string Fabricante, string Modelo)> DiskModelMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["SAMSUNG MZVLB256HAHQ-000H1"] = ("Samsung", "PM981a NVMe 256GB"),
            ["SKHynix_HFS256GEJ9X113N"] = ("SK Hynix", "BC501 NVMe 256GB"),
            ["KINGSTON SA400S37240G"] = ("Kingston", "A400 SATA 240GB"),
            ["KINGSTON SA400S37480G"] = ("Kingston", "A400 SATA 480GB"),
            ["KINGSTON SNVS250G"] = ("Kingston", "NV1 NVMe 250GB"),
            ["KINGSTON SNVS500G"] = ("Kingston", "NV1 NVMe 500GB"),
            ["CT250MX500SSD1"] = ("Crucial", "MX500 SATA 250GB"),
            ["CT500MX500SSD1"] = ("Crucial", "MX500 SATA 500GB"),
            ["WDC WD10EZEX-08WN4A0"] = ("Western Digital", "Blue HDD 1TB"),
            ["WDC WDS250G2B0A"] = ("Western Digital", "Blue SATA 250GB"),
            ["ST1000DM010-2EP102"] = ("Seagate", "Barracuda HDD 1TB")
        };

    public static string GetFriendlyManufacturer(string? rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return "N/D";
        string clean = rawName.Trim();
        if (GeneralManufacturerMap.TryGetValue(clean, out string? mapped))
        {
            return mapped;
        }
        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(clean.ToLowerInvariant());
    }

    public static string FormatRamManufacturer(string? rawMan)
    {
        if (string.IsNullOrWhiteSpace(rawMan)) return "N/D";
        string trim = rawMan.Trim();
        var match = Regex.Match(trim, @"^(?:0x)?([A-Fa-f0-9]+)$");
        if (match.Success)
        {
            try
            {
                int code = Convert.ToInt32(match.Groups[1].Value, 16);
                if (JedecMap.TryGetValue(code, out string? jedecName))
                {
                    return jedecName;
                }
                return $"Desconhecido (0x{code:X})";
            }
            catch
            {
                return trim;
            }
        }
        return trim;
    }

    public static string ConvertVideoOutputTech(int code)
    {
        return VideoTechMap.TryGetValue(code, out string? val) ? val : $"Desconhecido ({code})";
    }

    public static string ConvertMonitorManufacturer(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return "N/D";
        string trim = code.Trim();
        return MonitorManufacturerMap.TryGetValue(trim, out string? val) ? val : trim;
    }

    public static (string Fabricante, string Modelo) GetParsedDiskInfo(string? rawModel)
    {
        if (string.IsNullOrWhiteSpace(rawModel))
            return ("N/D", "N/D");

        string trimmed = rawModel.Trim();
        if (DiskModelMap.TryGetValue(trimmed, out var mapped))
            return mapped;

        string m = trimmed;
        string fab = "Desconhecido";

        var prefixes = new (string Pattern, string Brand, bool StripPrefix)[]
        {
            (@"^SAMSUNG\s+(.*)", "Samsung", true),
            (@"^SKHynix[_\s]+(.*)", "SK Hynix", true),
            (@"^KINGSTON\s+(.*)", "Kingston", true),
            (@"^WDC\s+(.*)", "Western Digital", true),
            (@"^ST\d+(.*)", "Seagate", false),
            (@"^CT\d+(.*)", "Crucial", false),
            (@"^ADATA\s+(.*)", "ADATA", true),
            (@"^INTEL\s+(.*)", "Intel", true),
            (@"^KIOXIA\s+(.*)", "Kioxia", true),
            (@"^TOSHIBA\s+(.*)", "Toshiba", true),
            (@"^SanDisk\s+(.*)", "SanDisk", true),
            (@"^CORSAIR\s+(.*)", "Corsair", true),
            (@"^CRUCIAL\s+(.*)", "Crucial", true),
            (@"^PATRIOT\s+(.*)", "Patriot", true)
        };

        foreach (var p in prefixes)
        {
            var match = Regex.Match(m, p.Pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                fab = p.Brand;
                if (p.StripPrefix && match.Groups.Count > 1)
                {
                    m = match.Groups[1].Value;
                }
                break;
            }
        }

        return (fab, m);
    }

    public static string FormatInstallDate(string? rawDate)
    {
        if (string.IsNullOrWhiteSpace(rawDate)) return string.Empty;
        string d = rawDate.Trim();
        if (Regex.IsMatch(d, @"^\d{8}$"))
        {
            return $"{d.Substring(6, 2)}/{d.Substring(4, 2)}/{d.Substring(0, 4)}";
        }
        var m = Regex.Match(d, @"^(\d{4})-(\d{2})-(\d{2})");
        if (m.Success)
        {
            return $"{m.Groups[3].Value}/{m.Groups[2].Value}/{m.Groups[1].Value}";
        }
        return d;
    }
}

