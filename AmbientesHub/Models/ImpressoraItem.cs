using System.Globalization;

namespace AmbientesHub.Models;

public class ImpressoraItem
{
    public string Id { get; set; } = string.Empty;
    public string Nome { get; set; } = string.Empty;
    public string Servidor { get; set; } = string.Empty;
    public string Tipo { get; set; } = string.Empty;
    public string Modelo { get; set; } = string.Empty;
    public string Localizacao { get; set; } = string.Empty;
    public string IpHost { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public long TotalPaginas { get; set; }
    public string FilasRelacionadas { get; set; } = string.Empty;
    public string DetalhesExtra { get; set; } = string.Empty;
    public string DataAtualizacao { get; set; } = string.Empty;
    public Dictionary<string, string> AllFields { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public string NomeDisplay => string.IsNullOrWhiteSpace(Nome) ? "Impressora sem nome" : Nome;
    public string TipoDisplay => string.IsNullOrWhiteSpace(Tipo) ? "Impressora" : Tipo;
    public string StatusDisplay => string.IsNullOrWhiteSpace(Status) ? "N/D" : Status;

    public bool IsStatusOk =>
        StatusDisplay.Equals("OK", StringComparison.OrdinalIgnoreCase) ||
        StatusDisplay.Contains("online", StringComparison.OrdinalIgnoreCase) ||
        StatusDisplay.Contains("pront", StringComparison.OrdinalIgnoreCase) ||
        StatusDisplay.Contains("normal", StringComparison.OrdinalIgnoreCase);

    public string StatusBadgeBackground => IsStatusOk ? "#064e3b" : "#7f1d1d";
    public string StatusBadgeForeground => IsStatusOk ? "#34d399" : "#fca5a5";

    public string SubtituloDisplay
    {
        get
        {
            var parts = new[] { Localizacao, Modelo, Servidor }
                .Where(p => !string.IsNullOrWhiteSpace(p) && p.Trim() != "-")
                .Distinct(StringComparer.OrdinalIgnoreCase);
            string joined = string.Join(" • ", parts);
            return string.IsNullOrWhiteSpace(joined) ? "Sem localização/modelo cadastrado" : joined;
        }
    }

    public string RedeEPaginasDisplay
    {
        get
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(IpHost) && IpHost.Trim() != "-")
                parts.Add($"🌐 {IpHost}");
            parts.Add($"📄 {TotalPaginas.ToString("N0", new CultureInfo("pt-BR"))} págs");
            if (!string.IsNullOrWhiteSpace(FilasRelacionadas) && FilasRelacionadas.Trim() != "-")
                parts.Add($"🖨️ Filas: {FilasRelacionadas}");
            return string.Join("  |  ", parts);
        }
    }

    public bool HasIpHost => !string.IsNullOrWhiteSpace(IpHost) && IpHost.Trim() != "-";
}

public class ImpressorasSummary
{
    public int Total { get; set; }
    public int OkCount { get; set; }
    public int AlertCount { get; set; }
    public long TotalPaginasSoma { get; set; }
}
