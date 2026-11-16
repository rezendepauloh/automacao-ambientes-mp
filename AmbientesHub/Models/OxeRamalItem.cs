namespace AmbientesHub.Models;

public class OxeRamalItem
{
    public string Ramal { get; set; } = string.Empty;
    public string Nome { get; set; } = string.Empty;
    public string Localidade { get; set; } = string.Empty;
    public string Setor { get; set; } = string.Empty;
    public string Tipo { get; set; } = string.Empty;
    public string Ip { get; set; } = string.Empty;
    public string Mac { get; set; } = string.Empty;
    public string GrupoCaptura { get; set; } = string.Empty;
    public string CatRedePublica { get; set; } = string.Empty;
    public string CategoriaDispositivo { get; set; } = string.Empty;
    public string TipoEstacao { get; set; } = string.Empty;
    public Dictionary<string, string> AllFields { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public string RamalDisplay => string.IsNullOrWhiteSpace(Ramal) ? "S/N" : Ramal;
    public string NomeDisplay => string.IsNullOrWhiteSpace(Nome) ? "Sem nome associado" : Nome;

    public string SubtituloDisplay
    {
        get
        {
            var parts = new[] { Localidade, Setor, Tipo }
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase);
            string joined = string.Join(" • ", parts);
            return string.IsNullOrWhiteSpace(joined) ? "Sem localização/tipo cadastrado" : joined;
        }
    }

    public string RedeDisplay
    {
        get
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(Ip)) parts.Add($"IP: {Ip}");
            if (!string.IsNullOrWhiteSpace(Mac)) parts.Add($"MAC: {Mac}");
            return parts.Count > 0 ? string.Join("  |  ", parts) : "Sem IP/MAC";
        }
    }

    public string OxeBadgesDisplay
    {
        get
        {
            string D(string v) => string.IsNullOrWhiteSpace(v) ? "-" : v;
            return $"👥 Captura: {D(GrupoCaptura)}  |  🌐 Rede Púb: {D(CatRedePublica)}  |  📱 Dispositivo: {D(CategoriaDispositivo)}  |  ⚙️ Estação: {D(TipoEstacao)}";
        }
    }
}

public class OxeSummary
{
    public int Total { get; set; }
    public int ComNome { get; set; }
    public int ComIpOuMac { get; set; }
}
