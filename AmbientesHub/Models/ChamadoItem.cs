using System.Windows;

namespace AmbientesHub.Models;

public class ChamadoItem
{
    public string Id { get; set; } = string.Empty;
    public string DataCriacao { get; set; } = string.Empty;
    public string Titulo { get; set; } = string.Empty;
    public string CidadePredio { get; set; } = string.Empty;
    public string Unidade { get; set; } = string.Empty;
    public string LocalidadeFisica { get; set; } = string.Empty;
    public string Usuario { get; set; } = string.Empty;
    public string IdCliente { get; set; } = string.Empty;
    public string Descricao { get; set; } = string.Empty;
    public string Tag { get; set; } = string.Empty;
    public string IpOrigem { get; set; } = string.Empty;
    public string Status { get; set; } = "Aberto";
    public string DataAtualizacao { get; set; } = string.Empty;
    public string Base { get; set; } = string.Empty;
    public string Andamento { get; set; } = string.Empty;
    public string Link { get; set; } = string.Empty;
    public string Hostname { get; set; } = string.Empty;

    public string TituloDisplay => string.IsNullOrWhiteSpace(Titulo) ? $"Chamado #{Id}" : Titulo;
    public string TagDisplay => string.IsNullOrWhiteSpace(Tag) ? "SEM TAG" : Tag.ToUpperInvariant();
    public string BaseDisplay => string.IsNullOrWhiteSpace(Base) ? "N/D" : Base.ToUpperInvariant();

    public string LocalizacaoDisplay
    {
        get
        {
            var parts = new[] { CidadePredio, Unidade, LocalidadeFisica }
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase);
            string joined = string.Join(" • ", parts);
            return string.IsNullOrWhiteSpace(joined) ? "Localidade não informada" : joined;
        }
    }

    public string SolicitanteDisplay => string.IsNullOrWhiteSpace(Usuario)
        ? (string.IsNullOrWhiteSpace(IdCliente) ? "Não informado" : IdCliente)
        : Usuario;

    public string BaseBadgeBackground => BaseDisplay.Contains("OTRS", StringComparison.OrdinalIgnoreCase)
        ? "#1e3a8a"
        : "#0f766e";

    public string BaseBadgeForeground => BaseDisplay.Contains("OTRS", StringComparison.OrdinalIgnoreCase)
        ? "#93c5fd"
        : "#5eead4";

    public string StatusBadgeBackground => Status.Equals("Aberto", StringComparison.OrdinalIgnoreCase)
        ? "#064e3b"
        : "#1e293b";

    public string StatusBadgeForeground => Status.Equals("Aberto", StringComparison.OrdinalIgnoreCase)
        ? "#34d399"
        : "#94a3b8";

    public Visibility LinkButtonVisibility => !string.IsNullOrWhiteSpace(Link) ? Visibility.Visible : Visibility.Collapsed;
    public Visibility HostnameButtonVisibility => !string.IsNullOrWhiteSpace(Hostname) ? Visibility.Visible : Visibility.Collapsed;
    public Visibility AndamentoVisibility => !string.IsNullOrWhiteSpace(Andamento) ? Visibility.Visible : Visibility.Collapsed;
    public string AndamentoDisplay => string.IsNullOrWhiteSpace(Andamento) ? string.Empty : Andamento.Trim();

    public string DataCriacaoBr
    {
        get
        {
            if (string.IsNullOrWhiteSpace(DataCriacao)) return "-";
            if (DateTime.TryParse(DataCriacao, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeLocal, out var dt) ||
                DateTime.TryParse(DataCriacao, new System.Globalization.CultureInfo("pt-BR"), System.Globalization.DateTimeStyles.AssumeLocal, out dt))
            {
                return dt.ToString("dd/MM/yyyy HH:mm:ss");
            }
            return DataCriacao.Trim();
        }
    }
}

public class ChamadosSummary
{
    public int Total { get; set; }
    public int Abertos { get; set; }
    public int OtrsCount { get; set; }
    public int CitSmartCount { get; set; }
}

public class ChamadoComentario
{
    public string Data { get; set; } = string.Empty;
    public string Autor { get; set; } = string.Empty;
    public string Texto { get; set; } = string.Empty;
}
