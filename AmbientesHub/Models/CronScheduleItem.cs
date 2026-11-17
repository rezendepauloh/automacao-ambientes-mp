using System.Globalization;

namespace AmbientesHub.Models;

public class CronScheduleItem
{
    public string TaskId { get; set; } = string.Empty;
    public string Nome { get; set; } = string.Empty;
    public string Categoria { get; set; } = string.Empty;
    public bool Ativo { get; set; }
    public string TipoAgendamento { get; set; } = "intervalo";
    public int IntervaloValor { get; set; } = 2;
    public string IntervaloUnidade { get; set; } = "horas";
    public string HorarioFixo { get; set; } = "12:00";
    public bool ApenasDiasUteis { get; set; }
    public string Descricao { get; set; } = string.Empty;
    public string UltimaExecucao { get; set; } = string.Empty;
    public string UltimoStatus { get; set; } = "pendente";
    public string UltimoLog { get; set; } = string.Empty;
    public bool RunRequested { get; set; }

    public bool IsExecuting => UltimoStatus.Equals("executando", StringComparison.OrdinalIgnoreCase);
    public bool CanTrigger => !IsExecuting && !RunRequested;

    public string StatusDisplay => RunRequested && !IsExecuting
        ? "AGUARDANDO CONTAINER..."
        : (string.IsNullOrWhiteSpace(UltimoStatus) ? "PENDENTE" : UltimoStatus.ToUpperInvariant());

    public string StatusBadgeBackground
    {
        get
        {
            if (IsExecuting || RunRequested) return "#1e3a8a";
            if (UltimoStatus.Equals("sucesso", StringComparison.OrdinalIgnoreCase)) return "#064e3b";
            if (UltimoStatus.Equals("erro", StringComparison.OrdinalIgnoreCase)) return "#7f1d1d";
            return "#78350f";
        }
    }

    public string StatusBadgeForeground
    {
        get
        {
            if (IsExecuting || RunRequested) return "#93c5fd";
            if (UltimoStatus.Equals("sucesso", StringComparison.OrdinalIgnoreCase)) return "#34d399";
            if (UltimoStatus.Equals("erro", StringComparison.OrdinalIgnoreCase)) return "#fca5a5";
            return "#fcd34d";
        }
    }

    public string AgendamentoDisplay
    {
        get
        {
            string baseFreq = TipoAgendamento.Equals("horario_fixo", StringComparison.OrdinalIgnoreCase)
                ? $"Diário às {HorarioFixo}"
                : $"A cada {IntervaloValor} {IntervaloUnidade}";
            return ApenasDiasUteis ? $"{baseFreq} (dias úteis)" : baseFreq;
        }
    }

    public string UltimaExecucaoBr
    {
        get
        {
            if (string.IsNullOrWhiteSpace(UltimaExecucao) ||
                UltimaExecucao.Equals("null", StringComparison.OrdinalIgnoreCase) ||
                UltimaExecucao.Equals("none", StringComparison.OrdinalIgnoreCase))
                return "Nunca";

            if (DateTime.TryParse(UltimaExecucao, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt) ||
                DateTime.TryParse(UltimaExecucao, new CultureInfo("pt-BR"), DateTimeStyles.AssumeLocal, out dt))
            {
                return dt.ToString("dd/MM/yyyy HH:mm:ss");
            }
            return UltimaExecucao.Trim();
        }
    }

    public string UltimoLogDisplay => string.IsNullOrWhiteSpace(UltimoLog) || UltimoLog.Equals("None", StringComparison.OrdinalIgnoreCase)
        ? "Sem registros recentes."
        : UltimoLog.Trim();
}

public class CronLogItem
{
    public int Id { get; set; }
    public string Rotina { get; set; } = string.Empty;
    public string Inicio { get; set; } = string.Empty;
    public string DuracaoSegundos { get; set; } = "0";
    public string Status { get; set; } = string.Empty;
    public string Mensagem { get; set; } = string.Empty;

    public string InicioBr
    {
        get
        {
            if (DateTime.TryParse(Inicio, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt))
                return dt.ToString("dd/MM/yyyy HH:mm:ss");
            return string.IsNullOrWhiteSpace(Inicio) ? "-" : Inicio;
        }
    }

    public string StatusDisplay => string.IsNullOrWhiteSpace(Status) ? "-" : Status.ToUpperInvariant();
}

public class CronSummary
{
    public int Total { get; set; }
    public int Ativas { get; set; }
    public int Executando { get; set; }
    public int ComErro { get; set; }
}
