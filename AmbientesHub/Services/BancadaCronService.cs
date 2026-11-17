using System.Text.Json.Serialization;
using AmbientesHub.Models;
using Npgsql;

namespace AmbientesHub.Services;

public class BancadaCronService
{
    public async Task<(bool Success, List<CronScheduleItem> Schedules, List<CronLogItem> Logs, CronSummary Summary, string Message)> QueryCronDataAsync()
    {
        if (BancadaApiClient.IsApiConfigured())
        {
            var (ok, res, _, err) = await BancadaApiClient.GetAsync<ApiCronResponse>("/api/v1/cron");
            if (!ok || res == null)
                return (false, [], [], new CronSummary(), err);

            var apiSchedules = (res.Schedules ?? []).Select(s => new CronScheduleItem
            {
                TaskId = s.TaskId ?? string.Empty,
                Nome = s.Nome ?? string.Empty,
                Categoria = s.Categoria ?? string.Empty,
                Ativo = s.Ativo,
                TipoAgendamento = string.IsNullOrWhiteSpace(s.TipoAgendamento) ? "intervalo" : s.TipoAgendamento,
                IntervaloValor = s.IntervaloValor > 0 ? s.IntervaloValor : 2,
                IntervaloUnidade = string.IsNullOrWhiteSpace(s.IntervaloUnidade) ? "horas" : s.IntervaloUnidade,
                HorarioFixo = string.IsNullOrWhiteSpace(s.HorarioFixo) ? "12:00" : s.HorarioFixo,
                ApenasDiasUteis = s.ApenasDiasUteis,
                Descricao = s.Descricao ?? string.Empty,
                UltimaExecucao = s.UltimaExecucao ?? string.Empty,
                UltimoStatus = string.IsNullOrWhiteSpace(s.UltimoStatus) ? "pendente" : s.UltimoStatus,
                UltimoLog = s.UltimoLog ?? string.Empty,
                RunRequested = s.RunRequested
            }).ToList();

            var apiLogs = (res.Logs ?? []).Select(l => new CronLogItem
            {
                Id = l.Id,
                Rotina = l.Rotina ?? string.Empty,
                Inicio = l.Inicio ?? string.Empty,
                DuracaoSegundos = l.DuracaoSegundos ?? "0",
                Status = l.Status ?? string.Empty,
                Mensagem = l.Mensagem ?? string.Empty
            }).ToList();

            var apiSummary = new CronSummary
            {
                Total = res.Summary?.Total ?? apiSchedules.Count,
                Ativas = res.Summary?.Ativas ?? apiSchedules.Count(x => x.Ativo),
                Executando = res.Summary?.Executando ?? apiSchedules.Count(x => x.IsExecuting || x.RunRequested),
                ComErro = res.Summary?.ComErro ?? apiSchedules.Count(x => x.UltimoStatus.Equals("erro", StringComparison.OrdinalIgnoreCase))
            };

            return (true, apiSchedules, apiLogs, apiSummary, $"{apiSchedules.Count} rotina(s) de agendamento carregada(s) via API REST.");
        }

        if (!BancadaChamadosService.TryBuildConnectionStringFromConfig(out string connStr, out string error))
        {
            return (false, [], [], new CronSummary(), error);
        }

        var schedules = new List<CronScheduleItem>();
        var logs = new List<CronLogItem>();

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            try
            {
                await using var alterCmd = new NpgsqlCommand(
                    "ALTER TABLE cron_schedules ADD COLUMN IF NOT EXISTS run_requested INTEGER DEFAULT 0;", conn);
                await alterCmd.ExecuteNonQueryAsync();
            }
            catch
            {
                // Ignora caso a tabela ainda não exista ou usuário seja apenas leitura
            }

            const string sqlSchedules = """
                SELECT
                    COALESCE(task_id, ''),
                    COALESCE(nome, ''),
                    COALESCE(categoria, ''),
                    COALESCE(ativo, 1),
                    COALESCE(tipo_agendamento, 'intervalo'),
                    COALESCE(intervalo_valor, 2),
                    COALESCE(intervalo_unidade, 'horas'),
                    COALESCE(horario_fixo, '12:00'),
                    COALESCE(apenas_dias_uteis, 0),
                    COALESCE(descricao, ''),
                    COALESCE(ultima_execucao::text, ''),
                    COALESCE(ultimo_status, 'pendente'),
                    COALESCE(ultimo_log, ''),
                    COALESCE(run_requested, 0)
                FROM cron_schedules
                ORDER BY COALESCE(categoria, ''), COALESCE(nome, '') ASC;
                """;

            await using (var cmd = new NpgsqlCommand(sqlSchedules, conn))
            await using (var reader = await cmd.ExecuteReaderAsync())
            {
                while (await reader.ReadAsync())
                {
                    schedules.Add(new CronScheduleItem
                    {
                        TaskId = reader.GetString(0),
                        Nome = reader.GetString(1),
                        Categoria = reader.GetString(2),
                        Ativo = reader.GetInt32(3) == 1,
                        TipoAgendamento = reader.GetString(4),
                        IntervaloValor = reader.GetInt32(5),
                        IntervaloUnidade = reader.GetString(6),
                        HorarioFixo = reader.GetString(7),
                        ApenasDiasUteis = reader.GetInt32(8) == 1,
                        Descricao = reader.GetString(9),
                        UltimaExecucao = reader.GetString(10),
                        UltimoStatus = reader.GetString(11),
                        UltimoLog = reader.GetString(12),
                        RunRequested = reader.GetInt32(13) == 1
                    });
                }
            }

            const string sqlLogs = """
                SELECT
                    l.id,
                    COALESCE(s.nome, l.task_id),
                    COALESCE(l.inicio::text, ''),
                    COALESCE(l.duracao_segundos::text, '0'),
                    COALESCE(l.status, ''),
                    COALESCE(l.mensagem, '')
                FROM cron_logs l
                LEFT JOIN cron_schedules s ON l.task_id = s.task_id
                ORDER BY l.id DESC
                LIMIT 25;
                """;

            await using (var logCmd = new NpgsqlCommand(sqlLogs, conn))
            await using (var logReader = await logCmd.ExecuteReaderAsync())
            {
                while (await logReader.ReadAsync())
                {
                    logs.Add(new CronLogItem
                    {
                        Id = logReader.GetInt32(0),
                        Rotina = logReader.GetString(1),
                        Inicio = logReader.GetString(2),
                        DuracaoSegundos = logReader.GetString(3),
                        Status = logReader.GetString(4),
                        Mensagem = logReader.GetString(5)
                    });
                }
            }

            var summary = new CronSummary
            {
                Total = schedules.Count,
                Ativas = schedules.Count(s => s.Ativo),
                Executando = schedules.Count(s => s.IsExecuting || s.RunRequested),
                ComErro = schedules.Count(s => s.UltimoStatus.Equals("erro", StringComparison.OrdinalIgnoreCase))
            };

            return (true, schedules, logs, summary, $"{schedules.Count} rotina(s) de agendamento carregada(s) do PostgreSQL.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao consultar agendamentos no PostgreSQL: {ex.Message}", nameof(BancadaCronService));
            return (false, [], [], new CronSummary(), $"Erro ao consultar Agendadores no PostgreSQL: {ex.Message}");
        }
    }

    public async Task<(bool Success, string Message)> TriggerTaskNowAsync(string taskId)
    {
        if (string.IsNullOrWhiteSpace(taskId))
            return (false, "ID da tarefa inválido.");

        if (BancadaApiClient.IsApiConfigured())
        {
            string encodedId = Uri.EscapeDataString(taskId.Trim());
            var (ok, res, _, err) = await BancadaApiClient.PostAsync<object, ApiCronActionResponse>($"/api/v1/cron/{encodedId}/trigger", new { });
            return ok
                ? (true, res?.Message ?? $"Execução solicitada para '{taskId}'.")
                : (false, err);
        }

        if (!BancadaChamadosService.TryBuildConnectionStringFromConfig(out string connStr, out string error))
            return (false, error);

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            await using (var alterCmd = new NpgsqlCommand(
                "ALTER TABLE cron_schedules ADD COLUMN IF NOT EXISTS run_requested INTEGER DEFAULT 0;", conn))
            {
                await alterCmd.ExecuteNonQueryAsync();
            }

            const string sql = """
                UPDATE cron_schedules
                SET run_requested = 1,
                    updated_at = CURRENT_TIMESTAMP
                WHERE task_id = @taskId
                  AND LOWER(COALESCE(ultimo_status, '')) <> 'executando';
                """;

            await using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@taskId", taskId.Trim());
            int rows = await cmd.ExecuteNonQueryAsync();

            return rows > 0
                ? (true, $"Execução solicitada para '{taskId}'. O container (WSL/RedHat) iniciará em segundos.")
                : (false, $"A tarefa '{taskId}' já está em execução ou não foi encontrada.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao disparar tarefa {taskId}: {ex.Message}", nameof(BancadaCronService));
            return (false, $"Erro ao solicitar execução: {ex.Message}");
        }
    }

    public async Task<(bool Success, string Message)> SetTaskActiveAsync(string taskId, bool ativo)
    {
        if (string.IsNullOrWhiteSpace(taskId))
            return (false, "ID da tarefa inválido.");

        if (BancadaApiClient.IsApiConfigured())
        {
            string encodedId = Uri.EscapeDataString(taskId.Trim());
            var (ok, res, _, err) = await BancadaApiClient.PatchAsync<ApiCronToggleRequest, ApiCronActionResponse>(
                $"/api/v1/cron/{encodedId}",
                new ApiCronToggleRequest { Ativo = ativo });
            return ok
                ? (true, res?.Message ?? $"Rotina '{taskId}' {(ativo ? "habilitada" : "desabilitada")}.")
                : (false, err);
        }

        if (!BancadaChamadosService.TryBuildConnectionStringFromConfig(out string connStr, out string error))
            return (false, error);

        try
        {
            await using var conn = new NpgsqlConnection(connStr);
            await conn.OpenAsync();

            const string sql = """
                UPDATE cron_schedules
                SET ativo = @ativo,
                    updated_at = CURRENT_TIMESTAMP
                WHERE task_id = @taskId;
                """;

            await using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@taskId", taskId.Trim());
            cmd.Parameters.AddWithValue("@ativo", ativo ? 1 : 0);
            int rows = await cmd.ExecuteNonQueryAsync();

            return rows > 0
                ? (true, $"Rotina '{taskId}' {(ativo ? "habilitada" : "desabilitada")}.")
                : (false, $"Rotina '{taskId}' não encontrada.");
        }
        catch (Exception ex)
        {
            AppLoggerService.Warning($"Falha ao alterar status da rotina {taskId}: {ex.Message}", nameof(BancadaCronService));
            return (false, $"Erro ao atualizar rotina: {ex.Message}");
        }
    }

    private sealed class ApiCronResponse
    {
        [JsonPropertyName("schedules")] public List<ApiCronScheduleDto>? Schedules { get; set; }
        [JsonPropertyName("logs")] public List<ApiCronLogDto>? Logs { get; set; }
        [JsonPropertyName("summary")] public ApiCronSummaryDto? Summary { get; set; }
    }

    private sealed class ApiCronScheduleDto
    {
        [JsonPropertyName("task_id")] public string? TaskId { get; set; }
        [JsonPropertyName("nome")] public string? Nome { get; set; }
        [JsonPropertyName("categoria")] public string? Categoria { get; set; }
        [JsonPropertyName("ativo")] public bool Ativo { get; set; }
        [JsonPropertyName("tipo_agendamento")] public string? TipoAgendamento { get; set; }
        [JsonPropertyName("intervalo_valor")] public int IntervaloValor { get; set; }
        [JsonPropertyName("intervalo_unidade")] public string? IntervaloUnidade { get; set; }
        [JsonPropertyName("horario_fixo")] public string? HorarioFixo { get; set; }
        [JsonPropertyName("apenas_dias_uteis")] public bool ApenasDiasUteis { get; set; }
        [JsonPropertyName("descricao")] public string? Descricao { get; set; }
        [JsonPropertyName("ultima_execucao")] public string? UltimaExecucao { get; set; }
        [JsonPropertyName("ultimo_status")] public string? UltimoStatus { get; set; }
        [JsonPropertyName("ultimo_log")] public string? UltimoLog { get; set; }
        [JsonPropertyName("run_requested")] public bool RunRequested { get; set; }
    }

    private sealed class ApiCronLogDto
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("rotina")] public string? Rotina { get; set; }
        [JsonPropertyName("inicio")] public string? Inicio { get; set; }
        [JsonPropertyName("duracao_segundos")] public string? DuracaoSegundos { get; set; }
        [JsonPropertyName("status")] public string? Status { get; set; }
        [JsonPropertyName("mensagem")] public string? Mensagem { get; set; }
    }

    private sealed class ApiCronSummaryDto
    {
        [JsonPropertyName("total")] public int Total { get; set; }
        [JsonPropertyName("ativas")] public int Ativas { get; set; }
        [JsonPropertyName("executando")] public int Executando { get; set; }
        [JsonPropertyName("com_erro")] public int ComErro { get; set; }
    }

    private sealed class ApiCronToggleRequest
    {
        [JsonPropertyName("ativo")] public bool Ativo { get; set; }
    }

    private sealed class ApiCronActionResponse
    {
        [JsonPropertyName("success")] public bool Success { get; set; }
        [JsonPropertyName("message")] public string? Message { get; set; }
    }
}
