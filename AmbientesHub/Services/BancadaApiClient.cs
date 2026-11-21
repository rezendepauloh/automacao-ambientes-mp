using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AmbientesHub.Services;

/// <summary>
/// Cliente HTTP compartilhado (JSON/REST) para comunicação aberta e segura
/// entre o AmbientesHub (Windows) e a API REST da Bancada STI (WSL ou RedHat).
/// Política fail-closed: exige BANCADA_API_URL (http/https sem userinfo) e BANCADA_API_KEY (>= 16 chars).
/// </summary>
public static class BancadaApiClient
{
    public const int MinApiKeyLength = 16;

    private static readonly HttpClient SharedClient = new()
    {
        Timeout = TimeSpan.FromSeconds(12)
    };

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static bool TryBuildApiEndpoint(
        string? baseUrl,
        string? apiKey,
        out string normalizedBaseUrl,
        out string cleanApiKey,
        out string error)
    {
        normalizedBaseUrl = string.Empty;
        cleanApiKey = string.Empty;
        error = string.Empty;

        string rawUrl = baseUrl?.Trim() ?? string.Empty;
        string rawKey = apiKey?.Trim() ?? string.Empty;

        if (!InputValidator.IsSafeHttpUrl(rawUrl))
        {
            error = "BANCADA_API_URL não configurada ou inválida (use http:// ou https:// sem credenciais embutidas).";
            return false;
        }

        if (string.IsNullOrWhiteSpace(rawKey) || rawKey.Length < MinApiKeyLength)
        {
            error = $"BANCADA_API_KEY ausente ou muito curta (mínimo {MinApiKeyLength} caracteres - fail-closed).";
            return false;
        }

        normalizedBaseUrl = rawUrl.TrimEnd('/');
        cleanApiKey = rawKey;
        return true;
    }

    public static bool IsApiConfigured()
    {
        return TryBuildApiEndpoint(
            ConfigService.BancadaApiUrl,
            ConfigService.BancadaApiKey,
            out _,
            out _,
            out _);
    }

    public static async Task<bool> CheckHealthAsync()
    {
        string rawUrl = ConfigService.BancadaApiUrl?.Trim() ?? string.Empty;
        if (!InputValidator.IsSafeHttpUrl(rawUrl))
        {
            AppLoggerService.Info("BancadaApiClient.CheckHealthAsync ignorado: BANCADA_API_URL ausente ou inválida.", nameof(BancadaApiClient));
            return false;
        }

        var sw = Stopwatch.StartNew();
        string healthUrl = $"{rawUrl.TrimEnd('/')}/health";
        try
        {
            using var req = new HttpRequestMessage(HttpMethod.Get, healthUrl);
            using var res = await SharedClient.SendAsync(req);
            AppLoggerService.Info(
                $"Healthcheck API Bancada '{healthUrl}': HTTP {(int)res.StatusCode} ({res.StatusCode}) em {AppLoggerService.FormatElapsed(sw)}.",
                nameof(BancadaApiClient));
            return res.IsSuccessStatusCode;
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaApiClient), $"Healthcheck API Bancada '{healthUrl}'", ex, sw.Elapsed, LogLevel.WARNING);
            return false;
        }
    }

    public static Task<(bool Success, T? Data, int StatusCode, string Error)> GetAsync<T>(string relativePath)
        => SendJsonAsync<object, T>(HttpMethod.Get, relativePath, null);

    public static Task<(bool Success, TRes? Data, int StatusCode, string Error)> PostAsync<TReq, TRes>(string relativePath, TReq body)
        => SendJsonAsync<TReq, TRes>(HttpMethod.Post, relativePath, body);

    public static Task<(bool Success, TRes? Data, int StatusCode, string Error)> PatchAsync<TReq, TRes>(string relativePath, TReq body)
        => SendJsonAsync<TReq, TRes>(HttpMethod.Patch, relativePath, body);

    private static async Task<(bool Success, TRes? Data, int StatusCode, string Error)> SendJsonAsync<TReq, TRes>(
        HttpMethod method,
        string relativePath,
        TReq? body)
    {
        if (!TryBuildApiEndpoint(
                ConfigService.BancadaApiUrl,
                ConfigService.BancadaApiKey,
                out string baseUrl,
                out string apiKey,
                out string cfgError))
        {
            return (false, default, 0, cfgError);
        }

        string path = relativePath.StartsWith('/') ? relativePath : $"/{relativePath}";
        string targetUrl = $"{baseUrl}{path}";
        var sw = Stopwatch.StartNew();

        try
        {
            using var req = new HttpRequestMessage(method, targetUrl);
            req.Headers.TryAddWithoutValidation("X-API-Key", apiKey);
            req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            if (body != null)
            {
                string json = JsonSerializer.Serialize(body, JsonOptions);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");
            }

            using var res = await SharedClient.SendAsync(req);
            string raw = await res.Content.ReadAsStringAsync();

            if (!res.IsSuccessStatusCode)
            {
                string detail = TryExtractDetail(raw);
                string msg = string.IsNullOrWhiteSpace(detail)
                    ? $"API Bancada retornou HTTP {(int)res.StatusCode}."
                    : detail;
                AppLoggerService.Warning(
                    $"API Bancada {method.Method} '{targetUrl}' falhou: HTTP {(int)res.StatusCode} ({msg}) em {AppLoggerService.FormatElapsed(sw)}.",
                    nameof(BancadaApiClient));
                return (false, default, (int)res.StatusCode, msg);
            }

            if (string.IsNullOrWhiteSpace(raw))
            {
                AppLoggerService.Info(
                    $"API Bancada {method.Method} '{targetUrl}' -> HTTP {(int)res.StatusCode} (corpo vazio) em {AppLoggerService.FormatElapsed(sw)}.",
                    nameof(BancadaApiClient));
                return (true, default, (int)res.StatusCode, string.Empty);
            }

            var parsed = JsonSerializer.Deserialize<TRes>(raw, JsonOptions);
            AppLoggerService.Info(
                $"API Bancada {method.Method} '{targetUrl}' -> HTTP {(int)res.StatusCode} ({raw.Length} bytes) em {AppLoggerService.FormatElapsed(sw)}.",
                nameof(BancadaApiClient));
            return (true, parsed, (int)res.StatusCode, string.Empty);
        }
        catch (TaskCanceledException ex)
        {
            AppLoggerService.Warning(
                $"Timeout (408) na API Bancada {method.Method} '{targetUrl}' após {AppLoggerService.FormatElapsed(sw)}: {ex.Message}",
                nameof(BancadaApiClient));
            return (false, default, 408, "Tempo limite excedido ao comunicar com a API REST da Bancada.");
        }
        catch (Exception ex)
        {
            AppLoggerService.LogException(nameof(BancadaApiClient), $"API Bancada {method.Method} '{targetUrl}'", ex, sw.Elapsed, LogLevel.WARNING);
            return (false, default, 0, $"Falha ao comunicar com a API REST ({baseUrl}): {ex.Message}");
        }
    }

    private static string TryExtractDetail(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return string.Empty;

        try
        {
            using var doc = JsonDocument.Parse(raw);
            if (doc.RootElement.TryGetProperty("message", out var msgProp) && msgProp.ValueKind == JsonValueKind.String)
                return msgProp.GetString() ?? string.Empty;
            if (doc.RootElement.TryGetProperty("detail", out var detProp))
                return detProp.ValueKind == JsonValueKind.String ? (detProp.GetString() ?? string.Empty) : detProp.ToString();
        }
        catch
        {
            // Ignora erro de parse em respostas HTML/texto
        }
        return string.Empty;
    }
}
