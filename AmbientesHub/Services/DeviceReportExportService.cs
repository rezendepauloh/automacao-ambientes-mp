using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Xml.Linq;
using AmbientesHub.Models;

namespace AmbientesHub.Services;

public static class DeviceReportExportService
{
    public static string BuildHtmlReport(DeviceReportData data)
    {
        string compName = E(string.IsNullOrWhiteSpace(data.HardwareInfo.Computador) ? data.ComputerName : data.HardwareInfo.Computador);
        string hostName = E(string.IsNullOrWhiteSpace(data.OS.NomeMaquina) ? data.ComputerName : data.OS.NomeMaquina);
        string nowStr = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

        var hwTable = BuildTable(
            ["Computador", "Modelo", "Fabricante", "Memória Física (GB)", "Processadores Lógicos", "Processadores Físicos", "Serial Number", "Portas de Vídeo Ativas", "Entradas de vídeo"],
            [[data.HardwareInfo.Computador, data.HardwareInfo.Modelo, data.HardwareInfo.Fabricante, data.HardwareInfo.MemoriaFisicaGb, data.HardwareInfo.ProcessadoresLogicos.ToString(), data.HardwareInfo.ProcessadoresFisicos.ToString(), data.HardwareInfo.SerialNumber, data.HardwareInfo.PortasVideoAtivas, data.HardwareInfo.EntradasVideo]]);

        var userTable = BuildTable(
            ["Usuário Logado agora", "DisplayName AD", "Domínio", "Nome do Computador"],
            [[data.CurrentUser, data.CurrentUserDisplayName, data.CurrentUserDomain, data.HardwareInfo.Computador]]);

        var osTable = BuildTable(
            ["SO", "Versão WMI (Version)", "Versão do Windows (HnH)", "Número da Build", "Nome da Máquina", "Último Boot"],
            [[data.OS.SO, data.OS.VersaoWmi, data.OS.VersaoWindowsHnH, data.OS.NumeroBuild, data.OS.NomeMaquina, data.OS.UltimoBoot]]);

        var biosTable = BuildTable(
            ["Fabricante", "Versão da BIOS", "Número de Serial", "Data de lançamento"],
            [[data.BIOS.Fabricante, data.BIOS.VersaoBios, data.BIOS.NumeroSerial, data.BIOS.DataLancamento]]);

        var cpuTable = BuildTable(
            ["Nome", "Número de Cores", "Número de processadores lógicos", "Clock máximo", "Voltagem", "L1 Cache", "L2 Cache", "L3 Cache"],
            [[data.Processor.Nome, data.Processor.NumeroCores.ToString(), data.Processor.ProcessadoresLogicos.ToString(), data.Processor.ClockMaximo, data.Processor.Voltagem, data.Processor.L1Cache, data.Processor.L2Cache, data.Processor.L3Cache]]);

        var mbTable = BuildTable(
            ["Fabricante", "Produto", "Versão"],
            [[data.Motherboard.Fabricante, data.Motherboard.Produto, data.Motherboard.Versao]]);

        var chassisTable = BuildTable(
            ["Tipo de Chassi", "Fabricante"],
            [[data.Chassis.TipoChassi, data.Chassis.Fabricante]]);

        var ramTable = BuildTable(
            ["Fabricante", "Capacidade (GB)", "Velocidade (MHz)", "Part Number", "Número de Série"],
            data.MemoryModules.Select(m => new[] { m.Fabricante, m.CapacidadeGb, m.VelocidadeMhz, m.PartNumber, m.NumeroSerie }));

        var diskTable = BuildTable(
            ["Fabricante", "Modelo", "Tipo de interface", "Tamanho em GB", "Tipo de mídia", "Número de Série"],
            data.DiskDrives.Select(d => new[] { d.Fabricante, d.Modelo, d.TipoInterface, d.TamanhoGb, d.TipoMidia, d.NumeroSerie }));

        var ldiskTable = BuildTable(
            ["DeviceID", "Espaço Livre", "Espaço Usado", "Espaço Total", "Nome do Volume", "Sistema de Arquivos"],
            data.LogicalDisks.Select(d => new[] { d.DeviceId, d.EspacoLivre, d.EspacoUsado, d.EspacoTotal, d.NomeVolume, d.SistemaArquivos }));

        var netTable = BuildTable(
            ["Nome da placa", "MAC Address", "Endereços IPs", "Gateway"],
            data.NetworkConfigs.Select(n => new[] { n.NomePlaca, n.MacAddress, n.EnderecosIp, n.Gateway }));

        var monTable = BuildTable(
            ["Nome", "Polegada", "Resolução", "Tamanho", "Ativo", "Serial Number", "Nome amigável", "Fabricante", "Entradas de vídeo", "Porta Ativa"],
            data.Monitors.Select(m => new[] { m.Nome, m.Polegada, m.Resolucao, m.Tamanho, m.Ativo, m.SerialNumber, m.NomeAmigavel, m.Fabricante, m.EntradasVideo, m.PortaAtiva }));

        var gpuTable = BuildTable(
            ["Nome", "Processador de Vídeo", "Adaptador de compatibilidade", "Capacidade (GB)", "DeviceID", "Resolução Horizontal", "Resolução Vertical", "Número de Cores", "Versão do Driver", "Data do Driver Instalado"],
            data.VideoControllers.Select(v => new[] { v.Nome, v.ProcessadorVideo, v.AdaptadorCompatibilidade, v.CapacidadeGb, v.DeviceId, v.ResolucaoHorizontal, v.ResolucaoVertical, v.NumeroCores, v.VersaoDriver, v.DataDriverInstalado }));

        var soundTable = BuildTable(
            ["Nome", "Fabricante", "Status", "Informação do Status"],
            data.SoundDevices.Select(s => new[] { s.Nome, s.Fabricante, s.Status, s.InformacaoStatus }));

        var printLocalTable = BuildTable(
            ["Nome", "Estado da impressão", "Status da impressão", "Nome de compartilhamento", "Porta"],
            data.PrintersLocal.Select(p => new[] { p.Nome, p.EstadoImpressao, p.StatusImpressao, p.NomeCompartilhamento, p.Porta }));

        string mappedPrintersHtml = data.PrintersMapped.Count > 0
            ? "<h3 style=\"margin-top:15px; font-size:14px; color:#475569;\">Impressoras Mapeadas de Usuários</h3>" +
              BuildTable(
                  ["ComputerName", "UserSID", "UserName", "Type", "Server", "PrinterName"],
                  data.PrintersMapped.Select(p => new[] { p.ComputerName, p.UserSid, p.UserName, p.Type, p.Server, p.PrinterName }))
            : string.Empty;

        string extraLinks = string.Empty;
        string extraSections = string.Empty;

        if (!data.SkipMajorData)
        {
            extraLinks = """
        <a href="#sec-users-all">👥 Usuários & Perfis</a>
        <a href="#sec-users-chart" class="sub-link">📊 Gráfico de Perfis</a>
        <a href="#sec-programs">📦 Programas Instalados</a>
        <a href="#sec-programs-chart" class="sub-link">📊 Gráfico de Programas</a>
        <a href="#sec-services">⚙️ Serviços em Execução</a>
        <a href="#sec-services-chart" class="sub-link">📊 Gráfico de Serviços</a>
        <a href="#sec-processes">📊 Processos Ativos</a>
        <a href="#sec-processes-chart" class="sub-link">📊 Gráfico de Processos</a>
        <a href="#sec-drivers">🔌 Drivers do Sistema</a>
        <a href="#sec-drivers-chart" class="sub-link">📊 Gráfico de Drivers</a>
""";

            var usersTable = BuildTable(
                ["Nome", "Usuário", "Descrição", "Escritório", "Caminho do Perfil", "SID", "Último Uso", "Loaded", "Special", "Espaço em Disco (GB)"],
                data.UserProfiles.Select(u => new[] { u.Nome, u.Usuario, u.Descricao, u.Escritorio, u.CaminhoPerfil, u.Sid, u.UltimoUso, u.StatusSessao, u.Special ? "Sim" : "Não", u.EspacoDiscoGb.ToString("0.00") }));

            var progsTable = BuildTable(
                ["Nome do programa", "Versão", "Fabricante", "Data da instalação"],
                data.InstalledPrograms.Select(p => new[] { p.NomePrograma, p.Versao, p.Fabricante, p.DataInstalacao }));

            var svcsTable = BuildTable(
                ["Nome técnico", "Nome de exibição", "Status", "Tipo de inicialização", "Conta de logon", "Caminho do executável"],
                data.RunningServices.Select(s => new[] { s.NomeTecnico, s.NomeExibicao, s.Status, s.TipoInicializacao, s.ContaLogon, s.CaminhoExecutavel }));

            var procsTable = BuildTable(
                ["PID", "Nome do executável", "Usuário", "Uso de Memória (MB)", "Threads", "Handles", "Caminho do arquivo", "Linha de comando"],
                data.ActiveProcesses.Select(p => new[] { p.Pid.ToString(), p.NomeExecutavel, p.Usuario, p.UsoMemoriaMb.ToString("0.00"), p.Threads.ToString(), p.Handles.ToString(), p.CaminhoArquivo, p.LinhaComando }));

            var drvTable = BuildTable(
                ["Nome do dispositivo", "Fabricante", "Versão do driver", "Data do driver", "Arquivo INF", "Assinado"],
                data.InstalledDrivers.Select(d => new[] { d.NomeDispositivo, d.Fabricante, d.VersaoDriver, d.DataDriver, d.ArquivoInf, d.Assinado }));

            extraSections = $$"""
        <div id="sec-users-all" class="section-card">
            <div class="card-header">
                <h2>👥 Usuários & Perfis no Disco</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                {{usersTable}}
            </div>
        </div>

        <div id="sec-users-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Consumo por Perfil (GB)</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="userDiskChartContainer" style="position: relative; height: 260px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-programs" class="section-card">
            <div class="card-header">
                <div class="section-header-row">
                    <h2>📦 Programas Instalados</h2>
                    <div class="search-box" onclick="event.stopPropagation()">
                        <span class="search-icon">🔍</span>
                        <input type="text" class="table-filter" data-target="sec-programs" placeholder="Pesquisar programa...">
                        <span class="clear-icon" title="Limpar pesquisa">✖</span>
                    </div>
                </div>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                {{progsTable}}
            </div>
        </div>

        <div id="sec-programs-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Programas por Ano de Instalação</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="programsChartContainer" style="position: relative; height: 260px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-services" class="section-card">
            <div class="card-header">
                <div class="section-header-row">
                    <h2>⚙️ Serviços em Execução</h2>
                    <div class="search-box" onclick="event.stopPropagation()">
                        <span class="search-icon">🔍</span>
                        <input type="text" class="table-filter" data-target="sec-services" placeholder="Pesquisar serviço...">
                        <span class="clear-icon" title="Limpar pesquisa">✖</span>
                    </div>
                </div>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                {{svcsTable}}
            </div>
        </div>

        <div id="sec-services-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Serviços por Tipo de Inicialização</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="servicesChartContainer" style="position: relative; height: 220px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-processes" class="section-card">
            <div class="card-header">
                <div class="section-header-row">
                    <h2>📊 Processos Ativos</h2>
                    <div class="search-box" onclick="event.stopPropagation()">
                        <span class="search-icon">🔍</span>
                        <input type="text" class="table-filter" data-target="sec-processes" placeholder="Pesquisar processo...">
                        <span class="clear-icon" title="Limpar pesquisa">✖</span>
                    </div>
                </div>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                {{procsTable}}
            </div>
        </div>

        <div id="sec-processes-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Métricas e Uso de RAM por Processo</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div class="chart-stats" style="margin-bottom: 15px;">
                            <div class="stat-box"><span class="stat-val" id="stat-total-proc">0</span><span class="stat-lbl">Processos Totais</span></div>
                            <div class="stat-box"><span class="stat-val" id="stat-total-threads">0</span><span class="stat-lbl">Threads Totais</span></div>
                            <div class="stat-box"><span class="stat-val" id="stat-total-handles">0</span><span class="stat-lbl">Handles Totais</span></div>
                        </div>
                        <div id="processesChartContainer" style="position: relative; height: 260px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-drivers" class="section-card">
            <div class="card-header">
                <div class="section-header-row">
                    <h2>🔌 Drivers do Sistema</h2>
                    <div class="search-box" onclick="event.stopPropagation()">
                        <span class="search-icon">🔍</span>
                        <input type="text" class="table-filter" data-target="sec-drivers" placeholder="Pesquisar driver...">
                        <span class="clear-icon" title="Limpar pesquisa">✖</span>
                    </div>
                </div>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                {{drvTable}}
            </div>
        </div>

        <div id="sec-drivers-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Origem dos Drivers (Microsoft vs Terceiros)</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="driversChartContainer" style="position: relative; height: 220px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>
""";
        }

        return $$"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Relatório de {{compName}} ({{hostName}})</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #3b82f6; }
        * { scrollbar-width: thin; scrollbar-color: #cbd5e1 #f1f5f9; }
        html { scroll-behavior: smooth; }
        body { font-family: 'Inter', system-ui, -apple-system, sans-serif; margin: 0; padding: 0; display: flex; background-color: #f8fafc; color: #1e293b; transition: margin-left 0.3s cubic-bezier(0.4, 0, 0.2, 1); }
        body.sidebar-hidden .sidebar { transform: translateX(-100%); }
        body.sidebar-hidden .main-content { margin-left: 20px; width: calc(100% - 20px); }
        #btn-toggle { position: fixed; top: 15px; left: 15px; width: 36px; height: 36px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; cursor: pointer; z-index: 1000; display: flex; flex-direction: column; justify-content: center; align-items: center; gap: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); transition: all 0.2s ease; }
        #btn-toggle span { display: block; width: 18px; height: 2px; background-color: #475569; border-radius: 1px; transition: all 0.2s ease; }
        #btn-toggle:hover { background-color: #f1f5f9; border-color: #cbd5e1; }
        .sidebar { width: 250px; background: #0f172a; position: fixed; top: 0; bottom: 0; left: 0; padding: 70px 16px 20px 16px; box-sizing: border-box; overflow-y: auto; transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1); z-index: 900; box-shadow: 2px 0 8px rgba(0,0,0,0.15); }
        .sidebar-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #38bdf8; margin: 0 0 16px 12px; border-bottom: 1px solid #1e293b; padding-bottom: 8px; }
        .sidebar a { display: flex; align-items: center; gap: 10px; color: #cbd5e1; text-decoration: none; padding: 8px 12px; margin-bottom: 2px; border-radius: 6px; font-size: 13px; font-weight: 500; transition: all 0.15s ease; }
        .sidebar a.sub-link { padding-left: 24px; font-size: 12px; color: #94a3b8; }
        .sidebar a:hover { color: #38bdf8; background-color: #1e293b; transform: translateX(2px); }
        .sidebar a.active { background-color: #1e293b; color: #38bdf8; font-weight: 600; border-left: 3px solid #38bdf8; padding-left: 9px; }
        .sidebar a.sub-link.active { padding-left: 21px; }
        .main-content { margin-left: 270px; padding: 30px 40px; width: calc(100% - 270px); box-sizing: border-box; transition: margin-left 0.3s cubic-bezier(0.4, 0, 0.2, 1), width 0.3s cubic-bezier(0.4, 0, 0.2, 1); }
        .report-header { background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); color: #ffffff; padding: 24px 30px; border-radius: 12px; margin-bottom: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
        .report-header h1 { margin: 0 0 8px 0; font-size: 24px; font-weight: 700; }
        .report-header .subtitle { margin: 0; color: #94a3b8; font-size: 14px; }
        .section-card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); overflow: hidden; transition: all 0.2s ease; }
        .card-header { padding: 16px 24px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; user-select: none; background-color: #ffffff; transition: background-color 0.15s ease; }
        .card-header:hover { background-color: #f8fafc; }
        .card-header h2 { margin: 0; font-size: 16px; font-weight: 600; color: #0f172a; display: flex; align-items: center; gap: 8px; border-bottom: none; padding-bottom: 0; flex: 1; }
        .accordion-icon { font-size: 12px; color: #64748b; transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1); margin-left: 12px; }
        .section-card.collapsed .accordion-icon { transform: rotate(-90deg); }
        .card-body { padding: 0 24px 20px 24px; border-top: 1px solid #f1f5f9; }
        .section-card.collapsed .card-body { display: none; }
        .section-header-row { display: flex; justify-content: space-between; align-items: center; width: 100%; }
        .search-box { position: relative; display: flex; align-items: center; width: 260px; }
        .search-box .search-icon { position: absolute; left: 12px; color: #64748b; font-size: 0.9em; pointer-events: none; }
        .search-box .table-filter { width: 100%; padding: 6px 30px 6px 34px; font-size: 0.85em; font-family: 'Inter', system-ui, sans-serif; color: #1e293b; background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 20px; outline: none; box-sizing: border-box; transition: all 0.2s ease; }
        .search-box .table-filter:focus { background-color: #ffffff; border-color: #3b82f6; box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15); }
        .search-box .clear-icon { position: absolute; right: 12px; color: #94a3b8; font-size: 0.9em; cursor: pointer; display: none; transition: color 0.15s ease; }
        .search-box .clear-icon:hover { color: #ef4444; }
        .no-results-row { text-align: center; padding: 16px !important; color: #64748b; font-style: italic; }
        .table-container { overflow-x: auto; margin-top: 15px; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; text-align: left; }
        th { background-color: #f8fafc; color: #475569; font-weight: 600; padding: 10px 12px; border-bottom: 1px solid #e2e8f0; white-space: nowrap; user-select: none; }
        td { padding: 10px 12px; border-bottom: 1px solid #f1f5f9; color: #334155; }
        tr:nth-child(even) { background-color: #f8fafc; }
        tr:hover td { background-color: #f1f5f9; }
        th.sortable { cursor: pointer; position: relative; padding-right: 28px; transition: background-color 0.15s ease, color 0.15s ease; }
        th.sortable:hover { background-color: #e2e8f0; color: #0f172a; }
        .sort-arrow { position: absolute; right: 10px; top: 50%; transform: translateY(-50%); font-size: 0.85em; opacity: 0.4; transition: all 0.15s ease; }
        th.sortable:hover .sort-arrow { opacity: 1; }
        th.sort-asc .sort-arrow, th.sort-desc .sort-arrow { opacity: 1; color: #2563eb; }
        .charts-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(350px, 1fr)); gap: 20px; margin-top: 15px; }
        .chart-card { background: #ffffff; border-radius: 8px; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); padding: 16px; box-sizing: border-box; display: flex; flex-direction: column; }
        .chart-stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
        .stat-box { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px 8px; text-align: center; transition: all 0.2s ease; }
        .stat-val { display: block; font-size: 1.4em; font-weight: 700; color: #2563eb; margin-bottom: 2px; }
        .stat-lbl { font-size: 0.7em; color: #64748b; text-transform: uppercase; font-weight: 600; letter-spacing: 0.05em; }
        .fallback-bars-list { display: flex; flex-direction: column; gap: 10px; }
        .bar-row { display: flex; align-items: center; gap: 10px; font-size: 0.85em; cursor: pointer; padding: 2px 0; }
        .bar-label { width: 140px; font-weight: 500; color: #334155; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .bar-track { flex: 1; height: 10px; background-color: #f1f5f9; border-radius: 5px; overflow: hidden; position: relative; }
        .bar-fill { height: 100%; background: linear-gradient(90deg, #3b82f6 0%, #2563eb 100%); border-radius: 5px; transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1); }
        .bar-value { width: 75px; text-align: right; font-weight: 600; color: #1e293b; }
        .report-footer { text-align: center; padding: 20px 0; color: #94a3b8; font-size: 12px; }
        @media print {
            .sidebar, #btn-toggle, .search-box, .accordion-icon { display: none !important; }
            .main-content { margin-left: 0 !important; width: 100% !important; padding: 0 !important; }
            .section-card { page-break-inside: avoid; border: 1px solid #ccc; box-shadow: none; margin-bottom: 15px; }
            .section-card.collapsed .card-body { display: block !important; }
        }
    </style>
    <script src="https://cdn.jsdelivr.net/npm/chart.js" defer></script>
</head>
<body>
    <button id="btn-toggle" onclick="toggleSidebar()" title="Alternar Menu">
        <span></span><span></span><span></span>
    </button>

    <div class="sidebar">
        <div class="sidebar-title">Navegação Rápida</div>
        <a href="#sec-hw">💻 Hardware Principal</a>
        <a href="#sec-user">👤 Usuário Atual</a>
        <a href="#sec-os">🖥️ Sistema Operacional</a>
        <a href="#sec-bios">⚙️ BIOS</a>
        <a href="#sec-cpu">🧠 Processador</a>
        <a href="#sec-mb">🎛️ Placa-Mãe</a>
        <a href="#sec-chassis">📦 Gabinete / Chassi</a>
        <a href="#sec-ram">⚡ Memória RAM</a>
        <a href="#sec-ram-chart" class="sub-link">📊 Gráfico de RAM</a>
        <a href="#sec-disks">💾 Discos Físicos</a>
        <a href="#sec-ldisks">💽 Discos Lógicos</a>
        <a href="#sec-disk-chart" class="sub-link">📊 Gráfico de Discos</a>
        <a href="#sec-net">🌐 Rede e Conexões</a>
        <a href="#sec-monitors">🖥️ Monitores</a>
        <a href="#sec-gpu">🎮 Vídeo / GPU</a>
        <a href="#sec-audio">🔊 Dispositivos de Áudio</a>
        <a href="#sec-printers">🖨️ Impressoras</a>
        {{extraLinks}}
    </div>

    <div class="main-content">
        <div class="report-header">
            <h1>Relatório de {{compName}} ({{hostName}})</h1>
            <div class="subtitle">Dispositivo: <strong>{{compName}}</strong> | Nome da Máquina: <strong>{{hostName}}</strong> | Gerado em: {{nowStr}}</div>
        </div>

        <div id="sec-hw" class="section-card">
            <div class="card-header"><h2>💻 Hardware Principal</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{hwTable}}</div>
        </div>

        <div id="sec-user" class="section-card">
            <div class="card-header"><h2>👤 Usuário Atualmente Logado</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{userTable}}</div>
        </div>

        <div id="sec-os" class="section-card">
            <div class="card-header"><h2>🖥️ Sistema Operacional</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{osTable}}</div>
        </div>

        <div id="sec-bios" class="section-card">
            <div class="card-header"><h2>⚙️ BIOS</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{biosTable}}</div>
        </div>

        <div id="sec-cpu" class="section-card">
            <div class="card-header"><h2>🧠 Processador</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{cpuTable}}</div>
        </div>

        <div id="sec-mb" class="section-card">
            <div class="card-header"><h2>🎛️ Placa-Mãe</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{mbTable}}</div>
        </div>

        <div id="sec-chassis" class="section-card">
            <div class="card-header"><h2>📦 Gabinete / Chassi</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{chassisTable}}</div>
        </div>

        <div id="sec-ram" class="section-card">
            <div class="card-header"><h2>⚡ Memória RAM</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{ramTable}}</div>
        </div>

        <div id="sec-ram-chart" class="section-card">
            <div class="card-header"><h2>📊 Gráfico: Uso de Slots de Memória RAM</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="memoryChartContainer" style="position: relative; height: 180px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-disks" class="section-card">
            <div class="card-header"><h2>💾 Discos Físicos</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{diskTable}}</div>
        </div>

        <div id="sec-ldisks" class="section-card">
            <div class="card-header"><h2>💽 Discos Lógicos</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{ldiskTable}}</div>
        </div>

        <div id="sec-disk-chart" class="section-card">
            <div class="card-header"><h2>📊 Gráfico: Armazenamento dos Discos</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="logicalDiskChartsContainer" style="display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; align-items: center; min-height: 240px;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-net" class="section-card">
            <div class="card-header"><h2>🌐 Rede e Conexões</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{netTable}}</div>
        </div>

        <div id="sec-monitors" class="section-card">
            <div class="card-header"><h2>🖥️ Monitores Conectados</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{monTable}}</div>
        </div>

        <div id="sec-gpu" class="section-card">
            <div class="card-header"><h2>🎮 Adaptador de Vídeo / GPU</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{gpuTable}}</div>
        </div>

        <div id="sec-audio" class="section-card">
            <div class="card-header"><h2>🔊 Dispositivos de Áudio</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{soundTable}}</div>
        </div>

        <div id="sec-printers" class="section-card">
            <div class="card-header"><h2>🖨️ Impressoras Locais e Mapeadas</h2><span class="accordion-icon">▼</span></div>
            <div class="card-body">{{printLocalTable}}{{mappedPrintersHtml}}</div>
        </div>

        {{extraSections}}

        <div class="report-footer">
            Relatório gerado nativamente em C# pelo AmbientesHub.
        </div>
    </div>

    <script>
        function toggleSidebar() { document.body.classList.toggle('sidebar-hidden'); }
        document.addEventListener('DOMContentLoaded', function() {
            // 0) Accordions
            document.querySelectorAll('.section-card .card-header').forEach(function(header) {
                header.addEventListener('click', function() {
                    var card = header.closest('.section-card');
                    if (card) card.classList.toggle('collapsed');
                });
            });

            // 1) Ordenação de tabelas
            document.querySelectorAll('.main-content table').forEach(function(table) {
                var headers = table.querySelectorAll('th');
                headers.forEach(function(header, colIndex) {
                    header.classList.add('sortable');
                    var arrow = document.createElement('span');
                    arrow.className = 'sort-arrow';
                    arrow.innerHTML = '↕';
                    header.appendChild(arrow);
                    header.addEventListener('click', function(e) {
                        e.stopPropagation();
                        var currentOrder = header.getAttribute('data-sort') || 'none';
                        var nextOrder = currentOrder === 'asc' ? 'desc' : 'asc';
                        headers.forEach(function(h) {
                            h.removeAttribute('data-sort');
                            h.classList.remove('sort-asc', 'sort-desc');
                            var a = h.querySelector('.sort-arrow');
                            if (a) a.innerHTML = '↕';
                        });
                        header.setAttribute('data-sort', nextOrder);
                        header.classList.add(nextOrder === 'asc' ? 'sort-asc' : 'sort-desc');
                        arrow.innerHTML = nextOrder === 'asc' ? '▲' : '▼';
                        var rows = Array.from(table.querySelectorAll('tr')).filter(function(r) { return !r.querySelector('th') && !r.classList.contains('no-results-row'); });
                        if (rows.length <= 1) return;
                        var isAsc = nextOrder === 'asc';
                        rows.sort(function(rowA, rowB) {
                            var cellA = rowA.cells[colIndex] ? rowA.cells[colIndex].textContent.trim() : '';
                            var cellB = rowB.cells[colIndex] ? rowB.cells[colIndex].textContent.trim() : '';
                            if (cellA === cellB) return 0;
                            var numA = parseFloat(cellA.replace(/[^\d.,-]/g, '').replace(',', '.'));
                            var numB = parseFloat(cellB.replace(/[^\d.,-]/g, '').replace(',', '.'));
                            if (cellA.indexOf('/') === -1 && !isNaN(numA) && !isNaN(numB)) return isAsc ? numA - numB : numB - numA;
                            return isAsc ? cellA.localeCompare(cellB, undefined, { numeric: true, sensitivity: 'base' }) : cellB.localeCompare(cellA, undefined, { numeric: true, sensitivity: 'base' });
                        });
                        var parent = rows[0].parentNode;
                        rows.forEach(function(r) { parent.appendChild(r); });
                    });
                });
            });

            // 2) ScrollSpy Inteligente com destaque azul (.active) na sidebar
            var sections = document.querySelectorAll('.main-content .section-card[id]');
            var navLinks = document.querySelectorAll('.sidebar a[href^="#"]');

            var updateActiveSection = function() {
                var scrollPosition = window.scrollY || document.documentElement.scrollTop;
                var windowHeight = window.innerHeight;
                var docHeight = document.documentElement.scrollHeight;
                var currentSectionId = '';

                if (scrollPosition + windowHeight >= docHeight - 15) {
                    if (sections.length > 0) {
                        currentSectionId = sections[sections.length - 1].getAttribute('id');
                    }
                } else {
                    for (var i = 0; i < sections.length; i++) {
                        var section = sections[i];
                        var sectionTop = section.offsetTop - 120;
                        if (scrollPosition >= sectionTop) {
                            currentSectionId = section.getAttribute('id');
                        } else {
                            break;
                        }
                    }
                }

                if (currentSectionId) {
                    navLinks.forEach(function(link) {
                        if (link.getAttribute('href') === '#' + currentSectionId) {
                            link.classList.add('active');
                            var sidebar = document.querySelector('.sidebar');
                            if (sidebar && (link.offsetTop < sidebar.scrollTop || link.offsetTop > sidebar.scrollTop + sidebar.clientHeight)) {
                                link.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                            }
                        } else {
                            link.classList.remove('active');
                        }
                    });
                }
            };

            window.addEventListener('scroll', updateActiveSection, { passive: true });
            updateActiveSection();

            // 3) Filtro de busca nas tabelas
            var filterInputs = document.querySelectorAll('.table-filter');
            filterInputs.forEach(function(input) {
                var clearBtn = input.nextElementSibling;
                var targetId = input.getAttribute('data-target');
                var cardDiv = document.getElementById(targetId);
                var table = cardDiv ? cardDiv.querySelector('table') : null;
                if (!table) return;
                var tbody = table.querySelector('tbody') || table;
                var rows = tbody.querySelectorAll('tr');
                var filterTable = function() {
                    var q = input.value.trim().toLowerCase();
                    if (clearBtn) clearBtn.style.display = q.length > 0 ? 'inline' : 'none';
                    if (q.length > 0 && cardDiv.classList.contains('collapsed')) cardDiv.classList.remove('collapsed');
                    rows.forEach(function(r) {
                        if (r.querySelector('th')) return;
                        r.style.display = (q === '' || r.textContent.toLowerCase().indexOf(q) !== -1) ? '' : 'none';
                    });
                    if (targetId === 'sec-processes' && typeof updateProcessesChart === 'function') {
                        updateProcessesChart();
                    }
                };
                input.addEventListener('input', filterTable);
                if (clearBtn) clearBtn.addEventListener('click', function(e) { e.stopPropagation(); input.value = ''; filterTable(); });
            });

            // 4) Gráficos Completos (RAM, Discos, Usuários, Programas, Serviços, Processos, Drivers)
            var renderFallbackBars = function(containerId, items, unit, color) {
                var el = document.getElementById(containerId);
                if (!el || items.length === 0) return;
                var max = Math.max.apply(null, items.map(function(i) { return i.val; })) || 1;
                var html = '<div class="fallback-bars-list">';
                items.forEach(function(i) {
                    var pct = Math.min(100, (i.val / max) * 100);
                    html += '<div class="bar-row"><div class="bar-label" title="' + i.label + '">' + i.label + '</div><div class="bar-track"><div class="bar-fill" style="width:' + pct + '%;background:' + (color || '#2563eb') + ';"></div></div><div class="bar-value">' + i.val + ' ' + unit + '</div></div>';
                });
                html += '</div>';
                el.innerHTML = html;
            };

            // a) Gráfico de RAM
            var secRam = document.getElementById('sec-ram');
            var tableRam = secRam ? secRam.querySelector('table') : null;
            if (tableRam) {
                var memoryData = [];
                var modIdx = 1;
                tableRam.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 2) {
                        var capText = cells[1].innerText.trim();
                        var capVal = parseFloat(capText.replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
                        memoryData.push({ label: 'Módulo ' + modIdx++, capacity: capVal });
                    }
                });

                var memContainer = document.getElementById('memoryChartContainer');
                if (memContainer && memoryData.length > 0) {
                    if (typeof Chart !== 'undefined') {
                        memContainer.innerHTML = '<canvas id="memoryChart" style="width: 100%; height: 180px;"></canvas>';
                        var ctxM = document.getElementById('memoryChart').getContext('2d');
                        new Chart(ctxM, {
                            type: 'bar',
                            data: {
                                labels: memoryData.map(function(m) { return m.label; }),
                                datasets: [{
                                    label: 'Capacidade (GB)',
                                    data: memoryData.map(function(m) { return m.capacity; }),
                                    backgroundColor: 'rgba(59, 130, 246, 0.2)',
                                    borderColor: '#3b82f6',
                                    borderWidth: 2,
                                    borderRadius: 4
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: {
                                    x: { grid: { display: false } },
                                    y: { beginAtZero: true, grace: '15%', ticks: { callback: function(v) { return v + ' GB'; } } }
                                }
                            }
                        });
                    } else {
                        renderFallbackBars('memoryChartContainer', memoryData.map(function(m) { return { label: m.label, val: m.capacity }; }), 'GB', '#3b82f6');
                    }
                }
            }

            // b) Gráfico de Discos Lógicos (Pie por disco)
            var secLdisks = document.getElementById('sec-ldisks');
            var tableDisks = secLdisks ? secLdisks.querySelector('table') : null;
            if (tableDisks) {
                var diskData = [];
                tableDisks.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 4) {
                        var devId = cells[0].innerText.trim();
                        var freeVal = parseFloat(cells[1].innerText.replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
                        var usedVal = parseFloat(cells[2].innerText.replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
                        var totalVal = parseFloat(cells[3].innerText.replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
                        var volName = cells[4] ? cells[4].innerText.trim() : '';
                        if (totalVal > 0) {
                            diskData.push({ devId: devId, volName: volName, free: freeVal, used: usedVal, total: totalVal });
                        }
                    }
                });

                var diskContainer = document.getElementById('logicalDiskChartsContainer');
                if (diskContainer && diskData.length > 0) {
                    diskContainer.innerHTML = '';
                    diskData.forEach(function(disk, idx) {
                        var cardId = 'diskPieChart_' + idx;
                        var wrapper = document.createElement('div');
                        wrapper.style.cssText = 'flex: 1 1 280px; max-width: 380px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 15px; text-align: center;';
                        var title = document.createElement('h4');
                        title.style.cssText = 'margin: 0 0 12px 0; color: #1e293b; font-size: 15px; font-weight: 600;';
                        title.innerText = 'Disco ' + disk.devId + (disk.volName && disk.volName !== 'Windows' ? ' (' + disk.volName + ')' : '');
                        wrapper.appendChild(title);

                        if (typeof Chart !== 'undefined') {
                            var canvasWrapper = document.createElement('div');
                            canvasWrapper.style.cssText = 'position: relative; width: 100%; height: 200px;';
                            canvasWrapper.innerHTML = '<canvas id="' + cardId + '"></canvas>';
                            wrapper.appendChild(canvasWrapper);
                            diskContainer.appendChild(wrapper);

                            var ctxD = document.getElementById(cardId).getContext('2d');
                            new Chart(ctxD, {
                                type: 'pie',
                                data: {
                                    labels: ['Usado (GB)', 'Livre (GB)'],
                                    datasets: [{
                                        data: [disk.used, disk.free],
                                        backgroundColor: ['#0284c7', '#38bdf8'],
                                        borderWidth: 2,
                                        borderColor: '#ffffff'
                                    }]
                                },
                                options: {
                                    responsive: true,
                                    maintainAspectRatio: false,
                                    plugins: { legend: { position: 'bottom' } }
                                }
                            });
                        } else {
                            var pctUsed = (disk.used / disk.total) * 100;
                            var pctFree = 100 - pctUsed;
                            var fallbackHtml = '<div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 180px;">' +
                                '<div style="width: 80px; height: 80px; border-radius: 50%; background: conic-gradient(#0284c7 0% ' + pctUsed + '%, #38bdf8 ' + pctUsed + '% 100%); margin-bottom: 15px;"></div>' +
                                '<div style="display: flex; justify-content: space-around; width: 100%; font-size: 12px;">' +
                                    '<div>Usado: ' + disk.used.toFixed(1) + ' GB (' + pctUsed.toFixed(0) + '%)</div>' +
                                    '<div>Livre: ' + disk.free.toFixed(1) + ' GB (' + pctFree.toFixed(0) + '%)</div>' +
                                '</div></div>';
                            wrapper.innerHTML += fallbackHtml;
                            diskContainer.appendChild(wrapper);
                        }
                    });
                }
            }

            // c) Gráfico de Espaço em Disco por Perfil
            var secUsers = document.getElementById('sec-users-all');
            var tableUsers = secUsers ? secUsers.querySelector('table') : null;
            if (tableUsers) {
                var userData = [];
                tableUsers.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 10) {
                        var name = cells[0].innerText.trim();
                        var user = cells[1].innerText.trim();
                        var sizeVal = parseFloat(cells[9].innerText.replace(/[^\d.,-]/g, '').replace(',', '.')) || 0;
                        if (sizeVal > 0) {
                            userData.push({ name: name || user, user: user, size: sizeVal });
                        }
                    }
                });

                var userContainer = document.getElementById('userDiskChartContainer');
                if (userContainer && userData.length > 0) {
                    userData.sort(function(a, b) { return b.size - a.size; });
                    if (typeof Chart !== 'undefined') {
                        userContainer.innerHTML = '<canvas id="userDiskChart" style="width: 100%; height: 260px;"></canvas>';
                        var ctxU = document.getElementById('userDiskChart').getContext('2d');
                        new Chart(ctxU, {
                            type: 'bar',
                            data: {
                                labels: userData.map(function(u) { return u.name; }),
                                datasets: [{
                                    label: 'Espaço em Disco (GB)',
                                    data: userData.map(function(u) { return u.size; }),
                                    backgroundColor: 'rgba(14, 116, 144, 0.2)',
                                    borderColor: '#0e7490',
                                    borderWidth: 2,
                                    borderRadius: 4
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: {
                                    x: { grid: { display: false } },
                                    y: { beginAtZero: true, grace: '15%', ticks: { callback: function(v) { return v + ' GB'; } } }
                                }
                            }
                        });
                    } else {
                        renderFallbackBars('userDiskChartContainer', userData.map(function(u) { return { label: u.name, val: u.size }; }), 'GB', '#0e7490');
                    }
                } else if (userContainer) {
                    userContainer.innerHTML = '<p style="text-align: center; color: #64748b; font-style: italic;">Nenhum consumo de disco significativo nos perfis listados.</p>';
                }
            }

            // d) Gráfico de Programas por Ano
            var secProg = document.getElementById('sec-programs');
            var tableProg = secProg ? secProg.querySelector('table') : null;
            if (tableProg) {
                var yearCounts = {};
                tableProg.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 4) {
                        var dateText = cells[3].innerText.trim();
                        var year = dateText.substring(dateText.length - 4);
                        if (/^\d{4}$/.test(year)) {
                            yearCounts[year] = (yearCounts[year] || 0) + 1;
                        }
                    }
                });

                var years = Object.keys(yearCounts).sort();
                var progContainer = document.getElementById('programsChartContainer');
                if (progContainer && years.length > 0) {
                    if (typeof Chart !== 'undefined') {
                        progContainer.innerHTML = '<canvas id="programsChart" style="width: 100%; height: 260px;"></canvas>';
                        var ctxP = document.getElementById('programsChart').getContext('2d');
                        new Chart(ctxP, {
                            type: 'bar',
                            data: {
                                labels: years,
                                datasets: [{
                                    label: 'Programas Instalados',
                                    data: years.map(function(y) { return yearCounts[y]; }),
                                    backgroundColor: 'rgba(16, 185, 129, 0.2)',
                                    borderColor: '#10b8a5',
                                    borderWidth: 2,
                                    borderRadius: 4
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: {
                                    x: { grid: { display: false } },
                                    y: { beginAtZero: true, grace: '15%', ticks: { stepSize: 1 } }
                                }
                            }
                        });
                    } else {
                        renderFallbackBars('programsChartContainer', years.map(function(y) { return { label: 'Ano ' + y, val: yearCounts[y] }; }), 'progs', '#10b8a5');
                    }
                }
            }

            // e) Gráfico de Serviços por Tipo de Inicialização
            var secServ = document.getElementById('sec-services');
            var tableServ = secServ ? secServ.querySelector('table') : null;
            if (tableServ) {
                var modeCounts = {};
                tableServ.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 4) {
                        var mode = cells[3].innerText.trim();
                        if (mode) modeCounts[mode] = (modeCounts[mode] || 0) + 1;
                    }
                });

                var modes = Object.keys(modeCounts);
                var servContainer = document.getElementById('servicesChartContainer');
                if (servContainer && modes.length > 0) {
                    if (typeof Chart !== 'undefined') {
                        servContainer.innerHTML = '<canvas id="servicesChart" style="width: 100%; height: 220px;"></canvas>';
                        var ctxS = document.getElementById('servicesChart').getContext('2d');
                        new Chart(ctxS, {
                            type: 'pie',
                            data: {
                                labels: modes,
                                datasets: [{
                                    data: modes.map(function(m) { return modeCounts[m]; }),
                                    backgroundColor: ['#f59e0b', '#10b981', '#3b82f6', '#ef4444', '#8b5cf6'],
                                    borderWidth: 2,
                                    borderColor: '#ffffff'
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { position: 'bottom' } }
                            }
                        });
                    } else {
                        renderFallbackBars('servicesChartContainer', modes.map(function(m) { return { label: m, val: modeCounts[m] }; }), 'serviços', '#3b82f6');
                    }
                }
            }

            // f) Gráfico de Processos (Top Consumidores de RAM)
            window.updateProcessesChart = function() {
                var secProc = document.getElementById('sec-processes');
                var tableProc = secProc ? secProc.querySelector('table') : null;
                if (!tableProc) return;

                var procData = [];
                tableProc.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th') || row.style.display === 'none' || row.classList.contains('no-results-row')) return;
                    var cells = row.cells;
                    if (cells.length >= 6) {
                        var name = cells[1].innerText.trim();
                        var user = cells[2].innerText.trim();
                        var memoryVal = parseFloat(cells[3].innerText.replace(/[^\d.,-]/g, '').replace(',', '.')) || 0;
                        var threads = parseInt(cells[4].innerText.trim(), 10) || 0;
                        var handles = parseInt(cells[5].innerText.trim(), 10) || 0;
                        procData.push({ name: name, user: user, memory: memoryVal, threads: threads, handles: handles });
                    }
                });

                var totalProcEl = document.getElementById('stat-total-proc');
                var totalThreadsEl = document.getElementById('stat-total-threads');
                var totalHandlesEl = document.getElementById('stat-total-handles');
                var chartContainer = document.getElementById('processesChartContainer');

                if (procData.length > 0) {
                    if (totalProcEl) totalProcEl.innerText = procData.length;
                    if (totalThreadsEl) totalThreadsEl.innerText = procData.reduce(function(acc, p) { return acc + p.threads; }, 0).toLocaleString('pt-BR');
                    if (totalHandlesEl) totalHandlesEl.innerText = procData.reduce(function(acc, p) { return acc + p.handles; }, 0).toLocaleString('pt-BR');

                    var topMemory = procData.slice().sort(function(a, b) { return b.memory - a.memory; }).slice(0, 10);
                    if (chartContainer) {
                        if (typeof Chart !== 'undefined') {
                            if (window.procMemoryChartObj) window.procMemoryChartObj.destroy();
                            chartContainer.innerHTML = '<canvas id="procMemoryChart" style="width: 100%; height: 260px;"></canvas>';
                            var ctxProc = document.getElementById('procMemoryChart').getContext('2d');
                            window.procMemoryChartObj = new Chart(ctxProc, {
                                type: 'bar',
                                data: {
                                    labels: topMemory.map(function(p) { return p.name; }),
                                    datasets: [{
                                        label: 'Uso de Memória (MB)',
                                        data: topMemory.map(function(p) { return p.memory; }),
                                        backgroundColor: 'rgba(37, 99, 235, 0.15)',
                                        borderColor: '#2563eb',
                                        borderWidth: 2,
                                        borderRadius: 4
                                    }]
                                },
                                options: {
                                    responsive: true,
                                    maintainAspectRatio: false,
                                    plugins: { legend: { display: false } },
                                    scales: {
                                        x: { grid: { display: false } },
                                        y: { beginAtZero: true, grace: '15%', ticks: { callback: function(v) { return v + ' MB'; } } }
                                    }
                                }
                            });
                        } else {
                            renderFallbackBars('processesChartContainer', topMemory.map(function(p) { return { label: p.name, val: p.memory }; }), 'MB', '#2563eb');
                        }
                    }
                }
            };
            updateProcessesChart();

            // g) Gráfico de Drivers (Microsoft vs Terceiros)
            var secDrv = document.getElementById('sec-drivers');
            var tableDrv = secDrv ? secDrv.querySelector('table') : null;
            if (tableDrv) {
                var msCount = 0;
                var otherCount = 0;
                tableDrv.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 2) {
                        var mfr = cells[1].innerText.trim().toLowerCase();
                        if (mfr.indexOf('microsoft') !== -1) msCount++;
                        else otherCount++;
                    }
                });

                var drvContainer = document.getElementById('driversChartContainer');
                if (drvContainer && (msCount > 0 || otherCount > 0)) {
                    if (typeof Chart !== 'undefined') {
                        drvContainer.innerHTML = '<canvas id="driversChart" style="width: 100%; height: 220px;"></canvas>';
                        var ctxDr = document.getElementById('driversChart').getContext('2d');
                        new Chart(ctxDr, {
                            type: 'pie',
                            data: {
                                labels: ['Microsoft', 'Terceiros'],
                                datasets: [{
                                    data: [msCount, otherCount],
                                    backgroundColor: ['#0284c7', '#f43f5e'],
                                    borderWidth: 2,
                                    borderColor: '#ffffff'
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { position: 'bottom' } }
                            }
                        });
                    } else {
                        renderFallbackBars('driversChartContainer', [
                            { label: 'Microsoft', val: msCount },
                            { label: 'Terceiros', val: otherCount }
                        ], 'drivers', '#0284c7');
                    }
                }
            }
        });
    </script>
</body>
</html>
""";
    }

    public static bool GeneratePdfReport(string htmlFilePath, string pdfFilePath)
    {
        string[] browserCandidates =
        [
            @"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            @"C:\Program Files\Google\Chrome\Application\chrome.exe",
            @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ];

        string? browserExe = browserCandidates.FirstOrDefault(File.Exists);
        if (string.IsNullOrEmpty(browserExe))
            return false;

        try
        {
            if (File.Exists(pdfFilePath)) File.Delete(pdfFilePath);

            using var proc = Process.Start(new ProcessStartInfo
            {
                FileName = browserExe,
                Arguments = $"--headless=new --disable-gpu --no-first-run --no-default-browser-check --disable-extensions --disable-logging --log-level=3 --print-to-pdf=\"{pdfFilePath}\" \"{htmlFilePath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true
            });

            if (proc == null) return false;
            proc.WaitForExit(35000);
            return File.Exists(pdfFilePath);
        }
        catch
        {
            return false;
        }
    }

    public static void GenerateExcelReport(DeviceReportData data, string excelFilePath)
    {
        var allRows = new List<(bool IsHeader, string[] Cells)>();

        void AddSection(string[] headers, IEnumerable<string[]> rows)
        {
            allRows.Add((true, headers));
            var list = rows.ToList();
            if (list.Count == 0)
            {
                allRows.Add((false, headers.Select(_ => "-").ToArray()));
            }
            else
            {
                foreach (var r in list)
                    allRows.Add((false, r));
            }
            allRows.Add((false, [])); // Linha em branco separadora (igual ao Analisador.ps1)
        }

        // 1) HardwareInfo
        AddSection(
            ["Computador", "Modelo", "Fabricante", "Memória Física (GB)", "Processadores Lógicos", "Processadores Físicos", "Serial Number", "Portas de Vídeo Ativas", "Entradas de vídeo"],
            [[data.HardwareInfo.Computador, data.HardwareInfo.Modelo, data.HardwareInfo.Fabricante, data.HardwareInfo.MemoriaFisicaGb, data.HardwareInfo.ProcessadoresLogicos.ToString(), data.HardwareInfo.ProcessadoresFisicos.ToString(), data.HardwareInfo.SerialNumber, data.HardwareInfo.PortasVideoAtivas, data.HardwareInfo.EntradasVideo]]);

        // 2) Usuário Atual
        AddSection(
            ["Usuário Logado agora", "DisplayName AD", "Domínio", "Nome do computador"],
            [[data.CurrentUser, data.CurrentUserDisplayName, data.CurrentUserDomain, data.HardwareInfo.Computador]]);

        // 3) Sistema Operacional
        AddSection(
            ["SO", "Versão WMI (Version)", "Versão do Windows (HnH)", "Número da Build", "Nome da Máquina", "Último Boot"],
            [[data.OS.SO, data.OS.VersaoWmi, data.OS.VersaoWindowsHnH, data.OS.NumeroBuild, data.OS.NomeMaquina, data.OS.UltimoBoot]]);

        // 4) BIOS
        AddSection(
            ["Fabricante", "Versão da BIOS", "Número de Serial", "Data de lançamento"],
            [[data.BIOS.Fabricante, data.BIOS.VersaoBios, data.BIOS.NumeroSerial, data.BIOS.DataLancamento]]);

        // 5) Processador
        AddSection(
            ["Nome", "Número de Cores", "Número de processadores lógicos", "Clock máximo", "Voltagem", "L1 Cache", "L2 Cache", "L3 Cache"],
            [[data.Processor.Nome, data.Processor.NumeroCores.ToString(), data.Processor.ProcessadoresLogicos.ToString(), data.Processor.ClockMaximo, data.Processor.Voltagem, data.Processor.L1Cache, data.Processor.L2Cache, data.Processor.L3Cache]]);

        // 6) Placa-Mãe
        AddSection(
            ["Fabricante", "Produto", "Versão"],
            [[data.Motherboard.Fabricante, data.Motherboard.Produto, data.Motherboard.Versao]]);

        // 7) Chassi
        AddSection(
            ["Tipo de Chassi", "Fabricante"],
            [[data.Chassis.TipoChassi, data.Chassis.Fabricante]]);

        // 8) Memória
        AddSection(
            ["Fabricante", "Capacidade (GB)", "Velocidade (MHz)", "Part Number", "Número de Série"],
            data.MemoryModules.Select(m => new[] { m.Fabricante, m.CapacidadeGb, m.VelocidadeMhz, m.PartNumber, m.NumeroSerie }));

        // 9) Discos Físicos
        AddSection(
            ["Fabricante", "Modelo", "Tipo de interface", "Tamanho em GB", "Tipo de mídia", "Número de Série"],
            data.DiskDrives.Select(d => new[] { d.Fabricante, d.Modelo, d.TipoInterface, d.TamanhoGb, d.TipoMidia, d.NumeroSerie }));

        // 10) Discos Lógicos
        AddSection(
            ["DeviceID", "Espaço Livre", "Espaço Usado", "Espaço Total", "Nome do Volume", "Sistema de Arquivos"],
            data.LogicalDisks.Select(d => new[] { d.DeviceId, d.EspacoLivre, d.EspacoUsado, d.EspacoTotal, d.NomeVolume, d.SistemaArquivos }));

        // 11) Rede
        AddSection(
            ["Nome da placa", "MAC Address", "Endereços IPs", "Gateway"],
            data.NetworkConfigs.Select(n => new[] { n.NomePlaca, n.MacAddress, n.EnderecosIp, n.Gateway }));

        // 12) Controlador de Vídeo
        AddSection(
            ["Nome", "Processador de Vídeo", "Adaptador de compatibilidade", "Capacidade (GB)", "DeviceID", "Resolução Horizontal", "Resolução Vertical", "Número de Cores", "Versão do Driver", "Data do Driver Instalado"],
            data.VideoControllers.Select(v => new[] { v.Nome, v.ProcessadorVideo, v.AdaptadorCompatibilidade, v.CapacidadeGb, v.DeviceId, v.ResolucaoHorizontal, v.ResolucaoVertical, v.NumeroCores, v.VersaoDriver, v.DataDriverInstalado }));

        // 13) Som
        AddSection(
            ["Nome", "Fabricante", "Status", "Informação do Status"],
            data.SoundDevices.Select(s => new[] { s.Nome, s.Fabricante, s.Status, s.InformacaoStatus }));

        // 14) Impressoras Locais
        AddSection(
            ["Nome", "Estado da impressão", "Status da impressão", "Nome de compartilhamento", "Porta"],
            data.PrintersLocal.Select(p => new[] { p.Nome, p.EstadoImpressao, p.StatusImpressao, p.NomeCompartilhamento, p.Porta }));

        // 15) Impressoras Mapeadas
        AddSection(
            ["ComputerName", "UserSID", "UserName", "Type", "Server", "PrinterName"],
            data.PrintersMapped.Select(p => new[] { p.ComputerName, p.UserSid, p.UserName, p.Type, p.Server, p.PrinterName }));

        // 16) Monitores
        AddSection(
            ["Nome", "Polegada", "Resolução", "Tamanho", "Ativo", "Serial Number", "Nome amigável", "Fabricante", "Entradas de vídeo", "Porta Ativa"],
            data.Monitors.Select(m => new[] { m.Nome, m.Polegada, m.Resolucao, m.Tamanho, m.Ativo, m.SerialNumber, m.NomeAmigavel, m.Fabricante, m.EntradasVideo, m.PortaAtiva }));

        if (!data.SkipMajorData)
        {
            // 17) Usuários na Máquina
            AddSection(
                ["Nome", "Usuário", "Descrição", "Escritório", "Caminho do Perfil", "SID", "Último Uso", "Loaded", "Special", "Espaço em Disco (GB)"],
                data.UserProfiles.Select(u => new[] { u.Nome, u.Usuario, u.Descricao, u.Escritorio, u.CaminhoPerfil, u.Sid, u.UltimoUso, u.StatusSessao, u.Special ? "Sim" : "Não", u.EspacoDiscoGb.ToString("0.00") }));

            // 18) Programas, Serviços, Processos e Drivers
            AddSection(
                ["Nome do programa", "Versão", "Fabricante", "Data da instalação"],
                data.InstalledPrograms.Select(p => new[] { p.NomePrograma, p.Versao, p.Fabricante, p.DataInstalacao }));

            AddSection(
                ["Nome técnico", "Nome de exibição", "Status", "Tipo de inicialização", "Conta de logon", "Caminho do executável"],
                data.RunningServices.Select(s => new[] { s.NomeTecnico, s.NomeExibicao, s.Status, s.TipoInicializacao, s.ContaLogon, s.CaminhoExecutavel }));

            AddSection(
                ["PID", "Nome do executável", "Usuário", "Uso de Memória (MB)", "Threads", "Handles", "Caminho do arquivo", "Linha de comando"],
                data.ActiveProcesses.Select(p => new[] { p.Pid.ToString(), p.NomeExecutavel, p.Usuario, p.UsoMemoriaMb.ToString("0.00"), p.Threads.ToString(), p.Handles.ToString(), p.CaminhoArquivo, p.LinhaComando }));

            AddSection(
                ["Nome do dispositivo", "Fabricante", "Versão do driver", "Data do driver", "Arquivo INF", "Assinado"],
                data.InstalledDrivers.Select(d => new[] { d.NomeDispositivo, d.Fabricante, d.VersaoDriver, d.DataDriver, d.ArquivoInf, d.Assinado }));
        }

        WriteOpenXmlWorkbook(excelFilePath, "Relatorio", allRows);
    }

    private static void WriteOpenXmlWorkbook(string filePath, string sheetName, List<(bool IsHeader, string[] Cells)> rows)
    {
        if (File.Exists(filePath)) File.Delete(filePath);

        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);

        WriteZipText(zip, "[Content_Types].xml", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
""");

        WriteZipText(zip, "_rels/.rels", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
""");

        WriteZipText(zip, "xl/workbook.xml", $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="{SecurityElementEscape(sheetName)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
""");

        WriteZipText(zip, "xl/_rels/workbook.xml.rels", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
""");

        WriteZipText(zip, "xl/styles.xml", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>
""");

        var sb = new StringBuilder();
        sb.AppendLine("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>""");
        sb.AppendLine("""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">""");
        sb.AppendLine("""  <cols><col min="1" max="12" width="26" customWidth="1"/></cols>""");
        sb.AppendLine("""  <sheetData>""");

        for (int rIdx = 0; rIdx < rows.Count; rIdx++)
        {
            int rowNum = rIdx + 1;
            var (isHeader, cells) = rows[rIdx];
            if (cells.Length == 0)
            {
                sb.AppendLine($"    <row r=\"{rowNum}\"/>");
                continue;
            }

            sb.Append($"    <row r=\"{rowNum}\">");
            for (int cIdx = 0; cIdx < cells.Length; cIdx++)
            {
                string colRef = $"{GetExcelColumnName(cIdx + 1)}{rowNum}";
                string styleAttr = isHeader ? " s=\"1\"" : "";
                string val = SanitizeXmlString(cells[cIdx] ?? string.Empty);
                sb.Append($"<c r=\"{colRef}\"{styleAttr} t=\"inlineStr\"><is><t>{SecurityElementEscape(val)}</t></is></c>");
            }
            sb.AppendLine("</row>");
        }

        sb.AppendLine("  </sheetData>");
        sb.AppendLine("</worksheet>");

        WriteZipText(zip, "xl/worksheets/sheet1.xml", sb.ToString());
    }

    private static void WriteZipText(ZipArchive zip, string entryName, string content)
    {
        var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static string GetExcelColumnName(int columnNumber)
    {
        string columnName = string.Empty;
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnName;
    }

    private static string SanitizeXmlString(string input)
    {
        if (string.IsNullOrEmpty(input)) return string.Empty;
        var sb = new StringBuilder(input.Length);
        foreach (char c in input)
        {
            if (c == 0x9 || c == 0xA || c == 0xD || (c >= 0x20 && c <= 0xD7FF) || (c >= 0xE000 && c <= 0xFFFD))
                sb.Append(c);
        }
        return sb.ToString();
    }

    private static string SecurityElementEscape(string input) => WebUtility.HtmlEncode(input ?? string.Empty);

    private static string BuildTable(string[] headers, IEnumerable<string[]> rows)
    {
        var sb = new StringBuilder();
        sb.Append("<div class='table-container'><table><thead><tr>");
        foreach (var h in headers)
            sb.Append($"<th>{E(h)}</th>");
        sb.Append("</tr></thead><tbody>");

        var list = rows.ToList();
        if (list.Count == 0)
        {
            sb.Append($"<tr><td colspan='{headers.Length}' class='no-results-row'>Nenhum registro encontrado</td></tr>");
        }
        else
        {
            foreach (var r in list)
            {
                sb.Append("<tr>");
                foreach (var c in r)
                    sb.Append($"<td>{E(c)}</td>");
                sb.Append("</tr>");
            }
        }

        sb.Append("</tbody></table></div>");
        return sb.ToString();
    }

    private static string E(string? s) => WebUtility.HtmlEncode(s ?? string.Empty);
}

