using AmbientesHub.Models;
using Xunit;

namespace AmbientesHub.Tests.Services
{
    public class BancadaCronServiceTests
    {
        [Fact]
        public void CronScheduleItem_IntervalAndStatusDisplay_WorkCorrectly()
        {
            var item = new CronScheduleItem
            {
                TaskId = "sync_oxe",
                Nome = "Central OXE",
                TipoAgendamento = "intervalo",
                IntervaloValor = 7,
                IntervaloUnidade = "dias",
                ApenasDiasUteis = false,
                UltimoStatus = "sucesso",
                UltimaExecucao = "2026-11-16 14:30:00"
            };

            Assert.Equal("A cada 7 dias", item.AgendamentoDisplay);
            Assert.Equal("SUCESSO", item.StatusDisplay);
            Assert.Equal("16/11/2026 14:30:00", item.UltimaExecucaoBr);
            Assert.True(item.CanTrigger);
        }

        [Fact]
        public void CronScheduleItem_FixedTimeAndExecutingState_DisableTrigger()
        {
            var item = new CronScheduleItem
            {
                TaskId = "whatsapp_d1",
                TipoAgendamento = "horario_fixo",
                HorarioFixo = "12:00",
                ApenasDiasUteis = true,
                UltimoStatus = "executando"
            };

            Assert.Equal("Diário às 12:00 (dias úteis)", item.AgendamentoDisplay);
            Assert.True(item.IsExecuting);
            Assert.False(item.CanTrigger);
        }
    }
}
