using AmbientesHub.Models;
using Xunit;

namespace AmbientesHub.Tests.Unit;

public class DocEntryItemTests
{
    [Fact]
    public void IsSelected_RaisesPropertyChanged_AndUpdatesCardColors()
    {
        var item = new DocEntryItem
        {
            Title = "Guia de Operação",
            RelativePath = "guias/guia_operacao.md"
        };

        var changedProps = new List<string>();
        item.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName != null)
                changedProps.Add(e.PropertyName);
        };

        Assert.False(item.IsSelected);
        Assert.Equal("#1e293b", item.BorderHex);
        Assert.Equal("#0f172a", item.CardBgHex);

        item.IsSelected = true;

        Assert.True(item.IsSelected);
        Assert.Equal("#38bdf8", item.BorderHex);
        Assert.Equal("#172554", item.CardBgHex);
        Assert.Contains(nameof(DocEntryItem.IsSelected), changedProps);
        Assert.Contains(nameof(DocEntryItem.BorderHex), changedProps);
        Assert.Contains(nameof(DocEntryItem.CardBgHex), changedProps);
    }

    [Fact]
    public void LastUpdatedDisplay_And_HasSummary_ReturnExpectedValues()
    {
        var emptyItem = new DocEntryItem();
        Assert.Equal(string.Empty, emptyItem.LastUpdatedDisplay);
        Assert.False(emptyItem.HasSummary);

        var dt = new DateTime(2026, 11, 25, 15, 30, 0);
        var populated = new DocEntryItem
        {
            Summary = "Resumo técnico do módulo.",
            LastModifiedLocal = dt
        };

        Assert.Equal("25/11/2026 15:30", populated.LastUpdatedDisplay);
        Assert.True(populated.HasSummary);
    }
}
