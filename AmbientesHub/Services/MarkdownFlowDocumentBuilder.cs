using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using AmbientesHub.Models;
using Wpf.Ui.Controls;
using WControls = System.Windows.Controls;
using WDocs = System.Windows.Documents;
using WMedia = System.Windows.Media;

namespace AmbientesHub.Services;

/// <summary>
/// Constrói um FlowDocument nativo WPF estilizado no padrão Fluent Dark do AmbientesHub
/// a partir da árvore de blocos gerada pelo DocumentationService.ParseMarkdown.
/// </summary>
public static class MarkdownFlowDocumentBuilder
{
    public static WDocs.FlowDocument Build(
        IReadOnlyList<MarkdownBlock> blocks,
        double baseFontSize = 13.5,
        string codeFontName = "Consolas",
        Action<string>? onLinkClicked = null,
        Action<string>? onCodeCopied = null)
    {
        var codeFont = CreateSafeFont(codeFontName, "Consolas, Cascadia Code, Courier New");
        var bodyFont = new WMedia.FontFamily("Segoe UI");

        var doc = new WDocs.FlowDocument
        {
            FontFamily = bodyFont,
            FontSize = baseFontSize,
            Foreground = Brush("#e2e8f0"),
            Background = WMedia.Brushes.Transparent,
            PagePadding = new Thickness(22, 16, 22, 26),
            ColumnWidth = double.PositiveInfinity,
            IsOptimalParagraphEnabled = false,
            IsHyphenationEnabled = false
        };

        foreach (var block in blocks)
        {
            switch (block.Kind)
            {
                case MarkdownBlockKind.Heading:
                    AppendHeading(doc, block, baseFontSize, codeFont, onLinkClicked);
                    break;

                case MarkdownBlockKind.Paragraph:
                    var para = new WDocs.Paragraph
                    {
                        Margin = new Thickness(0, 0, 0, 10),
                        LineHeight = baseFontSize * 1.55
                    };
                    AppendInlines(para.Inlines, block.Inlines, codeFont, baseFontSize, onLinkClicked);
                    doc.Blocks.Add(para);
                    break;

                case MarkdownBlockKind.HorizontalRule:
                    doc.Blocks.Add(new WDocs.BlockUIContainer(new WControls.Border
                    {
                        Height = 1,
                        Background = Brush("#1e293b"),
                        Margin = new Thickness(0, 10, 0, 14)
                    }));
                    break;

                case MarkdownBlockKind.BlockQuote:
                    doc.Blocks.Add(BuildBlockQuote(block, baseFontSize, codeFont, onLinkClicked));
                    break;

                case MarkdownBlockKind.CodeBlock:
                    doc.Blocks.Add(BuildCodeBlock(block, baseFontSize, codeFont, onCodeCopied));
                    break;

                case MarkdownBlockKind.List:
                    AppendList(doc, block, baseFontSize, codeFont, onLinkClicked);
                    break;

                case MarkdownBlockKind.Table:
                    doc.Blocks.Add(BuildTable(block, baseFontSize, codeFont, onLinkClicked));
                    break;
            }
        }

        return doc;
    }

    private static void AppendHeading(
        WDocs.FlowDocument doc,
        MarkdownBlock block,
        double baseFontSize,
        WMedia.FontFamily codeFont,
        Action<string>? onLinkClicked)
    {
        var (scale, fgHex, topMargin, bottomMargin, showUnderline) = block.HeadingLevel switch
        {
            1 => (1.62, "#f8fafc", 4.0, 6.0, true),
            2 => (1.32, "#38bdf8", 16.0, 6.0, true),
            3 => (1.14, "#93c5fd", 12.0, 4.0, false),
            _ => (1.04, "#cbd5e1", 10.0, 4.0, false)
        };

        var p = new WDocs.Paragraph
        {
            FontSize = Math.Round(baseFontSize * scale, 1),
            FontWeight = FontWeights.SemiBold,
            Foreground = Brush(fgHex),
            Margin = new Thickness(0, topMargin, 0, showUnderline ? 4 : bottomMargin)
        };

        AppendInlines(p.Inlines, block.Inlines, codeFont, p.FontSize, onLinkClicked);
        doc.Blocks.Add(p);

        if (showUnderline)
        {
            doc.Blocks.Add(new WDocs.BlockUIContainer(new WControls.Border
            {
                Height = 1,
                Background = Brush(block.HeadingLevel == 1 ? "#334155" : "#1e293b"),
                Margin = new Thickness(0, 0, 0, 10)
            }));
        }
    }

    private static WDocs.Block BuildBlockQuote(
        MarkdownBlock block,
        double baseFontSize,
        WMedia.FontFamily codeFont,
        Action<string>? onLinkClicked)
    {
        var (accentHex, bgHex, badgeLabel) = block.AlertKind switch
        {
            "NOTE" => ("#38bdf8", "#0c192c", "ℹ️ NOTA"),
            "TIP" => ("#34d399", "#071f1a", "💡 DICA"),
            "IMPORTANT" => ("#a78bfa", "#1a1433", "📌 IMPORTANTE"),
            "WARNING" => ("#fbbf24", "#271c09", "⚠️ ATENÇÃO"),
            "CAUTION" => ("#f87171", "#2a1215", "🛑 CUIDADO"),
            _ => ("#38bdf8", "#0f172a", string.Empty)
        };

        var stack = new WControls.StackPanel();

        if (!string.IsNullOrEmpty(badgeLabel))
        {
            stack.Children.Add(new WControls.TextBlock
            {
                Text = badgeLabel,
                FontSize = Math.Max(11, baseFontSize - 2),
                FontWeight = FontWeights.Bold,
                Foreground = Brush(accentHex),
                Margin = new Thickness(0, 0, 0, 4)
            });
        }

        var tb = new WControls.TextBlock
        {
            FontSize = Math.Max(11.5, baseFontSize - 0.5),
            Foreground = Brush("#cbd5e1"),
            TextWrapping = TextWrapping.Wrap,
            LineHeight = baseFontSize * 1.45
        };
        AppendInlines(tb.Inlines, block.Inlines, codeFont, tb.FontSize, onLinkClicked);
        stack.Children.Add(tb);

        var border = new WControls.Border
        {
            Background = Brush(bgHex),
            BorderBrush = Brush(accentHex),
            BorderThickness = new Thickness(3.5, 1, 1, 1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(12, 9, 12, 9),
            Margin = new Thickness(0, 2, 0, 12),
            Child = stack
        };

        return new WDocs.BlockUIContainer(border);
    }

    private static WDocs.Block BuildCodeBlock(
        MarkdownBlock block,
        double baseFontSize,
        WMedia.FontFamily codeFont,
        Action<string>? onCodeCopied)
    {
        string langDisplay = string.IsNullOrWhiteSpace(block.CodeLanguage)
            ? "CÓDIGO"
            : block.CodeLanguage.ToUpperInvariant();

        var headerGrid = new WControls.Grid
        {
            Margin = new Thickness(0, 0, 0, 6)
        };
        headerGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });

        var langBadge = new WControls.Border
        {
            Background = Brush("#1e293b"),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(7, 2, 7, 2),
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Center,
            Child = new WControls.TextBlock
            {
                Text = langDisplay,
                FontSize = 10.5,
                FontWeight = FontWeights.Bold,
                FontFamily = codeFont,
                Foreground = Brush("#38bdf8")
            }
        };
        WControls.Grid.SetColumn(langBadge, 0);
        headerGrid.Children.Add(langBadge);

        var copyBtn = new Wpf.Ui.Controls.Button
        {
            Content = "📋 Copiar",
            Appearance = ControlAppearance.Secondary,
            Padding = new Thickness(8, 3, 8, 3),
            FontSize = 11,
            ToolTip = "Copiar bloco de código para a área de transferência"
        };

        string rawCode = block.Text;
        copyBtn.Click += (_, _) =>
        {
            if (ClipboardHelper.Copy(rawCode))
            {
                copyBtn.Content = "✅ Copiado!";
                onCodeCopied?.Invoke(langDisplay);
                var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.8) };
                timer.Tick += (_, _) =>
                {
                    copyBtn.Content = "📋 Copiar";
                    timer.Stop();
                };
                timer.Start();
            }
        };
        WControls.Grid.SetColumn(copyBtn, 1);
        headerGrid.Children.Add(copyBtn);

        var codeBox = new WControls.TextBox
        {
            Text = rawCode,
            IsReadOnly = true,
            FontFamily = codeFont,
            FontSize = Math.Max(11, baseFontSize - 1.2),
            Foreground = Brush("#e2e8f0"),
            Background = WMedia.Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Padding = new Thickness(0),
            TextWrapping = TextWrapping.NoWrap,
            HorizontalScrollBarVisibility = WControls.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = WControls.ScrollBarVisibility.Disabled
        };

        var stack = new WControls.StackPanel();
        stack.Children.Add(headerGrid);
        stack.Children.Add(codeBox);

        var card = new WControls.Border
        {
            Background = Brush("#070b14"),
            BorderBrush = Brush("#1e293b"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(12, 8, 12, 10),
            Margin = new Thickness(0, 2, 0, 12),
            Child = stack
        };

        return new WDocs.BlockUIContainer(card);
    }

    private static void AppendList(
        WDocs.FlowDocument doc,
        MarkdownBlock block,
        double baseFontSize,
        WMedia.FontFamily codeFont,
        Action<string>? onLinkClicked)
    {
        foreach (var item in block.ListItems)
        {
            double leftIndent = 8 + (item.IndentLevel * 18);
            var p = new WDocs.Paragraph
            {
                Margin = new Thickness(leftIndent, 0, 0, 5),
                LineHeight = baseFontSize * 1.5
            };

            p.Inlines.Add(new WDocs.Run($"{item.Marker} ")
            {
                Foreground = Brush("#38bdf8"),
                FontWeight = FontWeights.SemiBold
            });

            AppendInlines(p.Inlines, item.Inlines, codeFont, baseFontSize, onLinkClicked);
            doc.Blocks.Add(p);
        }

        doc.Blocks.Add(new WDocs.Paragraph { Margin = new Thickness(0, 0, 0, 4), FontSize = 4 });
    }

    private static WDocs.Block BuildTable(
        MarkdownBlock block,
        double baseFontSize,
        WMedia.FontFamily codeFont,
        Action<string>? onLinkClicked)
    {
        int colCount = Math.Max(
            block.TableHeaders.Count,
            block.TableRows.Count > 0 ? block.TableRows.Max(r => r.Count) : 0);

        var table = new WDocs.Table
        {
            CellSpacing = 0,
            BorderBrush = Brush("#334155"),
            BorderThickness = new Thickness(1),
            Margin = new Thickness(0, 4, 0, 14)
        };

        for (int c = 0; c < colCount; c++)
        {
            table.Columns.Add(new WDocs.TableColumn
            {
                Width = c == 0 && colCount == 2
                    ? new GridLength(1, GridUnitType.Star)
                    : new GridLength(colCount == 2 ? 2.2 : 1, GridUnitType.Star)
            });
        }

        var group = new WDocs.TableRowGroup();

        if (block.TableHeaders.Count > 0)
        {
            var headerRow = new WDocs.TableRow
            {
                Background = Brush("#1e293b")
            };

            for (int c = 0; c < colCount; c++)
            {
                var p = new WDocs.Paragraph
                {
                    Margin = new Thickness(0),
                    FontWeight = FontWeights.SemiBold,
                    FontSize = Math.Max(11.5, baseFontSize - 1),
                    Foreground = Brush("#f8fafc")
                };

                if (c < block.TableHeaders.Count)
                    AppendInlines(p.Inlines, block.TableHeaders[c], codeFont, p.FontSize, onLinkClicked);

                headerRow.Cells.Add(new WDocs.TableCell(p)
                {
                    Padding = new Thickness(10, 7, 10, 7),
                    BorderBrush = Brush("#334155"),
                    BorderThickness = new Thickness(0, 0, c < colCount - 1 ? 1 : 0, 1)
                });
            }

            group.Rows.Add(headerRow);
        }

        for (int r = 0; r < block.TableRows.Count; r++)
        {
            var rowData = block.TableRows[r];
            var tr = new WDocs.TableRow
            {
                Background = Brush(r % 2 == 0 ? "#0f172a" : "#111827")
            };

            for (int c = 0; c < colCount; c++)
            {
                var p = new WDocs.Paragraph
                {
                    Margin = new Thickness(0),
                    FontSize = Math.Max(11.5, baseFontSize - 1),
                    Foreground = Brush("#cbd5e1"),
                    LineHeight = baseFontSize * 1.4
                };

                if (c < rowData.Count)
                    AppendInlines(p.Inlines, rowData[c], codeFont, p.FontSize, onLinkClicked);

                tr.Cells.Add(new WDocs.TableCell(p)
                {
                    Padding = new Thickness(10, 6, 10, 6),
                    BorderBrush = Brush("#1e293b"),
                    BorderThickness = new Thickness(0, 0, c < colCount - 1 ? 1 : 0, r < block.TableRows.Count - 1 ? 1 : 0)
                });
            }

            group.Rows.Add(tr);
        }

        table.RowGroups.Add(group);
        return table;
    }

    private static void AppendInlines(
        WDocs.InlineCollection target,
        IReadOnlyList<MarkdownInlineSpan> spans,
        WMedia.FontFamily codeFont,
        double currentFontSize,
        Action<string>? onLinkClicked)
    {
        foreach (var span in spans)
        {
            switch (span.Kind)
            {
                case MarkdownInlineKind.Text:
                    target.Add(new WDocs.Run(span.Text));
                    break;

                case MarkdownInlineKind.Bold:
                    target.Add(new WDocs.Bold(new WDocs.Run(span.Text))
                    {
                        Foreground = Brush("#f8fafc")
                    });
                    break;

                case MarkdownInlineKind.Italic:
                    target.Add(new WDocs.Italic(new WDocs.Run(span.Text))
                    {
                        Foreground = Brush("#cbd5e1")
                    });
                    break;

                case MarkdownInlineKind.Strikethrough:
                    target.Add(new WDocs.Run(span.Text)
                    {
                        TextDecorations = TextDecorations.Strikethrough,
                        Foreground = Brush("#64748b")
                    });
                    break;

                case MarkdownInlineKind.Code:
                    target.Add(new WDocs.Run($" {span.Text} ")
                    {
                        FontFamily = codeFont,
                        FontSize = Math.Max(11, currentFontSize - 1.2),
                        Background = Brush("#1e293b"),
                        Foreground = Brush("#7dd3fc")
                    });
                    break;

                case MarkdownInlineKind.Link:
                    string href = span.TargetUrl ?? string.Empty;
                    var link = new WDocs.Hyperlink(new WDocs.Run(span.Text))
                    {
                        Foreground = Brush("#38bdf8"),
                        Cursor = Cursors.Hand,
                        ToolTip = string.IsNullOrWhiteSpace(href) ? null : href
                    };
                    link.Click += (_, e) =>
                    {
                        e.Handled = true;
                        onLinkClicked?.Invoke(href);
                    };
                    target.Add(link);
                    break;
            }
        }
    }

    private static WMedia.FontFamily CreateSafeFont(string preferred, string fallbackSpec)
    {
        if (string.IsNullOrWhiteSpace(preferred))
            return new WMedia.FontFamily(fallbackSpec);
        return new WMedia.FontFamily($"{preferred}, {fallbackSpec}");
    }

    private static WMedia.SolidColorBrush Brush(string hex)
    {
        var brush = (WMedia.SolidColorBrush)new WMedia.BrushConverter().ConvertFromString(hex)!;
        if (brush.CanFreeze) brush.Freeze();
        return brush;
    }
}
