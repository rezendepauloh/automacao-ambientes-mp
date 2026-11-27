using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using AmbientesHub.Models;
using Wpf.Ui.Controls;
using WControls = System.Windows.Controls;
using WDocs = System.Windows.Documents;
using WMedia = System.Windows.Media;

namespace AmbientesHub.Services;

/// <summary>
/// Constrói um FlowDocument nativo WPF estilizado no padrão Fluent Dark do AmbientesHub
/// a partir da árvore de blocos gerada pelo DocumentationService.ParseMarkdown,
/// com suporte a imagens inline sob demanda (OneDrive-friendly) e cards de vídeo/anexos.
/// </summary>
public static class MarkdownFlowDocumentBuilder
{
    private const long MaxUiImageBytes = 25 * 1024 * 1024; // 25 MB

    public static WDocs.FlowDocument Build(
        IReadOnlyList<MarkdownBlock> blocks,
        double baseFontSize = 13.5,
        string codeFontName = "Consolas",
        Action<string>? onLinkClicked = null,
        Action<string>? onCodeCopied = null,
        string? sourceRootPath = null,
        string? currentDocFullPath = null)
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
            TextAlignment = TextAlignment.Left,
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

                case MarkdownBlockKind.Image:
                    doc.Blocks.Add(BuildImageBlock(block, sourceRootPath, currentDocFullPath, onLinkClicked));
                    break;

                case MarkdownBlockKind.AttachmentCard:
                    doc.Blocks.Add(BuildAttachmentCard(block, codeFont, sourceRootPath, currentDocFullPath, onLinkClicked));
                    break;
            }
        }

        return doc;
    }

    private static WDocs.Block BuildImageBlock(
        MarkdownBlock block,
        string? sourceRootPath,
        string? currentDocFullPath,
        Action<string>? onLinkClicked)
    {
        string caption = string.IsNullOrWhiteSpace(block.ImageAlt)
            ? Path.GetFileName(block.ImageSource)
            : block.ImageAlt;

        BitmapImage? bitmap = null;
        string? resolvedPath = null;

        if (!string.IsNullOrWhiteSpace(sourceRootPath) && !string.IsNullOrWhiteSpace(currentDocFullPath))
        {
            var res = DocumentationService.ResolveMarkdownLinkTarget(sourceRootPath, currentDocFullPath, block.ImageSource);
            if (res.Kind == MarkdownLinkKind.LocalImage && File.Exists(res.Target))
            {
                resolvedPath = res.Target;
                bitmap = TryLoadBitmapOnDemand(res.Target);
            }
        }

        var stack = new WControls.StackPanel();

        if (bitmap != null)
        {
            var img = new WControls.Image
            {
                Source = bitmap,
                MaxHeight = 460,
                Stretch = WMedia.Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Center,
                Cursor = Cursors.Hand,
                ToolTip = "Clique para abrir a imagem em tamanho original"
            };
            WMedia.RenderOptions.SetBitmapScalingMode(img, WMedia.BitmapScalingMode.HighQuality);

            string imgSrc = block.ImageSource;
            img.MouseLeftButtonUp += (_, _) => onLinkClicked?.Invoke(imgSrc);
            stack.Children.Add(img);
        }
        else
        {
            stack.Children.Add(new WControls.TextBlock
            {
                Text = $"🖼️ Imagem indisponível localmente: {block.ImageSource}",
                FontSize = 12,
                Foreground = Brush("#fbbf24"),
                TextWrapping = TextWrapping.Wrap,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 6, 0, 4)
            });
        }

        if (!string.IsNullOrWhiteSpace(caption))
        {
            stack.Children.Add(new WControls.TextBlock
            {
                Text = resolvedPath != null ? $"🖼️ {caption} (clique na imagem para ampliar)" : caption,
                FontSize = 11.5,
                Foreground = Brush("#94a3b8"),
                HorizontalAlignment = HorizontalAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 8, 0, 0)
            });
        }

        var border = new WControls.Border
        {
            Background = Brush("#070b14"),
            BorderBrush = Brush("#1e293b"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(12, 10, 12, 10),
            Margin = new Thickness(0, 4, 0, 14),
            Child = stack
        };

        return new WDocs.BlockUIContainer(border);
    }

    private static BitmapImage? TryLoadBitmapOnDemand(string fullPath)
    {
        try
        {
            var fi = new FileInfo(fullPath);
            if (!fi.Exists || fi.Length <= 0 || fi.Length > MaxUiImageBytes)
                return null;

            // Lê em memória (OnLoad) para não manter handle aberto bloqueando a sincronização do OneDrive.
            byte[] bytes = File.ReadAllBytes(fullPath);
            using var ms = new MemoryStream(bytes);
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.CacheOption = BitmapCacheOption.OnLoad;
            bmp.StreamSource = ms;
            bmp.EndInit();
            bmp.Freeze();
            return bmp;
        }
        catch
        {
            return null;
        }
    }

    private static WDocs.Block BuildAttachmentCard(
        MarkdownBlock block,
        WMedia.FontFamily codeFont,
        string? sourceRootPath,
        string? currentDocFullPath,
        Action<string>? onLinkClicked)
    {
        string ext = Path.GetExtension(block.AttachmentSource).ToLowerInvariant();
        bool isVideo = ext is ".mp4" or ".mkv" or ".webm" or ".mov" or ".avi";
        var (emoji, badge, accentBg, accentFg, actionText) = ext switch
        {
            ".mp4" or ".mkv" or ".webm" or ".mov" or ".avi" => ("🎬", $"VÍDEO {ext.TrimStart('.').ToUpperInvariant()}", "#1e3a8a", "#93c5fd", "🪟 Externo"),
            ".pdf" => ("📄", "PDF", "#7f1d1d", "#fca5a5", "📖 Abrir PDF"),
            ".docx" => ("📝", "WORD DOCX", "#1e3a8a", "#bfdbfe", "📝 Abrir Documento"),
            ".xlsx" => ("📊", "EXCEL", "#064e3b", "#6ee7b7", "📊 Abrir Planilha"),
            _ => ("📎", "ANEXO", "#1e293b", "#cbd5e1", "📂 Abrir Anexo")
        };

        var outerStack = new WControls.StackPanel();

        var headerGrid = new WControls.Grid();
        headerGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });

        var infoStack = new WControls.StackPanel { Margin = new Thickness(0, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center };

        var topRow = new WControls.WrapPanel { Orientation = WControls.Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 3) };
        topRow.Children.Add(new WControls.Border
        {
            Background = Brush(accentBg),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(6, 2, 6, 2),
            Margin = new Thickness(0, 0, 8, 0),
            Child = new WControls.TextBlock
            {
                Text = badge,
                FontSize = 10,
                FontWeight = FontWeights.Bold,
                Foreground = Brush(accentFg)
            }
        });
        topRow.Children.Add(new WControls.TextBlock
        {
            Text = $"{emoji} {block.AttachmentLabel}",
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Foreground = Brush("#f8fafc"),
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = VerticalAlignment.Center
        });
        infoStack.Children.Add(topRow);

        infoStack.Children.Add(new WControls.TextBlock
        {
            Text = $"{block.AttachmentSource}  •  Sob demanda (OneDrive só baixa ao reproduzir/abrir)",
            FontSize = 11,
            FontFamily = codeFont,
            Foreground = Brush("#94a3b8"),
            TextWrapping = TextWrapping.Wrap
        });

        WControls.Grid.SetColumn(infoStack, 0);
        headerGrid.Children.Add(infoStack);

        var buttonsPanel = new WControls.StackPanel
        {
            Orientation = WControls.Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center
        };

        string targetSrc = block.AttachmentSource;
        WControls.StackPanel? playerContainer = null;

        if (isVideo)
        {
            var inlinePlayBtn = new Wpf.Ui.Controls.Button
            {
                Content = "▶ Assistir Aqui",
                Appearance = ControlAppearance.Primary,
                Padding = new Thickness(11, 6, 11, 6),
                Margin = new Thickness(0, 0, 6, 0),
                FontSize = 12,
                ToolTip = "Abre o player embutido e baixa o vídeo sob demanda pelo OneDrive"
            };
            buttonsPanel.Children.Add(inlinePlayBtn);

            playerContainer = BuildEmbeddedVideoPlayer(
                sourceRootPath,
                currentDocFullPath,
                targetSrc,
                codeFont,
                inlinePlayBtn);
        }

        var openBtn = new Wpf.Ui.Controls.Button
        {
            Content = actionText,
            Appearance = isVideo ? ControlAppearance.Secondary : ControlAppearance.Primary,
            Padding = new Thickness(11, 6, 11, 6),
            FontSize = 12,
            VerticalAlignment = VerticalAlignment.Center,
            ToolTip = isVideo ? "Abrir no reprodutor padrão do Windows" : "Abrir arquivo sob demanda"
        };
        openBtn.Click += (_, _) => onLinkClicked?.Invoke(targetSrc);
        buttonsPanel.Children.Add(openBtn);

        WControls.Grid.SetColumn(buttonsPanel, 1);
        headerGrid.Children.Add(buttonsPanel);

        outerStack.Children.Add(headerGrid);
        if (playerContainer != null)
            outerStack.Children.Add(playerContainer);

        var card = new WControls.Border
        {
            Background = Brush("#0f172a"),
            BorderBrush = Brush("#334155"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(14, 10, 14, 10),
            Margin = new Thickness(0, 4, 0, 12),
            Child = outerStack
        };

        return new WDocs.BlockUIContainer(card);
    }

    /// <summary>
    /// Constrói um player de vídeo embutido (WPF MediaElement) 100% sob demanda:
    /// nenhum byte é lido ou baixado do OneDrive até o usuário clicar em "▶ Assistir Aqui".
    /// </summary>
    private static WControls.StackPanel BuildEmbeddedVideoPlayer(
        string? sourceRootPath,
        string? currentDocFullPath,
        string targetSrc,
        WMedia.FontFamily codeFont,
        Wpf.Ui.Controls.Button triggerBtn)
    {
        var playerPanel = new WControls.StackPanel
        {
            Visibility = Visibility.Collapsed,
            Margin = new Thickness(0, 10, 0, 0)
        };

        var statusTxt = new WControls.TextBlock
        {
            Text = "⏳ Carregando vídeo (OneDrive sob demanda)...",
            FontSize = 11.5,
            Foreground = Brush("#38bdf8"),
            Margin = new Thickness(0, 0, 0, 6)
        };

        var media = new WControls.MediaElement
        {
            LoadedBehavior = WControls.MediaState.Manual,
            UnloadedBehavior = WControls.MediaState.Stop,
            Stretch = WMedia.Stretch.Uniform,
            MaxHeight = 420,
            ScrubbingEnabled = true,
            Volume = 0.85
        };

        var videoBorder = new WControls.Border
        {
            Background = WMedia.Brushes.Black,
            BorderBrush = Brush("#1e293b"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            MinHeight = 200,
            Child = media
        };

        var controlsGrid = new WControls.Grid { Margin = new Thickness(0, 8, 0, 0) };
        controlsGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });
        controlsGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        controlsGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });
        controlsGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });
        controlsGrid.ColumnDefinitions.Add(new WControls.ColumnDefinition { Width = GridLength.Auto });

        var playPauseBtn = new Wpf.Ui.Controls.Button
        {
            Content = "⏸ Pausar",
            Appearance = ControlAppearance.Secondary,
            Padding = new Thickness(10, 4, 10, 4),
            Margin = new Thickness(0, 0, 8, 0),
            FontSize = 11.5
        };

        var seekSlider = new WControls.Slider
        {
            Minimum = 0,
            Maximum = 1,
            Value = 0,
            IsMoveToPointEnabled = true,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 8, 0)
        };

        var timeTxt = new WControls.TextBlock
        {
            Text = "00:00 / 00:00",
            FontFamily = codeFont,
            FontSize = 11,
            Foreground = Brush("#94a3b8"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 8, 0)
        };

        var muteBtn = new Wpf.Ui.Controls.Button
        {
            Content = "🔊",
            Appearance = ControlAppearance.Transparent,
            Padding = new Thickness(7, 4, 7, 4),
            Margin = new Thickness(0, 0, 6, 0),
            FontSize = 12,
            ToolTip = "Silenciar / Ativar áudio"
        };

        var closeBtn = new Wpf.Ui.Controls.Button
        {
            Content = "⏹ Fechar",
            Appearance = ControlAppearance.Secondary,
            Padding = new Thickness(9, 4, 9, 4),
            FontSize = 11.5,
            ToolTip = "Para o vídeo e libera o arquivo imediatamente"
        };

        WControls.Grid.SetColumn(playPauseBtn, 0);
        WControls.Grid.SetColumn(seekSlider, 1);
        WControls.Grid.SetColumn(timeTxt, 2);
        WControls.Grid.SetColumn(muteBtn, 3);
        WControls.Grid.SetColumn(closeBtn, 4);

        controlsGrid.Children.Add(playPauseBtn);
        controlsGrid.Children.Add(seekSlider);
        controlsGrid.Children.Add(timeTxt);
        controlsGrid.Children.Add(muteBtn);
        controlsGrid.Children.Add(closeBtn);

        playerPanel.Children.Add(statusTxt);
        playerPanel.Children.Add(videoBorder);
        playerPanel.Children.Add(controlsGrid);

        bool isPlaying = false;
        bool isDragging = false;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };

        static string Fmt(TimeSpan ts) => ts.TotalHours >= 1
            ? ts.ToString(@"hh\:mm\:ss")
            : ts.ToString(@"mm\:ss");

        void StopAndRelease()
        {
            timer.Stop();
            isPlaying = false;
            try
            {
                media.Stop();
                media.Close();
                media.Source = null;
            }
            catch { }
            playerPanel.Visibility = Visibility.Collapsed;
            triggerBtn.Content = "▶ Assistir Aqui";
            seekSlider.Value = 0;
            timeTxt.Text = "00:00 / 00:00";
        }

        timer.Tick += (_, _) =>
        {
            if (isDragging || !media.NaturalDuration.HasTimeSpan) return;
            var total = media.NaturalDuration.TimeSpan;
            var cur = media.Position;
            seekSlider.Maximum = Math.Max(1, total.TotalSeconds);
            seekSlider.Value = Math.Clamp(cur.TotalSeconds, 0, seekSlider.Maximum);
            timeTxt.Text = $"{Fmt(cur)} / {Fmt(total)}";
        };

        media.MediaOpened += (_, _) =>
        {
            statusTxt.Visibility = Visibility.Collapsed;
            if (media.NaturalDuration.HasTimeSpan)
            {
                var total = media.NaturalDuration.TimeSpan;
                seekSlider.Maximum = Math.Max(1, total.TotalSeconds);
                timeTxt.Text = $"00:00 / {Fmt(total)}";
            }
            timer.Start();
        };

        media.MediaEnded += (_, _) =>
        {
            timer.Stop();
            isPlaying = false;
            playPauseBtn.Content = "▶ Reproduzir";
            media.Position = TimeSpan.Zero;
            media.Pause();
        };

        media.MediaFailed += (_, e) =>
        {
            timer.Stop();
            isPlaying = false;
            statusTxt.Visibility = Visibility.Visible;
            statusTxt.Foreground = Brush("#fbbf24");
            statusTxt.Text = $"⚠️ Não foi possível reproduzir no player embutido ({e.ErrorException?.Message ?? "codec/arquivo indisponível"}). Use '🪟 Externo'.";
        };

        media.Unloaded += (_, _) => StopAndRelease();

        seekSlider.PreviewMouseLeftButtonDown += (_, _) => isDragging = true;
        seekSlider.PreviewMouseLeftButtonUp += (_, _) =>
        {
            isDragging = false;
            if (media.NaturalDuration.HasTimeSpan)
            {
                media.Position = TimeSpan.FromSeconds(seekSlider.Value);
            }
        };

        playPauseBtn.Click += (_, _) =>
        {
            if (isPlaying)
            {
                media.Pause();
                isPlaying = false;
                playPauseBtn.Content = "▶ Continuar";
            }
            else
            {
                media.Play();
                isPlaying = true;
                playPauseBtn.Content = "⏸ Pausar";
                timer.Start();
            }
        };

        muteBtn.Click += (_, _) =>
        {
            media.IsMuted = !media.IsMuted;
            muteBtn.Content = media.IsMuted ? "🔇" : "🔊";
        };

        closeBtn.Click += (_, _) => StopAndRelease();

        triggerBtn.Click += (_, _) =>
        {
            if (playerPanel.Visibility == Visibility.Visible)
            {
                StopAndRelease();
                return;
            }

            if (string.IsNullOrWhiteSpace(sourceRootPath) || string.IsNullOrWhiteSpace(currentDocFullPath))
                return;

            var res = DocumentationService.ResolveMarkdownLinkTarget(sourceRootPath, currentDocFullPath, targetSrc);
            if (res.Kind != MarkdownLinkKind.LocalAttachment || !File.Exists(res.Target))
            {
                statusTxt.Visibility = Visibility.Visible;
                statusTxt.Foreground = Brush("#fbbf24");
                statusTxt.Text = "⚠️ Vídeo não localizado na pasta sincronizada do OneDrive.";
                playerPanel.Visibility = Visibility.Visible;
                return;
            }

            statusTxt.Visibility = Visibility.Visible;
            statusTxt.Foreground = Brush("#38bdf8");
            statusTxt.Text = "⏳ Carregando vídeo (OneDrive sob demanda)...";
            playerPanel.Visibility = Visibility.Visible;
            triggerBtn.Content = "⏹ Ocultar Player";
            playPauseBtn.Content = "⏸ Pausar";
            isPlaying = true;
            media.Source = new Uri(res.Target, UriKind.Absolute);
            media.Play();
        };

        return playerPanel;
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
                case MarkdownInlineKind.Image:
                    string href = span.TargetUrl ?? string.Empty;
                    string label = span.Kind == MarkdownInlineKind.Image ? $"🖼️ {span.Text}" : span.Text;
                    var link = new WDocs.Hyperlink(new WDocs.Run(label))
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
