using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace AmbientesHub.Services;

public static class ClipboardHelper
{
    public static bool Copy(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;

        for (int i = 0; i < 5; i++)
        {
            try
            {
                Clipboard.Clear();
                Clipboard.SetDataObject(text.Trim(), true);
                return true;
            }
            catch (COMException)
            {
                Thread.Sleep(60);
            }
            catch
            {
                try
                {
                    Clipboard.SetText(text.Trim());
                    return true;
                }
                catch
                {
                    Thread.Sleep(60);
                }
            }
        }
        AppLoggerService.Warning($"ClipboardHelper.Copy falhou após 5 tentativas (tamanho={text.Trim().Length} chars).", nameof(ClipboardHelper));
        return false;
    }

    /// <summary>
    /// Copia um fragmento HTML no formato nativo CF_HTML do Windows (com fallback em texto puro),
    /// permitindo colar texto rico formatado (títulos, listas, tabelas, links) diretamente na Web Part Texto do SharePoint, Word ou Teams.
    /// </summary>
    public static bool CopyHtml(string htmlFragment, string plainTextFallback)
    {
        if (string.IsNullOrWhiteSpace(htmlFragment)) return false;

        string cfHtml = BuildCfHtmlPayload(htmlFragment);

        for (int i = 0; i < 5; i++)
        {
            try
            {
                var data = new DataObject();
                data.SetData(DataFormats.Html, cfHtml);
                data.SetData(DataFormats.UnicodeText, plainTextFallback ?? string.Empty);
                data.SetData(DataFormats.Text, plainTextFallback ?? string.Empty);

                Clipboard.Clear();
                Clipboard.SetDataObject(data, true);
                return true;
            }
            catch (COMException)
            {
                Thread.Sleep(60);
            }
            catch
            {
                Thread.Sleep(60);
            }
        }

        AppLoggerService.Warning("ClipboardHelper.CopyHtml falhou após 5 tentativas.", nameof(ClipboardHelper));
        return false;
    }

    public static string BuildCfHtmlPayload(string htmlFragment)
    {
        const string headerTemplate =
            "Version:0.9\r\n" +
            "StartHTML:{0:D10}\r\n" +
            "EndHTML:{1:D10}\r\n" +
            "StartFragment:{2:D10}\r\n" +
            "EndFragment:{3:D10}\r\n";

        const string prefix = "<html><body><!--StartFragment-->";
        const string suffix = "<!--EndFragment--></body></html>";

        int headerBytes = Encoding.UTF8.GetByteCount(string.Format(headerTemplate, 0, 0, 0, 0));
        int prefixBytes = Encoding.UTF8.GetByteCount(prefix);
        int fragmentBytes = Encoding.UTF8.GetByteCount(htmlFragment);
        int suffixBytes = Encoding.UTF8.GetByteCount(suffix);

        int startHtml = headerBytes;
        int startFragment = startHtml + prefixBytes;
        int endFragment = startFragment + fragmentBytes;
        int endHtml = endFragment + suffixBytes;

        return string.Format(headerTemplate, startHtml, endHtml, startFragment, endFragment) + prefix + htmlFragment + suffix;
    }
}

