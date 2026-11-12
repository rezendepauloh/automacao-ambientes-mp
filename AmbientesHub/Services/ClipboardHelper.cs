using System.Runtime.InteropServices;
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
        return false;
    }
}

