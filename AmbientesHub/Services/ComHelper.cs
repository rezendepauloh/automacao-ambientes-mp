using System.Runtime.InteropServices;

namespace AmbientesHub.Services;

public static class ComHelper
{
    public static object? GetActiveObject(string progId, bool throwOnError = false)
    {
        if (progId == null) throw new ArgumentNullException(nameof(progId));

        int hr = CLSIDFromProgIDEx(progId, out Guid clsid);
        if (hr < 0)
        {
            if (throwOnError) Marshal.ThrowExceptionForHR(hr);
            return null;
        }

        hr = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
        if (hr < 0)
        {
            if (throwOnError) Marshal.ThrowExceptionForHR(hr);
            return null;
        }

        return obj;
    }

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    private static extern int CLSIDFromProgIDEx(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid lpclsid
    );

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(
        [MarshalAs(UnmanagedType.LPStruct)] ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.Interface)] out object ppunk
    );
}

