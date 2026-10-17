using System;
using System.Runtime.InteropServices;

public static class ComProbe
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr LoadLibraryW(string p);
    [DllImport("kernel32.dll")]
    static extern bool FreeLibrary(IntPtr h);

    public static string Run(string dllPath)
    {
        IntPtr h = LoadLibraryW(dllPath);
        if (h == IntPtr.Zero) return "LoadLibrary failed err=" + Marshal.GetLastWin32Error();
        try
        {
            var clsid = new Guid("2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21");
            var iid = new Guid("00000001-0000-0000-C000-000000000046");
            object o = Activator.CreateInstance(Type.GetTypeFromCLSID(clsid));
            string s = "ok: " + o.GetType().FullName;
            Marshal.FinalReleaseComObject(o);
            return s;
        }
        catch (Exception ex) { return "InprocServer32 ok, but " + ex.GetType().Name + ": " + ex.Message; }
        finally { FreeLibrary(h); }
    }
}