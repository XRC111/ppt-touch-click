using System;
using System.Runtime.InteropServices;

public static class ComProbe
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr LoadLibraryW(string p);
    [DllImport("kernel32.dll")]
    static extern IntPtr GetProcAddress(IntPtr h, string name);
    [DllImport("kernel32.dll")]
    static extern bool FreeLibrary(IntPtr h);
    [DllImport("kernel32.dll")]
    static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll")]
    static extern bool IsWow64Process(IntPtr h, out bool wow);

    static string Wide(IntPtr p) { return p == IntPtr.Zero ? "(null)" : Marshal.PtrToStringUni(p); }

    [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
    delegate int GetClassObjectFn(IntPtr clsid, IntPtr iid, out IntPtr ppv);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int CanUnloadFn();

    const int S_OK = 0;
    const int S_FALSE = 1;

    public static string Run(string dllPath)
    {
        IntPtr h = LoadLibraryW(dllPath);
        if (h == IntPtr.Zero) return "LoadLibrary failed err=" + Marshal.GetLastWin32Error();
        try
        {
            bool wow; IsWow64Process(GetCurrentProcess(), out wow);

            IntPtr pClass = GetProcAddress(h, "DllGetClassObject");
            IntPtr pUnload = GetProcAddress(h, "DllCanUnloadNow");
            if (pClass == IntPtr.Zero) return "DllGetClassObject export missing (32/64 位不匹配？进程 wow64=" + wow + ")";

            var clsid = new byte[16];
            var iid = new byte[16];
            Guid.Parse("2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21").TryWriteBytes(clsid);
            Guid.Parse("00000001-0000-0000-C000-000000000046").TryWriteBytes(iid);

            IntPtr pClsid = Marshal.AllocHGlobal(16), pIid = Marshal.AllocHGlobal(16), ppv;
            try
            {
                Marshal.Copy(clsid, 0, pClsid, 16);
                Marshal.Copy(iid, 0, pIid, 16);
                var fn = (GetClassObjectFn)Marshal.GetDelegateForFunctionPointer(pClass, typeof(GetClassObjectFn));
                int hr = fn(pClsid, pIid, out ppv);

                int unload = -1;
                if (pUnload != IntPtr.Zero)
                    unload = ((CanUnloadFn)Marshal.GetDelegateForFunctionPointer(pUnload, typeof(CanUnloadFn)))();

                return "wow64=" + wow + " DllGetClassObject=0x" + hr.ToString("X8") + " DllCanUnloadNow=0x" + unload.ToString("X8");
            }
            finally { Marshal.FreeHGlobal(pClsid); Marshal.FreeHGlobal(pIid); }
        }
        finally { FreeLibrary(h); }
    }
}