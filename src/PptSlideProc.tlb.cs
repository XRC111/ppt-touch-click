// 类型库生成源：CI 用 SDK 的 tlbexp.exe 导出为 tlb（本机没有 SDK，只保留源）
// 仅用于让组件可以被早期绑定调用；PptTouchClick.exe 走 IDispatch，不依赖类型库。
using System;
using System.Runtime.InteropServices;

namespace PptSlideProc
{
    [ComVisible(true)]
    [Guid("7D2E1B44-03C5-4F88-B172-9E4D1860C29F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISlideWatcher
    {
        bool Start();
        string GetStatus();
        bool Stop();
    }

    [ComVisible(true)]
    [Guid("2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21")]
    [ProgId("PptTouchClick.SlideWatcher")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SlideWatcher : ISlideWatcher
    {
        public bool Start() { return true; }
        public string GetStatus() { return "stub"; }
        public bool Stop() { return true; }
    }
}