// PptSlideProc - PowerPoint "当前放映页"读取 + 全局触摸手势过滤（给 PptTouchClick 服务）
//
// 输出：PptSlideProc.ocx（32 位，进程内 COM）与 PptSlideProc.x64.dll（64 位，进程内 COM）
//
// 原理：
//   PowerPoint 放映/编辑时会把"最近访问过的幻灯片"顺序写进
//     HKCU\Software\Microsoft\Office\16.0\PowerPoint\Document MRU\<文件>
//       MRUList        "最后访问顺序"，每个字符对应一个 "Item X"（0 = 最近）
//       Item X         X 对应的 PowerPoint 内部 slideID
//       Settings       二进制：逐个"幻灯片区段" = [uint8 段内张数 v][v 个 uint8 访问序号]
//       Settings.k     第 k 个区段里各幻灯片的内部 slideID
//   用 MRUList 的先后把"当前页"算出来，写进
//     HKCU\Software\PptTouchClick\Slide  {Count, Current, AgeMs}
//   由 PptTouchClick.exe 读取（C# 端不解析任何二进制）。
//
// 为什么要原生：C# 托管进程调 SetWindowsHookEx(WH_POINTER) 时拿不到模块句柄，钩子装不上；
// 这里用原生代码装全局指针钩子，吞掉"滑动 / 长按"，只放行"点按"，
// 避免触摸模式下被当成换页手势，也避免误触双指缩放。

#include "SlideWatcher.h"

#include <algorithm>
#include <atomic>
#include <chrono>
#include <climits>
#include <cstdint>
#include <fstream>
#include <mutex>
#include <string>
#include <thread>
#include <unordered_map>
#include <vector>

#include <olectl.h>    // RegisterTypeLibForUser / LoadTypeLibFromResource / RT_TYPELIB

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "ole32.lib")

static const GUID kClsidSlideWatcher = CLSID_SlideWatcher;

// 指针设备钩子常量（WH_POINTER 需要 WinUser.h 中较新的定义，这里显式给出）
#ifndef WH_POINTER
#define WH_POINTER 24
#endif
static const DWORD kPTTouch = 2;
static const DWORD kFlagDown = 0x00010000;
static const DWORD kFlagUpdate = 0x00020000;
static const DWORD kFlagUp = 0x00040000;
static const DWORD kFlagCanceled = 0x80000000;

// ---------------------------------------------------------------------------
// 小工具
// ---------------------------------------------------------------------------
static std::mutex& LogMutex()
{
    static std::mutex m;
    return m;
}

static std::wstring ToLower(const std::wstring& s)
{
    std::wstring r = s;
    for (auto& c : r) c = (wchar_t)towlower(c);
    return r;
}

static void WriteLog(const std::wstring& msg)
{
    wchar_t path[MAX_PATH] = {};
    DWORD n = GetModuleFileNameW(nullptr, path, MAX_PATH);
    if (!n) return;
    std::wstring p(path, n);
    size_t slash = p.find_last_of(L"\\/");
    if (slash == std::wstring::npos) return;
    p = p.substr(0, slash + 1) + L"PptSlideProc.log";

    SYSTEMTIME st;
    GetLocalTime(&st);
    wchar_t ts[48];
    swprintf_s(ts, L"[%02d:%02d:%02d.%03d] ", st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);

    std::lock_guard<std::mutex> lock(LogMutex());
    std::ofstream f(p.c_str(), std::ios::app);
    if (!f.is_open()) return;
    int len = WideCharToMultiByte(CP_UTF8, 0, ts, -1, nullptr, 0, nullptr, nullptr);
    std::string s(len > 0 ? len - 1 : 0, 0);
    if (len > 1) WideCharToMultiByte(CP_UTF8, 0, ts, -1, &s[0], len, nullptr, nullptr);
    len = WideCharToMultiByte(CP_UTF8, 0, msg.c_str(), (int)msg.size(), nullptr, 0, nullptr, nullptr);
    std::string m(len > 0 ? len : 0, 0);
    if (len > 0) WideCharToMultiByte(CP_UTF8, 0, msg.c_str(), (int)msg.size(), &m[0], len, nullptr, nullptr);
    f << s << m << "\n";
}

static bool RegReadRaw(HKEY root, const std::wstring& sub, const std::wstring& name,
                       std::vector<BYTE>& out, DWORD* typeOut)
{
    HKEY k = nullptr;
    if (RegOpenKeyExW(root, sub.c_str(), 0, KEY_READ, &k) != ERROR_SUCCESS) return false;
    DWORD type = 0, size = 0;
    bool ok = false;
    if (RegQueryValueExW(k, name.c_str(), nullptr, &type, nullptr, &size) == ERROR_SUCCESS && size > 0)
    {
        out.assign(size + 8, 0);
        if (RegQueryValueExW(k, name.c_str(), nullptr, &type, out.data(), &size) == ERROR_SUCCESS)
        {
            out.resize(size);
            if (typeOut) *typeOut = type;
            ok = true;
        }
    }
    RegCloseKey(k);
    return ok;
}

static bool RegReadStr(HKEY root, const std::wstring& sub, const std::wstring& name, std::wstring& out)
{
    std::vector<BYTE> raw;
    DWORD type = 0;
    if (!RegReadRaw(root, sub, name, raw, &type)) return false;
    if ((type != REG_SZ && type != REG_EXPAND_SZ) || raw.size() < sizeof(wchar_t)) return false;
    out.assign(reinterpret_cast<const wchar_t*>(raw.data()), raw.size() / sizeof(wchar_t));
    while (!out.empty() && out.back() == L'\0') out.pop_back();
    return !out.empty();
}

// ---------------------------------------------------------------------------
// Settings 解析：若干"区段"，每段 [uint8 count][count 个 uint8 mark]
// ---------------------------------------------------------------------------
struct Segment
{
    int count = 0;
    std::vector<uint8_t> marks;
};

static bool ParseSettings(const std::vector<BYTE>& data, std::vector<Segment>& segs)
{
    segs.clear();
    size_t pos = 0;
    while (pos < data.size())
    {
        uint8_t v = data[pos];
        if (v == 0) return false;
        if (pos + 1 + (size_t)v > data.size()) return false;
        Segment s;
        s.count = (int)v;
        s.marks.assign(data.begin() + (ptrdiff_t)(pos + 1), data.begin() + (ptrdiff_t)(pos + 1 + v));
        segs.push_back(s);
        pos += 1 + (size_t)v;
        if (segs.size() > 512) return false;
    }
    return !segs.empty();
}

struct SetItem
{
    std::wstring sid;
    int mark;
    int segIndex;
    int inSeg;
};

// ---------------------------------------------------------------------------
// 选出"正在放映"的那个演示文稿并算出当前页
// ---------------------------------------------------------------------------
static bool PickPresentation(const std::wstring& baseKey, std::wstring& subOut, std::wstring& title)
{
    std::wstring mruList;
    if (!RegReadStr(HKEY_CURRENT_USER, baseKey, L"MRUList", mruList)) return false;
    for (wchar_t c : mruList)
    {
        std::wstring item = L"Item ";
        item += c;
        std::wstring path;
        if (!RegReadStr(HKEY_CURRENT_USER, baseKey, item, path)) continue;
        std::wstring sub = baseKey + L"\\" + path;
        std::vector<BYTE> st;
        if (RegReadRaw(HKEY_CURRENT_USER, sub, L"Settings", st, nullptr) && !st.empty())
        {
            subOut = sub;
            title = path;
            return true;
        }
    }
    return false;
}

static bool ComputeCurrent(int& current, int& count, std::wstring* detail)
{
    static const wchar_t* kVersionKeys[] = {
        L"Software\\Microsoft\\Office\\16.0\\PowerPoint\\Document MRU",
        L"Software\\Microsoft\\Office\\15.0\\PowerPoint\\Document MRU",
        L"Software\\Microsoft\\Office\\14.0\\PowerPoint\\Document MRU",
    };
    for (const wchar_t* base : kVersionKeys)
    {
        std::wstring sub, title;
        if (!PickPresentation(base, sub, title)) continue;

        std::vector<BYTE> settings;
        if (!RegReadRaw(HKEY_CURRENT_USER, sub, L"Settings", settings, nullptr)) continue;
        std::vector<Segment> segs;
        if (!ParseSettings(settings, segs)) continue;

        int total = 0;
        for (auto& s : segs) total += s.count;

        std::wstring order;
        RegReadStr(HKEY_CURRENT_USER, sub, L"MRUList", order);

        int best = -1, bestIter = INT_MAX;
        int visited = 0;
        for (size_t k = 0; k < segs.size(); ++k)
        {
            std::wstring name = L"Settings." + std::to_wstring(k);
            std::wstring ids;
            RegReadStr(HKEY_CURRENT_USER, sub, name, ids);

            std::vector<std::wstring> list;
            std::wstring cur;
            for (wchar_t ch : ids)
            {
                if (ch == L'\n' || ch == L'\r' || ch == L'|')
                {
                    if (!cur.empty()) list.push_back(cur);
                    cur.clear();
                }
                else if (ch != L' ' && ch != L'\t') cur += ch;
            }
            if (!cur.empty()) list.push_back(cur);

            for (int i = 0; i < segs[k].count; ++i, ++visited)
            {
                if (i >= (int)list.size()) continue;
                size_t it = order.find(list[i]);
                if (it != std::wstring::npos && (int)it < bestIter)
                {
                    bestIter = (int)it;
                    best = visited;
                }
            }
        }

        count = total;
        if (best >= 0 && total > 0)
        {
            current = best;
            if (detail) *detail = title + L" " + std::to_wstring(best + 1) + L"/" + std::to_wstring(total);
            return true;
        }
        if (detail) *detail = title + L" (无法定位当前页)";
    }
    return false;
}

// ---------------------------------------------------------------------------
// 全局指针钩子：吞掉"滑动 / 长按"，只放行"点按"
// ---------------------------------------------------------------------------
static HHOOK g_pHook = nullptr;
static std::atomic<bool> g_hookOk{ false };
static std::atomic<int> g_swallowed{ 0 };
static std::atomic<bool> g_hookFailed{ false };   // 钩子没装成（通常是没以管理员注册/系统限制）

struct TouchRec
{
    DWORD tick;
    POINT origin;
    bool swallow;
};

static std::mutex g_touchMu;
static std::unordered_map<UINT32, TouchRec> g_touch;
static int g_moveTol = 26;

static LRESULT CALLBACK PointerProc(int code, WPARAM wParam, LPARAM lParam)
{
    if (code < 0) return CallNextHookEx(g_pHook, code, wParam, lParam);
    const PPTPOINTER_INFO* info = reinterpret_cast<const PPTPOINTER_INFO*>(lParam);
    if (!info || info->pointerType != kPTTouch) return CallNextHookEx(g_pHook, code, wParam, lParam);

    const UINT32 f = info->pointerFlags;
    const UINT32 id = info->pointerId;

    std::lock_guard<std::mutex> lock(g_touchMu);
    if (f & kFlagDown)
    {
        if (g_touch.size() > 24) g_touch.clear();
        TouchRec r;
        r.tick = GetTickCount();
        r.origin = info->ptPixelLocation;
        r.swallow = false;
        g_touch[id] = r;
        return CallNextHookEx(g_pHook, code, wParam, lParam);
    }
    if (f & kFlagCanceled)
    {
        g_touch.erase(id);
        return 1;
    }
    auto it = g_touch.find(id);
    if (it == g_touch.end()) return CallNextHookEx(g_pHook, code, wParam, lParam);

    const LONG dx = info->ptPixelLocation.x - it->second.origin.x;
    const LONG dy = info->ptPixelLocation.y - it->second.origin.y;
    if (dx * dx + dy * dy > (LONG)g_moveTol * g_moveTol) it->second.swallow = true;

    if (it->second.swallow)
    {
        if (f & kFlagUp) g_touch.erase(id);
        ++g_swallowed;
        return 1;
    }
    if (f & kFlagUp)
    {
        DWORD held = GetTickCount() - it->second.tick;
        g_touch.erase(id);
        if (held >= 1500) { ++g_swallowed; return 1; }
    }
    return CallNextHookEx(g_pHook, code, wParam, lParam);
}

static bool InstallHook()
{
    if (g_pHook) return true;
    g_pHook = SetWindowsHookExW(WH_POINTER, PointerProc, GetModuleHandleW(L"PptSlideProc.x64.dll"), 0);
    if (!g_pHook) g_pHook = SetWindowsHookExW(WH_POINTER, PointerProc, GetModuleHandleW(L"PptSlideProc.ocx"), 0);
    if (!g_pHook)
    {
        g_hookFailed = true;
        WriteLog(L"SetWindowsHookEx(WH_POINTER) 失败 err=" + std::to_wstring(GetLastError()));
        return false;
    }
    g_hookFailed = false;
    g_hookOk = true;
    WriteLog(L"WH_POINTER hook installed");
    return true;
}

static void RemoveHook()
{
    if (g_pHook) UnhookWindowsHookEx(g_pHook);
    g_pHook = nullptr;
    g_hookOk = false;
}

// ---------------------------------------------------------------------------
// 页面监视线程：把当前页写进 HKCU\Software\PptTouchClick\Slide
// ---------------------------------------------------------------------------
static std::atomic<bool> g_watchRun{ false };
static std::thread g_watchThread;
static std::wstring g_lastDetail;

static void PublishSlide()
{
    int current = -1, count = 0;
    std::wstring detail;
    bool ok = ComputeCurrent(current, count, &detail);
    if (detail != g_lastDetail)
    {
        g_lastDetail = detail;
        WriteLog(L"page: " + detail);
    }

    HKEY k = nullptr;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\PptTouchClick\\Slide", 0, nullptr, 0,
                        KEY_WRITE, nullptr, &k, nullptr) != ERROR_SUCCESS)
        return;
    DWORD c = (DWORD)(ok ? current : -1);
    DWORD n = (DWORD)count;
    DWORD age = 0;
    RegSetValueExW(k, L"Current", 0, REG_DWORD, reinterpret_cast<const BYTE*>(&c), sizeof(c));
    RegSetValueExW(k, L"Count", 0, REG_DWORD, reinterpret_cast<const BYTE*>(&n), sizeof(n));
    RegSetValueExW(k, L"AgeMs", 0, REG_DWORD, reinterpret_cast<const BYTE*>(&age), sizeof(age));
    RegCloseKey(k);
}

static void WatchLoop()
{
    while (g_watchRun.load())
    {
        PublishSlide();
        for (int i = 0; i < 8 && g_watchRun.load(); ++i)
            std::this_thread::sleep_for(std::chrono::milliseconds(100));
    }
}

// ---------------------------------------------------------------------------
// 最小 IDispatch：Start() / GetStatus() / Stop()
// ---------------------------------------------------------------------------
class SlideWatcher : public IDispatch
{
public:
    SlideWatcher() : m_ref(1) {}

    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch))
        {
            *ppv = static_cast<IDispatch*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef() override { return ++m_ref; }
    STDMETHODIMP_(ULONG) Release() override
    {
        ULONG r = --m_ref;
        if (r == 0) delete this;
        return r;
    }

    STDMETHODIMP GetTypeInfoCount(UINT* p) override { if (p) *p = 0; return S_OK; }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo** pp) override { if (pp) *pp = nullptr; return E_NOTIMPL; }

    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* names, UINT count, LCID, DISPID* ids) override
    {
        if (!names || !ids || count != 1) return E_INVALIDARG;
        std::wstring n = ToLower(names[0]);
        if (n == L"start") ids[0] = 1;
        else if (n == L"getstatus") ids[0] = 2;
        else if (n == L"stop") ids[0] = 3;
        else return DISP_E_UNKNOWNNAME;
        return S_OK;
    }

    STDMETHODIMP Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS*, VARIANT* result,
                        EXCEPINFO*, UINT*) override
    {
        if (result) VariantInit(result);
        switch (id)
        {
        case 1:   // Start
        {
            InstallHook();   // 失败也不影响页码监视
            if (!g_watchRun.load())
            {
                g_watchRun = true;
                if (g_watchThread.joinable()) g_watchThread.join();
                g_watchThread = std::thread(WatchLoop);
            }
            WriteLog(std::wstring(L"Start(): hook=") + (g_hookOk.load() ? L"on" : L"off"));
            if (result) { result->vt = VT_BOOL; result->boolVal = VARIANT_TRUE; }
            return S_OK;
        }
        case 2:   // GetStatus
        {
            int current = -1, count = 0;
            std::wstring detail;
            ComputeCurrent(current, count, &detail);
            std::wstring s = L"ver=";
            s += PPTSLIDEPROC_VERSION;
            s += L" hook=";
            if (!g_hookOk.load()) InstallHook();            // 首次失败也没关系，这里再试一次
            s += g_hookOk.load() ? L"on" : (g_hookFailed.load() ? L"unavailable" : L"off");
            s += L" swallowed=" + std::to_wstring(g_swallowed.load());
            s += L" slide=" + std::to_wstring(current + 1) + L"/" + std::to_wstring(count);
            s += L" [" + detail + L"]";
            if (result) { result->vt = VT_BSTR; result->bstrVal = ::SysAllocString(s.c_str()); }
            return S_OK;
        }
        case 3:   // Stop
            g_watchRun = false;
            RemoveHook();
            if (result) { result->vt = VT_BOOL; result->boolVal = VARIANT_TRUE; }
            return S_OK;
        default:
            return DISP_E_MEMBERNOTFOUND;
        }
    }

private:
    std::atomic<ULONG> m_ref;
};

// ---------------------------------------------------------------------------
// 类工厂 / DLL 导出
// ---------------------------------------------------------------------------
static LONG g_serverLock = 0;
static std::atomic<long> g_objectCount{ 0 };

class ClassFactory : public IClassFactory
{
public:
    ClassFactory() : m_ref(1) {}
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory))
        {
            *ppv = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef() override { return ++m_ref; }
    STDMETHODIMP_(ULONG) Release() override
    {
        ULONG r = --m_ref;
        if (r == 0) delete this;
        return r;
    }
    STDMETHODIMP CreateInstance(IUnknown* outer, REFIID riid, void** ppv) override
    {
        if (outer) return CLASS_E_NOAGGREGATION;
        auto* w = new SlideWatcher();
        HRESULT hr = w->QueryInterface(riid, ppv);
        w->Release();
        if (SUCCEEDED(hr)) ++g_objectCount;
        return hr;
    }
    STDMETHODIMP LockServer(BOOL lock) override
    {
        if (lock) InterlockedIncrement(&g_serverLock);
        else InterlockedDecrement(&g_serverLock);
        return S_OK;
    }

private:
    std::atomic<ULONG> m_ref;
};

extern "C" __declspec(dllexport) HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    if (!ppv) return E_POINTER;
    if (!IsEqualCLSID(rclsid, kClsidSlideWatcher)) return CLASS_E_CLASSNOTAVAILABLE;
    auto* f = new ClassFactory();
    HRESULT hr = f->QueryInterface(riid, ppv);
    f->Release();
    return hr;
}

extern "C" __declspec(dllexport) HRESULT __stdcall DllCanUnloadNow()
{
    return (g_objectCount.load() == 0 && g_serverLock == 0) ? S_OK : S_FALSE;
}

static std::wstring ModulePath()
{
    wchar_t path[MAX_PATH] = {};
    DWORD n = GetModuleFileNameW(nullptr, path, MAX_PATH);
    return std::wstring(path, n);
}

static LONG RegSetSz(HKEY root, const std::wstring& sub, const std::wstring& name, const std::wstring& value)
{
    HKEY k = nullptr;
    LONG r = RegCreateKeyExW(root, sub.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &k, nullptr);
    if (r != ERROR_SUCCESS) return r;
    r = RegSetValueExW(k, name.c_str(), 0, REG_SZ,
                       reinterpret_cast<const BYTE*>(value.c_str()),
                       (DWORD)((value.size() + 1) * sizeof(wchar_t)));
    RegCloseKey(k);
    return r;
}

/// 注册 COM 类：HKCU\Software\Classes\CLSID\{...}（按当前用户，无需管理员；安装脚本会用管理员跑一次）
extern "C" __declspec(dllexport) HRESULT __stdcall DllRegisterServer()
{
    const std::wstring mod = ModulePath();
    const std::wstring clsidKey = L"Software\\Classes\\CLSID\\{2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21}";
    const std::wstring progIdKey = L"Software\\Classes\\PptTouchClick.SlideWatcher";

    RegSetSz(HKEY_CURRENT_USER, clsidKey, L"", L"PptTouchClick SlideWatcher");
    RegSetSz(HKEY_CURRENT_USER, clsidKey + L"\\InprocServer32", L"", mod);
    RegSetSz(HKEY_CURRENT_USER, clsidKey + L"\\InprocServer32", L"ThreadingModel", L"Both");
    RegSetSz(HKEY_CURRENT_USER, clsidKey + L"\\ProgID", L"", L"PptTouchClick.SlideWatcher");
    RegSetSz(HKEY_CURRENT_USER, clsidKey + L"\\TypeLib", L"", L"{7D2E1B44-03C5-4F88-B172-9E4D1860C29F}");
    RegSetSz(HKEY_CURRENT_USER, clsidKey + L"\\TypeLib", L"Version", L"1.0");

    RegSetSz(HKEY_CURRENT_USER, progIdKey, L"", L"PptTouchClick SlideWatcher");
    RegSetSz(HKEY_CURRENT_USER, progIdKey + L"\\CLSID", L"", L"{2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21}");

    // 类型库（可选；失败也不影响 CreateObject，C# 走 IDispatch 即可）
    bool typelibOk = false;
    HMODULE self = GetModuleHandleW(nullptr);
    HRSRC res = FindResourceW(self, MAKEINTRESOURCEW(1), RT_TYPELIB);
    if (res)
    {
        HGLOBAL mem = LoadResource(self, res);
        DWORD size = SizeofResource(self, res);
        void* data = LockResource(mem);
        if (data && size)
        {
            ITypeLib* tlb = nullptr;
            if (SUCCEEDED(LoadTypeLibFromResource(data, size, &tlb)) && tlb)
            {
                std::wstring tlbPath = mod.substr(0, mod.find_last_of(L"\\/") + 1) + L"PptSlideProc.tlb";
                HRESULT hr = RegisterTypeLibForUser(tlb, const_cast<LPOLESTR>(tlbPath.c_str()),
                                                   const_cast<LPOLESTR>(L"PptSlideProc 1.0 Type Library"));
                typelibOk = SUCCEEDED(hr);
                tlb->Release();
            }
        }
    }
    WriteLog(std::wstring(L"DllRegisterServer ") + mod + (typelibOk ? L" (+typelib)" : L""));
    return S_OK;
}

static LONG RegDeleteTree(HKEY root, const std::wstring& sub)
{
    // 逐层删除（值 + 子键）
    HKEY k = nullptr;
    if (RegOpenKeyExW(root, sub.c_str(), 0, KEY_READ | KEY_WRITE, &k) != ERROR_SUCCESS) return ERROR_SUCCESS;
    std::vector<std::wstring> kids;
    wchar_t nameBuf[512];
    DWORD idx = 0;
    for (;;)
    {
        DWORD len = 512;
        if (RegEnumKeyExW(k, idx, nameBuf, &len, nullptr, nullptr, nullptr, nullptr) != ERROR_SUCCESS) break;
        kids.emplace_back(nameBuf, len);
        ++idx;
    }
    RegCloseKey(k);
    for (auto& kid : kids) RegDeleteTree(root, sub + L"\\" + kid);
    RegDeleteKeyW(root, sub.c_str());
    return ERROR_SUCCESS;
}

extern "C" __declspec(dllexport) HRESULT __stdcall DllUnregisterServer()
{
    RegDeleteTree(HKEY_CURRENT_USER, L"Software\\Classes\\CLSID\\{2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21}");
    RegDeleteTree(HKEY_CURRENT_USER, L"Software\\Classes\\PptTouchClick.SlideWatcher");
    HKEY k = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Classes\\Interface\\{7D2E1B44-03C5-4F88-B172-9E4D1860C29F}",
                      0, KEY_WRITE, &k) == ERROR_SUCCESS)
    {
        RegCloseKey(k);
        RegDeleteTree(HKEY_CURRENT_USER, L"Software\\Classes\\Interface\\{7D2E1B44-03C5-4F88-B172-9E4D1860C29F}");
    }
    WriteLog(L"DllUnregisterServer");
    return S_OK;
}

extern "C" __declspec(dllexport) BOOL WINAPI SlideProcGetCurrent(int* current0, int* total)
{
    int current = -1, count = 0;
    bool ok = ComputeCurrent(current, count, nullptr);
    if (current0) *current0 = ok ? current : -1;
    if (total) *total = count;
    return ok ? TRUE : FALSE;
}

extern "C" __declspec(dllexport) DWORD WINAPI SlideProcRunSelfTest()
{
    // 供 CI 用：能加载、能取类对象、能算一次页码即视为通过
    IClassFactory* cf = nullptr;
    HRESULT hr = DllGetClassObject(kClsidSlideWatcher, IID_IClassFactory, reinterpret_cast<void**>(&cf));
    if (FAILED(hr) || !cf) return 1;
    IDispatch* d = nullptr;
    hr = cf->CreateInstance(nullptr, IID_IDispatch, reinterpret_cast<void**>(&d));
    cf->Release();
    if (FAILED(hr) || !d) return 2;
    DISPID id = 0;
    LPOLESTR name = const_cast<LPOLESTR>(L"GetStatus");
    hr = d->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) { d->Release(); return 3; }
    VARIANT res;
    VariantInit(&res);
    DISPPARAMS dp = {};
    hr = d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &res, nullptr, nullptr);
    bool ok = SUCCEEDED(hr) && res.vt == VT_BSTR;
    if (ok) WriteLog(std::wstring(L"self-test: ") + res.bstrVal);
    VariantClear(&res);
    d->Release();
    return ok ? 0 : 4;
}

extern "C" BOOL WINAPI DllMain(HINSTANCE h, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        DisableThreadLibraryCalls(h);
        WriteLog(L"PptSlideProc loaded " PPTSLIDEPROC_VERSION);
    }
    else if (reason == DLL_PROCESS_DETACH)
    {
        g_watchRun = false;
        if (g_watchThread.joinable()) g_watchThread.detach();
        if (g_pHook) { UnhookWindowsHookEx(g_pHook); g_pHook = nullptr; }
    }
    return TRUE;
}
