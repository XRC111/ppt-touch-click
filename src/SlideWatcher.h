// PptSlideProc - PowerPoint 放映页码 / 触摸手势 原生代理（为 PptTouchClick 服务）
// 输出：PptSlideProc.ocx（32 位 COM 组件，进程内）与 PptSlideProc.x64.dll（64 位，进程内）
//
// 对外：
//   - COM 类 PptTouchClick.SlideWatcher（ThreadingModel=Both）
//       Start()       -> bool   安装全局 WH_POINTER 钩子（吞掉滑动/长按），开始监视当前页
//       GetStatus()   -> string 形如 "hook=on slide=3/12 total=12"
//       Stop()        -> bool
//   - 导出 SlideProcGetCurrent(int* current1based, int* total)
//       current1based = 放映/当前页（1 起）；没有可用的 PowerPoint MRU 时返回 0
//
// 为什么需要它：C# 托管进程调用 SetWindowsHookEx(WH_POINTER) 时 hMod 为 NULL，钩子装不上；
// 由本原生模块安装即可，同时把"当前页"从 PowerPoint 的 MRU 注册表里算出来，C# 端零解析。

#pragma once

#include <windows.h>
#include <cstdint>

#define PPTSLIDEPROC_VERSION L"1.0.0"

// CLSID: 2F4C7A19-8B63-4E21-9A6D-1C3B550E7A21
#define CLSID_SlideWatcher { 0x2F4C7A19, 0x8B63, 0x4E21, { 0x9A, 0x6D, 0x1C, 0x3B, 0x55, 0x0E, 0x7A, 0x21 } }
