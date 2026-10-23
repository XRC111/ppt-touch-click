<div align="center">

# PPT 触控单击翻页（PptTouchClick）

**让 Windows 10 上的 PowerPoint 放映回到「手指轻点 = 下一页」**

点空白处翻页，点文字/图片/按钮不拦截，滑动和长按也不会误翻。

[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D4?style=flat-square)](#系统要求)
[![.NET](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?style=flat-square)](#从源码构建)
[![License](https://img.shields.io/badge/License-MIT-blue?style=flat-square)](./LICENSE)

</div>

---

## 为什么需要它

Windows 10 上的 PowerPoint 会自动判断设备类型：一旦认定是纯触摸设备，放映时就切到**触摸模式**——换页方式从「单击」变成「手指滑动」，而且没有开关可以改。微软官方答复明确表示不提供「选择翻页方式」的设置，也无法单独关闭滑动换页或双指缩放。

既然改不了 PowerPoint，就从外面补一层：检测到全屏放映时，把「点空白处」翻译成一次翻页；手指滑动、长按则不计为单击。

## 特性

| 能力 | 说明 |
| --- | --- |
| 点空白翻页 | 点页面空白处 → 下一页；点在文字/图片/按钮上 → 交给 PPT 处理 |
| 触摸过滤 | 放映时滑动与长按不算单击（不会多翻一页） |
| 纯触摸可用 | 不依赖系统是否把设备识别为「有鼠标」 |
| 键盘让路 | 用键盘/翻页笔翻页时自动避让 1.2 秒 |
| 降级可用 | 原生组件未注册也能用：滑动过滤在主程序里，缺的只是“精确页码” |
| 托盘常驻 | 状态、暂停、日志、开机自启 |

## 快速开始

1. 解压全部文件到同一目录（例如 `D:\Tools\PptTouchClick`，别放桌面）
2. **双击 `install.cmd`**（注册写 HKCU，不需要管理员；直接双击 exe 也能用，程序会自行注册）
3. 托盘出现黄色 ▶ 图标，右键可查看状态、暂停、打开日志、设置开机自启
4. 打开 PPT，F5 放映；手指点空白处翻页

卸载：双击 `uninstall.cmd`。

> 也可以直接双击 `PptTouchClick.exe`：组件会在首次启动时就地注册到 HKCU；即使组件缺失，「点空白 → 翻页」也照常工作，只是换页后要重扫页面（约 0.3 秒）。

### 系统要求

- Windows 8 / 10 / 11（触摸注入 API 自 Windows 8 起提供）
- .NET Framework 4.8（Win10 自带）
- 仅 PowerPoint 桌面版；WPS、PowerPoint Online 不适用

## 组成

| 文件 | 说明 |
| --- | --- |
| `PptTouchClick.exe` | 主程序：桌面代理，托盘常驻，不是 PPT 加载项 |
| `PptSlideProc.x64.dll` | 原生 COM 组件（64 位）：读取 PowerPoint 当前页码 |
| `PptSlideProc.ocx` | 同上，32 位版本（仅 32 位 Windows 使用） |
| `PptSlideProc.tlb` | 类型库（已内嵌在 DLL 资源中） |
| `install.cmd` / `uninstall.cmd` | 安装 / 卸载（按用户注册，无需管理员） |
| `PptTouchClick.log` / `PptSlideProc.log` | 运行日志 |

> 主程序**不依赖**原生组件也能完成「点空白 → 翻页」。组件只额外提供一件事：放映时精确知道当前第几页（换页即重扫页面元素）。滑动/长按过滤在主程序自己的低层鼠标钩子里完成。

## 工作原理

1. **找到放映窗口**：PowerPoint 全屏放映窗口类名为 `screenClass`、无标题、铺满显示器，每 0.3 秒扫描一次
2. **判断是否空白**：`PrintWindow` 离屏截图 → 求非黑区域外接矩形作为内容区（排除左右控制条）→ 内容区内做局部对比度分析，得到文字/图片/图表的包围盒，按页缓存
3. **翻页**：先向放映窗口注入一次真实触摸点击；确认页码变化后即完成；否则补发 `PageDown` 兜底
4. **防误触**：低层鼠标钩子按 `dwExtraInfo` 区分手指与鼠标，触摸时位移超过 `MoveTol`（滑动）或按住超过 `TapMs`（长按）不算单击；另有同点去抖、键盘让路、翻页后屏蔽

页码来源：PowerPoint 放映时会把「最近访问的幻灯片」写进 `HKCU\Software\Microsoft\Office\16.0\PowerPoint\Document MRU\<文件>`，原生组件据此换算出当前页。

## 参数调优

命令行参数：

| 参数 | 说明 |
| --- | --- |
| `--mouse-only` | 关闭触摸注入，只用 PageDown 翻页 |
| `--verbose` | 记录每页扫描的详细日志 |
| `--install-autostart` / `--remove-autostart` | 安装脚本内部使用 |

注册表 `HKEY_CURRENT_USER\Software\PptTouchClick`（DWORD）：

| 名称 | 默认 | 含义 |
| --- | --- | --- |
| `MoveTol` | 26 | 触摸抬起允许的最大位移(px)，超过视为滑动并吞掉 |
| `TapMs` | 1500 | 按住超过此值不算单击 |
| `DebounceMs` | 800 | 同页同键位两次点击的最短间隔 |
| `MarkMs` | 1200 | 翻页后屏蔽该点周围的时长 |
| `AdvanceDelayMs` | 160 | 注入触摸后等待确认的时间 |

改完重启主程序生效。

## 排障

| 现象 | 原因与处理 |
| --- | --- |
| 托盘显示「组件不可用」 | 看日志里的 `com:` 行：组件缺失就重新解压；注册失败就运行 `install.cmd` |
| 点空白不翻页 | 看 `PptTouchClick.log` 是否识别到放映窗口；整页无留白时点四周 |
| 点按钮也被翻页 | 该页扫描未框住按钮（装饰图形过多），可调 `MarkMs` 或改用翻页笔 |
| 一次点击翻两页 | 调大 `DebounceMs`；确认没有同时运行其他翻页工具 |
| 非放映时系统触摸异常 | 确认是当前版本（触摸过滤只在放映时生效） |
| 高缩放下判定错位 | 确认使用带 `app.manifest` 的新版本 |

## 已知限制

- 只有 PowerPoint 桌面版生效
- 「空白」是画面识别结果：整页背景图且无留白的页可能被整体当作内容
- 按钮保护是按位置生效：点在按钮上不会翻页，但按钮本身无动作时也不会翻页
- 组件按当前用户注册（HKCU），换 Windows 账户需重新安装

## 从源码构建

```powershell
# 仅编译 exe（无需 VS）
.\build.ps1

# 原生组件由 GitHub Actions 产出（MSVC x64/x86 + COM 自检）
```

推送 main 自动构建并发布 Release；也可手动打 tag：`git tag v1.0.0 && git push origin v1.0.0`。

## License

MIT（见 [LICENSE](./LICENSE)）。本项目按现状提供，不做任何担保。
