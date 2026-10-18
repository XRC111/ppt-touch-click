# PPT 触控单击翻页（PptTouchClick）

让 **Windows 10 上的 PowerPoint 放映**回到“**手指轻点 = 下一页**”，并且：

- 点 **空白处** → 下一页（点动画/视频等会先走“下一步动画”）
- 点 **文字、图片、按钮、超链接** → **不拦截**，交给 PowerPoint 自己处理
- **滑动**不再被当成换页手势；**长按**也不会误翻页
- 纯触摸、没插鼠标也能用（不依赖 PowerPoint 自己是否把设备当成“有鼠标”）

---

## 背景：为什么 PowerPoint 自己做不到

Windows 10 上的 PowerPoint 会**自动判断**当前是“触摸设备”还是“鼠标设备”：

- 一旦认定是纯触摸设备，放映时就切到**触摸模式**：换页方式从“单击”变成“手指滑动”；
- 这个行为没有开关。微软官方答复里明确说过：**PowerPoint 不提供“选择翻页方式”的设置**，也无法单独关掉滑动换页 / 双指缩放
  （[官方答复 1](https://learn.microsoft.com/zh-cn/answers/questions/5403153/win10-ppt)、[官方答复 2](https://learn.microsoft.com/en-us/answers/questions/5143263/how-can-i-disable-advance-to-next-slide-on-swipe)、[官方答复 3](https://learn.microsoft.com/en-us/answers/questions/5373818/cannot-disable-the-double-tap-zoom-or-pinch-zoom-t)）。

所以只能从**外面补一层**：检测到 PowerPoint 全屏放映时，把“点空白处”翻译成一次翻页；同时拦住滑动。这正是本项目做的事。

---

## 组成

| 文件 | 说明 |
| --- | --- |
| `PptTouchClick.exe` | 主程序：桌面代理，托盘里出现黄色图标；不是 PowerPoint 加载项，重启 PPT 不受影响 |
| `PptSlideProc.x64.dll` | 原生 COM 组件（64 位）：安装**全局触摸钩子**（放映时吞掉滑动/长按），并读取 PowerPoint “最近访问幻灯片”给出当前页 |
| `PptSlideProc.ocx` | 同上，32 位进程内 COM 版本（仅 32 位 Windows 使用；64 位系统不注册） |
| `PptSlideProc.tlb` | 类型库（注册用，已内嵌在 DLL 资源里） |
| `install.cmd` / `uninstall.cmd` | 安装 / 卸载（需管理员） |
| `PptTouchClick.log`、`PptSlideProc.log` | 运行日志（排错用） |

> **重要**：主程序**不依赖**原生组件也能“点空白 → 翻页”。
> 组件只额外提供两件事：① 滑动/长按过滤；② 放映时精确知道当前是第几页（换页即重新识别页面元素）。
> 组件没装好时，程序会自动降级：按“页面变化”重扫，偶尔慢半拍，但不会崩。

> 页码来源：PowerPoint 放映时会把“最近访问过的幻灯片”顺序写进注册表 `HKCU\Software\Microsoft\Office\16.0\PowerPoint\Document MRU\<文件>`。
> 组件把它换算成“当前是第几页”，主程序据此判断“这一页上哪些地方是内容、哪些地方是空白”。

---

## 安装

1. 解压全部文件到同一个文件夹（例如 `D:\Tools\PptTouchClick`，**别放桌面**）。
2. **右键 `install.cmd` → 以管理员身份运行**（注册 COM 组件需要管理员）。脚本会按系统位数自动选择 x64 组件或 32 位 ocx（`HKCU\Software\Classes\CLSID` 不区分 32/64 位，两个同时注册会互相覆盖）。
3. 运行后托盘里出现黄色 ▶ 图标，右键可：
   - 查看状态（放映中显示 `放映中 3/12`）
   - 暂停/恢复翻页代理
   - 打开日志
   - 设为开机自动启动
   - 退出

卸载：右键 `uninstall.cmd` → 以管理员身份运行。

> 如果只是自己临时用，也可以**直接双击 `PptTouchClick.exe`**（不需要管理员）：
> 这时触摸钩子可能装不上（`PptSlideProc` 未注册），但“点击空白 → 翻页”仍然可用，
> 只是滑动/长按不会被拦住，且换页后需要 0.3 秒左右重扫页面。

---

## 使用

1. 打开 PPT，**F5** 开始放映（建议用“演讲者放映”）。
2. 手指点屏幕空白处 → 下一页；点文字/图片/按钮 → 不翻页。
3. 用键盘、翻页笔、鼠标翻页时，代理会自动让路 1.2 秒，不会打扰。
4. 想翻页但屏幕上几乎没有空白（整页都是图），可以点页面四周留白，或临时用翻页笔。

---

## 参数与调优

命令���参数：

| 参数 | 说明 |
| --- | --- |
| `--mouse-only` | 关闭触摸注入，只用 PageDown 翻页（某些机器上注入无效时用） |
| `--verbose` | 保留（当前版本已默认记录日志） |
| `--install-autostart` / `--remove-autostart` | 写入/删除开机启动项（安装脚本内部使用） |

注册表可调（`HKEY_CURRENT_USER\Software\PptTouchClick`，DWORD）：

| 名称 | 默认 | 含义 |
| --- | --- | --- |
| `MoveTol` | 26 | 触摸抬起时允许的最大位移(px)，超过即视为“滑动”并吞掉 |
| `TapMs` | 1500 | 按住超过这么久不算单击 |
| `DebounceMs` | 800 | 同一窗口、同一页、同一键位的两次点击最短间隔（防止一次触摸翻两页） |
| `MarkMs` | 1200 | 翻页后屏蔽该点周围的时长 |
| `AdvanceDelayMs` | 160 | 让 PowerPoint 自己处理触摸点击后的等待确认时间 |

改完重启 `PptTouchClick.exe` 生效（托盘菜单“重新载入设置”只重载程序内参数）。

---

## 工作方式（原理）

1. **找到放映窗口**：PowerPoint 全屏放映窗口的类名是 `screenClass`、无标题、铺满一块显示器。代理每 0.3 秒扫描一次。
2. **判断“是不是空白”**：对该窗口用 `PrintWindow` 离屏截图 → 求“非黑区域”的外接矩形作为幻灯片内容区（自动排除左右两侧的放映控制条）→ 在内容区里做局部对比度分析，找出文字/图片/图表/按钮的包围盒，按页缓存。点击落在这些包围盒里就放行，否则视为空白。
3. **翻页**：先向放映窗口注入一次**真实触摸点击**（`InitializeTouchInjection`），PowerPoint 在触摸模式下会把它当作“前进”，与它自己的换页逻辑一致；若 160ms 内没有翻页，再补发 `PageDown` 兜底。
4. **防误触**：全局触摸钩子（`WH_POINTER`）在原生组件里，滑动和长按直接吞掉；主程序侧还有“同点去抖”和“键盘让路”。

---

## 已知限制

- 只有 **PowerPoint 桌面版**生效（WPS、PowerPoint Online 不适用）。
- “空白”是按画面识别出来的：**整页背景图且没有留白**的页、或页面上大量装饰图形，可能被整体当作内容而不翻页。此时点页面四周留白或改用翻页笔。
- 按钮/超链接的“保护”是**按位置**生效的：点在按钮上交给 PPT 处理，不会翻页；如果按钮本身没有设置任何动作，那么点了也不会翻页。
- 需要 Windows 8/10 及以上（触摸注入 API 自 Windows 8 起提供）；Windows 7 不可用。
- **高 DPI（125% / 150% / 200%）已处理**：程序声明了 PerMonitorV2 DPI 感知，点击坐标与截图坐标在同一坐标系。
- 触摸过滤**只在全屏放映时生效**；不开放映时不会影响系统的触摸滚动和双指缩放。
- 组件按**当前用户**注册（`HKCU\Software\Classes`）；换了 Windows 账户需要重新安装。
- 自检/验证已在 GitHub Actions（Windows Server 2022，有触摸 API）通过；**真机触摸屏上的“点空白翻页”手感需要你实机确认**，
  如果不跟手，优先调 `AdvanceDelayMs`（默认 160ms）和 `DebounceMs`（默认 800ms）。

---

## 编译 / 发布

**推荐：推送即编译。** `.github/workflows/build.yml` 在 `windows-latest` 上：

1. 用系统自带的 `csc.exe`（.NET Framework 4.8）编 `PptTouchClick.exe`；
2. 用 `tlbexp.exe` 从桩程序集 `src/PptSlideProc.TlbGen.dll` 生成 `PptSlideProc.tlb`（仓库里也放了一份预生成的，生成失败时自动回退）；
3. 用 MSVC 分别编 x64 / x86 的 `PptSlideProc.x64.dll` 与 `PptSlideProc.ocx`；
4. 自检：直接调用 `DllGetClassObject`（`src/ComProbe.cs`）验证 COM 工厂可用，再走 `regsvr32 /i:user` 验证 `DllRegisterServer`；
5. 打包成 `PptTouchClick.zip` 作为构建产物上传。

**每次 main 构建成功，runner 会自动打 tag（`v1.0.<构建号>`）并发布 Release，无需任何本地操作。** 若想手动发特定版本，也可以打一个 `v*` 标签：

```bat
git tag v1.0.0 && git push origin v1.0.0
```

**本地只编 exe**（没有 Visual Studio 也能跑，只是没有滑动过滤和页码组件）：

```powershell
.\build.ps1        # 输出 dist\PptTouchClick.exe
```

**本地生成类型库**：`.\gen-tlb.ps1`（或 Git Bash 下的 `./gen-tlb.sh`）。

> 坑记录：CI 里用 PowerShell 的 `New-Object -ComObject` 创建本组件会报 `CO_E_ERRORINDLL`，
> 原因是 CoreCLR（PowerShell 7 / .NET 平台）对只有 IDispatch、没有 TypeLib 的进程内组件会拒绝 `CoCreateInstance`；
> 同一个组件在桌面版 PowerShell 5.1、VBScript、C# 里都能正常创建。CI 因此改为“直接调用导出函数 + regsvr32”来验证。

---

## 本次修复

- 页面形状缓存失效：之前每次刷新都新建状态对象，导致放映期间每 300ms 全屏重扫一次（CPU 高、日志暴涨、去抖被清空）；现在按页缓存，只有换页/窗口切换才重扫
- 右键/中键点击也会触发翻页；现在只响应左键/单指
- 触摸过滤在非放映状态下也全局生效，会吞掉系统的滑动/缩放；现在只在放映时接管
- 原生触摸钩子装在无消息循环的线程上，回调永远收不到；改为在该线程上跑消息泵
- 高 DPI 下点击坐标与截图坐标错位；增加 PerMonitorV2 DPI 感知（manifest + API 兜底）
- `install.cmd` 用 64 位 regsvr32 注册 32 位 ocx 必然失败；改为按系统位数选择
- 页码 `AgeMs` 恒为 0，过期判断失效；现在按"距上次成功读取的时长"写入
- CI 编译未加 `/utf-8`，原生组件中文日志乱码；已修正
- 托盘图标每次状态变化都新建且不释放（GDI 泄漏）；改为复用同一个图标，并在退出时释放
- 托盘更新从工作线程直接改控件（跨线程）；改为切回 UI 线程
- 工作线程处理点击无 try/catch（异常会直接杀进程）；已加保护
- 页码未知时"同页去抖"会吃掉快速连点；现在只在知道页码时去抖
- 其他：触摸状态泄漏、日志仅启动时裁剪、COM 对象计数不减、卸载时线程未 join、`DllMain` 里做文件 I/O、卸载删错 Interface 键、右边界扫描越界等

## 许可

个人/教学用途自用脚本，按现状提供，不做任何担保。
