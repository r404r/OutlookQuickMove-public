<p align="right">
  <a href="README.md">English</a> | <strong>简体中文</strong>
</p>

# OutlookQuickMove

一款键盘优先的 VSTO 加载项，适用于 **Windows 版经典 Outlook**。它可以把选中的邮件
移动到你通过输入找到的文件夹，无需拖拽，也无需展开文件夹树。

## 功能特性

- **Quick Move 对话框**：输入文字即可按完整路径筛选文件夹，使用 `Up`/`Down` 选择，
  按 `Enter` 移动。支持一封或多封已选邮件。
- **Go to Folder**：搜索同一批文件夹，并把当前 Outlook 窗口切换到选中的文件夹，
  不移动任何邮件。
- **常用文件夹优先**：最常使用的目标文件夹会浮到顶部，让常见移动操作只需一次按键
  就能完成（搜索框为空时会预选最常用的文件夹）。
- **移动前标记为已读**：可选，并且会在重启后记住你的选择。
- **按数据文件过滤**：可选择要搜索哪些 Outlook 数据文件（stores）。
- **安全设计**：不会删除邮件；会跳过非邮件项目；会跳过移动到同一文件夹的操作；
  会确定性地释放 Outlook COM 对象；遇到问题会记录日志，而不是静默失败。

## 环境要求

- Windows，安装 **经典 Win32 Outlook**（Microsoft 365 / 2016+；不支持“新 Outlook”
  或网页版 Outlook）。开发环境为 `Microsoft Outlook for Microsoft 365 MSO (Version 2605 Build
  16.0.20026.20076) 64-bit`。
- **.NET Framework 4.8**。
- 构建需要：安装带有 *Office/SharePoint development* 工作负载（VSTO）的
  **Visual Studio 2022**。请参见 `DEV_SETUP.md`。

## 构建

- 用 Visual Studio 打开 `OutlookQuickMove.sln` 并构建。按 **F5** 会启动 Outlook，
  并加载该加载项用于调试。
- 或者从命令行构建：`msbuild OutlookQuickMove.sln /p:Configuration=Debug`。

ClickOnce 清单签名默认 **关闭**（`SignManifests` 为 `false`），仓库也不包含签名密钥，
因此干净克隆后无需额外设置即可构建。在分发到其他机器之前，请提供你自己的代码签名
证书，并在 `OutlookQuickMove/OutlookQuickMove.csproj` 中设置 `SignManifests` /
`ManifestKeyFile`（以及 thumbprint）。请勿提交私钥（`*.pfx` 已加入 git ignore）。

## 安装

请先构建项目，让 `.vsto` 部署清单出现在输出目录中。

Debug 构建清单：

```powershell
C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto
```

使用带 `file:///` URI 的 `VSTOInstaller.exe`。这样可以避免本地路径中包含换行、复制进来的
提示符标记或其他隐藏字符时出现路径解析问题。

```powershell
$installer = Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'
$vsto = 'C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto'
$vstoUri = ([System.Uri](Get-Item -LiteralPath $vsto).FullName).AbsoluteUri

& $installer /Install $vstoUri
```

不要把 `.vsto` 路径拆成多行放在带引号的字符串里。

安装完成后，重启 Outlook。加载项会创建自己的顶层功能区选项卡：

`Quick Move` 功能区选项卡 -> `Actions` 组

可用按钮：

- `Quick Move`：搜索并移动选中的邮件项目。
- `Go to Folder`：搜索文件夹并把当前 Outlook 窗口切换到该文件夹（只导航，不移动邮件）。
  它使用与 Quick Move 相同的数据文件选择和输入筛选对话框。
- `Undo Quick Move...`：打开最近移动记录清单，把选中的项目移回原文件夹。仅当存在
  可撤销的移动历史时启用。
- `Settings`：一个带选项卡的对话框，包含 `Data Files`（搜索哪些 Outlook 数据文件）、
  `Frequent Folders`（记忆目标文件夹的数量上限和列表）和 `Undo History`（记忆移动记录
  的数量上限和清空选项）。

在 Quick Move 对话框中，输入文字筛选文件夹，使用 `Up` / `Down` 切换高亮候选项，
按 `Enter` 确认。

## 跳转到文件夹

`Go to Folder` 复用 Quick Move 的文件夹选择器，但执行的是导航而不是移动：点击按钮，
输入文字筛选，使用 `Up` / `Down` 选择，然后按 `Enter`（或点击 `Go`）即可把当前
Outlook 窗口切换到选中的文件夹。它搜索的文件夹范围与 Quick Move 相同（保存的 Data Files
选择），并会把最常用的 Quick Move 目标文件夹置顶以便快速访问；但它不会修改任何邮件，
也不会记录一次移动，因此不会影响常用文件夹统计或撤销历史。

## 常用文件夹

Quick Move 会记住你最常移动到的文件夹，并把它们置于搜索结果顶部（按使用次数从高到低），
因此常用目标可以一次按键直达；当搜索框为空时，最常用的文件夹已经被选中。每次确认移动
都会更新使用次数。系统会跟踪你使用过的每一个文件夹，所以新使用的文件夹会随着次数累积
逐步上升，而不会因为列表已满就永远排除在外。

在 `Settings` -> `Frequent Folders` 中，你可以设置 Quick Move 列表顶部要**显示**多少个
文件夹（0-100，默认 20；`0` 会关闭此功能），查看所有按使用次数排序的已跟踪文件夹，
以及删除或清空条目。数据会在 Outlook 重启后保留，存储在：

```powershell
$env:APPDATA\OutlookQuickMove\frequent-targets.txt
$env:APPDATA\OutlookQuickMove\frequent-targets-max.txt
```

## 撤销移动

Outlook 原生的 `Ctrl+Z` 无法撤销由加载项执行的移动操作（对象模型没有 API 可以写入
Outlook 的撤销栈），因此 Quick Move 会维护自己的移动历史。每次移动都会记录每个项目来自
哪里、去了哪里，以及原始已读/未读状态。

点击 `Undo Quick Move...` 会打开最近移动记录清单（最新在前）。最近一次操作中的项目会默认
勾选，适合常见的“撤销刚才那次移动”场景；你也可以跨多次操作自由勾选或取消勾选，然后点击
`Undo Selected` 将它们移回原文件夹（同时恢复已读/未读状态）。如果某个项目已经被删除或再次
移动导致无法找到，它会被报告并从列表中移除；其他失败会保留在历史中，便于稍后重试。
`Clear History` 会清空列表，但不会移动任何邮件。

在 `Settings` -> `Undo History` 中，你可以设置记住多少条最近移动记录（0-500，默认 50；
`0` 会关闭记录，但保留现有历史直到清空），也可以清空历史。数据会在 Outlook 重启后保留，
存储在：

```powershell
$env:APPDATA\OutlookQuickMove\undo-history.txt
$env:APPDATA\OutlookQuickMove\undo-max.txt
```

如果尚未保存 store 过滤设置，Quick Move 会搜索所有 Outlook 数据文件中的文件夹。
Settings 对话框支持选择多个数据文件。

选中的数据文件会持久化到：

```powershell
$env:APPDATA\OutlookQuickMove\store-filter.txt
```

PST/OST 支持的 store 会按文件路径匹配，因此该选择可以在 Outlook 重启和 VSTO 部署刷新后保留。

`Mark as read before moving` 复选框会记住上一次确认后的状态，并在关闭对话框和重启 Outlook
后继续保留。它与 store 过滤设置一起存储在：

```powershell
$env:APPDATA\OutlookQuickMove\mark-as-read.txt
```

对话框窗口使用嵌入的图标资源：`OutlookQuickMove\Assets\QuickMove.ico`。

修改并重新构建加载项后，请使用同一个 `.vsto` 路径执行一次干净重装，确保 Outlook 刷新
VSTO 部署缓存。见下方的 `干净重装`。

如果按钮不可见，请检查 Outlook 是否加载了该加载项：

`File` -> `Options` -> `Add-ins` -> `Manage: COM Add-ins` -> `Go...`

该加载项在 Debug 和 Release 构建中都会写入轻量级诊断日志（启动/功能区 trace 以及任何
已处理的错误），位置为：

```powershell
$env:TEMP\OutlookQuickMove.log
```

有用的条目包括 `ThisAddIn_Startup`、`CreateRibbonExtensibilityObject` 和
`GetCustomUI ribbonId=Microsoft.Outlook.Explorer`，以及任何文件夹枚举或移动失败记录。
该文件采用 best-effort 写入方式，绝不会阻塞加载项，并且上限约为 1 MB（超过该大小后会重建）。
注意：日志中可能包含文件夹路径，以及移动失败邮件的主题。

## 干净重装

在代码变更后刷新本地 Debug 构建时使用此流程。

1. 关闭 Outlook。
2. 使用同一个 `.vsto` URI 卸载当前 VSTO 部署。
3. 清理当前用户的 ClickOnce 应用缓存。
4. 再次安装。

```powershell
$installer = Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'
$vsto = 'C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto'
$vstoUri = ([System.Uri](Get-Item -LiteralPath $vsto).FullName).AbsoluteUri
$mage = 'C:\Program Files (x86)\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\mage.exe'

Get-Process Outlook,VSTOInstaller,dfsvc -ErrorAction SilentlyContinue |
    Stop-Process -Force

& $installer /Uninstall $vstoUri
& $mage -cc
& $installer /Install $vstoUri
```

如果卸载时报 customization 未安装或已经卸载，这在清理损坏的本地安装时是可以接受的。
继续执行缓存清理和安装步骤即可。

## 卸载

使用与安装时相同的 `.vsto` 部署清单位置。

```powershell
$installer = Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'
$vsto = 'C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto'
$vstoUri = ([System.Uri](Get-Item -LiteralPath $vsto).FullName).AbsoluteUri

& $installer /Uninstall $vstoUri
```

卸载后重启 Outlook。

如果本地构建输出已被清理或移动，请先重新构建同一配置，确保 `.vsto` 清单路径再次存在，
然后运行卸载命令。

## 重装故障排查

如果重装仍然失败并显示此错误：

```text
Unable to install this application because an application with the same identity is already installed.
```

并且第二次卸载显示：

```text
Attempting to uninstall a customization that has not been installed on this computer or has already been uninstalled from this computer.
```

这说明 Outlook 的加载项注册已经不存在，但 ClickOnce/VSTO 缓存中仍保留旧的部署身份。
缓存身份包含诸如 `OutlookQuickMove.vsto`、`version="1.0.0.0"` 和清单 public key token
这样的值。

首先确认没有普通 Outlook 加载项注册残留：

```powershell
Get-ItemProperty -Path `
  'HKLM:\Software\Microsoft\Office\Outlook\Addins\*',`
  'HKLM:\Software\WOW6432Node\Microsoft\Office\Outlook\Addins\*',`
  'HKCU:\Software\Microsoft\Office\Outlook\Addins\*' `
  -ErrorAction SilentlyContinue |
  Where-Object {
      $_.PSChildName -like '*OutlookQuickMove*' -or
      $_.FriendlyName -like '*OutlookQuickMove*' -or
      $_.Manifest -like '*OutlookQuickMove*'
  } |
  Select-Object PSPath,PSChildName,FriendlyName,LoadBehavior,Manifest
```

然后查找过期的 ClickOnce/VSTO 缓存条目：

```powershell
$cacheRoot = Join-Path $env:LOCALAPPDATA 'Apps\2.0'

Get-ChildItem $cacheRoot -Recurse -ErrorAction SilentlyContinue |
  Where-Object {
      $_.Name -like '*.vsto' -or
      $_.Name -like '*.manifest' -or
      $_.Name -like '*.cdf-ms'
  } |
  Select-String -Pattern 'OutlookQuickMove' -List -ErrorAction SilentlyContinue |
  Select-Object Path,LineNumber,Line

reg query "HKCU\Software\Microsoft\VSTO" /f OutlookQuickMove /s
```

如果运行 `mage.exe -cc` 后仍有过期条目，请只移除与此加载项匹配的条目。**不要**删除整个
`%LOCALAPPDATA%\Apps\2.0` 文件夹，因为其中可能包含其他 Office 加载项。

下面第一个脚本会从当前 `.vsto` 文件中提取清单 public key token，并预览匹配到的缓存目标。

```powershell
$installer = Join-Path $env:CommonProgramFiles 'Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'
$vsto = 'C:\path\to\OutlookQuickMove\OutlookQuickMove\bin\Debug\OutlookQuickMove.vsto'
$vstoUri = ([System.Uri](Get-Item -LiteralPath $vsto).FullName).AbsoluteUri
$cacheRoot = (Resolve-Path (Join-Path $env:LOCALAPPDATA 'Apps\2.0')).Path

[xml]$vstoManifest = Get-Content -LiteralPath $vsto
$identity = $vstoManifest.assembly.assemblyIdentity
$token = $identity.publicKeyToken
if ([string]::IsNullOrWhiteSpace($token)) {
    throw 'Could not read publicKeyToken from the .vsto manifest.'
}

Get-Process Outlook,VSTOInstaller,dfsvc -ErrorAction SilentlyContinue |
    Stop-Process -Force

$manifestTargets = Get-ChildItem $cacheRoot -Recurse -File -ErrorAction SilentlyContinue |
  Where-Object {
      ($_.Name -like '*.manifest' -or $_.Name -like '*.cdf-ms') -and
      $_.Name -like "*$token*"
  }

$payloadTargets = Get-ChildItem $cacheRoot -Recurse -Directory -ErrorAction SilentlyContinue |
  Where-Object { $_.Name -like "outl*$token*" }

$targets = @($manifestTargets.FullName) + @($payloadTargets.FullName) |
  Sort-Object -Unique

$targets
```

请检查预览结果。如果其中只列出了 OutlookQuickMove 缓存条目，再运行清理：

```powershell
foreach ($target in $targets) {
    $resolved = (Resolve-Path -LiteralPath $target).Path
    if (-not $resolved.StartsWith($cacheRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to delete outside ClickOnce cache: $resolved"
    }
    if ($resolved -notlike "*$token*") {
        throw "Refusing to delete a path that does not contain the manifest token: $resolved"
    }

    Remove-Item -LiteralPath $resolved -Recurse -Force
}

Get-ChildItem 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion' -ErrorAction SilentlyContinue |
  Where-Object {
      (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue).Url -eq $vstoUri
  } |
  Remove-Item -Recurse -Force

& $installer /Install $vstoUri
```

作为仅限开发环境的 workaround，递增部署清单版本也会改变 ClickOnce identity，从而避开冲突。
建议优先清理过期缓存，让本地安装状态保持清晰可理解。

## 贡献

欢迎贡献。请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)，了解开发环境设置、手动测试清单，
以及项目约定（COM hygiene、绝不删除邮件、记录日志而不是静默失败）。

## 许可证

[MIT](LICENSE) © r404r
