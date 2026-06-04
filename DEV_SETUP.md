# DEV_SETUP: OutlookQuickMove 开发环境

## Outlook 版本

当前确认版本：

`Microsoft Outlook for Microsoft 365 MSO (Version 2605 Build 16.0.20026.20076) 64-bit`

该版本可走经典 Outlook VSTO 插件开发路线。

## Visual Studio 安装

安装 Visual Studio 2022 Community / Professional / Enterprise。

工作负载选择：

`Office/SharePoint development`

建议保留：

- Office Developer Tools for Visual Studio
- Visual Studio Tools for Office (VSTO)
- .NET Framework 4.8 development tools
- .NET Framework 4.7.2 development tools
- IntelliCode

可以不装：

- Web Deploy
- GitHub Copilot
- .NET Framework 4.6.2-4.7.1 development tools
- .NET Framework 4.8.1 development tools
- Windows Identity Foundation 3.5

## 创建项目

1. 打开 Visual Studio 2022。
2. 选择 `Create a new project`。
3. 搜索 `Outlook VSTO Add-in`。
4. 选择 `Outlook VSTO Add-in`，不要选择 `Outlook Web Add-in`。
5. Language 选择 `C#`。
6. Framework 选择 `.NET Framework 4.8`。
7. Project name 建议 `OutlookQuickMove`。
8. 不勾选 `Place solution and project in the same directory`。

推荐目录结构：

```text
OutlookQuickMove/
  OutlookQuickMove.sln
  OutlookQuickMove/
    OutlookQuickMove.csproj
    ThisAddIn.cs
```

## 首次验证

创建后按 `F5`。

期望结果：

- Visual Studio 启动 Outlook。
- Outlook 正常打开。
- 没有 VSTO 加载错误。

检查插件：

`File` -> `Options` -> `Add-ins`

确认插件出现在 Add-ins 列表中。

## 添加 Ribbon

1. 右键项目。
2. `Add` -> `New Item`。
3. 搜索 `Ribbon`。
4. 选择 `Ribbon (Visual Designer)`。
5. 命名为 `QuickMoveRibbon.cs`。
6. 添加 Group。
7. 添加 Button。
8. Button 文本设为 `Quick Move`。
9. 点击事件中先写测试弹窗。

测试代码：

```csharp
private void buttonQuickMove_Click(object sender, RibbonControlEventArgs e)
{
    System.Windows.Forms.MessageBox.Show("Quick Move triggered.");
}
```

再次 `F5`，确认 Outlook 中能看到按钮并成功弹窗。

## 添加 WinForms 对话框

新增 Windows Form：

`MoveToFolderForm.cs`

建议控件：

- `TextBox` named `textSearch`
- `ListBox` named `listFolders`
- `CheckBox` named `checkMarkAsRead`
- `Button` named `buttonOk`
- `Button` named `buttonCancel`

建议属性：

- `AcceptButton = buttonOk`
- `CancelButton = buttonCancel`
- `StartPosition = CenterParent`

Form 打开后：

```csharp
textSearch.Focus();
```

## 发布部署

开发调试阶段：

- 使用 F5 调试。

个人使用阶段：

- 可先用 ClickOnce。

企业稳定使用：

- 建议使用 MSI / WiX。
- 建议使用代码签名证书。
- 确认目标电脑已安装 VSTO Runtime。

