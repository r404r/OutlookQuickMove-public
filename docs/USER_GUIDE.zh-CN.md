# OutlookQuickMove 安装和使用指南

[English](USER_GUIDE.md) | 中文

OutlookQuickMove 是一个经典版 Outlook for Windows 的 VSTO 加载项，用于快速把当前选中的邮件移动到目标文件夹，也可以快速跳转到文件夹或跳转到邮件所在文件夹。

## 适用环境

- 经典版 Outlook for Windows。
- 不适用于 New Outlook、Outlook Web 或 Mac Outlook。
- 安装前建议先关闭 Outlook。

## 安装

1. 打开 GitHub Releases 页面，下载最新的发布压缩包，例如 `OutlookQuickMove-1.0.0.9.7z`。
2. 将压缩包解压到一个稳定的位置。不要直接从压缩包内运行安装程序。
3. 确认 Outlook 已关闭。
4. 双击解压后的 `setup.exe`。
5. 按安装向导完成安装。
6. 安装完成后打开 Outlook。

![GitHub Releases 中的 OutlookQuickMove 发布条目](assets/user-guide/01-release-download.png)

## 打开 Quick Move

安装完成并重新打开 Outlook 后，功能区会出现 `Quick Move` 选项卡。

![Outlook 功能区中的 Quick Move 选项卡](assets/user-guide/02-quick-move-tab.png)

在这个选项卡中可以看到以下按钮：

- `Quick Move`
- `Go to Folder`
- `Go to Mail Folder`
- `Undo Quick Move...`
- `Settings`

## Quick Move

`Quick Move` 用于把当前选中的邮件移动到目标文件夹。

1. 在 Outlook 中选中一封或多封邮件。
2. 点击 `Quick Move`。
3. 对话框打开后，输入框会自动获得焦点。
4. 输入目标文件夹的关键字，列表会实时过滤匹配的文件夹。
5. 使用方向键选择目标文件夹，或直接保留当前高亮项。
6. 按 `Enter`，或点击确认按钮，将邮件移动到所选文件夹。

如果 Outlook 配置了很多数据文件，重启 Outlook 后第一次打开 `Quick Move` 可能会慢一些。插件需要建立文件夹列表，后续会使用缓存，通常会快很多。

如果需要在移动前把邮件标记为已读，勾选 `Mark as read before moving`。

![Quick Move 对话框和文件夹搜索结果](assets/user-guide/03-quick-move-dialog.png)

## Go to Folder

`Go to Folder` 用于快速跳转到指定文件夹，不会移动邮件。

1. 点击 `Go to Folder`。
2. 输入文件夹关键字。
3. 选择目标文件夹。
4. 按 `Enter`，Outlook 会切换到该文件夹。

它使用与 `Quick Move` 相同的文件夹候选范围，也会复用常用文件夹排序。

## Go to Mail Folder

`Go to Mail Folder` 用于从当前选中的邮件跳转到该邮件所在的文件夹。

这个功能适合在 Outlook 搜索结果中使用：当你搜索到一封邮件后，如果想回到它所在的原始文件夹，选中该邮件并点击 `Go to Mail Folder`。

如果选中了多封邮件，插件会提示确认，并使用当前选择中的第一封邮件作为跳转目标。

## Undo Quick Move

`Undo Quick Move...` 用于回撤通过 `Quick Move` 执行的移动操作。

1. 点击 `Undo Quick Move...`。
2. 在列表中选择要回撤的移动记录。
3. 点击 `Undo Selected`。
4. 插件会尝试把邮件移回原始文件夹，并恢复原来的已读或未读状态。

![Undo Quick Move 对话框](assets/user-guide/04-undo-quick-move.png)

## Settings

`Settings` 用于调整 Quick Move 的候选数据文件、常用文件夹记录数量和撤回历史记录数量。

### Data Files

当 Outlook 中有多个数据文件时，可以在 `Data Files` 中选择哪些数据文件参与候选文件夹搜索。

如果只想在某一个邮箱或 PST 文件中查找目标文件夹，只勾选对应的数据文件即可。这样可以减少候选列表，也能降低第一次建立文件夹列表时的开销。

![Settings 中的 Data Files 设置页](assets/user-guide/05-settings-data-files.png)

### Frequent Folders

`Frequent Folders` 用于控制常用目标文件夹在候选列表顶部显示的数量。

常用文件夹会根据你通过 `Quick Move` 实际移动邮件的次数自动排序。设置为 `0` 时会关闭常用文件夹显示。

![Settings 中的 Frequent Folders 设置页](assets/user-guide/06-settings-frequent-folders.png)

### Undo History

`Undo History` 用于设置保留多少条可撤回的移动记录。

设置为 `0` 时会停止记录新的撤回历史，但已有历史仍会保留，直到手动清除。

![Settings 中的 Undo History 设置页](assets/user-guide/07-settings-undo-history.png)

## 快捷键

OutlookQuickMove 不注册全局快捷键。推荐使用 Outlook 自带的 Quick Access Toolbar 来获得 `Alt + 数字键` 的快捷入口。

1. 打开 Quick Access Toolbar 的下拉菜单。
2. 点击 `More Commands...`。
3. 在 `Choose commands from` 中选择 `Quick Move Tab`。
4. 选择 `Quick Move`，点击 `Add >>`。
5. 点击 `OK` 保存。
6. 根据该按钮在 Quick Access Toolbar 中的位置，使用 `Alt + 数字键` 快速打开。

例如，如果 `Quick Move` 是 Quick Access Toolbar 上的第一个按钮，可以用 `Alt + 1` 打开。

![Quick Access Toolbar 菜单中的 More Commands](assets/user-guide/08-add-to-quick-access-toolbar.png)

![在设置界面中把 Quick Move 添加到 Quick Access Toolbar](assets/user-guide/09-quick-access-toolbar-shortcut.png)

## 截图隐私检查

同步到公共库前，请确认截图中不包含以下信息：

- 真实姓名、邮箱地址、公司账号或头像。
- 邮件主题、发件人、收件人、正文预览。
- 私有邮箱、PST、OST 或共享邮箱名称。
- 私人文件夹名称、客户名称、项目名称、工单号。
- Windows 用户名、本地路径、网络路径。
- 任何可反推出个人身份或工作内容的信息。

如果截图中需要展示列表内容，请使用测试邮箱、测试文件夹，或先对敏感内容打码后再提交。
