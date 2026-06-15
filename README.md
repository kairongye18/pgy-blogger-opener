# 小红书蒲公英博主详情页批量打开工具

这个工具会读取 Excel 里的「小红书昵称」和「主页链接」列，解析小红书 userId，然后在默认浏览器里连续打开蒲公英博主详情页。默认每次打开 10 个链接，并且每 10 个链接新开一个浏览器窗口。

## 使用前准备

1. Mac 用户直接使用 `蒲公英博主打开工具.app`，不需要另外安装 Python。
2. 浏览器里已登录蒲公英账号。
3. Excel 表里有「小红书昵称」和「主页链接」两列。

## 图形界面程序使用

Mac：

1. 双击 `蒲公英博主打开工具.app`。
2. 如果 macOS 提示无法打开，可以右键点击它，选择「打开」。
3. 当前打包版本适用于 Apple Silicon Mac。
4. 如果需要重新生成这个一键启动程序，可以在这个文件夹里执行：

   ```bash
   python3 make_mac_app.py
   ```

5. 在窗口里选择 Excel 文件，确认参数后点击「开始打开」。

Windows：

1. 在 GitHub Releases 下载 `pgy-blogger-opener-windows.exe`。
2. 双击 `.exe` 即可打开，不需要安装 Python。
3. 如果 Windows Defender 弹出未知发布者提示，选择「更多信息」后确认运行。

源码运行方式：

   ```bat
   python -m pip install -r requirements.txt
   python 蒲公英博主打开工具.py
   ```

在窗口里选择 Excel 文件，确认参数后点击「开始打开」。

备用方式：

1. Mac 可运行 `python3 蒲公英博主打开工具.py` 或 `run_mac.command`。
2. Windows 可运行 `run_windows.bat`。

## Windows EXE 发布

GitHub Actions 工作流位于 `.github/workflows/windows-release.yml`。

- 手动运行工作流会生成可下载的 Actions Artifact。
- 推送 `v*` 标签会自动创建 GitHub Release，并上传 `pgy-blogger-opener-windows.exe`。

窗口里的常用设置：

- 「从第几个有效博主开始」：默认 1；下一批可填 11、21、31。
- 「打开几个」：默认 10。
- 「每几个新开一个窗口」：默认 10。
- 「只预览链接，不打开浏览器」：用于检查解析结果。

## 命令行使用

在这个文件夹里执行：

```bash
python3 -m pip install -r requirements.txt
python3 pgy_opener.py "/path/to/博主表.xlsx" --start 1 --count 10
```

常用参数：

```bash
# 打开第 11 到第 20 个有效博主
python3 pgy_opener.py "/path/to/博主表.xlsx" --start 11 --count 10

# 打开 30 个；第 1、11、21 个会分别新开窗口
python3 pgy_opener.py "/path/to/博主表.xlsx" --start 1 --count 30

# 改成每 20 个链接新开一个窗口
python3 pgy_opener.py "/path/to/博主表.xlsx" --start 1 --count 40 --window-size 20

# 只打印链接，不打开浏览器
python3 pgy_opener.py "/path/to/博主表.xlsx" --dry-run

# 指定工作表名
python3 pgy_opener.py "/path/to/博主表.xlsx" --sheet "收集结果"
```

## 注意

- `xhslink.com` 短链需要联网解析。
- 如果浏览器没登录蒲公英，打开的详情页会跳到登录页；先登录后再运行即可。
- 如果某一行缺少主页链接或短链解析失败，工具会跳过并在窗口日志里列出原因。
