# 开发维护笔记

## 本地运行

```powershell
python gui.py
```

运行要求：

- Windows 环境。
- Python 环境中安装 `keyboard`、`pynput`。
- 需要管理员权限，否则程序会尝试以管理员权限重启。

## 打包

```powershell
.\pack.bat
```

打包脚本会：

1. 创建 `pack` 目录。
2. 执行 `pyinstaller --onefile --noconsole gui.py`。
3. 将 `dist\gui.exe` 移动到 `pack\gui.exe`。
4. 删除 `dist` 和 `build`。

`.gitignore` 已忽略根目录运行时生成的 `quick.json`、`setting.json`、`pack/gui.exe`、`*.spec` 和 `__pycache__`。

## 重要维护点

- `gui.py` 同时承担入口、UI 和协调逻辑。后续如果功能增长，可以优先把 `mgr` 拆到独立模块。
- `QuickCastManager` 在初始化时立即启动鼠标监听器，生命周期和程序一致。
- `stop_listener()` 会清理键盘 hotkey，但鼠标监听器不会停止，只会清空 `mouse_combo` 映射。
- `run_combo()` 的 `lock` 没有使用 `try/finally`。如果执行过程中异常，理论上可能导致锁无法释放。
- JSON 文件路径是相对路径。源码运行和打包后运行时，实际读写位置可能不同。
- 当前没有测试用例。若要补测试，优先覆盖纯数据逻辑，例如 token 校验、方案增删、配置读写；实际键鼠模拟建议隔离为可替换适配器后再测。

## 适合 AI 修改前先看的位置

- 新增或调整 UI：先看 `gui.ui.__init__()` 和对应的 `on_*` 状态刷新方法。
- 新增用户操作：先看 `gui.mgr` 中的事件处理方法。
- 调整宏执行语义：先看 `quick_mgr.QuickCastManager.run_combo()`。
- 调整持久化结构：同时更新 `load_*`、`save_*`、UI 展示逻辑和本目录的数据模型文档。
