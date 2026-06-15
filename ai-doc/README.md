# quick_skill AI 文档

这个目录用于给 AI 助手和后续维护者快速理解项目。内容基于当前代码结构整理，不改变项目运行逻辑。

## 文档索引

- [基础架构](./architecture.md)：项目分层、模块职责、启动流程和运行时关系。
- [数据模型](./data-model.md)：`quick.json`、`setting.json` 的结构，以及宏序列 token 规则。
- [开发维护笔记](./maintenance.md)：本地运行、打包、权限要求和当前代码中的维护注意点。

## 项目一句话概览

`quick_skill` 是一个 Windows 桌面键盘/鼠标宏工具。它使用 Tkinter 提供图形界面，使用 `keyboard` 和 `pynput` 监听/模拟输入，通过 JSON 文件持久化宏方案和全局延迟设置。
