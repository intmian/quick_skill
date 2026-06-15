# README Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rewrite `readme.md` so new users can quickly understand what the tool does, how to install it, how to configure macros, and what safety boundaries apply.

**Architecture:** Keep the README as a single user-facing document. Preserve existing screenshot URLs, replace ornate marketing copy with practical usage guidance, and add a concise macro syntax reference. Do not change runtime code in this plan.

**Tech Stack:** Markdown, GitHub README rendering, existing GitHub-hosted image links.

---

## File Structure

- Modify: `readme.md`
  - Replace the current introduction with direct product positioning.
  - Add quick-start steps.
  - Add macro syntax examples for keyboard keys, mouse keys, hold tokens, and delay tokens.
  - Keep existing feedback and safety sections, but make the language clearer and more responsible.
- No change: `gui.py`, `quick_mgr.py`, `pack.bat`, `pack/quick.json`, `pack/setting.json`

## Task 1: Capture Existing README Assets

**Files:**
- Read: `readme.md`
- Modify: none

- [ ] **Step 1: List existing README headings**

Run:

```bash
rg -n "^#" readme.md
```

Expected result:

```text
The command lists the current heading structure, including 项目简介, 功能特点, 项目速览, 使用方法, 例子, 反馈, 注意事项, and 交响之旅.
```

- [ ] **Step 2: List existing image URLs**

Run:

```bash
rg -n "!\\[image\\]\\(" readme.md
```

Expected result:

```text
The command lists all existing screenshot URLs so they can be preserved in the rewrite.
```

## Task 2: Replace README with a Practical User Guide

**Files:**
- Modify: `readme.md`

- [ ] **Step 1: Replace the README content**

Replace `readme.md` with this content:

````markdown
# quick_skill

`quick_skill` 是一个 Windows 上的简易键鼠宏管理工具。它可以为不同场景创建多套方案，并通过键盘热键或鼠标按键触发一段按键序列。

> 使用前请确认你的使用场景允许自动化输入。不要在违反软件、平台或游戏规则的场景中使用。

## 功能

- 多方案管理：为不同应用或操作流程保存不同宏方案。
- 键盘触发：配置普通键盘按键作为宏触发键。
- 鼠标触发：支持鼠标左键、右键、侧后键、侧前键作为触发键。
- 按键序列：按顺序模拟键盘按下、释放和鼠标按下、释放。
- 延迟控制：支持全局按键间隔、抬起间隔，以及序列中的自定义等待。
- 随机抖动：在延迟中加入小范围随机值，降低机械感。

## 界面预览

![image](https://github.com/intmian/quick_skill/assets/38103855/9b42afbd-6836-4eab-ad3d-2e27f6b58448)

![image](https://github.com/intmian/quick_skill/assets/38103855/d293064c-7407-4ad9-b87d-50f29263111a)

![image](https://github.com/intmian/quick_skill/assets/38103855/51709ee0-d170-417a-8d17-5ca861895ce6)

## 安装

1. 打开 [Releases](https://github.com/intmian/quick_skill/releases/)。
2. 下载最新正式版本。
3. 解压后运行程序。
4. 按提示以管理员权限启动。全局键盘监听通常需要管理员权限。

## 快速开始

1. 新建一个方案，例如“默认方案”。
2. 双击列表中的“新增”提示，添加一个快捷键。
3. 输入触发键，例如 `f`、`x1` 或 `MRight`。
4. 输入按键序列，例如 `上 下 左 右`。
5. 根据需要调整“按键间隔秒”和“抬起间隔秒”。
6. 点击“开始”后，在目标窗口中按触发键执行宏。
7. 点击“停止”后可以继续编辑方案。

## 触发键写法

| 写法 | 含义 |
| --- | --- |
| `a` | 键盘 `a` 键 |
| `f` | 键盘 `f` 键 |
| `x1` | 鼠标侧后键 |
| `x2` | 鼠标侧前键 |
| `MLeft` | 鼠标左键 |
| `MRight` | 鼠标右键 |

普通键盘触发键按 `keyboard` 库支持的名称填写。鼠标触发键使用表中的固定写法。

## 按键序列写法

按键序列使用空格分隔。

| 写法 | 含义 |
| --- | --- |
| `上` | 方向键上 |
| `下` | 方向键下 |
| `左` | 方向键左 |
| `右` | 方向键右 |
| `a` | 点击键盘 `a` 键 |
| `` `200 `` | 等待约 200 毫秒 |
| `lp500` | 按住鼠标左键约 500 毫秒 |
| `rp500` | 按住鼠标右键约 500 毫秒 |
| `x1500` | 按住鼠标侧后键约 500 毫秒 |
| `x2500` | 按住鼠标侧前键约 500 毫秒 |

示例：

```text
上 下 左 右
```

```text
`200 a b c
```

```text
lp800 `100 a
```

## 示例

地狱潜兵 2：一键输入战略配备方向序列，或为磁轨炮不安全模式设置按住时间。

![image](https://github.com/user-attachments/assets/4d968bbd-84d8-4e9a-8504-e67685043530)

某些动作流程：一键执行一段固定按键序列。

![image](https://github.com/intmian/quick_skill/assets/38103855/d0e9b57c-5ead-4b56-843b-43639871d05e)

## 数据文件

程序会在运行目录读写以下文件：

- `quick.json`：保存宏方案和按键序列。
- `setting.json`：保存全局按键间隔和抬起间隔。

如果使用打包版本，运行目录可能与源码目录不同。备份或迁移配置时，请确认实际生成的 JSON 文件位置。

## 常见问题

### 为什么需要管理员权限？

程序使用全局键盘监听和输入模拟。部分 Windows 环境下，如果程序不是管理员权限，监听或模拟输入可能不生效。

### 为什么触发后没有反应？

先检查以下项目：

- 是否已经点击“开始”。
- 当前方案是否包含对应触发键。
- 触发键写法是否正确，例如鼠标右键应写作 `MRight`。
- 目标窗口是否允许模拟输入。
- 是否有其他软件拦截了全局热键。

### 如何删除快捷键？

在方案列表中双击已有快捷键，确认后即可删除。

## 注意事项

- 请只在允许自动化输入的场景中使用本程序。
- 使用本程序前，请确认不会违反目标软件、平台或游戏的用户协议。
- 宏执行期间会模拟真实输入，请避免在聊天框、表单、支付页面等敏感输入区域误触。
- 当前版本没有按前台窗口过滤，开始监听后全局触发键都可能生效。
- 因使用本程序产生的账号、数据或其他风险，由使用者自行承担。

## 反馈

如果你遇到问题或有改进建议，可以通过 [Issues](https://github.com/intmian/quick_skill/issues) 反馈。

常见问题也可以参考 [issue 置顶](https://github.com/intmian/quick_skill/issues/11)。

## 开发

源码主要由两个文件组成：

- `gui.py`：Tkinter 界面、用户事件处理、管理员权限重启。
- `quick_mgr.py`：配置读写、键鼠监听注册、宏执行。

打包脚本为 `pack.bat`，使用 PyInstaller 生成单文件 Windows 程序。
````

- [ ] **Step 2: Verify the README contains the required user sections**

Run:

```bash
rg -n "^## (功能|界面预览|安装|快速开始|触发键写法|按键序列写法|示例|数据文件|常见问题|注意事项|反馈|开发)" readme.md
```

Expected result:

```text
All listed sections are present exactly once.
```

- [ ] **Step 3: Verify old unclear phrasing was removed**

Run:

```bash
rg -n "不为监测所知|交响之旅|琴箫|韵律|花语" readme.md
```

Expected result:

```text
No matches.
```

- [ ] **Step 4: Commit the README rewrite**

```bash
git add readme.md
git commit -m "docs: clarify README usage guide"
```

## Task 3: Final Markdown Review

**Files:**
- Read: `readme.md`

- [ ] **Step 1: Check image links remain present**

Run:

```bash
rg -n "github.com/.+/(assets|user-attachments)/" readme.md
```

Expected result:

```text
The command prints the existing screenshot and example image URLs.
```

- [ ] **Step 2: Check command and token examples render as code**

Run:

```bash
rg -n "`quick_skill`|`quick.json`|`setting.json`|`MLeft`|`MRight`|`x1`|`x2`|lp800" readme.md
```

Expected result:

```text
The command prints lines containing inline code examples and the hold-token example.
```

- [ ] **Step 3: Review the README in GitHub-style Markdown preview**

Open `readme.md` in an editor with Markdown preview and verify:

```text
Tables render correctly.
The backtick-delay examples render visibly.
All screenshots appear in the intended sections.
The safety notes are easy to find.
```

