# 数据模型

## `setting.json`

全局按键延迟设置。

```json
{
  "key_up_interval": 0.01,
  "key_interval": 0.08
}
```

| 字段 | 类型 | 含义 |
| --- | --- | --- |
| `key_interval` | number | 普通按键之间的基础间隔，单位是秒 |
| `key_up_interval` | number | 普通按键按下到抬起之间的基础间隔，单位是秒 |

执行时会叠加随机系数，让每次输入间隔略有差异。

## `quick.json`

宏方案集合。顶层 key 是方案名，value 是该方案下的 combo 列表。

```json
{
  "默认方案": [
    {
      "trigger_key": "f",
      "sequence": ["上", "下", "`120", "x150"]
    }
  ]
}
```

| 字段 | 类型 | 含义 |
| --- | --- | --- |
| `trigger_key` | string | 触发键。键盘键直接写键名；鼠标键使用特殊值 |
| `sequence` | string[] | 宏执行序列，每个元素是一个 token |

## 触发键规则

| 值 | 含义 |
| --- | --- |
| `x1` | 鼠标侧后键 |
| `x2` | 鼠标侧前键 |
| `MLeft` | 鼠标左键 |
| `MRight` | 鼠标右键 |
| 其他字符串 | 交给 `keyboard` 库识别的键盘热键 |

注意：`add_combo_to_cast(..., hotkey=True)` 中键盘触发键使用了 `"alt+" + trigger_key` 注册，但 `run_listener()` 中直接使用 `trigger_key` 注册。当前 UI 的“开始监听”路径走的是 `run_listener()`。

## 宏序列 token 规则

| token 形式 | 含义 | 示例 |
| --- | --- | --- |
| 单字符 | 普通键盘按键 | `q`、`e`、`1` |
| `上` / `下` / `左` / `右` | 方向键别名 | `上` 转为 `up` |
| `` `n`` | 空等待 n 毫秒 | `` `120`` 表示等待约 120ms |
| `lpn` | 按住鼠标左键 n 毫秒 | `lp50` |
| `rpn` | 按住鼠标右键 n 毫秒 | `rp50` |
| `x1n` | 按住鼠标侧后键 n 毫秒 | `x180` |
| `x2n` | 按住鼠标侧前键 n 毫秒 | `x280` |

UI 新增宏时会限制：

- token 用空格分隔。
- 普通键只能是长度为 1 的字符串，方向键中文别名也按普通 token 接收。
- 延迟和按住类 token 的毫秒值必须是数字，且不超过 10000。
- 触发键不能包含空格。

## 执行时间线

`run_combo()` 不会边解析边立即执行，而是先生成事件列表：

```text
(real_key, is_press, timestamp)
```

普通键会生成一次 press 和一次 release。按住类鼠标 token 也会生成 press/release，并把 release 放在 `timestamp + 按住时长`。最后按 timestamp 排序并依次 sleep 到目标时间点。
