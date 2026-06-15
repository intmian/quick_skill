# Performance Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reduce macro trigger latency and timing jitter while preserving the existing UI workflow and persisted JSON format.

**Architecture:** Keep `gui.py` as the UI coordinator and concentrate runtime performance changes in `quick_mgr.py`. Split the hot path into three parts: sequence compilation, timeline generation, and input execution. Add tests around lock release, debug output removal, precompiled sequence reuse, and listener callback behavior.

**Tech Stack:** Python 3, Tkinter, `keyboard`, `pynput.mouse`, `ctypes.windll`, `pytest`, `unittest.mock`.

---

## File Structure

- Modify: `quick_mgr.py`
  - Remove high-frequency debug output.
  - Add reusable constants for mouse trigger keys and direction key aliases.
  - Add sequence compilation helpers that do not mutate persisted combo dictionaries.
  - Add `try/finally` lock release around macro execution.
  - Move listener callbacks toward a queued worker so callback threads do minimal work.
- Create: `tests/test_quick_mgr.py`
  - Provide fake `keyboard`, `pynput.mouse`, and `ctypes.windll` dependencies.
  - Verify behavior without requiring administrator permissions or real input devices.
- No change: `gui.py`
  - UI event flow should keep calling the existing `QuickCastManager` methods.

## Task 1: Add a Test Harness and Lock-Release Regression Test

**Files:**
- Create: `tests/test_quick_mgr.py`
- Modify: none

- [ ] **Step 1: Add the failing test file**

Create `tests/test_quick_mgr.py` with this content:

```python
import importlib
import sys
import types

import pytest


class FakeButton:
    x1 = object()
    x2 = object()
    left = object()
    right = object()


class FakeController:
    def __init__(self):
        self.events = []

    def press(self, button):
        self.events.append(("press", button))

    def release(self, button):
        self.events.append(("release", button))


class FakeListener:
    def __init__(self, on_click):
        self.on_click = on_click
        self.started = False

    def start(self):
        self.started = True


class FakeUser32:
    pass


class FakeKernel32:
    pass


class FakePsapi:
    pass


class FakeWindll:
    user32 = FakeUser32()
    kernel32 = FakeKernel32()
    psapi = FakePsapi()


def import_quick_mgr(monkeypatch, tmp_path, keyboard_module):
    fake_mouse_module = types.SimpleNamespace(
        Button=FakeButton,
        Controller=FakeController,
        Listener=FakeListener,
    )
    fake_pynput_module = types.SimpleNamespace(mouse=fake_mouse_module)

    import ctypes

    monkeypatch.setattr(ctypes, "windll", FakeWindll(), raising=False)
    monkeypatch.setitem(sys.modules, "keyboard", keyboard_module)
    monkeypatch.setitem(sys.modules, "pynput", fake_pynput_module)
    monkeypatch.setitem(sys.modules, "pynput.mouse", fake_mouse_module)
    monkeypatch.chdir(tmp_path)

    sys.modules.pop("quick_mgr", None)
    return importlib.import_module("quick_mgr")


def make_keyboard(press=None, release=None, is_pressed=None):
    return types.SimpleNamespace(
        add_hotkey=lambda *args, **kwargs: None,
        unhook_all_hotkeys=lambda: None,
        is_pressed=is_pressed or (lambda key: False),
        press=press or (lambda key: None),
        release=release or (lambda key: None),
    )


def test_run_combo_releases_lock_when_input_fails(monkeypatch, tmp_path):
    def failing_press(key):
        raise RuntimeError("input backend failed")

    keyboard_module = make_keyboard(press=failing_press)
    quick_mgr = import_quick_mgr(monkeypatch, tmp_path, keyboard_module)
    manager = quick_mgr.QuickCastManager()

    with pytest.raises(RuntimeError, match="input backend failed"):
        manager.run_combo({"trigger_key": "a", "sequence": ["a"]})

    assert manager.lock is False
```

- [ ] **Step 2: Run the test and verify it fails for the expected reason**

Run:

```bash
pytest tests/test_quick_mgr.py::test_run_combo_releases_lock_when_input_fails -q
```

Expected result before implementation:

```text
FAILED tests/test_quick_mgr.py::test_run_combo_releases_lock_when_input_fails
E       assert True is False
```

- [ ] **Step 3: Commit the failing test**

```bash
git add tests/test_quick_mgr.py
git commit -m "test: cover macro lock release on input failure"
```

## Task 2: Release the Macro Lock with `try/finally`

**Files:**
- Modify: `quick_mgr.py:151-237`
- Test: `tests/test_quick_mgr.py`

- [ ] **Step 1: Wrap macro execution in `try/finally`**

In `quick_mgr.py`, change `run_combo()` so the lock is always released:

```python
    def run_combo(self, combo):
        if self.lock:
            return
        self.lock = True
        try:
            # 如果当前存在alt按键被按下，等待按键释放
            while keyboard.is_pressed('alt'):
                time.sleep(0.001)

            timeline_now = 0
            keys = []
            for key in combo['sequence']:
                if key[0] == "`":
                    delay = float(0)
                    delay += random.uniform(0.95, 1.051) * 0.001 * int(key[1:])
                    timeline_now += delay
                    continue
                if len(key) >= 3 and (key[:2] == "lp" or key[:2] == "rp" or key[:2] == "x1" or key[:2] == "x2"):
                    delay = float(0)
                    delay += random.uniform(0.99, 1.01) * 0.001 * int(key[2:])
                    if key[0] == "l":
                        keys.append(("MLeft", True, timeline_now))
                        keys.append(("MLeft", False, timeline_now + delay))
                    if key[0] == "r":
                        keys.append(("MRight", True, timeline_now))
                        keys.append(("MRight", False, timeline_now + delay))
                    if key[:2] == "x1":
                        keys.append(("x1", True, timeline_now))
                        keys.append(("x1", False, timeline_now + delay))
                    if key[:2] == "x2":
                        keys.append(("x2", True, timeline_now))
                        keys.append(("x2", False, timeline_now + delay))
                    timeline_now += delay
                    continue
                real_key = key
                if key == "上":
                    real_key = "up"
                elif key == "下":
                    real_key = "down"
                elif key == "左":
                    real_key = "left"
                elif key == "右":
                    real_key = "right"
                keys.append((real_key, True, timeline_now))
                delay = self.settings["key_up_interval"] * random.uniform(0.952, 1.05)
                keys.append((real_key, False, timeline_now + delay))
                timeline_now += self.settings["key_interval"] * random.uniform(0.82, 1.19)
            keys.sort(key=lambda x: x[2])
            for i, key in enumerate(keys):
                print(key)
                if i != 0:
                    time.sleep(key[2] - keys[i - 1][2])
                if key[0] == "MLeft":
                    if key[1]:
                        mouseController.press(mouse.Button.left)
                    else:
                        mouseController.release(mouse.Button.left)
                    continue
                elif key[0] == "MRight":
                    if key[1]:
                        mouseController.press(mouse.Button.right)
                    else:
                        mouseController.release(mouse.Button.right)
                    continue
                elif key[0] == "x1":
                    if key[1]:
                        mouseController.press(mouse.Button.x1)
                    else:
                        mouseController.release(mouse.Button.x1)
                    continue
                elif key[0] == "x2":
                    if key[1]:
                        mouseController.press(mouse.Button.x2)
                    else:
                        mouseController.release(mouse.Button.x2)
                    continue
                if key[1]:
                    keyboard.press(key[0])
                else:
                    keyboard.release(key[0])
        finally:
            self.lock = False
```

- [ ] **Step 2: Run the targeted test**

Run:

```bash
pytest tests/test_quick_mgr.py::test_run_combo_releases_lock_when_input_fails -q
```

Expected result:

```text
1 passed
```

- [ ] **Step 3: Commit the safety fix**

```bash
git add quick_mgr.py tests/test_quick_mgr.py
git commit -m "fix: release macro lock after input failures"
```

## Task 3: Remove High-Frequency Debug Output

**Files:**
- Modify: `quick_mgr.py:35`
- Modify: `quick_mgr.py:206`
- Modify: `tests/test_quick_mgr.py`

- [ ] **Step 1: Add failing stdout tests**

Append these tests to `tests/test_quick_mgr.py`:

```python
def test_mouse_click_without_registered_combo_does_not_print(monkeypatch, tmp_path, capsys):
    keyboard_module = make_keyboard()
    quick_mgr = import_quick_mgr(monkeypatch, tmp_path, keyboard_module)
    manager = quick_mgr.QuickCastManager()

    manager.on_click(0, 0, quick_mgr.mouse.Button.left, True)

    assert capsys.readouterr().out == ""


def test_run_combo_does_not_print_per_event(monkeypatch, tmp_path, capsys):
    keyboard_module = make_keyboard()
    quick_mgr = import_quick_mgr(monkeypatch, tmp_path, keyboard_module)
    manager = quick_mgr.QuickCastManager()

    manager.run_combo({"trigger_key": "a", "sequence": ["a"]})

    assert capsys.readouterr().out == ""
```

- [ ] **Step 2: Run the new tests and verify they fail**

Run:

```bash
pytest tests/test_quick_mgr.py::test_mouse_click_without_registered_combo_does_not_print tests/test_quick_mgr.py::test_run_combo_does_not_print_per_event -q
```

Expected result before implementation:

```text
2 failed
```

- [ ] **Step 3: Remove the debug prints**

In `quick_mgr.py`, delete these two lines:

```python
        print(self.mouse_combo)
```

```python
            print(key)
```

- [ ] **Step 4: Run all current tests**

Run:

```bash
pytest tests/test_quick_mgr.py -q
```

Expected result:

```text
3 passed
```

- [ ] **Step 5: Commit the debug output removal**

```bash
git add quick_mgr.py tests/test_quick_mgr.py
git commit -m "perf: remove debug output from macro hot path"
```

## Task 4: Extract Timeline Generation and Use Absolute Scheduling

**Files:**
- Modify: `quick_mgr.py`
- Modify: `tests/test_quick_mgr.py`

- [ ] **Step 1: Add a test for timeline event order**

Append this test:

```python
def test_build_timeline_keeps_expected_event_order(monkeypatch, tmp_path):
    keyboard_module = make_keyboard()
    quick_mgr = import_quick_mgr(monkeypatch, tmp_path, keyboard_module)
    manager = quick_mgr.QuickCastManager()
    manager.settings["key_interval"] = 0.08
    manager.settings["key_up_interval"] = 0.01
    monkeypatch.setattr(quick_mgr.random, "uniform", lambda low, high: 1.0)

    events = manager.build_timeline(["上", "`20", "lp50", "a"])

    assert events == [
        ("up", True, 0),
        ("up", False, 0.01),
        ("MLeft", True, 0.1),
        ("MLeft", False, 0.15),
        ("a", True, 0.15),
        ("a", False, 0.16),
    ]
```

- [ ] **Step 2: Run the test and verify it fails because `build_timeline` does not exist**

Run:

```bash
pytest tests/test_quick_mgr.py::test_build_timeline_keeps_expected_event_order -q
```

Expected result before implementation:

```text
AttributeError: 'QuickCastManager' object has no attribute 'build_timeline'
```

- [ ] **Step 3: Add constants and `build_timeline()`**

Near the top of `quick_mgr.py`, after the `psapi` assignment, add:

```python
MOUSE_TRIGGER_KEYS = {"x1", "x2", "MLeft", "MRight"}
KEY_ALIASES = {
    "上": "up",
    "下": "down",
    "左": "left",
    "右": "right",
}
HOLD_MOUSE_KEYS = {
    "lp": "MLeft",
    "rp": "MRight",
    "x1": "x1",
    "x2": "x2",
}
```

Inside `QuickCastManager`, add:

```python
    def build_timeline(self, sequence):
        timeline_now = 0
        events = []
        for key in sequence:
            if key[0] == "`":
                timeline_now += random.uniform(0.95, 1.051) * 0.001 * int(key[1:])
                continue

            prefix = key[:2]
            if len(key) >= 3 and prefix in HOLD_MOUSE_KEYS:
                delay = random.uniform(0.99, 1.01) * 0.001 * int(key[2:])
                mouse_key = HOLD_MOUSE_KEYS[prefix]
                events.append((mouse_key, True, timeline_now))
                events.append((mouse_key, False, timeline_now + delay))
                timeline_now += delay
                continue

            real_key = KEY_ALIASES.get(key, key)
            events.append((real_key, True, timeline_now))
            delay = self.settings["key_up_interval"] * random.uniform(0.952, 1.05)
            events.append((real_key, False, timeline_now + delay))
            timeline_now += self.settings["key_interval"] * random.uniform(0.82, 1.19)

        events.sort(key=lambda event: event[2])
        return events
```

- [ ] **Step 4: Add `execute_timeline()` with absolute scheduling**

Inside `QuickCastManager`, add:

```python
    def execute_timeline(self, events):
        started_at = time.perf_counter()
        for key, pressed, target_offset in events:
            wait_time = started_at + target_offset - time.perf_counter()
            if wait_time > 0:
                time.sleep(wait_time)
            self.send_input_event(key, pressed)
```

Inside `QuickCastManager`, add:

```python
    def send_input_event(self, key, pressed):
        if key == "MLeft":
            button = mouse.Button.left
            if pressed:
                mouseController.press(button)
            else:
                mouseController.release(button)
            return
        if key == "MRight":
            button = mouse.Button.right
            if pressed:
                mouseController.press(button)
            else:
                mouseController.release(button)
            return
        if key == "x1":
            button = mouse.Button.x1
            if pressed:
                mouseController.press(button)
            else:
                mouseController.release(button)
            return
        if key == "x2":
            button = mouse.Button.x2
            if pressed:
                mouseController.press(button)
            else:
                mouseController.release(button)
            return
        if pressed:
            keyboard.press(key)
        else:
            keyboard.release(key)
```

- [ ] **Step 5: Update `run_combo()` to use the extracted methods**

Replace the body of `run_combo()` with:

```python
    def run_combo(self, combo):
        if self.lock:
            return
        self.lock = True
        try:
            while keyboard.is_pressed('alt'):
                time.sleep(0.001)
            self.execute_timeline(self.build_timeline(combo["sequence"]))
        finally:
            self.lock = False
```

- [ ] **Step 6: Run all tests**

Run:

```bash
pytest tests/test_quick_mgr.py -q
```

Expected result:

```text
4 passed
```

- [ ] **Step 7: Commit the timeline extraction**

```bash
git add quick_mgr.py tests/test_quick_mgr.py
git commit -m "perf: extract macro timeline execution"
```

## Task 5: Cache Compiled Sequence Tokens Without Changing JSON Data

**Files:**
- Modify: `quick_mgr.py`
- Modify: `tests/test_quick_mgr.py`

- [ ] **Step 1: Add a test that compiled sequences are cached by sequence content**

Append this test:

```python
def test_compile_sequence_is_cached_by_sequence_content(monkeypatch, tmp_path):
    keyboard_module = make_keyboard()
    quick_mgr = import_quick_mgr(monkeypatch, tmp_path, keyboard_module)
    manager = quick_mgr.QuickCastManager()

    first = manager.get_compiled_sequence(["上", "`20", "lp50"])
    second = manager.get_compiled_sequence(["上", "`20", "lp50"])

    assert first is second
    assert all(not isinstance(item, dict) for item in first)
```

- [ ] **Step 2: Run the test and verify it fails because `get_compiled_sequence` does not exist**

Run:

```bash
pytest tests/test_quick_mgr.py::test_compile_sequence_is_cached_by_sequence_content -q
```

Expected result before implementation:

```text
AttributeError: 'QuickCastManager' object has no attribute 'get_compiled_sequence'
```

- [ ] **Step 3: Add compiled token support**

At the top of `quick_mgr.py`, add:

```python
from dataclasses import dataclass
```

After the constants, add:

```python
@dataclass(frozen=True)
class SequenceToken:
    kind: str
    key: str = ""
    delay_ms: int = 0
```

In `QuickCastManager.__init__()`, add this before creating the mouse listener:

```python
        self.compiled_sequences = {}
```

Inside `QuickCastManager`, add:

```python
    def get_compiled_sequence(self, sequence):
        cache_key = tuple(sequence)
        if cache_key not in self.compiled_sequences:
            self.compiled_sequences[cache_key] = self.compile_sequence(sequence)
        return self.compiled_sequences[cache_key]

    def compile_sequence(self, sequence):
        tokens = []
        for raw_key in sequence:
            if raw_key[0] == "`":
                tokens.append(SequenceToken("pause", delay_ms=int(raw_key[1:])))
                continue

            prefix = raw_key[:2]
            if len(raw_key) >= 3 and prefix in HOLD_MOUSE_KEYS:
                tokens.append(SequenceToken("hold_mouse", key=HOLD_MOUSE_KEYS[prefix], delay_ms=int(raw_key[2:])))
                continue

            tokens.append(SequenceToken("tap_key", key=KEY_ALIASES.get(raw_key, raw_key)))
        return tuple(tokens)
```

- [ ] **Step 4: Update `build_timeline()` to accept compiled tokens**

Replace `build_timeline()` with:

```python
    def build_timeline(self, sequence):
        timeline_now = 0
        events = []
        for token in self.get_compiled_sequence(sequence):
            if token.kind == "pause":
                timeline_now += random.uniform(0.95, 1.051) * 0.001 * token.delay_ms
                continue

            if token.kind == "hold_mouse":
                delay = random.uniform(0.99, 1.01) * 0.001 * token.delay_ms
                events.append((token.key, True, timeline_now))
                events.append((token.key, False, timeline_now + delay))
                timeline_now += delay
                continue

            events.append((token.key, True, timeline_now))
            delay = self.settings["key_up_interval"] * random.uniform(0.952, 1.05)
            events.append((token.key, False, timeline_now + delay))
            timeline_now += self.settings["key_interval"] * random.uniform(0.82, 1.19)

        events.sort(key=lambda event: event[2])
        return events
```

- [ ] **Step 5: Run all tests**

Run:

```bash
pytest tests/test_quick_mgr.py -q
```

Expected result:

```text
5 passed
```

- [ ] **Step 6: Commit the compile cache**

```bash
git add quick_mgr.py tests/test_quick_mgr.py
git commit -m "perf: cache compiled macro sequences"
```

## Task 6: Move Listener Callbacks to a Single Macro Worker

**Files:**
- Modify: `quick_mgr.py`
- Modify: `tests/test_quick_mgr.py`

- [ ] **Step 1: Add a test that callbacks enqueue instead of executing inline**

Append this test:

```python
def test_run_combo_enqueues_without_executing_inline(monkeypatch, tmp_path):
    keyboard_module = make_keyboard()
    quick_mgr = import_quick_mgr(monkeypatch, tmp_path, keyboard_module)
    manager = quick_mgr.QuickCastManager()
    executed = []
    monkeypatch.setattr(manager, "_execute_combo", lambda combo: executed.append(combo))

    manager.run_combo({"trigger_key": "a", "sequence": ["a"]})

    assert executed == []
    assert manager.combo_queue.qsize() == 1
```

- [ ] **Step 2: Run the test and verify it fails because `run_combo()` executes inline**

Run:

```bash
pytest tests/test_quick_mgr.py::test_run_combo_enqueues_without_executing_inline -q
```

Expected result before implementation:

```text
FAILED tests/test_quick_mgr.py::test_run_combo_enqueues_without_executing_inline
```

- [ ] **Step 3: Add queue and thread imports**

At the top of `quick_mgr.py`, add:

```python
import queue
import threading
```

- [ ] **Step 4: Initialize the worker queue**

In `QuickCastManager.__init__()`, add:

```python
        self.combo_queue = queue.Queue(maxsize=1)
        self.combo_worker = threading.Thread(target=self._combo_worker_loop, daemon=True)
        self.combo_worker.start()
```

- [ ] **Step 5: Split execution from enqueueing**

Replace `run_combo()` with:

```python
    def run_combo(self, combo):
        try:
            self.combo_queue.put_nowait(combo)
        except queue.Full:
            return
```

Add this method:

```python
    def _combo_worker_loop(self):
        while True:
            combo = self.combo_queue.get()
            try:
                self._execute_combo(combo)
            finally:
                self.combo_queue.task_done()
```

Add this method:

```python
    def _execute_combo(self, combo):
        if self.lock:
            return
        self.lock = True
        try:
            while keyboard.is_pressed('alt'):
                time.sleep(0.001)
            self.execute_timeline(self.build_timeline(combo["sequence"]))
        finally:
            self.lock = False
```

- [ ] **Step 6: Update the lock-release test to call `_execute_combo()`**

Change the failing-backend test call from:

```python
        manager.run_combo({"trigger_key": "a", "sequence": ["a"]})
```

to:

```python
        manager._execute_combo({"trigger_key": "a", "sequence": ["a"]})
```

- [ ] **Step 7: Run all tests**

Run:

```bash
pytest tests/test_quick_mgr.py -q
```

Expected result:

```text
6 passed
```

- [ ] **Step 8: Commit the worker queue**

```bash
git add quick_mgr.py tests/test_quick_mgr.py
git commit -m "perf: run macros from a worker queue"
```

## Task 7: Avoid Unnecessary Writes on Startup

**Files:**
- Modify: `quick_mgr.py:15-19`
- Modify: `tests/test_quick_mgr.py`

- [ ] **Step 1: Add tests for default file creation and existing-file preservation**

Append these tests:

```python
def test_manager_creates_missing_json_files(monkeypatch, tmp_path):
    keyboard_module = make_keyboard()
    quick_mgr = import_quick_mgr(monkeypatch, tmp_path, keyboard_module)

    quick_mgr.QuickCastManager()

    assert (tmp_path / "setting.json").exists()
    assert (tmp_path / "quick.json").exists()


def test_manager_does_not_rewrite_existing_json_files(monkeypatch, tmp_path):
    (tmp_path / "setting.json").write_text('{"key_up_interval": 0.02, "key_interval": 0.09}', encoding="utf-8")
    (tmp_path / "quick.json").write_text('{"default": []}', encoding="utf-8")
    keyboard_module = make_keyboard()
    quick_mgr = import_quick_mgr(monkeypatch, tmp_path, keyboard_module)

    setting_before = (tmp_path / "setting.json").read_text(encoding="utf-8")
    quick_before = (tmp_path / "quick.json").read_text(encoding="utf-8")
    quick_mgr.QuickCastManager()

    assert (tmp_path / "setting.json").read_text(encoding="utf-8") == setting_before
    assert (tmp_path / "quick.json").read_text(encoding="utf-8") == quick_before
```

- [ ] **Step 2: Run the tests and verify the preservation test fails**

Run:

```bash
pytest tests/test_quick_mgr.py::test_manager_creates_missing_json_files tests/test_quick_mgr.py::test_manager_does_not_rewrite_existing_json_files -q
```

Expected result before implementation:

```text
1 failed, 1 passed
```

- [ ] **Step 3: Track whether files were missing during load**

Replace `load_settings()` with:

```python
    def load_settings(self):
        try:
            with open("setting.json", "r") as file:
                self.settings_file_missing = False
                return json.load(file)
        except FileNotFoundError:
            self.settings_file_missing = True
            return {"key_up_interval": 0.01, "key_interval": 0.08}
```

Replace `load_quick_casts()` with:

```python
    def load_quick_casts(self):
        try:
            with open("quick.json", "r") as file:
                self.quick_casts_file_missing = False
                return json.load(file)
        except FileNotFoundError:
            self.quick_casts_file_missing = True
            return {}
```

Replace the two unconditional saves in `__init__()` with:

```python
        if self.settings_file_missing:
            self.save_settings()
        if self.quick_casts_file_missing:
            self.save_quick_casts()
```

- [ ] **Step 4: Run all tests**

Run:

```bash
pytest tests/test_quick_mgr.py -q
```

Expected result:

```text
8 passed
```

- [ ] **Step 5: Commit the startup write reduction**

```bash
git add quick_mgr.py tests/test_quick_mgr.py
git commit -m "perf: avoid rewriting config on startup"
```

## Task 8: Final Verification

**Files:**
- No new source changes expected

- [ ] **Step 1: Run the unit tests**

Run:

```bash
pytest tests/test_quick_mgr.py -q
```

Expected result:

```text
8 passed
```

- [ ] **Step 2: Run syntax compilation**

Run:

```bash
python -m py_compile gui.py quick_mgr.py
```

Expected result:

```text
exit code 0 with no output
```

- [ ] **Step 3: Review the hot path manually**

Run:

```bash
rg -n "print\\(|def run_combo|def _execute_combo|def build_timeline|def execute_timeline|combo_queue|compiled_sequences" quick_mgr.py
```

Expected result:

```text
No print() calls remain in quick_mgr.py.
run_combo() only enqueues.
_execute_combo(), build_timeline(), execute_timeline(), combo_queue, and compiled_sequences are present.
```

- [ ] **Step 4: Commit any final adjustments**

If Step 3 exposes a missing or extra change, fix it, rerun Steps 1-3, then commit:

```bash
git add quick_mgr.py tests/test_quick_mgr.py
git commit -m "test: verify macro performance path"
```

