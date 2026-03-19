from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import Error, TimeoutError as PlaywrightTimeoutError, sync_playwright


# ================= Main(Classroom) 配置 =================
NDM_PATH = r"D:\Neat Download Manager\NeatDM.exe"
FFMPEG_PATH = r"D:\ffmpeg\bin\ffmpeg.exe"
BOT_DATA = os.path.join(os.getcwd(), "zju_bot_data")
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

BASE_DIR = os.path.join(os.getcwd(), "study_data")
VIDEO_DIR = os.path.join(BASE_DIR, "Videos")
AUDIO_DIR = os.path.join(BASE_DIR, "Audios")
TRANSCRIPT_DIR = os.path.join(BASE_DIR, "Transcripts")


# ================= Tingwu 配置 =================
DEFAULT_URL = "https://tingwu.aliyun.com/"
DEFAULT_FOLDERS_URL = "https://tingwu.aliyun.com/folders/0"
DEFAULT_STATE_PATH = Path(".auth/tingwu_storage_state.json")
DEFAULT_BROWSER = "msedge"
DEBUG_SCREENSHOT_PATH = Path(".auth/tingwu_transcribe_debug.png")
COURSE_HISTORY_PATH = Path(".auth/classroom_course_history.json")
APP_CONFIG_PATH = Path(".auth/app_config.json")

DEFAULT_EXPORT_ACTION_RETRY_SECONDS = 3
DEFAULT_SUBMIT_CONFIRM_SECONDS = 15
DEFAULT_START_TRANSCRIBE_WAIT_SECONDS = 1800
DEFAULT_UPLOAD_CONTROL_WAIT_SECONDS = 45
DEFAULT_POLL_SECONDS = 10
DEFAULT_MP4_WAIT_SECONDS = 1800

PROCESSING_STATUS_KEYWORDS = [
    "转写中",
    "处理中",
    "排队中",
    "识别中",
    "转码中",
    "上传中",
    "摘要生成中",
    "生成中",
]

SPINNER_FRAMES = ["|", "/", "-", "\\"]


def update_wait_status_line(
    prefix: str,
    started_at: float,
    deadline: float | None = None,
    extra: str = "",
    show_elapsed: bool = True,
) -> None:
    spinner = SPINNER_FRAMES[int(time.time() * 5) % len(SPINNER_FRAMES)]
    elapsed = max(0, int(time.time() - started_at))
    message = f"\r\x1b[2K{prefix} {spinner}"
    if show_elapsed:
        message += f" 已等待 {elapsed}s"
    if deadline is not None:
        remaining = max(0, int(deadline - time.time()))
        message += f"，剩余约 {remaining}s"
    if extra:
        message += f"，{extra}"
    sys.stdout.write(message)
    sys.stdout.flush()


def clear_wait_status_line() -> None:
    sys.stdout.write("\r\x1b[2K")
    sys.stdout.flush()


def finish_wait_status_line(message: str) -> None:
    clear_wait_status_line()
    print(message)


def ensure_parent_dir(file_path: Path) -> None:
    file_path.parent.mkdir(parents=True, exist_ok=True)


def sanitize_filename_component(text: str) -> str:
    # Windows 文件名非法字符统一替换，避免重命名和转码写盘失败。
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", (text or "").strip())
    cleaned = re.sub(r"\s+", "_", cleaned)
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    return cleaned or "未命名"


def load_course_history(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(raw, list):
            return []
        result: list[dict[str, str]] = []
        for item in raw:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).strip()
            url = str(item.get("url", "")).strip()
            if name and url:
                result.append({"name": name, "url": url})
        return result
    except Exception:
        return []


def save_course_history(path: Path, items: list[dict[str, str]]) -> None:
    ensure_parent_dir(path)
    path.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")


def upsert_course_history(path: Path, course_name: str, course_url: str) -> None:
    name = (course_name or "").strip()
    url = (course_url or "").strip()
    if not name or not url:
        return

    items = load_course_history(path)
    updated = False
    for item in items:
        if item.get("name") == name:
            item["url"] = url
            updated = True
            break
    if not updated:
        items.append({"name": name, "url": url})
    save_course_history(path, items)


def parse_course_id_from_url(url: str) -> str | None:
    match = re.search(r"course_id=([^&]+)", url or "")
    return match.group(1).strip() if match else None


def parse_lesson_selection(selection_text: str, total: int) -> list[int] | None:
    """Parse user input like '1 2 5', '1,3,5', or '1-3,8' into zero-based indexes."""
    raw = (selection_text or "").strip()
    if not raw:
        return None

    normalized = (
        raw.replace("，", ",")
        .replace("、", ",")
        .replace("；", ",")
        .replace(";", ",")
    )
    normalized = re.sub(r"\s+", ",", normalized)
    parts = [part.strip() for part in normalized.split(",") if part.strip()]
    if not parts:
        return None

    picked: list[int] = []
    seen: set[int] = set()
    for part in parts:
        if "-" in part:
            start_text, end_text = [token.strip() for token in part.split("-", 1)]
            if not (start_text.isdigit() and end_text.isdigit()):
                return None
            start_num = int(start_text)
            end_num = int(end_text)
            if start_num <= 0 or end_num <= 0:
                return None
            if start_num > end_num:
                start_num, end_num = end_num, start_num
            for num in range(start_num, end_num + 1):
                idx = num - 1
                if idx < 0 or idx >= total:
                    return None
                if idx not in seen:
                    seen.add(idx)
                    picked.append(idx)
            continue

        if not part.isdigit():
            return None
        idx = int(part) - 1
        if idx < 0 or idx >= total:
            return None
        if idx not in seen:
            seen.add(idx)
            picked.append(idx)

    return picked if picked else None


def load_app_config(path: Path) -> dict[str, str]:
    defaults = {
        "ndm_path": NDM_PATH,
        "ffmpeg_path": FFMPEG_PATH,
        "base_dir": BASE_DIR,
        "video_dir": VIDEO_DIR,
        "audio_dir": AUDIO_DIR,
        "transcript_dir": TRANSCRIPT_DIR,
    }
    if not path.exists():
        return defaults

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return defaults
    if not isinstance(raw, dict):
        return defaults

    merged = defaults.copy()
    for key in defaults:
        value = raw.get(key)
        if isinstance(value, str) and value.strip():
            merged[key] = value.strip()
    return merged


def save_app_config(path: Path, config: dict[str, str]) -> None:
    ensure_parent_dir(path)
    path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def prompt_path_settings(config: dict[str, str]) -> dict[str, str]:
    prompts = [
        ("ndm_path", "NDM 可执行文件路径"),
        ("ffmpeg_path", "ffmpeg 可执行文件路径"),
        ("base_dir", "工作根目录(BASE_DIR)"),
        ("video_dir", "mp4 存放目录(VIDEO_DIR)"),
        ("audio_dir", "mp3 存放目录(AUDIO_DIR)"),
        ("transcript_dir", "文本存放目录(TRANSCRIPT_DIR)"),
    ]

    updated = config.copy()
    print("\n[CONFIG] 请输入路径，直接回车表示沿用当前值。")
    for key, title in prompts:
        current = updated.get(key, "")
        value = input(f"[CONFIG] {title} [{current}]: ").strip()
        if value:
            updated[key] = value
    return updated


def apply_app_config(config: dict[str, str]) -> None:
    global NDM_PATH, FFMPEG_PATH, BASE_DIR, VIDEO_DIR, AUDIO_DIR, TRANSCRIPT_DIR

    NDM_PATH = config["ndm_path"]
    FFMPEG_PATH = config["ffmpeg_path"]
    BASE_DIR = config["base_dir"]
    VIDEO_DIR = config["video_dir"]
    AUDIO_DIR = config["audio_dir"]
    TRANSCRIPT_DIR = config["transcript_dir"]


def print_runtime_config() -> None:
    print("\n[CONFIG] 当前运行配置:")
    print(f"[CONFIG] NDM_PATH       : {NDM_PATH}")
    print(f"[CONFIG] FFMPEG_PATH    : {FFMPEG_PATH}")
    print(f"[CONFIG] BASE_DIR       : {BASE_DIR}")
    print(f"[CONFIG] VIDEO_DIR      : {VIDEO_DIR}")
    print(f"[CONFIG] AUDIO_DIR      : {AUDIO_DIR}")
    print(f"[CONFIG] TRANSCRIPT_DIR : {TRANSCRIPT_DIR}")


def setup_runtime_config(force_prompt: bool = False) -> None:
    config = load_app_config(APP_CONFIG_PATH)
    if force_prompt or not APP_CONFIG_PATH.exists():
        config = prompt_path_settings(config)
        save_app_config(APP_CONFIG_PATH, config)

    apply_app_config(config)
    for directory in [BASE_DIR, VIDEO_DIR, AUDIO_DIR, TRANSCRIPT_DIR]:
        if directory:
            os.makedirs(directory, exist_ok=True)


def is_tingwu_login_state_valid(state_path: Path, browser_name: str) -> bool:
    if not state_path.exists():
        return False

    try:
        with sync_playwright() as p:
            browser = launch_browser(p, headless=True, browser_name=browser_name)
            context = browser.new_context(storage_state=str(state_path))
            page = context.new_page()
            safe_navigate(page, DEFAULT_URL)
            started_at = time.time()
            for _ in range(4):
                update_wait_status_line("[AUTH] 正在校验登录态", started_at, extra="等待页面稳定")
                page.wait_for_timeout(500)
            clear_wait_status_line()

            current_url = page.url.lower()
            if "login" in current_url:
                browser.close()
                return False

            body_text = ""
            try:
                body_text = page.locator("body").inner_text(timeout=1200)
            except Exception:
                pass

            login_hints = ["登录", "手机号", "验证码", "请先登录"]
            if any(hint in body_text for hint in login_hints):
                browser.close()
                return False

            browser.close()
            return True
    except Exception:
        return False


def ensure_tingwu_login_ready(state_path: Path, browser_name: str) -> bool:
    print("[AUTH] 正在检测听悟登录态...")
    if is_tingwu_login_state_valid(state_path, browser_name):
        print("[AUTH] 登录态有效。")
        return True

    print("[AUTH] 检测到听悟登录态缺失或失效，需要重新登录。")
    answer = input("[AUTH] 现在打开听悟登录页进行登录吗？(Y/n): ").strip().lower()
    if answer in {"n", "no"}:
        print("[AUTH] 已取消登录。")
        return False

    init_login_state(DEFAULT_URL, state_path, timeout_minutes=10, browser_name=browser_name)
    if not is_tingwu_login_state_valid(state_path, browser_name):
        print("[AUTH] 登录态仍不可用，请重试。")
        return False
    print("[AUTH] 登录态检查通过。")
    return True


def interactive_main_menu(state_path: Path) -> int:
    while True:
        print("\n[MENU] 请选择操作:")
        print("  1) 全链条：下载 mp4 -> 转 mp3 -> 上传转写并导出文本")
        print("  2) 仅下载 mp4 + 转 mp3")
        print("  3) 仅上传指定 mp3 并自动导出文本")
        print("  4) 修改并保存路径配置")
        print("  0) 退出")
        choice = input("[MENU] 输入编号: ").strip()

        if choice == "1":
            print_runtime_config()
            if not ensure_tingwu_login_ready(state_path, DEFAULT_BROWSER):
                continue
            return classroom_interactive_flow(auto_transcribe=True)
        if choice == "2":
            print_runtime_config()
            return classroom_interactive_flow(auto_transcribe=False)
        if choice == "3":
            print_runtime_config()
            if not ensure_tingwu_login_ready(state_path, DEFAULT_BROWSER):
                continue
            file_path = input("[MENU] 请输入 mp3 文件路径: ").strip()
            if not file_path:
                print("[MENU] 路径不能为空。")
                continue
            target_dir_input = input(f"[MENU] 文本输出目录(回车默认 {TRANSCRIPT_DIR}): ").strip()
            target_dir = Path(target_dir_input) if target_dir_input else Path(TRANSCRIPT_DIR)
            transcribe_then_export(
                transcribe_url=DEFAULT_URL,
                folders_url=DEFAULT_FOLDERS_URL,
                state_path=state_path,
                file_path=Path(file_path),
                record_title="",
                browser_name=DEFAULT_BROWSER,
                headless=True,
                kickoff_wait_seconds=DEFAULT_START_TRANSCRIBE_WAIT_SECONDS,
                wait_ready_seconds=1800,
                poll_seconds=DEFAULT_POLL_SECONDS,
                download_dir=target_dir,
            )
            return 0
        if choice == "4":
            setup_runtime_config(force_prompt=True)
            print_runtime_config()
            continue
        if choice == "0":
            print("[MENU] 已退出。")
            return 0

        print("[MENU] 输入无效，请重新选择。")


def get_latest_file(directory: str, extension: str = ".mp4") -> str | None:
    files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(extension)]
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def wait_for_new_or_updated_mp4(directory: Path, before_files: dict[str, tuple[float, int]], timeout_seconds: int = DEFAULT_MP4_WAIT_SECONDS) -> Path | None:
    deadline = time.time() + max(10, timeout_seconds)
    started_at = time.time()

    while time.time() < deadline:
        # 下载等待仅显示单行动画和已等待时长，不显示剩余时间。
        update_wait_status_line("[CLASSROOM] 等待 mp4 下载完成", started_at, show_elapsed=False)
        candidates: list[Path] = []
        for p in directory.glob("*.mp4"):
            if not p.is_file():
                continue
            sig = (p.stat().st_mtime, p.stat().st_size)
            if p.name not in before_files or before_files[p.name] != sig:
                candidates.append(p)

        if candidates:
            latest = max(candidates, key=lambda p: p.stat().st_mtime)
            first_sig = (latest.stat().st_mtime, latest.stat().st_size)
            time.sleep(2)
            second_sig = (latest.stat().st_mtime, latest.stat().st_size)
            if first_sig == second_sig:
                finish_wait_status_line("[CLASSROOM] 已检测到稳定的 mp4 下载结果。")
                return latest
        time.sleep(2)

    finish_wait_status_line("[CLASSROOM] 等待 mp4 下载结果超时。")
    return None


def move_file_to_dir(source: Path, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    destination = target_dir / source.name
    # 同名文件按用户期望直接覆盖。
    os.replace(str(source), str(destination))
    return destination


def launch_classroom_context(playwright, is_headless: bool):
    return playwright.chromium.launch_persistent_context(
        user_data_dir=BOT_DATA,
        executable_path=EDGE_EXE,
        headless=is_headless,
        args=["--remote-debugging-port=9222"],
    )


def launch_browser(playwright, headless: bool, browser_name: str):
    if browser_name == "chromium":
        return playwright.chromium.launch(headless=headless)
    return playwright.chromium.launch(headless=headless, channel=browser_name)


def click_locator_best_effort(locator) -> bool:
    try:
        count = locator.count()
        if count == 0:
            return False
        for idx in range(min(count, 10)):
            candidate = locator.nth(idx)
            try:
                if candidate.is_visible(timeout=500):
                    candidate.click(timeout=5000)
                    return True
            except Exception:
                continue
        locator.first.click(timeout=5000, force=True)
        return True
    except Exception:
        return False


def click_card_more_by_position(page, card_locator) -> bool:
    try:
        if card_locator.count() == 0:
            return False
        box = card_locator.first.bounding_box()
        if not box:
            return False
        x = box["x"] + box["width"] - 26
        y = box["y"] + box["height"] - 26
        page.mouse.click(x, y)
        return True
    except Exception:
        return False


def click_export_menu_item_with_retry(page, retries: int = 6, interval_ms: int = 250) -> bool:
    for _ in range(max(1, retries)):
        export_item = page.locator(
            ".ant-dropdown:visible span:has-text('导出'), .ant-dropdown:visible div:has-text('导出')"
        ).first
        if click_locator_best_effort(export_item):
            return True
        page.wait_for_timeout(max(50, interval_ms))
    return False


def click_first_visible_anywhere(page, labels: list[str]) -> bool:
    scopes = [page, *page.frames]
    for label in labels:
        pattern = re.compile(r"\s*".join(map(re.escape, label.split())), re.IGNORECASE)
        for scope in scopes:
            try:
                locator = scope.get_by_text(pattern)
                if click_locator_best_effort(locator):
                    return True
            except Exception:
                continue
            try:
                container = scope.locator(f"*:has-text('{label}')")
                if click_locator_best_effort(container):
                    return True
            except Exception:
                continue
    return False


def open_upload_entry(page) -> bool:
    for _ in range(10):
        if click_first_visible_anywhere(page, ["上传音视频"]):
            return True
        page.wait_for_timeout(500)
    return click_first_visible_anywhere(
        page,
        [
            "上传 本地 音视频 文件",
            "上传本地音视频文件",
            "点击 / 拖拽 本地音视频文件到这里",
        ],
    )


def find_file_input_anywhere(page):
    scopes = [page, *page.frames]
    for scope in scopes:
        try:
            locator = scope.locator("input[type='file']")
            if locator.count() > 0:
                return locator.first
        except Exception:
            continue
    return None


def upload_file_with_file_chooser(page, file_path: Path, timeout_seconds: int = 30) -> bool:
    triggers = [
        "点击 / 拖拽 本地音视频文件到这里",
        "上传本地音视频文件",
        "点击/拖拽本地音视频文件到这里",
    ]

    def has_trigger(label: str) -> bool:
        pattern = re.compile(r"\s*".join(map(re.escape, label.split())), re.IGNORECASE)
        for scope in [page, *page.frames]:
            try:
                locator = scope.get_by_text(pattern)
                if locator.count() > 0:
                    return True
            except Exception:
                pass
            try:
                container = scope.locator(f"*:has-text('{label}')")
                if container.count() > 0:
                    return True
            except Exception:
                pass
        return False

    deadline = time.time() + max(3, timeout_seconds)
    started_at = time.time()
    while time.time() < deadline:
        update_wait_status_line("[TRANSCRIBE] file chooser 兜底上传中", started_at, deadline)

        for label in triggers:
            if not has_trigger(label):
                continue
            try:
                with page.expect_file_chooser(timeout=2500) as chooser_info:
                    clicked = click_first_visible_anywhere(page, [label])
                if not clicked:
                    continue
                chooser = chooser_info.value
                chooser.set_files(str(file_path))
                finish_wait_status_line("[TRANSCRIBE] file chooser 上传已触发。")
                return True
            except PlaywrightTimeoutError:
                continue
            except Exception:
                continue

        open_upload_entry(page)
        page.wait_for_timeout(800)
    finish_wait_status_line("[TRANSCRIBE] file chooser 兜底上传超时。")
    return False


def wait_for_file_input_anywhere(page, timeout_seconds: int):
    deadline = time.time() + max(3, timeout_seconds)
    started_at = time.time()
    while time.time() < deadline:
        file_input = find_file_input_anywhere(page)
        if file_input is not None:
            finish_wait_status_line("[TRANSCRIBE] 已检测到上传控件。")
            return file_input
        update_wait_status_line("[TRANSCRIBE] 等待上传控件出现", started_at, deadline)
        page.wait_for_timeout(800)
    finish_wait_status_line("[TRANSCRIBE] 上传控件等待超时。")
    return None


def wait_for_transcribe_submission(page, timeout_seconds: int) -> bool:
    deadline = time.time() + max(1, timeout_seconds)
    started_at = time.time()
    print(f"[TRANSCRIBE] 提交确认中，最长等待 {max(1, timeout_seconds)} 秒。")
    upload_seen = False
    while time.time() < deadline:
        update_wait_status_line("[TRANSCRIBE] 等待提交生效", started_at, deadline)
        try:
            body_text = page.locator("body").inner_text(timeout=1200)
        except Exception:
            body_text = ""

        for keyword in PROCESSING_STATUS_KEYWORDS:
            if keyword in body_text:
                # "上传中" 仅表示文件仍在上传，不能作为可关闭浏览器的成功信号。
                if keyword == "上传中":
                    if not upload_seen:
                        clear_wait_status_line()
                        print("[TRANSCRIBE] 检测到上传中，继续等待进入转写阶段。")
                        upload_seen = True
                    break

                finish_wait_status_line(f"[TRANSCRIBE] 检测到任务状态: {keyword}，提交已生效。")
                return True

        try:
            start_btn = page.get_by_text(re.compile(r"开始\s*转写"))
            if start_btn.count() == 0:
                finish_wait_status_line("[TRANSCRIBE] 开始转写按钮已消失，视为提交已生效。")
                return True
        except Exception:
            pass

        time.sleep(1)

    finish_wait_status_line("[TRANSCRIBE] 未检测到明显提交信号，将继续后续导出轮询。")
    return False


def detect_processing_status(card) -> str | None:
    try:
        text = card.inner_text(timeout=1000)
    except Exception:
        return None

    for keyword in PROCESSING_STATUS_KEYWORDS:
        if keyword in text:
            return keyword
    return None


def build_record_title_candidates(record_title: str) -> list[str]:
    raw = (record_title or "").strip()
    if not raw:
        return []

    stem = Path(raw).stem
    variants = [
        raw,
        stem,
        f"{stem}.mp3",
        stem.replace("_", " "),
        stem.replace(" ", "_"),
    ]

    candidates: list[str] = []
    for item in variants:
        item = item.strip()
        if item and item not in candidates:
            candidates.append(item)
    return candidates


def normalize_record_key(text: str) -> str:
    stem = Path((text or "").strip()).stem
    return re.sub(r"[\s_\-]+", "", stem).lower()


def is_export_filename_match(record_title: str, filename: str) -> bool:
    target_key = normalize_record_key(record_title)
    file_key = normalize_record_key(filename)
    if not target_key:
        return True
    return target_key in file_key


def find_record_card(page, record_title: str):
    candidates = build_record_title_candidates(record_title)
    if not candidates:
        return None, None

    candidate_keys = [normalize_record_key(candidate) for candidate in candidates if candidate.strip()]

    try:
        card_locator = page.locator(
            "div:has(.ant-dropdown-trigger.edits), "
            "div:has(.edits.ant-dropdown-trigger), "
            "div:has(svg.edits)"
        )
        count = min(card_locator.count(), 200)
        for idx in range(count):
            card = card_locator.nth(idx)
            try:
                card_text = card.inner_text(timeout=800)
            except Exception:
                continue

            card_key = normalize_record_key(card_text)
            for candidate, candidate_key in zip(candidates, candidate_keys):
                if candidate_key and candidate_key in card_key:
                    return card, candidate
    except Exception:
        pass

    stem = Path(candidates[0]).stem
    escaped = re.escape(stem)
    escaped = escaped.replace(r"\\ ", r"[ _]+")
    escaped = escaped.replace(r"_", r"[ _]+")
    pattern = re.compile(rf"{escaped}(?:\\.mp3)?", re.IGNORECASE)
    try:
        card = page.locator("div", has_text=pattern).first
        if card.count() > 0:
            return card, f"regex:{pattern.pattern}"
    except Exception:
        pass

    return None, None


def safe_navigate(page, url: str, wait_until: str = "domcontentloaded", retries: int = 3) -> None:
    last_exc = None
    for attempt in range(1, max(1, retries) + 1):
        try:
            page.goto(url, wait_until=wait_until)
            return
        except Exception as exc:
            last_exc = exc
            message = str(exc)
            is_abort = "ERR_ABORTED" in message or "net::ERR_ABORTED" in message
            if attempt >= retries or not is_abort:
                raise
            print(f"[NAV] 导航被中断，重试 {attempt}/{retries}: {url}")
            page.wait_for_timeout(500)
    if last_exc is not None:
        raise last_exc


def upload_and_start_transcribe_on_page(page, file_path: Path, submit_timeout_seconds: int) -> None:
    print("[TRANSCRIBE] 打开上传入口。")
    opened = open_upload_entry(page)
    if not opened:
        ensure_parent_dir(DEBUG_SCREENSHOT_PATH)
        page.screenshot(path=str(DEBUG_SCREENSHOT_PATH), full_page=True)
        print("[TRANSCRIBE] 未找到上传入口按钮，请使用 --show 排查页面元素。")
        print(f"[TRANSCRIBE] 已保存调试截图: {DEBUG_SCREENSHOT_PATH}")
        sys.exit(1)

    page.wait_for_timeout(800)
    click_first_visible_anywhere(page, ["上传本地音视频文件", "上传 本地 音视频 文件"])

    print(f"[TRANSCRIBE] 上传文件: {file_path}")
    print(f"[TRANSCRIBE] 等待上传控件，最长 {DEFAULT_UPLOAD_CONTROL_WAIT_SECONDS} 秒。")
    file_input = wait_for_file_input_anywhere(page, timeout_seconds=DEFAULT_UPLOAD_CONTROL_WAIT_SECONDS)

    uploaded = False
    if file_input is not None:
        try:
            file_input.set_input_files(str(file_path))
            uploaded = True
        except Exception:
            uploaded = False

    if not uploaded:
        print("[TRANSCRIBE] 未找到 file input，尝试 file chooser 循环兜底上传。")
        uploaded = upload_file_with_file_chooser(
            page,
            file_path,
            timeout_seconds=DEFAULT_UPLOAD_CONTROL_WAIT_SECONDS,
        )

    if not uploaded:
        ensure_parent_dir(DEBUG_SCREENSHOT_PATH)
        page.screenshot(path=str(DEBUG_SCREENSHOT_PATH), full_page=True)
        print("[TRANSCRIBE] 上传控件等待与 file chooser 兜底均失败，请检查登录状态或页面结构。")
        print(f"[TRANSCRIBE] 已保存调试截图: {DEBUG_SCREENSHOT_PATH}")
        sys.exit(1)

    start_timeout = max(DEFAULT_SUBMIT_CONFIRM_SECONDS, submit_timeout_seconds)
    print(f"[TRANSCRIBE] 等待“开始转写”可点击，最长 {start_timeout} 秒。")
    deadline = time.time() + start_timeout
    while time.time() < deadline:
        if click_first_visible_anywhere(page, ["开始转写"]):
            print("[TRANSCRIBE] 已触发转写。")
            wait_for_transcribe_submission(page, timeout_seconds=start_timeout)
            return
        time.sleep(2)

    ensure_parent_dir(DEBUG_SCREENSHOT_PATH)
    page.screenshot(path=str(DEBUG_SCREENSHOT_PATH), full_page=True)
    print("[TRANSCRIBE] 超时：未等到“开始转写”可点击。")
    print(f"[TRANSCRIBE] 已保存调试截图: {DEBUG_SCREENSHOT_PATH}")
    sys.exit(1)


def init_login_state(url: str, state_path: Path, timeout_minutes: int, browser_name: str) -> None:
    ensure_parent_dir(state_path)
    with sync_playwright() as p:
        browser = launch_browser(p, headless=False, browser_name=browser_name)
        context = browser.new_context()
        page = context.new_page()

        print(f"[INIT] 打开登录页: {url}")
        print(f"[INIT] 浏览器通道: {browser_name}")
        safe_navigate(page, url)
        print("[INIT] 请在浏览器中手动完成登录。")
        print(f"[INIT] 参考登录时长: {timeout_minutes} 分钟。")
        print("[INIT] 登录完成后，回到终端按 Enter 保存登录状态。")

        try:
            page.wait_for_timeout(1500)
            input()
        except KeyboardInterrupt:
            browser.close()
            print("\n[INIT] 已取消。")
            return

        try:
            page.wait_for_timeout(1000)
            context.storage_state(path=str(state_path))
            print(f"[INIT] 登录状态已保存: {state_path}")
        finally:
            browser.close()


def run_in_background(url: str, state_path: Path, keep_alive_seconds: int, browser_name: str, headless: bool) -> None:
    if not state_path.exists():
        print(f"[RUN] 未找到登录状态文件: {state_path}")
        print("[RUN] 请先执行: python study.py init")
        sys.exit(1)

    with sync_playwright() as p:
        browser = launch_browser(p, headless=headless, browser_name=browser_name)
        context = browser.new_context(storage_state=str(state_path))
        page = context.new_page()

        mode = "无头后台" if headless else "可见窗口"
        print(f"[RUN] 浏览器已启动({mode})，访问: {url}")
        print(f"[RUN] 浏览器通道: {browser_name}")
        safe_navigate(page, url)
        print("[RUN] 已尝试复用登录态进入页面。")

        if keep_alive_seconds <= 0:
            print("[RUN] 后台挂起模式: 按 Ctrl+C 结束。")
            try:
                while True:
                    time.sleep(3)
            except KeyboardInterrupt:
                print("\n[RUN] 收到中断，正在退出。")
        else:
            print(f"[RUN] 将保持运行 {keep_alive_seconds} 秒。")
            time.sleep(keep_alive_seconds)

        browser.close()
        print("[RUN] 浏览器已关闭。")


def upload_and_start_transcribe(url: str, state_path: Path, file_path: Path, browser_name: str, headless: bool, wait_seconds: int) -> None:
    if not state_path.exists():
        print(f"[TRANSCRIBE] 未找到登录状态文件: {state_path}")
        print("[TRANSCRIBE] 请先执行: python study.py init")
        sys.exit(1)
    if not file_path.exists() or not file_path.is_file():
        print(f"[TRANSCRIBE] 音频文件不存在: {file_path}")
        sys.exit(1)
    if file_path.suffix.lower() != ".mp3":
        print(f"[TRANSCRIBE] 文件不是 mp3: {file_path}")
        sys.exit(1)

    with sync_playwright() as p:
        browser = launch_browser(p, headless=headless, browser_name=browser_name)
        context = browser.new_context(storage_state=str(state_path))
        page = context.new_page()

        mode = "无头后台" if headless else "可见窗口"
        print(f"[TRANSCRIBE] 浏览器已启动({mode})，通道: {browser_name}")
        safe_navigate(page, url)
        page.wait_for_timeout(2500)

        upload_and_start_transcribe_on_page(
            page,
            file_path=file_path,
            submit_timeout_seconds=DEFAULT_START_TRANSCRIBE_WAIT_SECONDS,
        )
        if wait_seconds > 0:
            print(f"[TRANSCRIBE] 额外等待 {wait_seconds} 秒后退出。")
            time.sleep(wait_seconds)
        else:
            print("[TRANSCRIBE] 常驻模式: 不自动退出，按 Ctrl+C 结束。")
            try:
                while True:
                    time.sleep(3)
            except KeyboardInterrupt:
                print("\n[TRANSCRIBE] 收到中断，正在退出。")

        browser.close()
        print("[TRANSCRIBE] 浏览器已关闭。")


def export_to_local_on_page(
    page,
    url: str,
    record_title: str,
    wait_ready_seconds: int,
    status_poll_seconds: int,
    action_retry_seconds: int,
    download_dir: Path,
    use_reload: bool = False,
) -> None:
    deadline = time.time() + wait_ready_seconds
    started_at = time.time()
    system_download_dir = Path.home() / "Downloads"

    while True:
        if time.time() > deadline:
            clear_wait_status_line()
            ensure_parent_dir(DEBUG_SCREENSHOT_PATH)
            page.screenshot(path=str(DEBUG_SCREENSHOT_PATH), full_page=True)
            print("[EXPORT] 超时：在限定时间内未完成导出。")
            print(f"[EXPORT] 调试截图: {DEBUG_SCREENSHOT_PATH}")
            sys.exit(1)

        update_wait_status_line("[EXPORT] 轮询导出状态", started_at, deadline)

        if use_reload:
            safe_navigate(page, page.url if page.url else url)
        else:
            safe_navigate(page, url)
        page.wait_for_timeout(1800)

        card, matched_title = find_record_card(page, record_title)
        if card is None:
            time.sleep(status_poll_seconds)
            continue

        if matched_title and matched_title != record_title:
            print(f"[EXPORT] 已匹配到标题变体: {matched_title}")

        processing_status = detect_processing_status(card)
        if processing_status is not None:
            update_wait_status_line("[EXPORT] 轮询导出状态", started_at, deadline, extra=f"记录处理中({processing_status})")
            time.sleep(status_poll_seconds)
            continue

        menu_opened = click_locator_best_effort(
            card.locator(".ant-dropdown-trigger.edits, .edits.ant-dropdown-trigger, svg.edits")
        )
        if not menu_opened:
            menu_opened = click_card_more_by_position(page, card)
        if not menu_opened:
            print("[EXPORT] 未打开三点菜单，继续重试。")
            time.sleep(action_retry_seconds)
            continue

        if not click_export_menu_item_with_retry(page):
            print("[EXPORT] 菜单中暂不可导出（已短时重试），继续下一轮。")
            time.sleep(action_retry_seconds)
            continue

        page.wait_for_timeout(150)
        export_modal = page.locator("div.ant-modal:visible").first
        if export_modal.count() == 0:
            print("[EXPORT] 导出弹窗未出现，继续重试。")
            time.sleep(action_retry_seconds)
            continue

        try:
            before_files = {
                p.name: (p.stat().st_mtime, p.stat().st_size)
                for p in download_dir.glob("*")
                if p.is_file()
            }
            before_system_files = {
                p.name: (p.stat().st_mtime, p.stat().st_size)
                for p in system_download_dir.glob("*")
                if p.is_file()
            }
            with page.expect_download(timeout=30000) as dl_info:
                clicked = click_locator_best_effort(
                    export_modal.locator(
                        "button:has-text('导出到本地'), button:has-text('导出到 本地'), .PanelFooterCombo_DownloadBtn"
                    )
                )
                if not clicked:
                    raise RuntimeError("未找到“导出到本地”按钮")
            download = dl_info.value
            target = download_dir / download.suggested_filename
            if target.exists():
                target.unlink()
            download.save_as(str(target))
            if not is_export_filename_match(record_title, target.name):
                clear_wait_status_line()
                print(
                    f"[EXPORT] 下载文件与目标记录不匹配，继续等待。"
                    f" 目标: {record_title}，实际: {target.name}"
                )
                try:
                    target.unlink()
                except Exception:
                    pass
                time.sleep(action_retry_seconds)
                continue
            finish_wait_status_line(f"[EXPORT] 下载完成: {target}")
            return
        except Exception as exc:
            for _ in range(15):
                after_files: list[Path] = []
                for p in download_dir.glob("*"):
                    if not p.is_file():
                        continue
                    current_sig = (p.stat().st_mtime, p.stat().st_size)
                    if p.name not in before_files or before_files[p.name] != current_sig:
                        after_files.append(p)
                if after_files:
                    latest = max(after_files, key=lambda p: p.stat().st_mtime)
                    if not is_export_filename_match(record_title, latest.name):
                        clear_wait_status_line()
                        print(
                            f"[EXPORT] 检测到下载文件但与目标记录不匹配，继续等待。"
                            f" 目标: {record_title}，实际: {latest.name}"
                        )
                        time.sleep(1)
                        continue
                    finish_wait_status_line(f"[EXPORT] 检测到下载文件(新增或更新): {latest}")
                    print(f"[EXPORT] 检测到下载文件(新增或更新): {latest}")
                    return

                # 某些浏览器通道会忽略下载目录，直接落到系统默认下载目录。
                system_after_files: list[Path] = []
                for p in system_download_dir.glob("*"):
                    if not p.is_file():
                        continue
                    current_sig = (p.stat().st_mtime, p.stat().st_size)
                    if p.name not in before_system_files or before_system_files[p.name] != current_sig:
                        system_after_files.append(p)
                if system_after_files:
                    latest_system = max(system_after_files, key=lambda p: p.stat().st_mtime)
                    if not is_export_filename_match(record_title, latest_system.name):
                        clear_wait_status_line()
                        print(
                            f"[EXPORT] 系统下载目录出现不匹配文件，继续等待。"
                            f" 目标: {record_title}，实际: {latest_system.name}"
                        )
                        time.sleep(1)
                        continue
                    if latest_system.parent.resolve() != download_dir.resolve():
                        moved = move_file_to_dir(latest_system, download_dir)
                        finish_wait_status_line(f"[EXPORT] 检测到文件下载到默认目录，已迁移: {moved}")
                        print(f"[EXPORT] 检测到文件下载到默认目录，已迁移: {moved}")
                    else:
                        finish_wait_status_line(f"[EXPORT] 检测到下载文件(新增或更新): {latest_system}")
                        print(f"[EXPORT] 检测到下载文件(新增或更新): {latest_system}")
                    return
                time.sleep(1)

            clear_wait_status_line()
            print(f"[EXPORT] 导出触发失败，继续重试: {exc}")
            time.sleep(action_retry_seconds)


def export_to_local(
    url: str,
    state_path: Path,
    record_title: str,
    browser_name: str,
    headless: bool,
    wait_ready_seconds: int,
    poll_seconds: int,
    download_dir: Path,
) -> None:
    if not state_path.exists():
        print(f"[EXPORT] 未找到登录状态文件: {state_path}")
        print("[EXPORT] 请先执行: python study.py init")
        sys.exit(1)
    if not record_title.strip():
        print("[EXPORT] 请提供 --record-title，用于定位默认文件夹中的记录卡片。")
        sys.exit(1)

    download_dir.mkdir(parents=True, exist_ok=True)
    status_poll_seconds = max(1, poll_seconds)
    action_retry_seconds = DEFAULT_EXPORT_ACTION_RETRY_SECONDS

    with sync_playwright() as p:
        browser = launch_browser(p, headless=headless, browser_name=browser_name)
        context = browser.new_context(storage_state=str(state_path), accept_downloads=True)
        page = context.new_page()

        mode = "无头后台" if headless else "可见窗口"
        print(f"[EXPORT] 浏览器已启动({mode})，通道: {browser_name}")
        print(f"[EXPORT] 目标记录: {record_title}")
        print(f"[EXPORT] 下载目录: {download_dir}")
        export_to_local_on_page(
            page=page,
            url=url,
            record_title=record_title,
            wait_ready_seconds=wait_ready_seconds,
            status_poll_seconds=status_poll_seconds,
            action_retry_seconds=action_retry_seconds,
            download_dir=download_dir,
            use_reload=False,
        )

        browser.close()
        print("[EXPORT] 浏览器已关闭。")


def transcribe_then_export(
    transcribe_url: str,
    folders_url: str,
    state_path: Path,
    file_path: Path,
    record_title: str,
    browser_name: str,
    headless: bool,
    kickoff_wait_seconds: int,
    wait_ready_seconds: int,
    poll_seconds: int,
    download_dir: Path,
) -> None:
    uploaded_file_stem = file_path.stem.strip()
    user_record_title = record_title.strip()
    if not uploaded_file_stem:
        print("[AUTO] 无法确定记录标题，请通过 --record-title 指定。")
        sys.exit(1)

    # auto 流程是"刚上传 -> 立刻导出"，实际网页记录标题通常就是上传文件名。
    # 因此即使用户传了不同标题，也优先按上传文件名匹配，避免定位不到卡片而超时。
    resolved_title = uploaded_file_stem
    if user_record_title and user_record_title != uploaded_file_stem:
        print("[AUTO] 检测到自定义记录标题与上传文件名不一致。")
        print(f"[AUTO] 为提高成功率，将按上传文件名匹配: {uploaded_file_stem}")
    elif user_record_title:
        resolved_title = user_record_title

    if not state_path.exists():
        print(f"[AUTO] 未找到登录状态文件: {state_path}")
        print("[AUTO] 请先执行: python study.py init")
        sys.exit(1)
    if not file_path.exists() or not file_path.is_file():
        print(f"[AUTO] 音频文件不存在: {file_path}")
        sys.exit(1)
    if file_path.suffix.lower() != ".mp3":
        print(f"[AUTO] 文件不是 mp3: {file_path}")
        sys.exit(1)

    mode = "无头后台" if headless else "可见窗口"
    print(f"[AUTO] 同一会话执行：上传转写后直接进入默认文件夹轮询（{mode}）。")
    print(f"[AUTO] 目标记录: {resolved_title}")
    print(f"[AUTO] 下载目录: {download_dir}")

    download_dir.mkdir(parents=True, exist_ok=True)
    status_poll_seconds = max(1, poll_seconds)
    action_retry_seconds = DEFAULT_EXPORT_ACTION_RETRY_SECONDS
    start_timeout = max(DEFAULT_SUBMIT_CONFIRM_SECONDS, kickoff_wait_seconds)

    with sync_playwright() as p:
        browser = launch_browser(p, headless=headless, browser_name=browser_name)
        context = browser.new_context(storage_state=str(state_path), accept_downloads=True)
        page = context.new_page()

        print("[AUTO] 第一步：上传音频并启动转写。")
        safe_navigate(page, transcribe_url)
        page.wait_for_timeout(2500)
        upload_and_start_transcribe_on_page(
            page,
            file_path=file_path,
            submit_timeout_seconds=start_timeout,
        )

        print(f"[AUTO] 第二步：新开标签页直达默认文件夹，并等待“{resolved_title}”完成后导出。")
        export_page = context.new_page()
        safe_navigate(export_page, folders_url)
        export_to_local_on_page(
            page=export_page,
            url=folders_url,
            record_title=resolved_title,
            wait_ready_seconds=wait_ready_seconds,
            status_poll_seconds=status_poll_seconds,
            action_retry_seconds=action_retry_seconds,
            download_dir=download_dir,
            use_reload=True,
        )

        browser.close()
        print("[AUTO] 浏览器已关闭。")


def fetch_classroom_lessons(course_id: str) -> list[dict[str, str]]:
    lessons: list[dict[str, str]] = []
    api_url = f"https://classroom.zju.edu.cn/courseapi/v2/course/catalogue?course_id={course_id}"

    with sync_playwright() as p:
        print("[CLASSROOM] 正在后台提取课程清单...")
        context = launch_classroom_context(p, is_headless=True)
        page = context.new_page()

        page.goto("https://classroom.zju.edu.cn/index")
        try:
            response = page.request.get(api_url)
            data = response.json()
        except Exception:
            data = {}

        if not data.get("success"):
            print("[CLASSROOM] 登录可能过期，切换可见窗口等待登录。")
            context.close()
            context = launch_classroom_context(p, is_headless=False)
            page = context.new_page()
            try:
                page.goto("https://classroom.zju.edu.cn/index")
                page.wait_for_selector("text=我的学习", timeout=180000)
                response = page.request.get(api_url)
                data = response.json()
            except Exception as exc:
                print(f"[CLASSROOM] 重新登录后仍无法获取课程清单: {exc}")
                context.close()
                return []

        result_obj = data.get("result", {})
        raw_list = result_obj.get("data", []) if isinstance(result_obj, dict) else []
        for item in raw_list:
            try:
                content = json.loads(item.get("content", "{}"))
                url_raw = content.get("playback", {}).get("url") or content.get("url")
                url = url_raw[0] if isinstance(url_raw, list) else url_raw
                if url:
                    lessons.append({"title": item.get("title", "未命名课程"), "url": str(url).strip()})
            except Exception:
                continue

        context.close()
    return lessons


def submit_audio_to_tingwu_after_export(audio_path: str, state_path: Path) -> None:
    if not state_path.exists():
        print(f"[LINK] 未找到听悟登录态: {state_path}")
        print("[LINK] 已跳过自动转写。可先运行: python study.py init")
        return

    audio_file = Path(audio_path)
    print("[LINK] mp3 已生成，开始复用 AUTO 流程（上传 -> 新标签页默认文件夹轮询导出）。")
    transcribe_then_export(
        transcribe_url=DEFAULT_URL,
        folders_url=DEFAULT_FOLDERS_URL,
        state_path=state_path,
        file_path=audio_file,
        record_title=audio_file.stem,
        browser_name=DEFAULT_BROWSER,
        headless=True,
        kickoff_wait_seconds=DEFAULT_START_TRANSCRIBE_WAIT_SECONDS,
        wait_ready_seconds=1800,
        poll_seconds=DEFAULT_POLL_SECONDS,
        download_dir=Path(TRANSCRIPT_DIR),
    )


def classroom_interactive_flow(auto_transcribe: bool = True) -> int:
    if not os.path.exists(NDM_PATH):
        print(f"[CLASSROOM] 找不到 NDM，请检查路径: {NDM_PATH}")
        return 1
    if not os.path.exists(FFMPEG_PATH):
        print(f"[CLASSROOM] 找不到 ffmpeg，请检查路径: {FFMPEG_PATH}")
        return 1

    for directory in [VIDEO_DIR, AUDIO_DIR]:
        os.makedirs(directory, exist_ok=True)

    history_items = load_course_history(COURSE_HISTORY_PATH)
    if history_items:
        print("\n[CLASSROOM] 课程历史:")
        for idx, item in enumerate(history_items, start=1):
            print(f"  [{idx}] {item['name']} -> {item['url']}")

    course_input = input("[CLASSROOM] 请输入课程名(可直接输入历史编号): ").strip()
    course_name = course_input
    matched_url = ""
    selected_by_index = False

    if course_input.isdigit() and history_items:
        index = int(course_input) - 1
        if 0 <= index < len(history_items):
            selected = history_items[index]
            course_name = selected["name"]
            matched_url = selected.get("url", "")
            selected_by_index = True
            print(f"[CLASSROOM] 已按编号选择课程: {course_name}")
        else:
            print("[CLASSROOM] 历史编号无效，将按文本课程名处理。")

    if not matched_url and course_name:
        for item in history_items:
            if item.get("name") == course_name:
                matched_url = item.get("url", "")
                break

    if selected_by_index and matched_url:
        full_url = matched_url
        print(f"[CLASSROOM] 已自动使用历史网址: {full_url}")
    elif matched_url:
        full_url = input(f"[CLASSROOM] 请输入课程列表网址(回车使用历史): ").strip() or matched_url
    else:
        full_url = input("[CLASSROOM] 请输入课程列表网址: ").strip()

    if not course_name:
        course_name = "网课"

    course_id = parse_course_id_from_url(full_url)
    while not full_url or not course_id:
        print("[CLASSROOM] 网址无效，必须包含 course_id。")
        full_url = input("[CLASSROOM] 请重新输入课程列表网址: ").strip()
        course_id = parse_course_id_from_url(full_url)

    upsert_course_history(COURSE_HISTORY_PATH, course_name, full_url)

    lessons = fetch_classroom_lessons(course_id)
    if not lessons:
        print("[CLASSROOM] 未获取到可下载课程。")
        return 1

    print("\n" + "=" * 40)
    for i, lesson in enumerate(lessons, start=1):
        print(f"  [{i}] {lesson['title']}")
    print("=" * 40)

    raw_selection = input(
        "\n[CLASSROOM] 请输入编号(支持组合: 1 2 5 / 1,2,5 / 1-3,8): "
    ).strip()
    selected_indexes = parse_lesson_selection(raw_selection, total=len(lessons))
    if not selected_indexes:
        print("[CLASSROOM] 输入无效，请使用编号组合，例如: 1 2 5 或 1-3,8")
        return 1

    total_count = len(selected_indexes)
    course_prefix = sanitize_filename_component(course_name)

    for task_index, lesson_index in enumerate(selected_indexes, start=1):
        selected = lessons[lesson_index]
        lesson_part = sanitize_filename_component(selected["title"])
        target_name = f"{course_prefix}_{lesson_part}"

        print("\n" + "-" * 40)
        print(f"[CLASSROOM] 批量任务 {task_index}/{total_count}: {selected['title']}")

        before_mp4_files = {
            p.name: (p.stat().st_mtime, p.stat().st_size)
            for p in Path(VIDEO_DIR).glob("*.mp4")
            if p.is_file()
        }

        print("[CLASSROOM] 正在唤起 NDM 下载...")
        subprocess.Popen([NDM_PATH, selected["url"]])
        print("[CLASSROOM] NDM 已启动，开始自动等待下载完成。")

        downloaded_mp4 = wait_for_new_or_updated_mp4(
            Path(VIDEO_DIR),
            before_mp4_files,
            timeout_seconds=DEFAULT_MP4_WAIT_SECONDS,
        )
        if downloaded_mp4 is None:
            print("[CLASSROOM] 超时：未检测到新的 mp4 下载结果。")
            return 1

        final_video_path = os.path.join(VIDEO_DIR, f"{target_name}.mp4")
        try:
            if str(downloaded_mp4) != final_video_path:
                os.replace(str(downloaded_mp4), final_video_path)
                print(f"[CLASSROOM] 已自动更名为: {target_name}.mp4")
        except Exception as exc:
            print(f"[CLASSROOM] 改名失败: {exc}")
            return 1

        audio_path = os.path.join(AUDIO_DIR, f"{target_name}.mp3")
        print("[CLASSROOM] 正在提取音频...")
        result = subprocess.run(
            [
                FFMPEG_PATH,
                "-i",
                final_video_path,
                "-vn",
                "-acodec",
                "libmp3lame",
                "-q:a",
                "4",
                audio_path,
                "-y",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if result.returncode != 0:
            print("[CLASSROOM] 音频提取失败。")
            return 1

        print("[CLASSROOM] 处理成功，音频已转出。")
        if auto_transcribe:
            submit_audio_to_tingwu_after_export(audio_path, DEFAULT_STATE_PATH)

    print("[CLASSROOM] 全流程结束（保持后台，不弹出浏览器/文件夹）。")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="课堂下载 + 通义听悟自动化一体脚本")
    subparsers = parser.add_subparsers(dest="command", required=False)

    classroom_parser = subparsers.add_parser("classroom", help="课堂下载并转码，可选自动提交听悟")
    classroom_parser.add_argument(
        "--no-auto-transcribe",
        action="store_true",
        help="仅下载转码，不自动提交听悟转写",
    )

    init_parser = subparsers.add_parser("init", help="首次手动登录并保存听悟状态")
    init_parser.add_argument("--url", default=DEFAULT_URL, help="通义听悟地址")
    init_parser.add_argument("--state", default=str(DEFAULT_STATE_PATH), help="登录状态文件路径")
    init_parser.add_argument("--timeout-minutes", type=int, default=10, help="预留手动登录时长（分钟）")
    init_parser.add_argument("--browser", choices=["msedge", "chrome", "chromium"], default=DEFAULT_BROWSER)

    run_parser = subparsers.add_parser("run", help="后台运行并复用听悟登录状态")
    run_parser.add_argument("--url", default=DEFAULT_URL, help="通义听悟地址")
    run_parser.add_argument("--state", default=str(DEFAULT_STATE_PATH), help="登录状态文件路径")
    run_parser.add_argument("--keep-alive-seconds", type=int, default=0, help="保持运行秒数，0 表示持续运行")
    run_parser.add_argument("--show", action="store_true", help="显示浏览器窗口")
    run_parser.add_argument("--browser", choices=["msedge", "chrome", "chromium"], default=DEFAULT_BROWSER)

    transcribe_parser = subparsers.add_parser("transcribe", help="上传 mp3 并启动转写")
    transcribe_parser.add_argument("--url", default=DEFAULT_URL, help="通义听悟地址")
    transcribe_parser.add_argument("--state", default=str(DEFAULT_STATE_PATH), help="登录状态文件路径")
    transcribe_parser.add_argument("--file", required=True, help="待上传 mp3 路径")
    transcribe_parser.add_argument("--show", action="store_true", help="显示浏览器窗口")
    transcribe_parser.add_argument("--wait-seconds", type=int, default=0, help="触发转写后额外等待秒数")
    transcribe_parser.add_argument("--browser", choices=["msedge", "chrome", "chromium"], default=DEFAULT_BROWSER)

    export_parser = subparsers.add_parser("export", help="默认文件夹导出到本地")
    export_parser.add_argument("--url", default=DEFAULT_FOLDERS_URL, help="默认文件夹地址")
    export_parser.add_argument("--state", default=str(DEFAULT_STATE_PATH), help="登录状态文件路径")
    export_parser.add_argument("--record-title", required=True, help="记录标题")
    export_parser.add_argument("--show", action="store_true", help="显示浏览器窗口")
    export_parser.add_argument("--wait-ready-seconds", type=int, default=1800, help="最长等待导出秒数")
    export_parser.add_argument("--poll-seconds", type=int, default=DEFAULT_POLL_SECONDS, help="轮询间隔秒数")
    export_parser.add_argument("--download-dir", default=str(Path.home() / "Downloads"), help="下载目录")
    export_parser.add_argument("--browser", choices=["msedge", "chrome", "chromium"], default=DEFAULT_BROWSER)

    auto_parser = subparsers.add_parser("auto", help="上传转写并在完成后自动导出")
    auto_parser.add_argument("--url", default=DEFAULT_URL, help="上传转写页面地址")
    auto_parser.add_argument("--folders-url", default=DEFAULT_FOLDERS_URL, help="默认文件夹页面地址")
    auto_parser.add_argument("--state", default=str(DEFAULT_STATE_PATH), help="登录状态文件路径")
    auto_parser.add_argument("--file", required=True, help="待上传 mp3 路径")
    auto_parser.add_argument("--record-title", default="", help="记录标题(用于匹配页面卡片；建议与上传文件名一致，默认用文件名)")
    auto_parser.add_argument("--show", action="store_true", help="调试参数：auto 固定后台运行，此参数会被忽略")
    auto_parser.add_argument("--kickoff-wait-seconds", type=int, default=1800, help="等待开始转写最长秒数")
    auto_parser.add_argument("--wait-ready-seconds", type=int, default=1800, help="等待可导出最长秒数")
    auto_parser.add_argument("--poll-seconds", type=int, default=DEFAULT_POLL_SECONDS, help="轮询间隔秒数")
    auto_parser.add_argument("--download-dir", default=str(Path.home() / "Downloads"), help="下载目录")
    auto_parser.add_argument("--browser", choices=["msedge", "chrome", "chromium"], default=DEFAULT_BROWSER)

    return parser.parse_args()


def main() -> None:
    args = parse_args()
    state_path = Path(getattr(args, "state", str(DEFAULT_STATE_PATH)))

    try:
        setup_runtime_config(force_prompt=False)

        if args.command is None:
            code = interactive_main_menu(state_path)
            sys.exit(code)

        print_runtime_config()

        if args.command == "classroom":
            code = classroom_interactive_flow(auto_transcribe=not args.no_auto_transcribe)
            sys.exit(code)

        if args.command == "init":
            init_login_state(args.url, state_path, args.timeout_minutes, args.browser)
        elif args.command == "run":
            if not ensure_tingwu_login_ready(state_path, args.browser):
                sys.exit(1)
            run_in_background(
                args.url,
                state_path,
                args.keep_alive_seconds,
                args.browser,
                headless=not args.show,
            )
        elif args.command == "transcribe":
            if not ensure_tingwu_login_ready(state_path, args.browser):
                sys.exit(1)
            upload_and_start_transcribe(
                args.url,
                state_path,
                Path(args.file),
                args.browser,
                headless=not args.show,
                wait_seconds=args.wait_seconds,
            )
        elif args.command == "export":
            if not ensure_tingwu_login_ready(state_path, args.browser):
                sys.exit(1)
            export_to_local(
                args.url,
                state_path,
                args.record_title,
                args.browser,
                headless=not args.show,
                wait_ready_seconds=args.wait_ready_seconds,
                poll_seconds=args.poll_seconds,
                download_dir=Path(args.download_dir),
            )
        elif args.command == "auto":
            if args.show:
                print("[AUTO] 已固定为后台无头运行，忽略 --show。")
            if not ensure_tingwu_login_ready(state_path, args.browser):
                sys.exit(1)
            transcribe_then_export(
                transcribe_url=args.url,
                folders_url=args.folders_url,
                state_path=state_path,
                file_path=Path(args.file),
                record_title=args.record_title,
                browser_name=args.browser,
                headless=True,
                kickoff_wait_seconds=args.kickoff_wait_seconds,
                wait_ready_seconds=args.wait_ready_seconds,
                poll_seconds=args.poll_seconds,
                download_dir=Path(args.download_dir),
            )
    except PlaywrightTimeoutError as exc:
        print(f"[ERROR] 页面加载超时: {exc}")
        sys.exit(2)
    except Error as exc:
        print("[ERROR] Playwright 运行失败。")
        print("[ERROR] 可能原因: 指定浏览器通道不可用，或 Playwright 内核未安装。")
        print("[ERROR] 可尝试: --browser chromium 并执行 playwright install chromium")
        print(f"[ERROR] 详情: {exc}")
        sys.exit(3)
    except KeyboardInterrupt:
        print("\n[INFO] 用户中断执行（Ctrl+C），任务已终止。")
        sys.exit(130)


if __name__ == "__main__":
    main()
