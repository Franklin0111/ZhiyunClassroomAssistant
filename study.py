import os
import json
import subprocess
import time
import shutil
import sys
from urllib.parse import urlparse, parse_qs
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

# ================= ⚡ 配置区 ⚡ =================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def resolve_app_home():
    # 打包后使用 exe 所在目录，源码运行时使用脚本所在目录。
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return SCRIPT_DIR


APP_HOME = resolve_app_home()
BOT_DATA = os.path.join(APP_HOME, "zju_bot_data")
DEFAULT_VIDEO_DIR = os.path.join(APP_HOME, "videos")
DEFAULT_AUDIO_DIR = os.path.join(APP_HOME, "audios")
CONFIG_PATH = os.path.join(APP_HOME, "study_config.json")
HISTORY_PATH = os.path.join(APP_HOME, "course_history.json")
OLD_APP_HOME = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "zju_study_tool")
OLD_CONFIG_PATH = os.path.join(OLD_APP_HOME, "study_config.json")
OLD_HISTORY_PATH = os.path.join(OLD_APP_HOME, "course_history.json")
OLD_BOT_DATA = os.path.join(OLD_APP_HOME, "zju_bot_data")
LOGIN_WAIT_TIMEOUT_MS = 120000
API_TIMEOUT_MS = 20000
DOWNLOAD_TIMEOUT_SECONDS = 1800
DOWNLOAD_POLL_SECONDS = 2
DOWNLOAD_STABLE_SECONDS = 6
# ===============================================

def ensure_app_home():
    os.makedirs(APP_HOME, exist_ok=True)
    os.makedirs(BOT_DATA, exist_ok=True)
    os.makedirs(DEFAULT_VIDEO_DIR, exist_ok=True)
    os.makedirs(DEFAULT_AUDIO_DIR, exist_ok=True)

def migrate_legacy_data():
    # 兼容历史版本：仅迁移旧 APPDATA 目录中的数据，避免把程序目录里的历史配置带入分发版本。
    if not os.path.exists(CONFIG_PATH) and os.path.exists(OLD_CONFIG_PATH):
        try:
            shutil.copy2(OLD_CONFIG_PATH, CONFIG_PATH)
        except OSError:
            pass

    if not os.path.exists(HISTORY_PATH) and os.path.exists(OLD_HISTORY_PATH):
        try:
            shutil.copy2(OLD_HISTORY_PATH, HISTORY_PATH)
        except OSError:
            pass

    if not os.path.exists(BOT_DATA) and os.path.exists(OLD_BOT_DATA):
        try:
            shutil.copytree(OLD_BOT_DATA, BOT_DATA, dirs_exist_ok=True)
        except OSError:
            pass

def detect_edge_path(configured_path=""):
    if configured_path and os.path.exists(configured_path):
        return configured_path

    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return ""

def sanitize_filename(text):
    invalid = '<>:"/\\|?*'
    clean = (text or "未命名").strip()
    for c in invalid:
        clean = clean.replace(c, "_")
    clean = "_".join(clean.split())
    return clean or "未命名"

def extract_course_id(full_url):
    parsed = urlparse((full_url or "").strip())
    query = parse_qs(parsed.query)
    return (query.get("course_id") or [""])[0].strip()

def normalize_course_url(full_url):
    course_id = extract_course_id(full_url)
    if not course_id:
        return ""
    return f"https://classroom.zju.edu.cn/courseapi/v2/course/catalogue?course_id={course_id}"

def prompt_required_path(prompt_text, current_value=""):
    while True:
        hint = f" [{current_value}]" if current_value else ""
        raw = input(f"{prompt_text}{hint}: ").strip()
        value = raw or current_value
        if value:
            return value
        print("❌ 该项不能为空。")

def prompt_optional_path(prompt_text, current_value=""):
    hint = f" [{current_value}]" if current_value else ""
    raw = input(f"{prompt_text}{hint}（可留空）: ").strip()
    return raw if raw else current_value

def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_json(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return default

def setup_or_update_config(current=None, first_time=False):
    current = current or {}
    print("\n" + "=" * 50)
    print("🧭 文件路径配置")
    if first_time:
        print("首次使用，请先填写以下路径。")
    print("=" * 50)

    default_video = current.get("video_dir", "").strip() or DEFAULT_VIDEO_DIR
    default_audio = current.get("audio_dir", "").strip() or DEFAULT_AUDIO_DIR

    ndm_path = prompt_required_path("1) 请输入 NDM.exe 路径", current.get("ndm_path", ""))
    ffmpeg_path = prompt_optional_path("2) 请输入 ffmpeg.exe 路径", current.get("ffmpeg_path", ""))
    video_dir = prompt_required_path("3) 请输入 mp4 存放目录", default_video)
    audio_dir = prompt_optional_path("4) 请输入 mp3 存放目录", default_audio)
    edge_path = prompt_optional_path("5) 请输入 Edge.exe 路径", current.get("edge_path", ""))

    os.makedirs(video_dir, exist_ok=True)
    if audio_dir:
        os.makedirs(audio_dir, exist_ok=True)

    config = {
        "ndm_path": ndm_path,
        "ffmpeg_path": ffmpeg_path,
        "video_dir": video_dir,
        "audio_dir": audio_dir,
        "edge_path": edge_path,
    }
    save_json(CONFIG_PATH, config)
    print("✅ 路径配置已保存。")
    return config

def load_or_init_config():
    ensure_app_home()
    migrate_legacy_data()
    config = load_json(CONFIG_PATH, {})
    required = ["ndm_path", "video_dir"]
    if not config or any(not str(config.get(k, "")).strip() for k in required):
        return setup_or_update_config(config, first_time=True)
    if not str(config.get("video_dir", "")).strip():
        config["video_dir"] = DEFAULT_VIDEO_DIR
    if not str(config.get("audio_dir", "")).strip():
        config["audio_dir"] = DEFAULT_AUDIO_DIR
    if "edge_path" not in config:
        config["edge_path"] = ""
    save_json(CONFIG_PATH, config)
    return config

def load_history():
    data = load_json(HISTORY_PATH, [])
    if not isinstance(data, list):
        return []
    valid = []
    for item in data:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "")).strip()
        url = str(item.get("url", "")).strip()
        if name and url:
            valid.append({"name": name, "url": url})
    return valid

def save_history(history):
    save_json(HISTORY_PATH, history)

def add_or_update_history(history, name, url):
    normalized_url = normalize_course_url(url) or url
    merged = [{"name": name, "url": normalized_url}]
    for item in history:
        if item.get("url") != normalized_url:
            merged.append(item)
    return merged[:50]

def choose_course(history):
    print("\n📚 请选择网课来源：")
    if history:
        for idx, item in enumerate(history, 1):
            print(f"  [{idx}] {item['name']} | {item['url']}")
        print("  [0] 手动输入新网课")
        print("  [b] 返回主菜单")
        while True:
            raw = input("👉 输入历史编号（或 0 新建，b 返回）: ").strip().lower()
            if raw == "b":
                return None, None
            if not raw.isdigit():
                print("❌ 请输入数字。")
                continue
            num = int(raw)
            if num == 0:
                break
            if 1 <= num <= len(history):
                item = history[num - 1]
                return item["name"], item["url"]
            print("❌ 编号超出范围。")

    while True:
        course_name = input("📝 请输入网课名（输入 b 返回）: ").strip()
        if course_name.lower() == "b":
            return None, None
        course_url = input("🔗 请输入网课链接（输入 b 返回）: ").strip()
        if course_url.lower() == "b":
            return None, None
        if not course_name:
            print("❌ 网课名不能为空。")
            continue
        normalized_url = normalize_course_url(course_url)
        if not normalized_url:
            print("❌ 网课链接缺少 course_id 参数，请重新输入。")
            continue
        return course_name, normalized_url

def get_latest_file(directory, extension=".mp4"):
    """在指定目录下寻找最新创建的文件"""
    try:
        files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(extension)]
    except OSError:
        return None
    if not files:
        return None
    # 按修改时间排序，取最后一个
    return max(files, key=os.path.getmtime)

def unique_path(directory, stem, extension):
    base = os.path.join(directory, f"{stem}{extension}")
    if not os.path.exists(base):
        return base
    index = 1
    while True:
        candidate = os.path.join(directory, f"{stem}_{index}{extension}")
        if not os.path.exists(candidate):
            return candidate
        index += 1

def decode_process_stderr(stderr_bytes):
    if not stderr_bytes:
        return ""
    for encoding in ("utf-8", "gbk"):
        try:
            return stderr_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return stderr_bytes.decode("utf-8", errors="replace")

def wait_for_completed_download(directory, started_at):
    """等待下载目录出现新 mp4 且文件大小稳定，视为下载完成"""
    deadline = time.time() + DOWNLOAD_TIMEOUT_SECONDS
    required_stable_checks = max(1, DOWNLOAD_STABLE_SECONDS // DOWNLOAD_POLL_SECONDS)

    last_file = None
    last_size = -1
    stable_checks = 0

    while time.time() < deadline:
        latest_file = get_latest_file(directory)
        if latest_file:
            try:
                mtime = os.path.getmtime(latest_file)
                size = os.path.getsize(latest_file)
            except OSError:
                time.sleep(DOWNLOAD_POLL_SECONDS)
                continue

            # 只关注本次下载启动后出现/更新的文件
            if mtime >= started_at - 1 and size > 0:
                if latest_file == last_file and size == last_size:
                    stable_checks += 1
                else:
                    last_file = latest_file
                    last_size = size
                    stable_checks = 0

                if stable_checks >= required_stable_checks:
                    return latest_file

        time.sleep(DOWNLOAD_POLL_SECONDS)

    return None

def launch_context(p, is_headless, edge_path=""):
    launch_kwargs = {
        "user_data_dir": BOT_DATA,
        "headless": is_headless,
        "args": ["--remote-debugging-port=9222"],
    }

    detected_edge = detect_edge_path(edge_path)
    if detected_edge:
        launch_kwargs["executable_path"] = detected_edge
    else:
        # 无显式路径时使用 channel，兼容不同安装路径。
        launch_kwargs["channel"] = "msedge"

    return p.chromium.launch_persistent_context(**launch_kwargs)

def fetch_lessons(course_id, config):
    edge_path = config.get("edge_path", "").strip()
    try:
        with sync_playwright() as p:
            print("🕵️ 正在后台提取课程清单...")
            context = launch_context(p, is_headless=True, edge_path=edge_path)
            page = context.new_page()
            api_url = f"https://classroom.zju.edu.cn/courseapi/v2/course/catalogue?course_id={course_id}"

            try:
                try:
                    page.goto("https://classroom.zju.edu.cn/index")
                    response = page.request.get(api_url, timeout=API_TIMEOUT_MS)
                    data = response.json()
                except Exception as e:
                    print(f"❌ 后台请求课程接口失败：{e}")
                    return []

                if not data.get("success"):
                    print("⚠️ 登录已过期，正在唤起窗口...")
                    context.close()
                    context = launch_context(p, is_headless=False, edge_path=edge_path)
                    page = context.new_page()
                    try:
                        page.goto("https://classroom.zju.edu.cn/index")
                    except Exception as e:
                        print(f"❌ 打开课堂页面失败：{e}")
                        return []

                    try:
                        page.wait_for_selector("text=我的学习", timeout=LOGIN_WAIT_TIMEOUT_MS)
                    except PlaywrightTimeoutError:
                        print("❌ 登录等待超时（120 秒），请确认是否已登录课堂页面。")
                        return []

                    try:
                        response = page.request.get(api_url, timeout=API_TIMEOUT_MS)
                        data = response.json()
                    except Exception as e:
                        print(f"❌ 重新登录后接口请求失败：{e}")
                        return []

                    if not data.get("success"):
                        msg = data.get("message") or data.get("msg") or "未知原因"
                        print(f"❌ 重新登录后仍无法获取课程清单：{msg}")
                        return []

                lessons = []
                raw_list = data.get("result", {}).get("data", [])
                for item in raw_list:
                    try:
                        content = json.loads(item.get("content", "{}"))
                        url_raw = content.get("playback", {}).get("url") or content.get("url")
                        url = url_raw[0] if isinstance(url_raw, list) else url_raw
                        if url:
                            title = str(item.get("title") or "未命名课时").strip()
                            lessons.append({"title": title, "url": str(url).strip()})
                    except (TypeError, AttributeError, KeyError, IndexError, json.JSONDecodeError) as e:
                        print(f"⚠️ 跳过异常条目：{item.get('title', '未命名')}（{e}）")
                        continue
                return lessons
            finally:
                context.close()
    except Exception as e:
        err = str(e)
        print(f"❌ Playwright 启动失败：{err}")
        print("💡 若你运行的是 exe，请优先使用 PyInstaller onedir，并在打包时收集 playwright 资源。")
        print("💡 建议命令：pyinstaller --onedir --collect-all playwright study.py")
        return []

def choose_lesson(lessons):
    if not lessons:
        print("❌ 没有解析到可下载课程，请检查课程链接或登录状态。")
        return None

    print("\n" + "=" * 40)
    for i, lesson in enumerate(lessons, 1):
        print(f"  [{i}] {lesson['title']}")
    print("  [0] 返回上一步")
    print("=" * 40)

    while True:
        raw_choice = input("\n👉 请输入编号（0 返回）: ").strip()
        if not raw_choice.isdigit():
            print("❌ 请输入数字编号。")
            continue
        if raw_choice == "0":
            return "BACK"
        choice = int(raw_choice) - 1
        if 0 <= choice < len(lessons):
            return lessons[choice]
        print(f"❌ 编号超出范围，请输入 1 到 {len(lessons)}。")

def ensure_runtime_paths(config, with_audio):
    ndm_path = config.get("ndm_path", "").strip()
    ffmpeg_path = config.get("ffmpeg_path", "").strip()
    video_dir = config.get("video_dir", "").strip()
    audio_dir = config.get("audio_dir", "").strip()

    if not ndm_path or not os.path.exists(ndm_path):
        print(f"❌ 找不到 NDM，请修改路径: {ndm_path}")
        return None
    if not video_dir:
        print("❌ mp4 存放目录未配置。")
        return None
    os.makedirs(video_dir, exist_ok=True)

    if with_audio:
        if not ffmpeg_path or not os.path.exists(ffmpeg_path):
            print(f"❌ 找不到 ffmpeg，请修改路径: {ffmpeg_path}")
            return None
        if not audio_dir:
            print("❌ mp3 存放目录未配置。")
            return None
        os.makedirs(audio_dir, exist_ok=True)

    return {
        "ndm_path": ndm_path,
        "ffmpeg_path": ffmpeg_path,
        "video_dir": video_dir,
        "audio_dir": audio_dir,
    }

def run_download_flow(config, with_audio):
    runtime = ensure_runtime_paths(config, with_audio)
    if not runtime:
        return

    while True:
        history = load_history()
        course_name, full_url = choose_course(history)
        if not course_name or not full_url:
            print("↩️ 已返回主菜单。")
            return

        course_id = extract_course_id(full_url)
        if not course_id:
            print("❌ 网址解析错误：缺少 course_id 参数。")
            continue

        lessons = fetch_lessons(course_id, config)
        selected = choose_lesson(lessons)
        if selected == "BACK":
            print("↩️ 已返回网课选择。")
            continue
        if not selected:
            return

        history = add_or_update_history(history, course_name, full_url)
        save_history(history)

        lesson_name = selected["title"]
        file_stem = f"{sanitize_filename(course_name)}_{sanitize_filename(lesson_name)}"

        print("📥 正在唤起 NDM 下载...")
        download_started_at = time.time()
        try:
            subprocess.Popen([runtime["ndm_path"], selected["url"]])
        except OSError as e:
            print(f"❌ 启动 NDM 失败：{e}")
            return

        print("\n" + "!" * 50)
        print("💡 NDM 已启动！")
        print("💡 脚本正在自动检测下载完成。")
        print("!" * 50 + "\n")

        print("⏳ 正在自动检测下载完成（检测到文件稳定后将自动继续）...")
        downloaded_file = wait_for_completed_download(runtime["video_dir"], download_started_at)
        if not downloaded_file:
            print(f"❌ 等待下载超时（{DOWNLOAD_TIMEOUT_SECONDS} 秒），请检查网络或下载状态后重试。")
            return

        final_video_path = unique_path(runtime["video_dir"], file_stem, ".mp4")
        if os.path.abspath(downloaded_file) != os.path.abspath(final_video_path):
            print(f"🔍 检测到下载文件: {os.path.basename(downloaded_file)}")
            try:
                os.rename(downloaded_file, final_video_path)
                print(f"✨ 已自动更名为: {os.path.basename(final_video_path)}")
            except OSError as e:
                print(f"❌ 改名失败: {e}")
                return
        else:
            print(f"✅ 下载文件名已符合规则: {os.path.basename(final_video_path)}")

        print(f"✅ mp4 下载完成: {final_video_path}")

        if with_audio:
            audio_path = unique_path(runtime["audio_dir"], file_stem, ".mp3")
            export_audio(runtime["ffmpeg_path"], final_video_path, audio_path)
        return

def export_audio(ffmpeg_path, mp4_path, mp3_path):
    print("🎵 正在提取音频...")
    ffmpeg_result = subprocess.run(
        [ffmpeg_path, "-i", mp4_path, "-vn", "-acodec", "libmp3lame", "-q:a", "4", mp3_path, "-y"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
    )
    if ffmpeg_result.returncode != 0:
        stderr_text = decode_process_stderr(ffmpeg_result.stderr)
        err_msg = stderr_text.strip().splitlines()
        err_tail = err_msg[-1] if err_msg else "未知错误"
        print(f"❌ 音频提取失败（ffmpeg 返回码 {ffmpeg_result.returncode}）：{err_tail}")
        return False
    print(f"✅ mp3 导出完成: {mp3_path}")
    return True

def run_export_audio_flow(config):
    ffmpeg_path = config.get("ffmpeg_path", "").strip()
    audio_dir = config.get("audio_dir", "").strip()

    if not ffmpeg_path or not os.path.exists(ffmpeg_path):
        print(f"❌ 找不到 ffmpeg，请先在路径配置中填写正确路径: {ffmpeg_path}")
        return
    if not audio_dir:
        print("❌ mp3 存放目录未配置，请先在路径配置中填写。")
        return
    os.makedirs(audio_dir, exist_ok=True)

    mp4_path = input("📁 请输入要导出音频的 mp4 完整路径: ").strip().strip('"')
    if not mp4_path:
        print("❌ mp4 路径不能为空。")
        return
    if not os.path.exists(mp4_path):
        print(f"❌ 文件不存在: {mp4_path}")
        return
    if not os.path.isfile(mp4_path):
        print(f"❌ 这不是文件: {mp4_path}")
        return

    base_name = sanitize_filename(os.path.splitext(os.path.basename(mp4_path))[0])
    mp3_path = unique_path(audio_dir, base_name, ".mp3")
    export_audio(ffmpeg_path, mp4_path, mp3_path)

def clear_history_interactive():
    history = load_history()
    if not history:
        print("ℹ️ 当前没有可删除的历史记录。")
        return

    print("\n📚 当前网课历史记录：")
    for idx, item in enumerate(history, 1):
        print(f"  [{idx}] {item['name']} | {item['url']}")
    print("  [0] 返回主菜单")

    while True:
        raw = input("👉 请输入要删除的编号（支持逗号分隔，如 1,3）: ").strip()
        if raw == "0":
            print("↩️ 已取消删除。")
            return

        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if not parts:
            print("❌ 请输入有效编号。")
            continue

        indexes = set()
        valid = True
        for p in parts:
            if not p.isdigit():
                valid = False
                break
            num = int(p)
            if num < 1 or num > len(history):
                valid = False
                break
            indexes.add(num - 1)

        if not valid:
            print(f"❌ 编号无效，请输入 1 到 {len(history)} 之间的数字。")
            continue

        kept = [item for i, item in enumerate(history) if i not in indexes]
        removed = [item for i, item in enumerate(history) if i in indexes]
        save_history(kept)
        print("✅ 已删除以下历史记录：")
        for item in removed:
            print(f"  - {item['name']}")
        return

def clear_config():
    try:
        removed = []
        if os.path.exists(CONFIG_PATH):
            os.remove(CONFIG_PATH)
            removed.append(CONFIG_PATH)
        if os.path.exists(HISTORY_PATH):
            os.remove(HISTORY_PATH)
            removed.append(HISTORY_PATH)
        if os.path.exists(BOT_DATA):
            shutil.rmtree(BOT_DATA, ignore_errors=True)
            removed.append(BOT_DATA)

        if removed:
            print("✅ 配置、历史记录与运行数据已清空。")
        else:
            print("ℹ️ 未发现可清空的数据。")
    except OSError as e:
        print(f"❌ 清空配置失败: {e}")

def print_menu():
    print("\n" + "=" * 50)
    print("🎛️ 主菜单")
    print("  [1] 仅下载 mp4")
    print("  [2] 下载 mp4 并转 mp3")
    print("  [3] 导出音频（手动输入 mp4 路径）")
    print("  [4] 修改文件路径")
    print("  [5] 删除网课历史记录")
    print("  [6] 清空配置")
    print("  [0] 退出")
    print("=" * 50)

def main():
    config = load_or_init_config()
    while True:
        print_menu()
        choice = input("👉 请选择功能编号: ").strip()
        if choice == "1":
            run_download_flow(config, with_audio=False)
        elif choice == "2":
            run_download_flow(config, with_audio=True)
        elif choice == "3":
            run_export_audio_flow(config)
        elif choice == "4":
            config = setup_or_update_config(config, first_time=False)
        elif choice == "5":
            clear_history_interactive()
        elif choice == "6":
            confirm = input("⚠️ 确认清空配置、历史记录和运行数据？(y/n): ").strip().lower()
            if confirm == "y":
                clear_config()
                config = load_or_init_config()
            else:
                print("已取消清空配置。")
        elif choice == "0":
            print("👋 已退出。")
            return
        else:
            print("❌ 无效编号，请重试。")

if __name__ == "__main__":
    main()