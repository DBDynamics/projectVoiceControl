import re
import asyncio
import tempfile
import os

import speech_recognition as sr
import edge_tts
import subprocess
import threading
import time
import sounddevice as sd
import soundfile as sf
from DBDynamics import Ant

# 初始化串口控制器（将 'COM9' 改为设备管理器中的实际端口，如 'COM3'）。
# DBDynamics 默认以 1.5 Mbps 连接（serial: 1500000），端口被占用或无权限会报错。
m = Ant('COM9')

# 给 ID=1 的电机上电，使能控制（发送控制字 1）。
m.setPowerOn(1)

# 设置电机 1 的加/减速时间，单位毫秒（ms）。
# 建议范围 200–1000ms；此处设为 100ms，响应更“急”。
m.setAccTime(1, 100)

# 设置电机 1 的目标速度，单位为 pulse/ms（50000 脉冲/圈）。
# 步进电机合理速度通常在 0–300 之间，过高可能丢步。
m.setTargetVelocity(1, 150)

# 配置：选择更自然的中文神经语音（需要联网）
VOICE = "zh-CN-XiaoxiaoNeural"
REPLY_TEXT = "ok"  # 按需求回复“ok”，如需中文可改为"好的"


def chinese_to_int(text: str) -> int:
    """将中文数字（零一二两三四五六七八九十百千万）转换为整数。
    仅处理常见范围（0-99999）。解析失败返回 -1。
    """
    num_map = {
        "零": 0, "〇": 0,
        "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
        "五": 5, "六": 6, "七": 7, "八": 8, "九": 9,
    }
    unit_map = {"十": 10, "百": 100, "千": 1000}

    result = 0
    section = 0
    number = 0

    for ch in text:
        if ch in num_map:
            number = num_map[ch]
        elif ch in unit_map:
            # 如“十”前无数字，默认是1，例如“十三”=> 10 + 3
            if number == 0:
                number = 1
            section += number * unit_map[ch]
            number = 0
        elif ch == "万":
            if number == 0 and section == 0:
                # 仅“万”不合理，直接失败
                return -1
            result += (section + number) * 10000
            section = 0
            number = 0
        else:
            # 遇到非数字字符直接停止（外层应传入紧邻“度”的数字片段）
            break

    result += section + number
    return result if result != 0 else (0 if ("零" in text or "〇" in text) else -1)


def extract_angle(text: str) -> int | None:
    """从文本中提取角度整数。
    支持：
    - 阿拉伯数字（含负数）：如 "-90度"、"-45"、"90度"
    - 中文数字（含负数前缀“负”）：如 "负九十度"、"一百八十度"
    返回角度整数或 None。
    """
    if not text:
        return None
    text = text.strip()
    # 规范化：统一负号、容错常见ASR误识别（如“付”->“负”）
    text = (
        text.replace("−", "-")
            .replace("－", "-")
            .replace("付", "负")
    )

    # 优先匹配紧邻“度”的阿拉伯数字，支持：负前缀（负/minus/negative）与负号（-、－、−）
    m = re.search(r"(负(?:的)?|minus|negative)?\s*([\-]?\d+)\s*度", text, flags=re.IGNORECASE)
    if m:
        num = m.group(2).replace("−", "-").replace("－", "-")
        try:
            val = int(num)
            if m.group(1):
                return -abs(val)
            return val
        except ValueError:
            pass

    # 匹配未跟“度”的阿拉伯数字，支持：负前缀与负号
    m2 = re.search(r"(负(?:的)?|minus|negative)?\s*([\-]?\d+)", text, flags=re.IGNORECASE)
    if m2:
        num = m2.group(2).replace("−", "-").replace("－", "-")
        try:
            val = int(num)
            if m2.group(1):
                return -abs(val)
            return val
        except ValueError:
            pass

    # 匹配中文数字片段 + ‘度’，可带“负”前缀
    m3 = re.search(r"(负)?([零〇一二两三四五六七八九十百千万]+)\s*度", text)
    if m3:
        val = chinese_to_int(m3.group(2))
        if val >= 0:
            return -val if m3.group(1) else val

    # 匹配中文数字片段（未跟‘度’），可带“负”前缀
    m4 = re.search(r"(负)?([零〇一二两三四五六七八九十百千万]+)", text)
    if m4:
        val = chinese_to_int(m4.group(2))
        if val >= 0:
            return -val if m4.group(1) else val

    return None


async def speak_text(text: str, voice: str = VOICE) -> None:
    """使用 edge-tts 合成语音并播放（缓存文件，避免重复合成与写入冲突）。"""
    cache_dir = os.path.join(os.path.dirname(__file__), ".cache")
    os.makedirs(cache_dir, exist_ok=True)
    out_path = os.path.join(cache_dir, "reply_ok.mp3")
    # 若缓存不存在则合成一次
    if not os.path.exists(out_path):
        communicate = edge_tts.Communicate(text=text, voice=voice)
        await communicate.save(out_path)
    # 使用 Windows PowerShell 播放 mp3（依赖 Windows Media Player COM）
    ps_script = (
        f"$player = New-Object -ComObject WMPlayer.OCX; "
        f"$media = $player.newMedia('{out_path}'); "
        f"$player.controls.play(); "
        f"while($player.playState -ne 1) {{ Start-Sleep -Milliseconds 200 }}"
    )
    try:
        # 优先尝试 playsound（若环境已安装）
        try:
            from playsound import playsound as _playsound
            _playsound(out_path)
            return
        except Exception:
            pass

        # 次选：Windows Media Player COM（阻塞直到播放结束）
        subprocess.run([
            "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script
        ], check=True)
    except Exception:
        # 兜底1：明确调用 WMP 可执行文件
        try:
            subprocess.run([
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                f"Start-Process -FilePath 'wmplayer.exe' -ArgumentList '\"{out_path}\"'"
            ], check=True)
        except Exception:
            # 兜底2：用默认程序打开（不阻塞）
            subprocess.run([
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                f"Invoke-Item \"{out_path}\""
            ], check=False)


def play_ok_async(text: str = REPLY_TEXT):
    """后台播放回复，不阻塞后续聆听，支持连续对话。"""
    def _run():
        try:
            asyncio.run(speak_text(text))
        except Exception:
            pass

    t = threading.Thread(target=_run, daemon=True)
    t.start()


def main() -> None:
    r = sr.Recognizer()
    samplerate = 16000
    channels = 1

    print("请说出角度，如：‘运动到90度’（按Ctrl+C退出）")

    try:
        while True:
            print("正在聆听（录音约5秒）...")
            try:
                # 使用 sounddevice 录音到临时 WAV 文件
                duration = 5.0
                with tempfile.TemporaryDirectory() as td:
                    wav_path = os.path.join(td, "input.wav")
                    audio_data = sd.rec(
                        int(duration * samplerate),
                        samplerate=samplerate,
                        channels=channels,
                        dtype="int16",
                    )
                    sd.wait()
                    sf.write(wav_path, audio_data, samplerate, subtype="PCM_16")

                    print("识别中...")
                    with sr.AudioFile(wav_path) as source:
                        audio = r.record(source)
                        text = r.recognize_google(audio, language="zh-CN")
                        print(f"识别文本：{text}")

                angle = extract_angle(text)
                if angle is not None:
                    print(f"识别角度：{angle}")
                    # 后台播放“ok”，不阻塞，提升连续对话流畅度
                    play_ok_async(REPLY_TEXT)
                    m.setTargetPosition(1, angle*51200/360)
                else:
                    print("未识别到角度信息，请再试一次（例如：‘运动到90度’）。")

            except sr.UnknownValueError:
                print("语音无法识别，请再试一次。")
            except sr.RequestError as e:
                print(f"语音服务请求错误：{e}")
            except Exception as e:
                print(f"发生错误：{e}")

    except KeyboardInterrupt:
        print("\n已退出。")
        m.stop()


if __name__ == "__main__":
    main()