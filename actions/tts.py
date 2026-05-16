"""TTS (Text-to-Speech) - Windows SAPI/System.Speech kullanir."""

import subprocess
import threading


VOICE = ""


def _powershell_quote(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"


def speak_text(text: str, on_done=None, blocking: bool = False):
    if not text or not text.strip():
        if on_done:
            on_done()
        return

    max_len = 500
    if len(text) > max_len:
        text = text[:max_len] + "..."

    def _run():
        try:
            voice_select = ""
            if VOICE:
                voice_select = f"$s.SelectVoice({_powershell_quote(VOICE)});"
            command = (
                "Add-Type -AssemblyName System.Speech; "
                "$s=New-Object System.Speech.Synthesis.SpeechSynthesizer; "
                f"{voice_select} "
                f"$s.Speak({_powershell_quote(text)});"
            )
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", command],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        finally:
            if on_done:
                on_done()

    if blocking:
        _run()
    else:
        threading.Thread(target=_run, daemon=True).start()


def get_available_voices() -> list[str]:
    try:
        command = (
            "Add-Type -AssemblyName System.Speech; "
            "$s=New-Object System.Speech.Synthesis.SpeechSynthesizer; "
            "$s.GetInstalledVoices() | ForEach-Object { $_.VoiceInfo.Name }"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            timeout=8,
        )
        return [line.strip() for line in result.stdout.splitlines() if line.strip()]
    except Exception:
        return []
