"""Uygulama acma - Windows uygulamalari, URI semalari ve Start Menu kisayollari."""

from actions.windows_utils import open_url, search_start_menu, start_process


APP_ALIASES = {
    "edge": "msedge",
    "chrome": "chrome",
    "firefox": "firefox",
    "terminal": "wt",
    "powershell": "powershell",
    "cmd": "cmd",
    "explorer": "explorer",
    "finder": "explorer",
    "spotify": "spotify",
    "vscode": "code",
    "vs code": "code",
    "code": "code",
    "notion": "Notion",
    "slack": "Slack",
    "discord": "Discord",
    "whatsapp": "WhatsApp",
    "telegram": "Telegram",
    "zoom": "Zoom",
    "mail": "outlookmail:",
    "calendar": "outlookcal:",
    "takvim": "outlookcal:",
    "notes": "onenote:",
    "notlar": "onenote:",
    "music": "spotify",
    "muzik": "spotify",
    "photos": "ms-photos:",
    "fotograflar": "ms-photos:",
    "maps": "bingmaps:",
    "haritalar": "bingmaps:",
    "calculator": "calc",
    "hesap makinesi": "calc",
    "system preferences": "ms-settings:",
    "system settings": "ms-settings:",
    "ayarlar": "ms-settings:",
    "activity monitor": "taskmgr",
    "aktivite monitoru": "taskmgr",
    "preview": "ms-photos:",
    "onizleme": "ms-photos:",
    "textedit": "notepad",
    "notepad": "notepad",
    "word": "winword",
    "excel": "excel",
    "powerpoint": "powerpnt",
    "figma": "Figma",
    "postman": "Postman",
    "docker": "Docker",
    "tableplus": "TablePlus",
    "keyshort": "Keyshort",
    "ks": "Keyshort",
}


def open_app(app_name: str) -> str:
    """Uygulamayi acar, basari/hata mesaji dondurur."""
    if not app_name:
        return "Uygulama adi belirtilmedi."

    normalized = app_name.lower().strip()
    resolved = APP_ALIASES.get(normalized, app_name.strip())

    try:
        if str(resolved).endswith(":"):
            open_url(str(resolved))
            return f"{app_name} acildi."

        shortcut = search_start_menu(str(resolved))
        if shortcut and start_process(str(shortcut)):
            return f"{resolved} acildi."

        if start_process(str(resolved)):
            return f"{resolved} acildi."

        return f"'{app_name}' bulunamadi veya acilamadi."
    except Exception as exc:
        return f"Hata: {exc}"
