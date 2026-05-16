"""Ozel Sistem Makrolari - Stream Deck Mantigi"""

import time
from actions.open_app import open_app
from actions.shell import shell_run

def run_system_macro(mode_name: str) -> str:
    mode = mode_name.lower().strip()
    
    if "oyun" in mode or "game" in mode:
        # 1. Arka plani temizle (Ornek: Edge ve Chrome'u zorla kapatir, RAM bosaltir)
        shell_run("taskkill /F /IM chrome.exe /T")
        shell_run("taskkill /F /IM msedge.exe /T")
        time.sleep(1)
        
        # 2. Spotify'i ac
        open_app("spotify")
        time.sleep(2)
        
        # 3. Valorant'i ac
        open_app("valorant")
        
        return "Oyun modu aktif edildi: Arka plan temizlendi, Spotify ve Valorant baslatildi. Iyi oyunlar patron."
        
    elif "çalışma" in mode or "work" in mode:
        # Sadece muzik ve kodlama
        open_app("spotify")
        time.sleep(1)
        open_app("vscode")
        return "Çalisma modu aktif: VS Code ve Spotify baslatildi. Kolay gelsin."
        
    return f"'{mode_name}' adinda bir makro tanimli degil."