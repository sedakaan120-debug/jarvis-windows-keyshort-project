"""Medya oynatma - Windows'ta YouTube, Spotify ve Apple Music web aramasi."""

from __future__ import annotations

import shutil
import time
import urllib.parse

from actions.browser import browser_control
from actions.windows_utils import hotkey, open_url, press_key


def _spotify_available() -> bool:
    return shutil.which("spotify") is not None


def _play_youtube(query: str) -> str:
    return browser_control("play_youtube", query=query)


def _play_spotify(query: str, autoplay: bool = True) -> str:
    encoded_query = urllib.parse.quote(query.strip())
    search_url = f"spotify:search:{encoded_query}"

    try:
        open_url(search_url)
    except Exception as exc:
        return f"Spotify acilamadi: {exc}"

    if not autoplay:
        return f"Spotify icinde '{query}' aramasi acildi."

    time.sleep(2.0)
    press_key("tab")
    press_key("enter", delay=0.2)
    press_key("space", delay=0.4)
    return f"Spotify'da '{query}' aramasi acildi; otomatik oynatma denendi."


def _play_apple_music(query: str, autoplay: bool = True) -> str:
    encoded_query = urllib.parse.quote(query.strip())
    open_url(f"https://music.apple.com/search?term={encoded_query}")
    if autoplay:
        return f"Apple Music web aramasi acildi: {query}. Oynatmak icin sonucu sec."
    return f"Apple Music aramasi acildi: {query}"


def play_media(query: str, provider: str = "auto", autoplay: bool = True) -> str:
    if not query or not query.strip():
        return "Calinacak icerik belirtilmedi."

    normalized_provider = (provider or "auto").strip().lower()
    if normalized_provider in {"yt", "youtube music"}:
        normalized_provider = "youtube"
    elif normalized_provider in {"apple music", "music", "apple_music"}:
        normalized_provider = "apple_music"

    if normalized_provider == "spotify":
        return _play_spotify(query, autoplay=autoplay)
    if normalized_provider == "apple_music":
        return _play_apple_music(query, autoplay=autoplay)
    if normalized_provider == "youtube":
        return _play_youtube(query)

    if _spotify_available():
        result = _play_spotify(query, autoplay=autoplay)
        if "acilamadi" not in result:
            return result

    return _play_youtube(query)
