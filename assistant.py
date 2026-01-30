import argparse
import json
import os
import platform
import subprocess
import sys
import time
import webbrowser

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

MODEL_DEFAULT = "gpt-4o-mini"

APPS = {
    "chrome": r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "edge": r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "spotify": os.path.expandvars(r"%APPDATA%\\Spotify\\Spotify.exe"),
    "discord": os.path.expandvars(r"%LOCALAPPDATA%\\Discord\\Update.exe"),
    "notepad": "notepad.exe",
    "cmd": "cmd.exe",
    "explorer": "explorer.exe",
    "vscode": "code",
}

CONFIRM_ACTIONS = {"shutdown", "restart", "shell"}


def get_client() -> OpenAI:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY non trovato. Controlla il file .env.")
    return OpenAI(api_key=api_key)


def parse_command_with_gpt(client: OpenAI, text: str, model: str) -> dict:
    system_prompt = (
        "Sei un assistente che converte comandi in JSON per controllare il PC. "
        "Rispondi SOLO con JSON valido.\n\n"
        "Azioni supportate:\n"
        "- open_app: { \"action\":\"open_app\", \"app\":\"chrome|edge|spotify|discord|notepad|cmd|explorer|vscode\" }\n"
        "- open_url: { \"action\":\"open_url\", \"url\":\"https://...\" }\n"
        "- shutdown: { \"action\":\"shutdown\", \"delay_seconds\":0 }\n"
        "- restart: { \"action\":\"restart\", \"delay_seconds\":0 }\n"
        "- volume: { \"action\":\"volume\", \"level\":0-100 }\n"
        "- type_text: { \"action\":\"type_text\", \"text\":\"...\" }\n"
        "- press_keys: { \"action\":\"press_keys\", \"keys\":[\"ctrl\",\"c\"] }\n"
        "- shell: { \"action\":\"shell\", \"command\":\"...\" }\n\n"
        "Se non capisci:\n{ \"action\":\"unknown\", \"original\":\"...\" }"
    )

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text},
        ],
        response_format={"type": "json_object"},
        temperature=0,
    )

    raw = response.choices[0].message.content.strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"action": "unknown", "original": text, "raw": raw}


def set_volume(level: int):
    if platform.system().lower() != "windows":
        print("⚠️ Controllo volume disponibile solo su Windows.")
        return
    try:
        from ctypes import POINTER, cast

        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))

        min_vol, max_vol, _ = volume.GetVolumeRange()
        vol_db = min_vol + (max_vol - min_vol) * (level / 100.0)
        volume.SetMasterVolumeLevel(vol_db, None)
        print(f"🔊 Volume impostato a {level}%")
    except Exception as exc:
        print("❌ Errore volume:", exc)


def open_app(app: str):
    app_key = app.lower()
    path = APPS.get(app_key, app)
    try:
        subprocess.Popen(path, shell=True)
        print(f"✅ Aperto: {app}")
    except Exception as exc:
        print("❌ Errore apertura app:", exc)


def open_url(url: str):
    try:
        webbrowser.open(url)
        print(f"🌐 Aperto: {url}")
    except Exception as exc:
        print("❌ Errore apertura URL:", exc)


def shutdown_pc(delay_seconds: int = 0):
    if platform.system().lower() == "windows":
        subprocess.Popen(f"shutdown /s /t {int(delay_seconds)}", shell=True)
    else:
        minutes = max(1, int(delay_seconds / 60))
        subprocess.Popen(["shutdown", "-h", f"+{minutes}"])
    print(f"🛑 Spegnimento tra {delay_seconds}s")


def restart_pc(delay_seconds: int = 0):
    if platform.system().lower() == "windows":
        subprocess.Popen(f"shutdown /r /t {int(delay_seconds)}", shell=True)
    else:
        minutes = max(1, int(delay_seconds / 60))
        subprocess.Popen(["shutdown", "-r", f"+{minutes}"])
    print(f"🔁 Riavvio tra {delay_seconds}s")


def type_text(text: str):
    try:
        import pyautogui

        pyautogui.write(text, interval=0.02)
        print("⌨️ Testo scritto.")
    except Exception:
        print("⚠️ type_text richiede: py -m pip install pyautogui")


def press_keys(keys: list[str]):
    try:
        import pyautogui

        pyautogui.hotkey(*keys)
        print(f"⌨️ Combinazione eseguita: {'+'.join(keys)}")
    except Exception:
        print("⚠️ press_keys richiede: py -m pip install pyautogui")


def run_shell(command: str):
    try:
        subprocess.Popen(command, shell=True)
        print(f"🧩 Comando eseguito: {command}")
    except Exception as exc:
        print("❌ Errore comando shell:", exc)


def confirm_action(action: str) -> bool:
    while True:
        answer = input(f"Confermi l'azione '{action}'? (s/n): ").strip().lower()
        if answer in {"s", "si", "sì"}:
            return True
        if answer in {"n", "no"}:
            return False
        print("Risposta non valida. Digita 's' o 'n'.")


def execute_action(cmd: dict, require_confirm: bool):
    action = cmd.get("action")

    if require_confirm and action in CONFIRM_ACTIONS:
        if not confirm_action(action):
            print("🚫 Azione annullata.")
            return

    if action == "open_app":
        open_app(cmd.get("app", ""))
    elif action == "open_url":
        open_url(cmd.get("url", ""))
    elif action == "shutdown":
        shutdown_pc(int(cmd.get("delay_seconds", 0)))
    elif action == "restart":
        restart_pc(int(cmd.get("delay_seconds", 0)))
    elif action == "volume":
        lvl = int(cmd.get("level", 50))
        lvl = max(0, min(100, lvl))
        set_volume(lvl)
    elif action == "type_text":
        type_text(cmd.get("text", ""))
    elif action == "press_keys":
        press_keys(cmd.get("keys", []))
    elif action == "shell":
        run_shell(cmd.get("command", ""))
    else:
        print("🤷 Comando non capito:", cmd)


def setup_voice_dependencies():
    try:
        import speech_recognition  # noqa: F401

        return True
    except Exception:
        print("⚠️ Modalità voce richiede: pip install SpeechRecognition pyaudio")
        return False


def listen_for_command() -> str:
    import speech_recognition as sr

    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("🎙️ In ascolto...")
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        audio = recognizer.listen(source)

    try:
        text = recognizer.recognize_google(audio, language="it-IT")
        print(f"🗣️ Hai detto: {text}")
        return text
    except sr.UnknownValueError:
        print("⚠️ Non ho capito. Riprova.")
        return ""
    except sr.RequestError as exc:
        print("❌ Errore riconoscimento vocale:", exc)
        return ""


def speak(text: str):
    try:
        import pyttsx3

        engine = pyttsx3.init()
        engine.say(text)
        engine.runAndWait()
    except Exception:
        print("⚠️ TTS richiede: pip install pyttsx3")


def main():
    parser = argparse.ArgumentParser(description="Assistente vocale locale")
    parser.add_argument("--voice", action="store_true", help="Abilita input vocale")
    parser.add_argument("--tts", action="store_true", help="Abilita risposta vocale")
    parser.add_argument("--model", default=MODEL_DEFAULT, help="Modello OpenAI")
    parser.add_argument(
        "--no-confirm",
        action="store_true",
        help="Disattiva conferma per azioni pericolose",
    )
    args = parser.parse_args()

    try:
        client = get_client()
    except RuntimeError as exc:
        print(f"❌ {exc}")
        sys.exit(1)

    voice_enabled = args.voice and setup_voice_dependencies()

    print("✅ Assistente avviato.")
    print("Scrivi un comando tipo: 'apri chrome', 'volume 30', 'spegni pc'")
    print("Digita 'exit' per uscire.\n")

    while True:
        if voice_enabled:
            text = listen_for_command()
            if not text:
                time.sleep(0.5)
                continue
        else:
            text = input("👤 Tu: ").strip()

        if not text:
            continue
        if text.lower() in {"exit", "quit", "esci"}:
            print("👋 Ciao!")
            break

        cmd = parse_command_with_gpt(client, text, args.model)
        print("📦 JSON:", cmd)
        execute_action(cmd, require_confirm=not args.no_confirm)

        if args.tts:
            speak("Operazione completata.")
        print()


if __name__ == "__main__":
    main()
