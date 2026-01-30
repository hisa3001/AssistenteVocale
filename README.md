# Assistente Vocale Locale (Terminale)

Questo progetto avvia un assistente stile "Jarvis" che interpreta i comandi in italiano e li esegue in locale. Funziona sia in modalità testuale che vocale.

## Requisiti

- Python 3.10+
- Chiave OpenAI in `.env`

Esempio `.env`:

```
OPENAI_API_KEY=sk-...
```

## Installazione

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

> Nota: su Windows potrebbe servire installare anche `pipwin` + `pyaudio` per il microfono.

## Uso

Modalità testuale:

```bash
python assistant.py
```

Modalità vocale:

```bash
python assistant.py --voice
```

Risposta vocale (TTS):

```bash
python assistant.py --tts
```

Disattivare la conferma delle azioni pericolose:

```bash
python assistant.py --no-confirm
```

## Azioni supportate

- Aprire app (`chrome`, `edge`, `spotify`, `discord`, `notepad`, `cmd`, `explorer`, `vscode`)
- Aprire URL
- Volume 0-100
- Scrivere testo
- Tasti rapidi
- Spegnimento / riavvio
- Eseguire un comando shell

L'assistente chiede conferma per le azioni pericolose (spegnimento/riavvio/comandi shell) a meno che non si usi `--no-confirm`.
