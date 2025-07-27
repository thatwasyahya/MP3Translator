# MP3Translator

Traducteur vocal automatique anglais → français, avec transcription, traduction et synthèse vocale, optimisé pour Codespaces et les environnements à espace disque limité.

## Fonctionnalités
- Télécharge automatiquement des fichiers MP3 découpés depuis une API distante (ex: via ngrok)
- Transcrit l’audio anglais en texte (Whisper)
- Traduit le texte anglais en français (Argos Translate)
- Génère un fichier audio en français (TTS, voix féminine accent français)
- Upload le résultat sur l’API distante
- Interface web simple via Gradio
- Nettoyage automatique des fichiers temporaires

## Installation rapide (Codespaces ou Linux)

1. **Créer un environnement conda dans /tmp pour éviter les problèmes d’espace disque**
   ```bash
   conda create -p /tmp/ttsenv python=3.10 -y
   conda activate /tmp/ttsenv
   pip install -r requirements.txt
   ```

2. **Lancer l’application**
   ```bash
   python translator.py
   ```

## Utilisation
- Ouvrir l’interface web Gradio (lien affiché dans le terminal)
- Entrer l’URL de base de l’API (ex: https://xxxx.ngrok-free.app)
- Cliquer sur "Démarrer le processus"
- Les fichiers chunk_1.mp3 à chunk_10.mp3 seront traités automatiquement

## Dépendances principales
- [faster-whisper](https://github.com/SYSTRAN/faster-whisper)
- [TTS](https://github.com/coqui-ai/TTS)
- [argostranslate](https://github.com/argosopentech/argos-translate)
- [pydub](https://github.com/jiaaro/pydub)
- [ffmpeg-python](https://github.com/kkroening/ffmpeg-python)
- [gradio](https://gradio.app/)

## Modèles TTS disponibles
- Par défaut : `tts_models/fr/mai/tacotron2-DDC` (accent français féminin)
- Pour changer de voix, modifier la ligne `model_name` dans `synthesize_speech` dans `translator.py`.

## Conseils
- Si l’espace disque est saturé, supprimez les anciens environnements conda ou créez l’environnement dans /tmp.
- Pour un accent français plus neutre, essayez aussi `tts_models/fr/siwis/vits`.

## Auteur
- Projet initial par thatwasyahya, automatisation et adaptation Codespaces par GitHub Copilot.
