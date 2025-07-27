import os
import requests
from faster_whisper import WhisperModel
import threading
import urllib.request
import gradio as gr
from TTS.api import TTS
from pydub import AudioSegment
import argostranslate.package
import argostranslate.translate

# 📥 Installer modèle de traduction EN → FR une seule fois
def install_argos_translation():
    local_path = "translate-en_fr-1_9.argosmodel"
    if os.path.exists(local_path):
        argostranslate.package.install_from_path(local_path)
        os.remove(local_path)

# 🔊 Transcription (anglais)
def transcribe_audio(path):
    from concurrent.futures import ThreadPoolExecutor, as_completed
    # Parameters
    chunk_minutes = 10
    max_workers = os.cpu_count() or 2

    # Load audio
    audio = AudioSegment.from_file(path)
    chunk_length_ms = chunk_minutes * 60 * 1000
    total_length_ms = len(audio)
    num_chunks = (total_length_ms + chunk_length_ms - 1) // chunk_length_ms

    # Export chunks to temp files
    temp_files = []
    for i in range(num_chunks):
        start = i * chunk_length_ms
        end = min((i + 1) * chunk_length_ms, total_length_ms)
        chunk = audio[start:end]
        temp_path = f"chunk_{i}.wav"
        chunk.export(temp_path, format="wav")
        temp_files.append(temp_path)

    def transcribe_chunk(chunk_path):
        # Préférer le modèle local si présent, sinon Hugging Face
        if os.path.isdir("faster-whisper-base"):
            model_name = "faster-whisper-base"
            local_files_only = True
        else:
            model_name = "Systran/faster-whisper-base"
            local_files_only = False
        # Utiliser GPU si disponible, sinon CPU
        device = "cuda" if os.environ.get("CUDA_VISIBLE_DEVICES", None) or _gpu_available() else "cpu"
        compute_type = "float16" if device == "cuda" else "int8"
        model = WhisperModel(model_name, device=device, compute_type=compute_type, local_files_only=local_files_only)
        segments, info = model.transcribe(chunk_path)
        text = " ".join([segment.text for segment in segments])
        return text

    def _gpu_available():
        try:
            import torch
            return torch.cuda.is_available()
        except ImportError:
            return False

    results = [None] * num_chunks
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {executor.submit(transcribe_chunk, temp_files[i]): i for i in range(num_chunks)}
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                results[idx] = future.result()
            except Exception as e:
                results[idx] = f"[Erreur transcription chunk {idx}: {e}]"

    # Clean up temp files
    for temp_path in temp_files:
        try:
            os.remove(temp_path)
        except Exception:
            pass

    return " ".join(results)

# 🌍 Traduction (EN -> FR)
def translate_text(text):
    installed_languages = argostranslate.translate.get_installed_languages()
    en = next(lang for lang in installed_languages if lang.code == "en")
    fr = next(lang for lang in installed_languages if lang.code == "fr")
    translation = en.get_translation(fr)
    return translation.translate(text)

# 🗣️ Synthèse vocale (français)
def synthesize_speech(text, output_path):
    # Utiliser GPU si disponible, sinon CPU
    use_gpu = False
    try:
        import torch
        use_gpu = torch.cuda.is_available()
    except ImportError:
        pass
    tts = TTS(model_name="tts_models/fr/mai/tacotron2-DDC", progress_bar=False, gpu=use_gpu)
    tts.tts_to_file(text=text, file_path=output_path)

# 🔁 Pipeline complet
def process_pipeline(input_path, output_dir, log_callback):
    try:
        log_callback("⏳ Installation de la traduction...")
        install_argos_translation()

        # Itérer automatiquement sur chunk_1.mp3 à chunk_10.mp3
        api_base = input_path.rstrip('/')
        mp3_files = [f"chunk_{i}.mp3" for i in range(1, 11)]
        for mp3_file in mp3_files:
            log_callback(f"\n=== Téléchargement de {mp3_file} ===")
            url = f"{api_base}/get-chunk/{mp3_file}"
            r = requests.get(url)
            if r.status_code != 200:
                log_callback(f"Erreur téléchargement {mp3_file}: {r.status_code}")
                continue
            with open(mp3_file, "wb") as f:
                f.write(r.content)
            log_callback(f"Fichier {mp3_file} téléchargé.")

            log_callback("🎙️ Transcription en cours...")
            transcription = transcribe_audio(mp3_file)
            log_callback("📝 Transcription :\n" + transcription[:200] + "...")

            log_callback("🌍 Traduction en cours...")
            translation = translate_text(transcription)
            log_callback("📝 Traduction :\n" + translation[:200] + "...")

            output_path = f"{os.path.splitext(mp3_file)[0]}_fr.mp3"
            log_callback("🔊 Synthèse vocale...")
            synthesize_speech(translation, output_path)
            log_callback(f"✅ Terminé ! Fichier généré : {output_path}")

            # Uploader le fichier généré via l'API POST
            log_callback(f"⬆️ Upload de {output_path} vers l'API...")
            with open(output_path, "rb") as f:
                files = {'file': (output_path, f, 'audio/mpeg')}
                resp = requests.post(f"{api_base}/upload", files=files)
            if resp.status_code == 200:
                log_callback(f"Upload réussi pour {output_path}")
            else:
                log_callback(f"Erreur upload {output_path}: {resp.status_code}")

            # Supprimer les fichiers locaux immédiatement après traitement
            try:
                os.remove(mp3_file)
            except Exception:
                pass
            try:
                os.remove(output_path)
            except Exception:
                pass
    except Exception as e:
        log_callback(f"❌ Erreur : {e}")


# 🎛️ Interface web (Gradio)
def launch_gradio():

    def gradio_process(input_dir):
        logs = []
        def log_callback(msg):
            logs.append(msg)
        try:
            process_pipeline(input_dir, None, log_callback)
            return "\n".join(logs), []
        except Exception as e:
            return f"Erreur : {e}", []

    with gr.Blocks() as demo:
        gr.Markdown("# Traducteur vocal (EN → FR) - Local & Gratuit")
        with gr.Row():
            input_dir = gr.Textbox(label="URL de base de l'API (ex: https://xxxx.ngrok-free.app)", value="https://1b2d798c8b46.ngrok-free.app")
        run_btn = gr.Button("Démarrer le processus")
        log_output = gr.Textbox(label="Logs", lines=10)
        audio_output = gr.Gallery(label="Audios synthétisés (français)", type="audio", columns=1)

        def on_run(in_dir):
            if not in_dir:
                return "Veuillez indiquer l'URL de l'API.", []
            logs, audio_paths = gradio_process(in_dir)
            return logs, audio_paths

        run_btn.click(on_run, inputs=[input_dir], outputs=[log_output, audio_output])
    demo.launch()

# 🚀 Lancement de l'app
if __name__ == "__main__":
    launch_gradio()
