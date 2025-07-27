import os
from faster_whisper import WhisperModel
import threading
import urllib.request
import tkinter as tk
from tkinter import filedialog, messagebox
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
        model = WhisperModel("faster-whisper-base", device="cpu", compute_type="int8", local_files_only=True)
        segments, info = model.transcribe(chunk_path)
        text = " ".join([segment.text for segment in segments])
        return text

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
    tts = TTS(model_name="tts_models/fr/css10/vits", progress_bar=False, gpu=False)
    tts.tts_to_file(text=text, file_path=output_path)

# 🔁 Pipeline complet
def process_pipeline(input_path, output_dir, log_callback):
    try:
        log_callback("⏳ Installation de la traduction...")
        install_argos_translation()

        log_callback("🎙️ Transcription en cours...")
        transcription = transcribe_audio(input_path)
        log_callback("📝 Transcription :\n" + transcription[:200] + "...")

        log_callback("🌍 Traduction en cours...")
        translation = translate_text(transcription)
        log_callback("📝 Traduction :\n" + translation[:200] + "...")

        output_path = os.path.join(output_dir, "output_fr.mp3")
        log_callback("🔊 Synthèse vocale...")
        synthesize_speech(translation, output_path)

        log_callback(f"✅ Terminé ! Fichier généré : {output_path}")
        messagebox.showinfo("Succès", f"Fichier généré : {output_path}")
    except Exception as e:
        log_callback(f"❌ Erreur : {e}")
        messagebox.showerror("Erreur", str(e))

# 🎛️ Interface graphique (Tkinter)
def launch_gui():
    root = tk.Tk()
    root.title("Traducteur vocal (EN → FR) - Local & Gratuit")
    root.geometry("600x400")

    input_file = tk.StringVar()
    output_dir = tk.StringVar()

    def browse_input():
        file = filedialog.askopenfilename(filetypes=[("Fichiers audio", "*.mp3")])
        if file:
            input_file.set(file)

    def browse_output():
        folder = filedialog.askdirectory()
        if folder:
            output_dir.set(folder)

    log_text = tk.Text(root, height=15, wrap=tk.WORD)
    log_text.pack(pady=10)

    def log(msg):
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)

    def run():
        if not input_file.get() or not output_dir.get():
            messagebox.showwarning("Champs manquants", "Veuillez sélectionner un fichier et un dossier de sortie.")
            return
        threading.Thread(target=process_pipeline, args=(input_file.get(), output_dir.get(), log)).start()

    frame = tk.Frame(root)
    frame.pack(pady=10)

    tk.Label(frame, text="Fichier audio (.mp3)").grid(row=0, column=0, padx=5)
    tk.Entry(frame, textvariable=input_file, width=40).grid(row=0, column=1)
    tk.Button(frame, text="Parcourir", command=browse_input).grid(row=0, column=2)

    tk.Label(frame, text="Dossier de sortie").grid(row=1, column=0, padx=5)
    tk.Entry(frame, textvariable=output_dir, width=40).grid(row=1, column=1)
    tk.Button(frame, text="Parcourir", command=browse_output).grid(row=1, column=2)

    tk.Button(root, text="Démarrer le processus", command=run, bg="green", fg="white", height=2).pack(pady=10)

    root.mainloop()

# 🚀 Lancement de l'app
if __name__ == "__main__":
    launch_gui()
