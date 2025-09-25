import os
import time
import moviepy.editor as mp
import whisper
from tqdm import tqdm

def mostrar_progreso_mensaje(texto: str, duracion: int = 3):
    print(texto, end="", flush=True)
    for _ in range(duracion):
        time.sleep(0.5)
        print(".", end="", flush=True)
    print(" ✅")

def extraer_audio(video_path: str, audio_filename: str = "audio_extraido.wav"):
    print(f"\n🎬 Extrayendo audio desde: {video_path}")
    video_dir = os.path.dirname(video_path)
    audio_path = os.path.join(video_dir, audio_filename)

    mostrar_progreso_mensaje("🛠️ Procesando audio", duracion=6)
    clip = mp.VideoFileClip(video_path)
    clip.audio.write_audiofile(audio_path, logger=None)
    print(f"✅ Audio guardado como: {audio_path}")
    return audio_path

def formato_srt(segundos):
    horas = int(segundos // 3600)
    minutos = int((segundos % 3600) // 60)
    seg = int(segundos % 60)
    milis = int((segundos - int(segundos)) * 1000)
    return f"{horas:02}:{minutos:02}:{seg:02},{milis:03}"

def transcribir_audio_completo(audio_path: str, modelo: str = "large", idioma: str = "es"):
    print("\n🧠 Cargando modelo Whisper...")
    mostrar_progreso_mensaje("📦 Cargando modelo", duracion=6)
    model = whisper.load_model(modelo)

    print("✍️ Transcribiendo todo el audio (esto puede tomar varios minutos)...")
    result = model.transcribe(
        audio_path,
        language=idioma,
        fp16=False
    )

    texto_completo = result["text"]
    srt = ""

    if "segments" in result:
        print("📡 Generando subtítulos segmentados...")
        for i, seg in tqdm(enumerate(result["segments"]), total=len(result["segments"]), desc="Generando SRT"):
            srt += f"{i+1}\n{formato_srt(seg['start'])} --> {formato_srt(seg['end'])}\n{seg['text'].strip()}\n\n"
    else:
        print("⚠️ No se detectaron segmentos para subtítulos.")

    print("✅ Transcripción finalizada")
    return texto_completo, srt

def procesar_video(video_path: str, modelo: str = "large", idioma: str = "es"):
    audio_path = extraer_audio(video_path)
    texto, srt = transcribir_audio_completo(audio_path, modelo, idioma)
    return texto, srt

if __name__ == "__main__":
    ruta_video = r"G:\My Drive\UTP\Video de WhatsApp 2025-04-13 a las 19.12.55_6f642a14.mp4"
    modelo_whisper = "large"
    idioma_transcripcion = "es"

    texto_resultado, srt_resultado = procesar_video(
        ruta_video,
        modelo=modelo_whisper,
        idioma=idioma_transcripcion
    )

    carpeta_salida = os.path.dirname(ruta_video)
    txt_path = os.path.join(carpeta_salida, "transcripcion.txt")
    srt_path = os.path.join(carpeta_salida, "transcripcion.srt")

    print("\n💾 Guardando resultados...")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(texto_resultado)

    with open(srt_path, "w", encoding="utf-8") as f:
        f.write(srt_resultado)

    print(f"\n📝 Archivos generados:\n - {txt_path}\n - {srt_path}")
