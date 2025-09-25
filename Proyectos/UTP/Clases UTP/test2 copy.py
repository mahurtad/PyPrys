import os
import time
from tqdm import tqdm
import moviepy.editor as mp
import whisper

def extraer_audio(video_path: str, audio_filename: str = "audio_extraido.wav"):
    print(f"🎬 Extrayendo audio desde: {video_path}")
    video_dir = os.path.dirname(video_path)
    audio_path = os.path.join(video_dir, audio_filename)

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

def transcribir_audio(audio_path: str, modelo: str = "large", max_duracion_min: int = None, idioma: str = "es", chunk_duracion: int = 60):
    print("🧠 Cargando modelo Whisper...")
    model = whisper.load_model(modelo)

    audio_clip = mp.AudioFileClip(audio_path)
    duracion_total = audio_clip.duration
    duracion_limite = min(duracion_total, max_duracion_min * 60 if max_duracion_min else duracion_total)

    print(f"⏱️ Transcribiendo {round(duracion_limite / 60, 2)} minutos en bloques de {chunk_duracion} segundos...")

    texto_total = ""
    srt_total = ""
    num_chunks = int(duracion_limite // chunk_duracion) + 1

    for i in tqdm(range(num_chunks), desc="Transcripción", unit="chunk"):
        start = i * chunk_duracion
        end = min(start + chunk_duracion, duracion_limite)

        if end - start < 1:
            continue

        try:
            subclip = audio_clip.subclip(start, end)
            temp_audio = f"temp_chunk_{i}.wav"
            subclip.write_audiofile(temp_audio, logger=None)

            result = model.transcribe(
                temp_audio,
                language=idioma,
                fp16=False
            )

            texto = result["text"].strip()
            texto_total += f"[{start:.0f}s - {end:.0f}s] {texto}\n"

            # SRT output
            srt_total += f"{i+1}\n{formato_srt(start)} --> {formato_srt(end)}\n{texto}\n\n"

            os.remove(temp_audio)
        except Exception as e:
            print(f"⚠️ Error en bloque {i} ({start}-{end}s): {e}")

    print("✅ Transcripción completada")
    return texto_total, srt_total

def procesar_video(video_path: str, modelo: str = "large", max_duracion_min: int = None, idioma: str = "es", chunk_duracion: int = 60):
    audio_path = extraer_audio(video_path)
    texto, srt = transcribir_audio(audio_path, modelo, max_duracion_min, idioma, chunk_duracion)
    return texto, srt

if __name__ == "__main__":
    ruta_video = r"G:\My Drive\UTP\Video de WhatsApp 2025-04-13 a las 19.12.55_6f642a14.mp4"
    modelo_whisper = "large"              # Modelos posibles: tiny, base, small, medium, large
    max_minutos = 20                      # Máximo de minutos a transcribir
    idioma_transcripcion = "es"          # Idioma de la transcripción
    duracion_chunk = 60                  # Duración por bloque en segundos

    texto_resultado, srt_resultado = procesar_video(
        ruta_video,
        modelo=modelo_whisper,
        max_duracion_min=max_minutos,
        idioma=idioma_transcripcion,
        chunk_duracion=duracion_chunk
    )

    carpeta_salida = os.path.dirname(ruta_video)
    with open(os.path.join(carpeta_salida, "transcripcion.txt"), "w", encoding="utf-8") as f:
        f.write(texto_resultado)

    with open(os.path.join(carpeta_salida, "transcripcion.srt"), "w", encoding="utf-8") as f:
        f.write(srt_resultado)

    print("\n📝 Archivos generados: transcripcion.txt y transcripcion.srt")
