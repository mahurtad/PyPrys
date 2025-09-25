import moviepy.editor as mp
import whisper
import os
import time
from tqdm import tqdm

def extraer_audio(video_path: str, audio_filename: str = "audio_extraido.wav"):
    print(f"🎬 Extrayendo audio desde: {video_path}")
    video_dir = os.path.dirname(video_path)
    audio_path = os.path.join(video_dir, audio_filename)

    clip = mp.VideoFileClip(video_path)
    clip.audio.write_audiofile(audio_path, logger=None)
    print(f"✅ Audio guardado como: {audio_path}")
    return audio_path

def transcribir_audio(audio_path: str, modelo: str = "base", max_duracion_min: int = None, idioma: str = "es", usar_chunking: bool = True):
    print("🧠 Cargando modelo Whisper...")
    model = whisper.load_model(modelo)

    # Cargar el audio completo
    audio_clip = mp.AudioFileClip(audio_path)
    duracion_total = audio_clip.duration
    duracion_limite = min(duracion_total, max_duracion_min * 60 if max_duracion_min else duracion_total)

    print(f"⏱️ Duración a transcribir: {round(duracion_limite / 60, 2)} minutos")

    texto_total = ""

    if usar_chunking:
        chunk_duracion = 60  # segundos
        num_chunks = int(duracion_limite // chunk_duracion) + 1
        print(f"🔄 Transcribiendo en {num_chunks} bloques de {chunk_duracion}s usando idioma '{idioma}'")

        for i in tqdm(range(num_chunks), desc="Transcripción", unit="chunk"):
            start = i * chunk_duracion
            end = min(start + chunk_duracion, duracion_limite)

            if end - start < 1:
                continue  # Evitar subclips vacíos

            try:
                subclip = audio_clip.subclip(start, end)
                temp_filename = f"temp_chunk_{i}.wav"
                subclip.write_audiofile(temp_filename, logger=None)

                result = model.transcribe(temp_filename, language=idioma, fp16=False)
                texto_total += f"[{start:.0f}s - {end:.0f}s] " + result['text'].strip() + "\n"

                os.remove(temp_filename)
            except Exception as e:
                print(f"⚠️ Error procesando bloque {i} ({start}-{end}s): {e}")
    else:
        print("✍️ Transcribiendo audio completo sin chunking...")
        options = {
            "fp16": False,
            "language": idioma,
            "duration": duracion_limite
        }
        result = model.transcribe(audio_path, **options)
        texto_total = result['text']

    print("✅ Transcripción finalizada")
    return texto_total

def procesar_video(video_path: str, modelo: str = "base", max_duracion_min: int = None, idioma: str = "es", usar_chunking: bool = True):
    audio_path = extraer_audio(video_path)
    texto = transcribir_audio(audio_path, modelo, max_duracion_min, idioma, usar_chunking)
    return texto

if __name__ == "__main__":
    ruta_video = r"G:\My Drive\UTP\GMT20250401-010916_Recording_1280x720.mp4"
    modelo_whisper = "base"                # Modelos: tiny, base, small, medium, large
    max_minutos = 20                       # Transcribir los primeros 20 minutos
    idioma_transcripcion = "es"            # Idioma: 'es', 'en', etc.
    usar_transcripcion_por_bloques = True  # True para usar chunking

    texto_resultado = procesar_video(
        ruta_video,
        modelo=modelo_whisper,
        max_duracion_min=max_minutos,
        idioma=idioma_transcripcion,
        usar_chunking=usar_transcripcion_por_bloques
    )

    output_txt_path = os.path.join(os.path.dirname(ruta_video), "transcripcion.txt")
    with open(output_txt_path, "w", encoding="utf-8") as f:
        f.write(texto_resultado)

    print(f"\n📝 Transcripción guardada en: {output_txt_path}")
