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

    try:
        mostrar_progreso_mensaje("🛠️ Procesando audio", duracion=6)
        clip = mp.VideoFileClip(video_path)
        clip.audio.write_audiofile(audio_path, logger=None)
    except Exception as e:
        raise RuntimeError(f"❌ Error al extraer audio: {e}")

    if not os.path.exists(audio_path):
        raise FileNotFoundError(f"❌ El archivo de audio no fue generado correctamente: {audio_path}")
    
    print(f"✅ Audio guardado como: {audio_path}")
    return audio_path

def formato_srt(segundos):
    horas = int(segundos // 3600)
    minutos = int((segundos % 3600) // 60)
    seg = int(segundos % 60)
    milis = int((segundos - int(segundos)) * 1000)
    return f"{horas:02}:{minutos:02}:{seg:02},{milis:03}"

def transcribir_audio_completo(audio_path: str, modelo: str = "base", idioma: str = "es"):
    print("\n🧠 Cargando modelo Whisper...")
    mostrar_progreso_mensaje("📦 Cargando modelo", duracion=6)
    model = whisper.load_model(modelo)

    print("✍️ Transcribiendo todo el audio (esto puede tomar varios minutos)...")
    result = model.transcribe(
        audio_path,
        language=idioma,
        fp16=False
    )

    texto_completo = result.get("text", "").strip()
    srt = ""

    segments = result.get("segments", [])
    if segments:
        print("📡 Generando subtítulos segmentados con visibilidad de avance...")
        for i, seg in tqdm(enumerate(segments), total=len(segments), desc="Transcripción en progreso", unit="segmento"):
            srt += f"{i+1}\n{formato_srt(seg['start'])} --> {formato_srt(seg['end'])}\n{seg['text'].strip()}\n\n"
            # Mostrar resumen breve del segmento
            resumen = seg['text'].strip().replace("\n", " ")
            print(f"🪄 Segmento {i+1}/{len(segments)}: {formato_srt(seg['start'])} - {resumen[:80]}...")
    else:
        print("⚠️ No se encontraron segmentos de subtítulos. Solo se generará el archivo .txt.")

    print("✅ Transcripción finalizada")
    return texto_completo, srt

def procesar_video(video_path: str, modelo: str = "base", idioma: str = "es"):
    audio_path = extraer_audio(video_path)
    texto, srt = transcribir_audio_completo(audio_path, modelo, idioma)
    return texto, srt, audio_path

if __name__ == "__main__":
    # RUTA CENTRAL
    base_dir = r"G:\My Drive\UTP\test3"
    ruta_video = os.path.join(base_dir, "GMT20250410-010619_Recording_1280x720.mp4")
    modelo_whisper = "base"
    idioma_transcripcion = "es"

    # PROCESAR VIDEO
    texto_resultado, srt_resultado, ruta_audio = procesar_video(
        ruta_video,
        modelo=modelo_whisper,
        idioma=idioma_transcripcion
    )

    # GUARDAR RESULTADOS
    txt_path = os.path.join(base_dir, "transcripcion.txt")
    srt_path = os.path.join(base_dir, "transcripcion.srt")

    print("\n💾 Guardando archivos...")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(texto_resultado)

    with open(srt_path, "w", encoding="utf-8") as f:
        f.write(srt_resultado)

    # OPCIONAL: BORRAR AUDIO TEMPORAL
    try:
        os.remove(ruta_audio)
        print("🧹 Archivo de audio temporal eliminado.")
    except Exception as e:
        print(f"⚠️ No se pudo eliminar el archivo de audio: {e}")

    print(f"\n📝 Archivos generados:\n - {txt_path}\n - {srt_path}")
