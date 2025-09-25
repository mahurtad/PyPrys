import os
import time
import moviepy.editor as mp

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

if __name__ == "__main__":
    # RUTA CENTRAL
    base_dir = r"C:\Users\manue\Videos"
    nombre_video = "14770.mp4"  # Cambia según el nombre de tu archivo
    ruta_video = os.path.join(base_dir, nombre_video)
    
    # NOMBRE DE SALIDA PARA EL AUDIO
    nombre_audio = "grabacion_audio.wav"

    # EJECUTAR EXTRACCIÓN
    ruta_audio = extraer_audio(ruta_video, audio_filename=nombre_audio)

    print(f"\n🔊 Audio extraído disponible en:\n - {ruta_audio}")
