import os
import time
import whisper
from tqdm import tqdm

def mostrar_progreso_mensaje(texto: str, duracion: int = 3):
    print(texto, end="", flush=True)
    for _ in range(duracion):
        time.sleep(0.5)
        print(".", end="", flush=True)
    print(" ✅")

def formato_srt(segundos):
    horas = int(segundos // 3600)
    minutos = int((segundos % 3600) // 60)
    seg = int(segundos % 60)
    milis = int((segundos - int(segundos)) * 1000)
    return f"{horas:02}:{minutos:02}:{seg:02},{milis:03}"

def transcribir_audio(audio_path: str, modelo: str = "small", idioma: str = "es"):
    print("\n🧠 Cargando modelo Whisper...")
    mostrar_progreso_mensaje("📦 Cargando modelo", duracion=6)
    model = whisper.load_model(modelo)

    print("✍️ Transcribiendo audio (esto puede tomar varios minutos)...")
    start_time = time.time()
    result = model.transcribe(
        audio_path,
        language=idioma,
        fp16=False
    )
    end_time = time.time()
    duracion = end_time - start_time
    print(f"⏱️ Tiempo de transcripción: {duracion:.2f} segundos")

    texto_completo = result.get("text", "").strip()
    srt = ""

    segments = result.get("segments", [])
    if segments:
        print("📡 Generando subtítulos segmentados con visibilidad de avance...")
        for i, seg in tqdm(enumerate(segments), total=len(segments), desc="Transcripción en progreso", unit="segmento"):
            srt += f"{i+1}\n{formato_srt(seg['start'])} --> {formato_srt(seg['end'])}\n{seg['text'].strip()}\n\n"
            resumen = seg['text'].strip().replace("\n", " ")
            print(f"🪄 Segmento {i+1}/{len(segments)}: {formato_srt(seg['start'])} - {resumen[:80]}...")
    else:
        print("⚠️ No se encontraron segmentos de subtítulos. Solo se generará el archivo .txt.")

    print("✅ Transcripción finalizada")
    return texto_completo, srt

if __name__ == "__main__":
    # RUTA CENTRAL
    base_dir = r"C:\Users\manue\Videos"
    ruta_audio = os.path.join(base_dir, "2025-09-23 12-20-12.mp3")  # Cambia el nombre según tu archivo de entrada
    modelo_whisper = "small"  # Puedes usar "large", "medium", etc.
    idioma_transcripcion = "es"

    # PROCESAR AUDIO
    texto_resultado, srt_resultado = transcribir_audio(
        ruta_audio,
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

    print(f"\n📝 Archivos generados:\n - {txt_path}\n - {srt_path}")
