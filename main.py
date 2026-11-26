# MetaForge - Inspector y Editor Forense de Metadatos
# Copyright (C) 2026 Kevin R. Clas
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import yt_dlp
from faster_whisper import WhisperModel
import fitz  # PyMuPDF para PDFs
import docx  # python-docx para Word
from pptx import Presentation  # python-pptx para PowerPoint
from rapidocr_onnxruntime import RapidOCR

# Cachés globales para no recargar los modelos en memoria en cada ejecución
modelo_whisper_cache = {"nombre": None, "instancia": None}
motor_ocr_cache = None


def obtener_motor_ocr():
    """Inicializa RapidOCR una sola vez bajo demanda."""
    global motor_ocr_cache
    if motor_ocr_cache is None:
        motor_ocr_cache = RapidOCR()
    return motor_ocr_cache


def aplicar_ocr_bytes(imagen_bytes):
    """Extrae el texto legible dentro de los bytes de una imagen."""
    try:
        ocr = obtener_motor_ocr()
        resultado, _ = ocr(imagen_bytes)
        if not resultado:
            return ""
        # resultado es una lista de tuplas: [caja_coordenadas, texto, confianza]
        lineas = [linea[1].strip() for linea in resultado if linea[1].strip()]
        return "\n  ".join(lineas)
    except Exception as e:
        return f"[Error al ejecutar OCR: {e}]"


# ==========================================
# LÓGICA MÓDULO 1: YOUTUBE A MP3
# ==========================================
def convertir_a_segundos(tiempo_str):
    tiempo_str = tiempo_str.strip()
    if not tiempo_str:
        return None
    partes = tiempo_str.split(":")
    try:
        if len(partes) == 1:
            return int(partes[0])
        elif len(partes) == 2:
            return int(partes[0]) * 60 + int(partes[1])
        elif len(partes) == 3:
            return int(partes[0]) * 3600 + int(partes[1]) * 60 + int(partes[2])
        else:
            raise ValueError
    except ValueError:
        raise ValueError(f"Formato inválido: '{tiempo_str}'. Usa MM:SS o HH:MM:SS.")


def seleccionar_carpeta_yt():
    carpeta = filedialog.askdirectory()
    if carpeta:
        entrada_carpeta_yt.delete(0, tk.END)
        entrada_carpeta_yt.insert(0, carpeta)


def iniciar_descarga():
    url = entrada_url.get().strip()
    inicio_str = entrada_inicio.get().strip()
    fin_str = entrada_fin.get().strip()
    calidad = combo_calidad.get()
    carpeta_destino = entrada_carpeta_yt.get().strip()

    if not url:
        messagebox.showwarning("Campo vacío", "Por favor ingresa el enlace de YouTube.")
        return

    try:
        seg_inicio = convertir_a_segundos(inicio_str)
        seg_fin = convertir_a_segundos(fin_str)
    except ValueError as e:
        messagebox.showerror("Error de formato", str(e))
        return

    if seg_inicio is not None and seg_fin is not None and seg_inicio >= seg_fin:
        messagebox.showerror("Error en tiempos", "El tiempo de inicio debe ser menor que el final.")
        return

    boton_descargar.config(state=tk.DISABLED)
    etiqueta_estado_yt.config(text="Descargando y procesando audio... Espera.", foreground="#0284c7")

    threading.Thread(
        target=proceso_descarga,
        args=(url, seg_inicio, seg_fin, calidad, carpeta_destino),
        daemon=True
    ).start()


def proceso_descarga(url, seg_inicio, seg_fin, calidad, carpeta_destino):
    try:
        plantilla_salida = os.path.join(carpeta_destino, "%(title)s.%(ext)s")
        ydl_opts = {
            "format": "bestaudio/best",
            "outtmpl": plantilla_salida,
            "noplaylist": True,
            "extractor_args": {"youtube": {"player_client": ["android", "web"]}},
            "postprocessors": [{
                "key": "FFmpegExtractAudio",
                "preferredcodec": "mp3",
                "preferredquality": calidad,
            }],
        }

        if seg_inicio is not None or seg_fin is not None:
            inicio = seg_inicio if seg_inicio is not None else 0
            fin = seg_fin if seg_fin is not None else float("inf")
            ydl_opts["download_ranges"] = yt_dlp.utils.download_range_func(None, [(inicio, fin)])
            ydl_opts["force_keyframes_at_cuts"] = True

        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])

        ventana.after(0, lambda: finalizar_exito_yt(carpeta_destino))
    except Exception as e:
        mensaje_error = str(e)
        ventana.after(0, lambda msg=mensaje_error: finalizar_error_yt(msg))


def finalizar_exito_yt(carpeta):
    boton_descargar.config(state=tk.NORMAL)
    etiqueta_estado_yt.config(text="¡Descarga completada!", foreground="#16a34a")
    messagebox.showinfo("Éxito", f"MP3 guardado en:\n{carpeta}")


def finalizar_error_yt(mensaje_error):
    boton_descargar.config(state=tk.NORMAL)
    etiqueta_estado_yt.config(text="Error durante la descarga.", foreground="#dc2626")
    messagebox.showerror("Error", f"No se pudo completar:\n{mensaje_error}")


# ==========================================
# LÓGICA MÓDULO 2: AUDIO A TXT (WHISPER IA)
# ==========================================
def seleccionar_archivo_audio():
    tipos = [
        ("Archivos de Audio (WhatsApp, MP3, etc.)", "*.opus *.ogg *.mp3 *.wav *.m4a *.aac *.flac *.mp4"),
        ("Todos los archivos", "*.*")
    ]
    archivo = filedialog.askopenfilename(title="Seleccionar audio", filetypes=tipos)
    if archivo:
        entrada_archivo_audio.delete(0, tk.END)
        entrada_archivo_audio.insert(0, archivo)


def iniciar_transcripcion():
    ruta_audio = entrada_archivo_audio.get().strip()
    tamano_modelo = combo_modelo.get()
    idioma = combo_idioma.get()
    incluir_tiempos = var_marcas_tiempo.get()

    if not ruta_audio or not os.path.isfile(ruta_audio):
        messagebox.showwarning("Archivo inválido", "Selecciona un archivo de audio válido.")
        return

    boton_transcribir.config(state=tk.DISABLED)
    caja_resultado.delete("1.0", tk.END)
    etiqueta_estado_txt.config(
        text=f"Cargando modelo '{tamano_modelo}' y procesando...",
        foreground="#0284c7"
    )

    threading.Thread(
        target=proceso_transcripcion,
        args=(ruta_audio, tamano_modelo, idioma, incluir_tiempos),
        daemon=True
    ).start()


def formatear_segundo(seg):
    m, s = divmod(int(seg), 60)
    h, m = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{s:02d}" if h > 0 else f"{m:02d}:{s:02d}"


def proceso_transcripcion(ruta_audio, tamano_modelo, idioma, incluir_tiempos):
    try:
        if modelo_whisper_cache["nombre"] != tamano_modelo or modelo_whisper_cache["instancia"] is None:
            modelo_whisper_cache["instancia"] = WhisperModel(tamano_modelo, device="cpu", compute_type="int8")
            modelo_whisper_cache["nombre"] = tamano_modelo

        modelo = modelo_whisper_cache["instancia"]
        lang_code = None if idioma == "Auto" else idioma

        segmentos, _ = modelo.transcribe(ruta_audio, beam_size=5, language=lang_code)

        lineas = []
        for seg in segmentos:
            texto_limpio = seg.text.strip()
            if incluir_tiempos:
                marca = f"[{formatear_segundo(seg.start)} -> {formatear_segundo(seg.end)}] {texto_limpio}"
                lineas.append(marca)
            else:
                lineas.append(texto_limpio)

        separador = "\n" if incluir_tiempos else " "
        texto_final = separador.join(lineas)

        base, _ = os.path.splitext(ruta_audio)
        ruta_txt = f"{base}_transcripcion.txt"
        with open(ruta_txt, "w", encoding="utf-8") as f:
            f.write(texto_final)

        ventana.after(0, lambda: finalizar_exito_txt(texto_final, ruta_txt))
    except Exception as e:
        mensaje_error = str(e)
        ventana.after(0, lambda msg=mensaje_error: finalizar_error_txt(msg))


def finalizar_exito_txt(texto, ruta_txt):
    boton_transcribir.config(state=tk.NORMAL)
    caja_resultado.insert(tk.END, texto)
    etiqueta_estado_txt.config(text=f"¡Guardado en: {os.path.basename(ruta_txt)}!", foreground="#16a34a")
    messagebox.showinfo("Transcripción Lista", f"Archivo generado en:\n{ruta_txt}")


def finalizar_error_txt(mensaje_error):
    boton_transcribir.config(state=tk.NORMAL)
    etiqueta_estado_txt.config(text="Error al transcribir el audio.", foreground="#dc2626")
    messagebox.showerror("Error", f"Falló la transcripción:\n{mensaje_error}")


# ==========================================
# LÓGICA MÓDULO 3: DOCS (PDF/WORD/PPT) + OCR
# ==========================================
def seleccionar_documento():
    tipos = [
        ("Documentos soportados", "*.pdf *.docx *.pptx"),
        ("Archivos PDF", "*.pdf"),
        ("Documentos Word", "*.docx"),
        ("Presentaciones PowerPoint", "*.pptx")
    ]
    archivo = filedialog.askopenfilename(title="Seleccionar documento", filetypes=tipos)
    if archivo:
        entrada_archivo_doc.delete(0, tk.END)
        entrada_archivo_doc.insert(0, archivo)


def convertir_tabla_a_md(filas_datos):
    if not filas_datos:
        return ""
    md = []
    encabezado = [str(celda).replace("\n", " ").strip() for celda in filas_datos[0]]
    md.append("| " + " | ".join(encabezado) + " |")
    md.append("| " + " | ".join(["---"] * len(encabezado)) + " |")
    for fila in filas_datos[1:]:
        celdas = [str(celda).replace("\n", " ").strip() for celda in fila]
        md.append("| " + " | ".join(celdas) + " |")
    return "\n".join(md) + "\n"


def extraer_pdf(ruta, guardar_img, carpeta_img, usar_ocr):
    doc = fitz.open(ruta)
    bloques = [f"# Documento PDF: {os.path.basename(ruta)}\n"]
    contador_img = 0

    for num_pag, pagina in enumerate(doc, start=1):
        bloques.append(f"## Página {num_pag}\n")
        texto = pagina.get_text("text").strip()
        if texto:
            bloques.append(texto + "\n")

        lista_imagenes = pagina.get_images(full=True)
        for img_index, img in enumerate(lista_imagenes, start=1):
            contador_img += 1
            xref = img[0]
            ancho, alto = img[2], img[3]
            info_img = f"> **[Imagen #{contador_img} en Pág. {num_pag}]** (Resolución: {ancho}x{alto}px)"

            if guardar_img or usar_ocr:
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha > 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                img_bytes = pix.tobytes("png")

                if guardar_img:
                    nombre_img = f"pdf_pag{num_pag}_img{img_index}.png"
                    ruta_img = os.path.join(carpeta_img, nombre_img)
                    with open(ruta_img, "wb") as f_img:
                        f_img.write(img_bytes)
                    info_img += f" -> Guardada como `{nombre_img}`"

                if usar_ocr:
                    texto_ocr = aplicar_ocr_bytes(img_bytes)
                    if texto_ocr:
                        info_img += f"\n> **Texto extraído por OCR:**\n> {texto_ocr}"
                    else:
                        info_img += "\n> *(OCR: Sin texto legible detectado en la imagen)*"

            bloques.append(info_img + "\n")

    doc.close()
    return "\n".join(bloques)


def extraer_docx(ruta, guardar_img, carpeta_img, usar_ocr):
    doc = docx.Document(ruta)
    bloques = [f"# Documento Word: {os.path.basename(ruta)}\n"]

    for p in doc.paragraphs:
        texto = p.text.strip()
        if not texto:
            continue
        estilo = p.style.name.lower()
        if "heading 1" in estilo or "título 1" in estilo:
            bloques.append(f"## {texto}\n")
        elif "heading 2" in estilo or "título 2" in estilo:
            bloques.append(f"### {texto}\n")
        elif "heading" in estilo or "título" in estilo:
            bloques.append(f"#### {texto}\n")
        else:
            bloques.append(f"{texto}\n")

    if doc.tables:
        bloques.append("## Tablas del Documento\n")
        for idx, tabla in enumerate(doc.tables, start=1):
            filas = [[celda.text.strip() for celda in fila.cells] for fila in tabla.rows]
            bloques.append(f"**Tabla {idx}:**\n" + convertir_tabla_a_md(filas))

    contador_img = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            contador_img += 1
            nombre_orig = os.path.basename(rel.target_ref)
            img_bytes = rel.target_part.blob
            info_img = f"> **[Imagen incrustada #{contador_img}]** Referencia interna: `{nombre_orig}`"

            if guardar_img:
                nombre_salida = f"word_img{contador_img}_{nombre_orig}"
                ruta_img = os.path.join(carpeta_img, nombre_salida)
                with open(ruta_img, "wb") as f:
                    f.write(img_bytes)
                info_img += f" -> Extraída en `{nombre_salida}`"

            if usar_ocr:
                texto_ocr = aplicar_ocr_bytes(img_bytes)
                if texto_ocr:
                    info_img += f"\n> **Texto extraído por OCR:**\n> {texto_ocr}"
                else:
                    info_img += "\n> *(OCR: Sin texto legible detectado en la imagen)*"

            bloques.append(info_img + "\n")

    return "\n".join(bloques)


def extraer_pptx(ruta, guardar_img, carpeta_img, usar_ocr):
    prs = Presentation(ruta)
    bloques = [f"# Presentación PowerPoint: {os.path.basename(ruta)}\n"]
    contador_img = 0

    for num_slide, slide in enumerate(prs.slides, start=1):
        bloques.append(f"## Diapositiva {num_slide}\n")
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    t = p.text.strip()
                    if t:
                        bloques.append(f"- {t}")
                bloques.append("")

            if shape.has_table:
                filas = [[celda.text.strip() for celda in fila.cells] for fila in shape.table.rows]
                bloques.append(convertir_tabla_a_md(filas))

            if shape.shape_type == 13:  # PICTURE
                contador_img += 1
                alt_text = shape.name or "Sin descripción"
                descr = shape._element.xpath('.//p:cNvPr/@descr')
                if descr and descr[0].strip():
                    alt_text = descr[0].strip()

                img_bytes = shape.image.blob
                info_img = f"> **[Imagen #{contador_img} en Diapositiva {num_slide}]** Alt-Text: *{alt_text}*"

                if guardar_img:
                    ext = shape.image.ext
                    nombre_img = f"ppt_slide{num_slide}_img{contador_img}.{ext}"
                    ruta_img = os.path.join(carpeta_img, nombre_img)
                    with open(ruta_img, "wb") as f:
                        f.write(img_bytes)
                    info_img += f" -> Guardada como `{nombre_img}`"

                if usar_ocr:
                    texto_ocr = aplicar_ocr_bytes(img_bytes)
                    if texto_ocr:
                        info_img += f"\n> **Texto extraído por OCR:**\n> {texto_ocr}"
                    else:
                        info_img += "\n> *(OCR: Sin texto legible detectado en la imagen)*"

                bloques.append(info_img + "\n")

        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notas = slide.notes_slide.notes_text_frame.text.strip()
            if notas:
                bloques.append(f"**Notas del orador:** {notas}\n")

    return "\n".join(bloques)


def iniciar_extraccion_doc():
    ruta_doc = entrada_archivo_doc.get().strip()
    formato_salida = combo_formato_doc.get()
    guardar_img = var_extraer_img.get()
    usar_ocr = var_usar_ocr.get()

    if not ruta_doc or not os.path.isfile(ruta_doc):
        messagebox.showwarning("Archivo inválido", "Selecciona un documento PDF, Word o PPT válido.")
        return

    boton_extraer_doc.config(state=tk.DISABLED)
    caja_resultado_doc.delete("1.0", tk.END)
    msg_estado = "Extrayendo texto, tablas y ejecutando OCR en imágenes..." if usar_ocr else "Extrayendo texto y tablas..."
    etiqueta_estado_doc.config(text=msg_estado, foreground="#0284c7")

    threading.Thread(
        target=proceso_extraccion_doc,
        args=(ruta_doc, formato_salida, guardar_img, usar_ocr),
        daemon=True
    ).start()


def proceso_extraccion_doc(ruta_doc, formato_salida, guardar_img, usar_ocr):
    try:
        base, ext = os.path.splitext(ruta_doc)
        ext = ext.lower()
        carpeta_img = f"{base}_imagenes"

        if guardar_img and not os.path.exists(carpeta_img):
            os.makedirs(carpeta_img, exist_ok=True)

        if ext == ".pdf":
            contenido = extraer_pdf(ruta_doc, guardar_img, carpeta_img, usar_ocr)
        elif ext == ".docx":
            contenido = extraer_docx(ruta_doc, guardar_img, carpeta_img, usar_ocr)
        elif ext == ".pptx":
            contenido = extraer_pptx(ruta_doc, guardar_img, carpeta_img, usar_ocr)
        else:
            raise ValueError("Extensión no soportada. Usa .pdf, .docx o .pptx")

        ext_salida = ".md" if formato_salida == "Markdown (.md)" else ".txt"
        ruta_salida = f"{base}_extraido{ext_salida}"

        with open(ruta_salida, "w", encoding="utf-8") as f:
            f.write(contenido)

        ventana.after(0, lambda: finalizar_exito_doc(contenido, ruta_salida))
    except Exception as e:
        mensaje_error = str(e)
        ventana.after(0, lambda msg=mensaje_error: finalizar_error_doc(msg))


def finalizar_exito_doc(contenido, ruta_salida):
    boton_extraer_doc.config(state=tk.NORMAL)
    caja_resultado_doc.insert(tk.END, contenido)
    etiqueta_estado_doc.config(text=f"¡Extraído con éxito en: {os.path.basename(ruta_salida)}!", foreground="#16a34a")
    messagebox.showinfo("Extracción Completada", f"Documento optimizado para IA guardado en:\n{ruta_salida}")


def finalizar_error_doc(mensaje_error):
    boton_extraer_doc.config(state=tk.NORMAL)
    etiqueta_estado_doc.config(text="Error al procesar el documento.", foreground="#dc2626")
    messagebox.showerror("Error", f"No se pudo extraer el documento:\n{mensaje_error}")


# ==========================================
# INTERFAZ GRÁFICA PRINCIPAL (PESTAÑAS)
# ==========================================
ventana = tk.Tk()
ventana.title("Suite Multimedia & Documentos para IA")
ventana.geometry("640x530")
ventana.resizable(False, False)

notebook = ttk.Notebook(ventana)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# --- PESTAÑA 1: YOUTUBE A MP3 ---
tab_yt = ttk.Frame(notebook, padding=15)
notebook.add(tab_yt, text=" 📥 YouTube a MP3 ")

ttk.Label(tab_yt, text="Enlace de YouTube:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
entrada_url = ttk.Entry(tab_yt)
entrada_url.pack(fill="x", pady=(4, 12))

marco_tiempos = ttk.LabelFrame(tab_yt, text="Recorte de tiempo (Opcional - MM:SS o HH:MM:SS)", padding=10)
marco_tiempos.pack(fill="x", pady=(0, 12))

ttk.Label(marco_tiempos, text="Desde (Inicio):").grid(row=0, column=0, padx=(0, 5), sticky="w")
entrada_inicio = ttk.Entry(marco_tiempos, width=15)
entrada_inicio.grid(row=0, column=1, padx=(0, 20))
entrada_inicio.insert(0, "00:00")

ttk.Label(marco_tiempos, text="Hasta (Fin):").grid(row=0, column=2, padx=(0, 5), sticky="w")
entrada_fin = ttk.Entry(marco_tiempos, width=15)
entrada_fin.grid(row=0, column=3)

marco_calidad = ttk.Frame(tab_yt)
marco_calidad.pack(fill="x", pady=(0, 12))
ttk.Label(marco_calidad, text="Calidad de audio (kbps):", font=("Segoe UI", 10, "bold")).pack(side="left")
combo_calidad = ttk.Combobox(marco_calidad, values=["64", "128", "192", "256", "320"], state="readonly", width=10)
combo_calidad.set("192")
combo_calidad.pack(side="left", padx=10)

ttk.Label(tab_yt, text="Carpeta de salida:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
marco_carpeta_yt = ttk.Frame(tab_yt)
marco_carpeta_yt.pack(fill="x", pady=(4, 15))
entrada_carpeta_yt = ttk.Entry(marco_carpeta_yt)
entrada_carpeta_yt.insert(0, os.path.expanduser("~\\Downloads"))
entrada_carpeta_yt.pack(side="left", fill="x", expand=True, padx=(0, 8))
ttk.Button(marco_carpeta_yt, text="Examinar...", command=seleccionar_carpeta_yt).pack(side="right")

boton_descargar = ttk.Button(tab_yt, text="Descargar MP3", command=iniciar_descarga)
boton_descargar.pack(fill="x", ipady=5)

etiqueta_estado_yt = ttk.Label(tab_yt, text="Listo para descargar", font=("Segoe UI", 9))
etiqueta_estado_yt.pack(pady=(10, 0))


# --- PESTAÑA 2: TRANSCRIPTOR DE AUDIO A TXT ---
tab_txt = ttk.Frame(notebook, padding=15)
notebook.add(tab_txt, text=" 📝 Audio a TXT ")

ttk.Label(tab_txt, text="Archivo de audio (.opus, .ogg, .mp3, .wav, .m4a):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
marco_archivo = ttk.Frame(tab_txt)
marco_archivo.pack(fill="x", pady=(4, 12))
entrada_archivo_audio = ttk.Entry(marco_archivo)
entrada_archivo_audio.pack(side="left", fill="x", expand=True, padx=(0, 8))
ttk.Button(marco_archivo, text="Buscar Audio...", command=seleccionar_archivo_audio).pack(side="right")

marco_opciones_ia = ttk.LabelFrame(tab_txt, text="Configuración del Motor IA (Faster-Whisper)", padding=10)
marco_opciones_ia.pack(fill="x", pady=(0, 12))

ttk.Label(marco_opciones_ia, text="Precisión:").grid(row=0, column=0, padx=(0, 5), sticky="w")
combo_modelo = ttk.Combobox(marco_opciones_ia, values=["tiny", "base", "small", "medium"], state="readonly", width=10)
combo_modelo.set("small")
combo_modelo.grid(row=0, column=1, padx=(0, 15))

ttk.Label(marco_opciones_ia, text="Idioma:").grid(row=0, column=2, padx=(0, 5), sticky="w")
combo_idioma = ttk.Combobox(marco_opciones_ia, values=["es", "en", "Auto"], state="readonly", width=8)
combo_idioma.set("es")
combo_idioma.grid(row=0, column=3, padx=(0, 15))

var_marcas_tiempo = tk.BooleanVar(value=False)
ttk.Checkbutton(marco_opciones_ia, text="Incluir tiempos [MM:SS]", variable=var_marcas_tiempo).grid(row=0, column=4)

boton_transcribir = ttk.Button(tab_txt, text="Transcribir y Guardar .TXT", command=iniciar_transcripcion)
boton_transcribir.pack(fill="x", ipady=5)

etiqueta_estado_txt = ttk.Label(tab_txt, text="Selecciona un audio para comenzar", font=("Segoe UI", 9))
etiqueta_estado_txt.pack(pady=(6, 6))

ttk.Label(tab_txt, text="Vista previa del texto transcrito:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
caja_resultado = tk.Text(tab_txt, height=7, wrap="word", font=("Segoe UI", 9))
caja_resultado.pack(fill="both", expand=True, pady=(4, 0))


# --- PESTAÑA 3: DOCUMENTOS A MARKDOWN / TXT + OCR ---
tab_doc = ttk.Frame(notebook, padding=15)
notebook.add(tab_doc, text=" 📄 Docs a Markdown/TXT ")

ttk.Label(tab_doc, text="Documento (.pdf, .docx, .pptx):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
marco_doc = ttk.Frame(tab_doc)
marco_doc.pack(fill="x", pady=(4, 12))
entrada_archivo_doc = ttk.Entry(marco_doc)
entrada_archivo_doc.pack(side="left", fill="x", expand=True, padx=(0, 8))
ttk.Button(marco_doc, text="Buscar Documento...", command=seleccionar_documento).pack(side="right")

marco_opciones_doc = ttk.LabelFrame(tab_doc, text="Opciones de Extracción y OCR Local", padding=10)
marco_opciones_doc.pack(fill="x", pady=(0, 12))

ttk.Label(marco_opciones_doc, text="Salida:").grid(row=0, column=0, padx=(0, 5), sticky="w")
combo_formato_doc = ttk.Combobox(marco_opciones_doc, values=["Markdown (.md)", "Texto Plano (.txt)"], state="readonly", width=16)
combo_formato_doc.set("Markdown (.md)")
combo_formato_doc.grid(row=0, column=1, padx=(0, 15))

var_usar_ocr = tk.BooleanVar(value=True)
ttk.Checkbutton(
    marco_opciones_doc,
    text="OCR en imágenes (Leer texto)",
    variable=var_usar_ocr
).grid(row=0, column=2, padx=(0, 15))

var_extraer_img = tk.BooleanVar(value=False)
ttk.Checkbutton(
    marco_opciones_doc,
    text="Guardar imágenes aparte",
    variable=var_extraer_img
).grid(row=0, column=3)

boton_extraer_doc = ttk.Button(tab_doc, text="Extraer Contenido y Ejecutar OCR", command=iniciar_extraccion_doc)
boton_extraer_doc.pack(fill="x", ipady=5)

etiqueta_estado_doc = ttk.Label(tab_doc, text="Listo para comprimir documentos con OCR local", font=("Segoe UI", 9))
etiqueta_estado_doc.pack(pady=(6, 6))

ttk.Label(tab_doc, text="Vista previa del resultado generado:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
caja_resultado_doc = tk.Text(tab_doc, height=7, wrap="word", font=("Consolas", 9))
caja_resultado_doc.pack(fill="both", expand=True, pady=(4, 0))

ventana.mainloop()