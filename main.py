# Media & Doc AI Suite - Suite Multimedia y Documentos para IA
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

"""Suite de escritorio para descarga de audio, transcripción local y extracción de documentos."""

import gc
import os
import re
import shutil
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import docx
from docx.opc.constants import RELATIONSHIP_TYPE
import pymupdf
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import yt_dlp

# Constantes de configuración
EXTENSIONES_AUDIO_SOPORTADAS = (
    ".opus", ".ogg", ".mp3", ".wav", ".m4a", ".aac", ".flac", ".mp4"
)
EXTENSIONES_DOCUMENTO_SOPORTADAS = (".pdf", ".docx", ".pptx")
LIMITE_VISTA_PREVIA = 50_000

COLOR_PROGRESO = "#0284c7"
COLOR_EXITO = "#16a34a"
COLOR_ERROR = "#dc2626"

# Cachés globales protegidas por cerrojo para reutilizar modelos entre operaciones
_whisper_lock = threading.Lock()
_ocr_lock = threading.Lock()
modelo_whisper_cache = {"nombre": None, "instancia": None}
motor_ocr_cache = None


# ==========================================
# UTILIDADES COMPARTIDAS Y CACHÉS DE MODELOS
# ==========================================
def obtener_carpeta_descargas_predeterminada():
    """Devuelve la ruta portable de descargas del usuario."""
    return os.path.join(os.path.expanduser("~"), "Downloads")


def verificar_permisos_escritura(carpeta):
    """Comprueba que la carpeta exista y permita crear archivos."""
    if not carpeta or not os.path.isdir(carpeta):
        return False
    try:
        with tempfile.TemporaryFile(dir=carpeta):
            return True
    except OSError:
        return False


def sanitizar_nombre_archivo(nombre):
    """Elimina caracteres inválidos de un nombre de archivo."""
    base = os.path.basename(str(nombre or "imagen")).strip()
    limpio = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", base)
    return limpio or "imagen"


def obtener_modelo_whisper(tamano_modelo):
    """Inicializa o reemplaza el modelo Faster-Whisper liberando el anterior antes de cargar."""
    from faster_whisper import WhisperModel

    with _whisper_lock:
        if (
            modelo_whisper_cache["nombre"] != tamano_modelo
            or modelo_whisper_cache["instancia"] is None
        ):
            if modelo_whisper_cache["instancia"] is not None:
                modelo_whisper_cache["instancia"] = None
                modelo_whisper_cache["nombre"] = None
                gc.collect()
            modelo_whisper_cache["instancia"] = WhisperModel(
                tamano_modelo, device="cpu", compute_type="int8"
            )
            modelo_whisper_cache["nombre"] = tamano_modelo
        return modelo_whisper_cache["instancia"]


def obtener_motor_ocr():
    """Inicializa RapidOCR una sola vez bajo demanda con sincronización."""
    global motor_ocr_cache
    with _ocr_lock:
        if motor_ocr_cache is None:
            from rapidocr_onnxruntime import RapidOCR

            motor_ocr_cache = RapidOCR()
        return motor_ocr_cache


def aplicar_ocr_bytes(imagen_bytes):
    """Extrae texto de los bytes de una imagen y devuelve una tupla (texto, error)."""
    try:
        ocr = obtener_motor_ocr()
        resultado, _ = ocr(imagen_bytes)
        if not resultado:
            return "", None
        # resultado es una lista de listas: [caja_coordenadas, texto, confianza]
        lineas = [str(linea[1]).strip() for linea in resultado if str(linea[1]).strip()]
        return "\n".join(lineas), None
    except Exception as exc:
        return "", str(exc)


# ==========================================
# LÓGICA MÓDULO 1: YOUTUBE A MP3
# ==========================================
def convertir_a_segundos(tiempo_str):
    """Convierte SS, MM:SS o HH:MM:SS a segundos no negativos; devuelve None si está vacío."""
    tiempo_str = tiempo_str.strip()
    if not tiempo_str:
        return None

    partes = tiempo_str.split(":")
    if len(partes) > 3:
        raise ValueError(
            f"Formato inválido: '{tiempo_str}'. Usa SS, MM:SS o HH:MM:SS."
        )

    try:
        if any(not p.isdigit() for p in partes):
            raise ValueError
        valores = [int(p) for p in partes]
    except ValueError as exc:
        raise ValueError(
            f"Formato inválido: '{tiempo_str}'. Usa SS, MM:SS o HH:MM:SS (solo valores no negativos)."
        ) from exc

    if len(valores) == 1:
        return valores[0]
    if len(valores) == 2:
        minutos, segundos = valores
        if segundos >= 60:
            raise ValueError(
                f"Formato inválido: '{tiempo_str}'. Los segundos deben estar entre 00 y 59."
            )
        return minutos * 60 + segundos

    horas, minutos, segundos = valores
    if minutos >= 60 or segundos >= 60:
        raise ValueError(
            f"Formato inválido: '{tiempo_str}'. Minutos y segundos deben estar entre 00 y 59."
        )
    return horas * 3600 + minutos * 60 + segundos


def construir_opciones_yt_dlp(seg_inicio, seg_fin, calidad, carpeta_destino):
    """Construye la configuración de yt_dlp activando recorte solo cuando el usuario lo solicita."""
    plantilla_salida = os.path.join(carpeta_destino, "%(title)s.%(ext)s")
    ydl_opts = {
        "format": "bestaudio/best",
        "outtmpl": plantilla_salida,
        "noplaylist": True,
        "postprocessors": [{
            "key": "FFmpegExtractAudio",
            "preferredcodec": "mp3",
            "preferredquality": calidad,
        }],
    }

    tiene_recorte = (seg_inicio is not None and seg_inicio > 0) or seg_fin is not None
    if tiene_recorte:
        inicio = seg_inicio if seg_inicio is not None else 0
        fin = seg_fin if seg_fin is not None else float("inf")
        ydl_opts["download_ranges"] = yt_dlp.utils.download_range_func(
            None, [(inicio, fin)]
        )
        ydl_opts["force_keyframes_at_cuts"] = True

    return ydl_opts


def descargar_audio_youtube(url, seg_inicio, seg_fin, calidad, carpeta_destino):
    """Descarga y convierte a MP3 un enlace de YouTube usando yt_dlp y FFmpeg."""
    ydl_opts = construir_opciones_yt_dlp(seg_inicio, seg_fin, calidad, carpeta_destino)
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        ydl.download([url])


# ==========================================
# LÓGICA MÓDULO 2: AUDIO A TXT (WHISPER)
# ==========================================
def formatear_segundo(seg):
    """Formatea segundos no negativos como MM:SS o HH:MM:SS."""
    if seg < 0:
        raise ValueError("Los segundos no pueden ser negativos.")
    m, s = divmod(int(seg), 60)
    h, m = divmod(m, 60)
    return f"{h:02d}:{m:02d}:{s:02d}" if h > 0 else f"{m:02d}:{s:02d}"


def transcribir_audio(ruta_audio, tamano_modelo, idioma, incluir_tiempos):
    """Transcribe un archivo de audio con Faster-Whisper y guarda el resultado en .txt."""
    modelo = obtener_modelo_whisper(tamano_modelo)
    lang_code = None if idioma == "Auto" else idioma

    segmentos, _ = modelo.transcribe(ruta_audio, beam_size=5, language=lang_code)

    lineas = []
    for seg in segmentos:
        texto_limpio = seg.text.strip()
        if not texto_limpio:
            continue
        if incluir_tiempos:
            marca = f"[{formatear_segundo(seg.start)} -> {formatear_segundo(seg.end)}] {texto_limpio}"
            lineas.append(marca)
        else:
            lineas.append(texto_limpio)

    separador = "\n" if incluir_tiempos else " "
    texto_final = separador.join(lineas)

    base, _ = os.path.splitext(ruta_audio)
    ruta_txt = f"{base}_transcripcion.txt"
    with open(ruta_txt, "w", encoding="utf-8") as archivo:
        archivo.write(texto_final)

    advertencia = None if texto_final else "No se detectó voz legible en el archivo de audio."
    return texto_final, ruta_txt, advertencia


# ==========================================
# LÓGICA MÓDULO 3: DOCUMENTOS (PDF/DOCX/PPTX) + OCR
# ==========================================
def convertir_tabla_a_md(filas_datos):
    """Convierte una matriz de celdas a una tabla Markdown válida escapando barras verticales."""
    if not filas_datos:
        return ""

    num_cols = max((len(fila) for fila in filas_datos), default=0)
    if num_cols == 0:
        return ""

    def limpiar_celda(celda):
        """Normaliza saltos de línea, valores nulos y barras verticales de una celda."""
        if celda is None:
            return ""
        texto = str(celda).replace("\r\n", "\n").replace("\r", "\n").replace("\n", " ").strip()
        return texto.replace("|", r"\|")

    def normalizar_fila(fila):
        """Limpia las celdas de una fila y la rellena hasta completar el ancho máximo."""
        celdas = [limpiar_celda(celda) for celda in fila]
        if len(celdas) < num_cols:
            celdas.extend([""] * (num_cols - len(celdas)))
        return celdas

    encabezado = normalizar_fila(filas_datos[0])
    md = [
        "| " + " | ".join(encabezado) + " |",
        "| " + " | ".join(["---"] * num_cols) + " |",
    ]
    for fila in filas_datos[1:]:
        md.append("| " + " | ".join(normalizar_fila(fila)) + " |")
    return "\n".join(md) + "\n"


def formatear_cita_ocr(texto_ocr, error_ocr):
    """Formatea el bloque Markdown de resultado OCR preservando prefijo de cita en todas las líneas."""
    if error_ocr:
        return "\n> *(OCR no disponible para esta imagen)*"
    if texto_ocr:
        lineas_cita = "\n".join(f"> {linea}" for linea in texto_ocr.splitlines())
        return f"\n> **Texto extraído por OCR:**\n{lineas_cita}"
    return "\n> *(OCR: sin texto legible detectado en la imagen)*"


def procesar_imagen_extraida(
    img_bytes,
    encabezado_md,
    nombre_archivo,
    guardar_img,
    carpeta_img,
    usar_ocr,
    advertencias,
):
    """Guarda opcionalmente una imagen (creando carpeta bajo demanda) y ejecuta OCR."""
    info_img = encabezado_md
    nombre_guardado = None
    texto_ocr = ""
    error_ocr = None

    if guardar_img:
        os.makedirs(carpeta_img, exist_ok=True)
        nombre_seguro = sanitizar_nombre_archivo(nombre_archivo)
        ruta_img = os.path.join(carpeta_img, nombre_seguro)
        with open(ruta_img, "wb") as f_img:
            f_img.write(img_bytes)
        nombre_guardado = nombre_seguro
        info_img += f" -> Guardada como `{nombre_seguro}`"

    if usar_ocr:
        texto_ocr, error_ocr = aplicar_ocr_bytes(img_bytes)
        if error_ocr:
            advertencias.append(f"Error de OCR en {nombre_archivo}: {error_ocr}")
        info_img += formatear_cita_ocr(texto_ocr, error_ocr)

    return info_img, nombre_guardado, texto_ocr, error_ocr


def extraer_pdf(ruta, guardar_img, carpeta_img, usar_ocr, advertencias=None):
    """Extrae texto, tablas e imágenes (con deduplicación por xref y OCR opcional) de un PDF."""
    if advertencias is None:
        advertencias = []

    bloques = [f"# Documento PDF: {os.path.basename(ruta)}\n"]
    contador_img = 0
    xrefs_cache = {}

    with pymupdf.open(ruta) as doc:
        if doc.needs_pass:
            raise ValueError(
                "El archivo PDF está protegido con contraseña y no se puede procesar."
            )

        for num_pag, pagina in enumerate(doc, start=1):
            bloques.append(f"## Página {num_pag}\n")
            texto = pagina.get_text("text").strip()
            if texto:
                bloques.append(texto + "\n")

            if hasattr(pagina, "find_tables"):
                try:
                    tablas_pag = pagina.find_tables()
                    for idx_tabla, tabla in enumerate(tablas_pag.tables, start=1):
                        filas = tabla.extract()
                        tabla_md = convertir_tabla_a_md(filas)
                        if tabla_md:
                            bloques.append(
                                f"**Tabla {idx_tabla} (Página {num_pag}):**\n{tabla_md}"
                            )
                except Exception as exc:
                    advertencias.append(
                        f"No se pudieron extraer tablas en página {num_pag}: {exc}"
                    )

            lista_imagenes = pagina.get_images(full=True)
            for img in lista_imagenes:
                contador_img += 1
                xref = img[0]
                ancho, alto = img[2], img[3]
                encabezado = (
                    f"> **[Imagen #{contador_img} en Página {num_pag}]** "
                    f"(Resolución: {ancho}x{alto}px)"
                )

                if xref in xrefs_cache:
                    nombre_prev, texto_prev, error_prev = xrefs_cache[xref]
                    info_img = encabezado
                    if nombre_prev:
                        info_img += f" -> Guardada como `{nombre_prev}`"
                    if usar_ocr:
                        info_img += formatear_cita_ocr(texto_prev, error_prev)
                    bloques.append(info_img + "\n")
                    continue

                if not (guardar_img or usar_ocr):
                    bloques.append(encabezado + "\n")
                    continue

                pix = None
                try:
                    pix = pymupdf.Pixmap(doc, xref)
                    if pix.colorspace is None:
                        advertencias.append(
                            f"Imagen #{contador_img} en página {num_pag} omitida (máscara sin espacio de color)."
                        )
                        bloques.append(encabezado + "\n")
                        continue
                    if pix.n - pix.alpha > 3:
                        pix_rgb = pymupdf.Pixmap(pymupdf.csRGB, pix)
                        pix = pix_rgb
                    img_bytes = pix.tobytes("png")
                except Exception as exc:
                    advertencias.append(
                        f"No se pudo decodificar imagen #{contador_img} en página {num_pag}: {exc}"
                    )
                    bloques.append(encabezado + "\n")
                    continue
                finally:
                    pix = None

                nombre_img = f"pdf_pag{num_pag}_img{contador_img}.png"
                info_img, nombre_guardado, texto_ocr, error_ocr = procesar_imagen_extraida(
                    img_bytes=img_bytes,
                    encabezado_md=encabezado,
                    nombre_archivo=nombre_img,
                    guardar_img=guardar_img,
                    carpeta_img=carpeta_img,
                    usar_ocr=usar_ocr,
                    advertencias=advertencias,
                )
                xrefs_cache[xref] = (nombre_guardado, texto_ocr, error_ocr)
                bloques.append(info_img + "\n")

    return "\n".join(bloques)


def extraer_docx(ruta, guardar_img, carpeta_img, usar_ocr, advertencias=None):
    """Extrae párrafos, tablas e imágenes incrustadas (con OCR opcional) de un archivo .docx."""
    if advertencias is None:
        advertencias = []

    doc = docx.Document(ruta)
    bloques = [f"# Documento Word: {os.path.basename(ruta)}\n"]

    for p in doc.paragraphs:
        texto = p.text.strip()
        if not texto:
            continue
        estilo = (p.style.name or "").lower() if p.style is not None else ""
        if "heading 1" in estilo or "título 1" in estilo:
            bloques.append(f"## {texto}\n")
        elif "heading 2" in estilo or "título 2" in estilo:
            bloques.append(f"### {texto}\n")
        elif "heading" in estilo or "título" in estilo:
            bloques.append(f"#### {texto}\n")
        else:
            bloques.append(f"{texto}\n")

    if doc.tables:
        bloques.append("## Tablas del documento\n")
        for idx, tabla in enumerate(doc.tables, start=1):
            filas = [[celda.text.strip() for celda in fila.cells] for fila in tabla.rows]
            bloques.append(f"**Tabla {idx}:**\n" + convertir_tabla_a_md(filas))

    contador_img = 0
    for rel in doc.part.rels.values():
        if getattr(rel, "is_external", False):
            continue
        reltype = getattr(rel, "reltype", "")
        if reltype != RELATIONSHIP_TYPE.IMAGE and "image" not in reltype.lower():
            continue

        contador_img += 1
        nombre_orig = sanitizar_nombre_archivo(os.path.basename(rel.target_ref))
        try:
            img_bytes = rel.target_part.blob
        except Exception as exc:
            advertencias.append(
                f"No se pudo leer la imagen incrustada #{contador_img} ({nombre_orig}): {exc}"
            )
            continue

        encabezado = (
            f"> **[Imagen incrustada #{contador_img}]** Referencia interna: `{nombre_orig}`"
        )
        nombre_salida = f"docx_img{contador_img}_{nombre_orig}"
        info_img, _, _, _ = procesar_imagen_extraida(
            img_bytes=img_bytes,
            encabezado_md=encabezado,
            nombre_archivo=nombre_salida,
            guardar_img=guardar_img,
            carpeta_img=carpeta_img,
            usar_ocr=usar_ocr,
            advertencias=advertencias,
        )
        bloques.append(info_img + "\n")

    return "\n".join(bloques)


def _obtener_imagen_de_forma_pptx(shape):
    """Devuelve el objeto imagen de una forma PICTURE o PLACEHOLDER si existe, o None."""
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return shape.image
    if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and hasattr(shape, "image"):
        try:
            return shape.image
        except (AttributeError, ValueError):
            return None
    return None


def extraer_pptx(ruta, guardar_img, carpeta_img, usar_ocr, advertencias=None):
    """Extrae texto, tablas, imágenes (incluidos marcadores de posición) y notas de un .pptx."""
    if advertencias is None:
        advertencias = []

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
                filas = [
                    [celda.text.strip() for celda in fila.cells]
                    for fila in shape.table.rows
                ]
                bloques.append(convertir_tabla_a_md(filas))

            imagen = _obtener_imagen_de_forma_pptx(shape)
            if imagen is not None:
                contador_img += 1
                alt_text = shape.name or "Sin descripción"
                try:
                    descr = shape._element.xpath(".//p:cNvPr/@descr")
                    if descr and descr[0].strip():
                        alt_text = descr[0].strip()
                except Exception:
                    pass
                alt_text = alt_text.replace("\n", " ").replace("*", "")

                try:
                    img_bytes = imagen.blob
                    ext = re.sub(r"[^a-zA-Z0-9]", "", str(imagen.ext or "png")) or "png"
                except Exception as exc:
                    advertencias.append(
                        f"No se pudo leer imagen #{contador_img} en diapositiva {num_slide}: {exc}"
                    )
                    continue

                encabezado = (
                    f"> **[Imagen #{contador_img} en Diapositiva {num_slide}]** "
                    f"Texto alternativo: *{alt_text}*"
                )
                nombre_img = f"pptx_diap{num_slide}_img{contador_img}.{ext}"
                info_img, _, _, _ = procesar_imagen_extraida(
                    img_bytes=img_bytes,
                    encabezado_md=encabezado,
                    nombre_archivo=nombre_img,
                    guardar_img=guardar_img,
                    carpeta_img=carpeta_img,
                    usar_ocr=usar_ocr,
                    advertencias=advertencias,
                )
                bloques.append(info_img + "\n")

        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notas = slide.notes_slide.notes_text_frame.text.strip()
            if notas:
                bloques.append(f"**Notas del orador:** {notas}\n")

    return "\n".join(bloques)


def extraer_documento(ruta_doc, guardar_img, usar_ocr):
    """Valida la extensión, extrae contenido a Markdown (.md) y devuelve (contenido, ruta_salida, advertencias)."""
    base, ext = os.path.splitext(ruta_doc)
    ext = ext.lower()
    if ext not in EXTENSIONES_DOCUMENTO_SOPORTADAS:
        raise ValueError("Extensión no soportada. Usa .pdf, .docx o .pptx.")

    carpeta_img = f"{base}_imagenes"
    advertencias = []

    if ext == ".pdf":
        contenido = extraer_pdf(ruta_doc, guardar_img, carpeta_img, usar_ocr, advertencias)
    elif ext == ".docx":
        contenido = extraer_docx(ruta_doc, guardar_img, carpeta_img, usar_ocr, advertencias)
    else:
        contenido = extraer_pptx(ruta_doc, guardar_img, carpeta_img, usar_ocr, advertencias)

    ruta_salida = f"{base}_extraido.md"
    with open(ruta_salida, "w", encoding="utf-8") as archivo:
        archivo.write(contenido)

    return contenido, ruta_salida, advertencias


# ==========================================
# INTERFAZ GRÁFICA PRINCIPAL (PESTAÑAS)
# ==========================================
class AplicacionSuite:
    """Interfaz gráfica Tkinter que agrupa los tres módulos locales."""

    def __init__(self, ventana):
        """Configura la ventana principal, el estado de tareas activas y las tres pestañas."""
        self.ventana = ventana
        self._cerrando = False
        self._tareas_lock = threading.Lock()
        self._tareas_activas = 0

        self.ventana.title("Suite Multimedia y Documentos para IA")
        self.ventana.geometry("680x560")
        self.ventana.minsize(640, 530)
        self.ventana.resizable(True, True)
        self.ventana.protocol("WM_DELETE_WINDOW", self._al_cerrar_ventana)

        self.fuente_negrita = ("TkDefaultFont", 10, "bold")
        self.fuente_estado = ("TkDefaultFont", 9)
        self.fuente_mono = ("TkFixedFont", 9)

        self.notebook = ttk.Notebook(self.ventana)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self._construir_tab_youtube()
        self._construir_tab_audio()
        self._construir_tab_documentos()

    def _registrar_inicio_tarea(self):
        """Incrementa de forma segura el contador de tareas en segundo plano activas."""
        with self._tareas_lock:
            self._tareas_activas += 1

    def _registrar_fin_tarea(self):
        """Decrementa de forma segura el contador de tareas en segundo plano activas."""
        with self._tareas_lock:
            self._tareas_activas = max(0, self._tareas_activas - 1)

    def _notificar_ui(self, callback):
        """Ejecuta una actualización en el hilo de interfaz ignorando errores si la ventana cerró."""
        if self._cerrando:
            return
        try:
            self.ventana.after(0, callback)
        except (tk.TclError, RuntimeError):
            pass

    def _al_cerrar_ventana(self):
        """Solicita confirmación si hay tareas activas antes de destruir la ventana."""
        with self._tareas_lock:
            activas = self._tareas_activas
        if activas > 0:
            confirmar = messagebox.askyesno(
                "Operación en curso",
                "Hay tareas en segundo plano en ejecución. ¿Deseas cerrar la aplicación de todos modos?",
            )
            if not confirmar:
                return
        self._cerrando = True
        self.ventana.destroy()

    def _insertar_vista_previa(self, widget_texto, contenido):
        """Inserta texto en un widget tk.Text aplicando un límite seguro de vista previa."""
        widget_texto.delete("1.0", tk.END)
        if len(contenido) > LIMITE_VISTA_PREVIA:
            recorte = (
                contenido[:LIMITE_VISTA_PREVIA]
                + "\n\n[Vista previa truncada; consulta el archivo guardado para ver el contenido completo.]"
            )
            widget_texto.insert(tk.END, recorte)
        else:
            widget_texto.insert(tk.END, contenido)

    def _crear_caja_texto_con_scroll(self, padre, fuente):
        """Crea un cuadro tk.Text con barra de desplazamiento vertical integrada."""
        contenedor = ttk.Frame(padre)
        contenedor.pack(fill="both", expand=True, pady=(4, 0))
        barra = ttk.Scrollbar(contenedor, orient="vertical")
        caja = tk.Text(
            contenedor, height=7, wrap="word", font=fuente, yscrollcommand=barra.set
        )
        barra.config(command=caja.yview)
        caja.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        return caja

    # --- PESTAÑA 1: YOUTUBE A MP3 ---
    def _construir_tab_youtube(self):
        """Construye los controles de la pestaña de descarga y recorte de YouTube a MP3."""
        tab_yt = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab_yt, text=" YouTube a MP3 ")

        ttk.Label(tab_yt, text="Enlace de YouTube:", font=self.fuente_negrita).pack(anchor="w")
        self.entrada_url = ttk.Entry(tab_yt)
        self.entrada_url.pack(fill="x", pady=(4, 12))

        marco_tiempos = ttk.LabelFrame(
            tab_yt,
            text="Recorte de tiempo (opcional: SS, MM:SS o HH:MM:SS)",
            padding=10,
        )
        marco_tiempos.pack(fill="x", pady=(0, 12))

        ttk.Label(marco_tiempos, text="Desde (inicio):").grid(
            row=0, column=0, padx=(0, 5), sticky="w"
        )
        self.entrada_inicio = ttk.Entry(marco_tiempos, width=15)
        self.entrada_inicio.grid(row=0, column=1, padx=(0, 20))

        ttk.Label(marco_tiempos, text="Hasta (fin):").grid(
            row=0, column=2, padx=(0, 5), sticky="w"
        )
        self.entrada_fin = ttk.Entry(marco_tiempos, width=15)
        self.entrada_fin.grid(row=0, column=3)

        marco_calidad = ttk.Frame(tab_yt)
        marco_calidad.pack(fill="x", pady=(0, 12))
        ttk.Label(
            marco_calidad, text="Calidad de audio (kbps):", font=self.fuente_negrita
        ).pack(side="left")
        self.combo_calidad = ttk.Combobox(
            marco_calidad,
            values=["64", "128", "192", "256", "320"],
            state="readonly",
            width=10,
        )
        self.combo_calidad.set("192")
        self.combo_calidad.pack(side="left", padx=10)

        ttk.Label(tab_yt, text="Carpeta de salida:", font=self.fuente_negrita).pack(anchor="w")
        marco_carpeta_yt = ttk.Frame(tab_yt)
        marco_carpeta_yt.pack(fill="x", pady=(4, 15))
        self.entrada_carpeta_yt = ttk.Entry(marco_carpeta_yt)
        self.entrada_carpeta_yt.insert(0, obtener_carpeta_descargas_predeterminada())
        self.entrada_carpeta_yt.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(
            marco_carpeta_yt, text="Examinar...", command=self._seleccionar_carpeta_yt
        ).pack(side="right")

        self.boton_descargar = ttk.Button(
            tab_yt, text="Descargar MP3", command=self._iniciar_descarga
        )
        self.boton_descargar.pack(fill="x", ipady=5)

        self.etiqueta_estado_yt = ttk.Label(
            tab_yt, text="Listo para descargar.", font=self.fuente_estado
        )
        self.etiqueta_estado_yt.pack(pady=(10, 0))

    def _seleccionar_carpeta_yt(self):
        """Abre un diálogo de carpeta y actualiza la ruta de salida de YouTube."""
        carpeta = filedialog.askdirectory()
        if carpeta:
            self.entrada_carpeta_yt.delete(0, tk.END)
            self.entrada_carpeta_yt.insert(0, carpeta)

    def _iniciar_descarga(self):
        """Valida entradas, permisos y FFmpeg antes de lanzar el hilo de descarga."""
        url = self.entrada_url.get().strip()
        inicio_str = self.entrada_inicio.get().strip()
        fin_str = self.entrada_fin.get().strip()
        calidad = self.combo_calidad.get()
        carpeta_destino = self.entrada_carpeta_yt.get().strip()

        if not url:
            messagebox.showwarning("Campo vacío", "Ingresa el enlace de YouTube.")
            return

        if not verificar_permisos_escritura(carpeta_destino):
            messagebox.showerror(
                "Carpeta inválida",
                "Selecciona una carpeta de salida existente con permisos de escritura.",
            )
            return

        if not shutil.which("ffmpeg"):
            messagebox.showerror(
                "FFmpeg no encontrado",
                "No se encontró FFmpeg en el PATH del sistema. Instálalo para exportar en MP3.",
            )
            return

        try:
            seg_inicio = convertir_a_segundos(inicio_str)
            seg_fin = convertir_a_segundos(fin_str)
        except ValueError as exc:
            messagebox.showerror("Error de formato", str(exc))
            return

        inicio_efectivo = seg_inicio if seg_inicio is not None else 0
        if seg_fin is not None and seg_fin <= inicio_efectivo:
            messagebox.showerror(
                "Error en tiempos",
                "El tiempo de fin debe ser mayor que el tiempo de inicio.",
            )
            return

        self.boton_descargar.config(state=tk.DISABLED)
        self.etiqueta_estado_yt.config(
            text="Descargando y procesando audio...", foreground=COLOR_PROGRESO
        )
        self._registrar_inicio_tarea()

        threading.Thread(
            target=self._proceso_descarga,
            args=(url, seg_inicio, seg_fin, calidad, carpeta_destino),
            daemon=True,
        ).start()

    def _proceso_descarga(self, url, seg_inicio, seg_fin, calidad, carpeta_destino):
        """Ejecuta la descarga en segundo plano y notifica el resultado a la interfaz."""
        try:
            descargar_audio_youtube(url, seg_inicio, seg_fin, calidad, carpeta_destino)
            self._notificar_ui(lambda: self._finalizar_exito_yt(carpeta_destino))
        except Exception as exc:
            mensaje_error = str(exc)
            self._notificar_ui(lambda msg=mensaje_error: self._finalizar_error_yt(msg))
        finally:
            self._registrar_fin_tarea()

    def _finalizar_exito_yt(self, carpeta):
        """Restaura el botón de descarga y muestra la confirmación de archivo guardado."""
        self.boton_descargar.config(state=tk.NORMAL)
        self.etiqueta_estado_yt.config(text="Descarga completada.", foreground=COLOR_EXITO)
        messagebox.showinfo("Descarga completada", f"MP3 guardado en:\n{carpeta}")

    def _finalizar_error_yt(self, mensaje_error):
        """Restaura el botón de descarga e informa el error ocurrido al usuario."""
        self.boton_descargar.config(state=tk.NORMAL)
        self.etiqueta_estado_yt.config(
            text="Error durante la descarga.", foreground=COLOR_ERROR
        )
        messagebox.showerror(
            "Error de descarga", f"No se pudo completar la descarga:\n{mensaje_error}"
        )

    # --- PESTAÑA 2: TRANSCRIPTOR DE AUDIO A TXT ---
    def _construir_tab_audio(self):
        """Construye los controles de la pestaña de transcripción de audio a texto."""
        tab_txt = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab_txt, text=" Audio a TXT ")

        extensiones_texto = ", ".join(EXTENSIONES_AUDIO_SOPORTADAS)
        ttk.Label(
            tab_txt,
            text=f"Archivo de audio ({extensiones_texto}):",
            font=self.fuente_negrita,
        ).pack(anchor="w")
        marco_archivo = ttk.Frame(tab_txt)
        marco_archivo.pack(fill="x", pady=(4, 12))
        self.entrada_archivo_audio = ttk.Entry(marco_archivo)
        self.entrada_archivo_audio.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(
            marco_archivo, text="Examinar...", command=self._seleccionar_archivo_audio
        ).pack(side="right")

        marco_opciones_ia = ttk.LabelFrame(
            tab_txt, text="Configuración de Faster-Whisper", padding=10
        )
        marco_opciones_ia.pack(fill="x", pady=(0, 12))

        ttk.Label(marco_opciones_ia, text="Tamaño de modelo:").grid(
            row=0, column=0, padx=(0, 5), sticky="w"
        )
        self.combo_modelo = ttk.Combobox(
            marco_opciones_ia,
            values=["tiny", "base", "small", "medium"],
            state="readonly",
            width=10,
        )
        self.combo_modelo.set("small")
        self.combo_modelo.grid(row=0, column=1, padx=(0, 15))

        ttk.Label(marco_opciones_ia, text="Idioma:").grid(
            row=0, column=2, padx=(0, 5), sticky="w"
        )
        self.combo_idioma = ttk.Combobox(
            marco_opciones_ia,
            values=["es", "en", "Auto"],
            state="readonly",
            width=8,
        )
        self.combo_idioma.set("es")
        self.combo_idioma.grid(row=0, column=3, padx=(0, 15))

        self.var_marcas_tiempo = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            marco_opciones_ia,
            text="Incluir marcas de tiempo",
            variable=self.var_marcas_tiempo,
        ).grid(row=0, column=4)

        self.boton_transcribir = ttk.Button(
            tab_txt, text="Transcribir y guardar .txt", command=self._iniciar_transcripcion
        )
        self.boton_transcribir.pack(fill="x", ipady=5)

        self.etiqueta_estado_txt = ttk.Label(
            tab_txt, text="Selecciona un audio para comenzar.", font=self.fuente_estado
        )
        self.etiqueta_estado_txt.pack(pady=(6, 6))

        ttk.Label(
            tab_txt, text="Vista previa del texto transcrito:", font=self.fuente_negrita
        ).pack(anchor="w")
        self.caja_resultado = self._crear_caja_texto_con_scroll(tab_txt, self.fuente_estado)

    def _seleccionar_archivo_audio(self):
        """Abre un selector de archivos de audio y carga la ruta elegida en el formulario."""
        patron_audio = " ".join(f"*{ext}" for ext in EXTENSIONES_AUDIO_SOPORTADAS)
        tipos = [
            ("Archivos de audio (WhatsApp, MP3, etc.)", patron_audio),
            ("Todos los archivos", "*.*"),
        ]
        archivo = filedialog.askopenfilename(title="Seleccionar audio", filetypes=tipos)
        if archivo:
            self.entrada_archivo_audio.delete(0, tk.END)
            self.entrada_archivo_audio.insert(0, archivo)

    def _iniciar_transcripcion(self):
        """Valida el archivo de audio y permisos de escritura antes de lanzar la transcripción."""
        ruta_audio = self.entrada_archivo_audio.get().strip()
        tamano_modelo = self.combo_modelo.get()
        idioma = self.combo_idioma.get()
        incluir_tiempos = self.var_marcas_tiempo.get()

        if not ruta_audio:
            messagebox.showwarning("Campo vacío", "Selecciona un archivo de audio.")
            return
        if not os.path.isfile(ruta_audio):
            messagebox.showwarning("Archivo inválido", "El archivo de audio seleccionado no existe.")
            return

        carpeta_salida = os.path.dirname(os.path.abspath(ruta_audio))
        if not verificar_permisos_escritura(carpeta_salida):
            messagebox.showerror(
                "Sin permisos de escritura",
                f"No hay permisos para guardar la transcripción en:\n{carpeta_salida}",
            )
            return

        self.boton_transcribir.config(state=tk.DISABLED)
        self.caja_resultado.delete("1.0", tk.END)
        self.etiqueta_estado_txt.config(
            text=f"Cargando modelo '{tamano_modelo}' y procesando...",
            foreground=COLOR_PROGRESO,
        )
        self._registrar_inicio_tarea()

        threading.Thread(
            target=self._proceso_transcripcion,
            args=(ruta_audio, tamano_modelo, idioma, incluir_tiempos),
            daemon=True,
        ).start()

    def _proceso_transcripcion(self, ruta_audio, tamano_modelo, idioma, incluir_tiempos):
        """Ejecuta la transcripción en segundo plano y envía el resultado al hilo principal."""
        try:
            texto_final, ruta_txt, advertencia = transcribir_audio(
                ruta_audio, tamano_modelo, idioma, incluir_tiempos
            )
            self._notificar_ui(
                lambda: self._finalizar_exito_txt(texto_final, ruta_txt, advertencia)
            )
        except Exception as exc:
            mensaje_error = str(exc)
            self._notificar_ui(lambda msg=mensaje_error: self._finalizar_error_txt(msg))
        finally:
            self._registrar_fin_tarea()

    def _finalizar_exito_txt(self, texto, ruta_txt, advertencia):
        """Actualiza la vista previa, reactiva el botón y muestra éxito o advertencia de silencio."""
        self.boton_transcribir.config(state=tk.NORMAL)
        self._insertar_vista_previa(self.caja_resultado, texto)
        if advertencia:
            self.etiqueta_estado_txt.config(text=advertencia, foreground=COLOR_PROGRESO)
            messagebox.showwarning(
                "Transcripción vacía",
                f"{advertencia}\nArchivo creado en:\n{ruta_txt}",
            )
        else:
            self.etiqueta_estado_txt.config(
                text=f"Guardado en: {os.path.basename(ruta_txt)}.",
                foreground=COLOR_EXITO,
            )
            messagebox.showinfo(
                "Transcripción completada", f"Archivo guardado en:\n{ruta_txt}"
            )

    def _finalizar_error_txt(self, mensaje_error):
        """Restaura el botón de transcripción y muestra el detalle del fallo."""
        self.boton_transcribir.config(state=tk.NORMAL)
        self.etiqueta_estado_txt.config(
            text="Error al transcribir el audio.", foreground=COLOR_ERROR
        )
        messagebox.showerror(
            "Error de transcripción", f"Falló la transcripción:\n{mensaje_error}"
        )

    # --- PESTAÑA 3: DOCUMENTOS A MARKDOWN + OCR ---
    def _construir_tab_documentos(self):
        """Construye los controles de la pestaña de extracción de documentos a Markdown con OCR."""
        tab_doc = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab_doc, text=" Documentos a Markdown ")

        ttk.Label(
            tab_doc,
            text="Documento (.pdf, .docx, .pptx):",
            font=self.fuente_negrita,
        ).pack(anchor="w")
        marco_doc = ttk.Frame(tab_doc)
        marco_doc.pack(fill="x", pady=(4, 12))
        self.entrada_archivo_doc = ttk.Entry(marco_doc)
        self.entrada_archivo_doc.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(
            marco_doc, text="Examinar...", command=self._seleccionar_documento
        ).pack(side="right")

        marco_opciones_doc = ttk.LabelFrame(
            tab_doc, text="Opciones de extracción a Markdown (.md) y OCR local", padding=10
        )
        marco_opciones_doc.pack(fill="x", pady=(0, 12))

        self.var_usar_ocr = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            marco_opciones_doc,
            text="Reconocimiento de texto en imágenes (OCR)",
            variable=self.var_usar_ocr,
        ).grid(row=0, column=0, padx=(0, 20), sticky="w")

        self.var_extraer_img = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            marco_opciones_doc,
            text="Guardar imágenes aparte",
            variable=self.var_extraer_img,
        ).grid(row=0, column=1, sticky="w")

        self.boton_extraer_doc = ttk.Button(
            tab_doc, text="Extraer contenido", command=self._iniciar_extraccion_doc
        )
        self.boton_extraer_doc.pack(fill="x", ipady=5)

        self.etiqueta_estado_doc = ttk.Label(
            tab_doc,
            text="Listo para extraer contenido de documentos.",
            font=self.fuente_estado,
        )
        self.etiqueta_estado_doc.pack(pady=(6, 6))

        ttk.Label(
            tab_doc, text="Vista previa del resultado generado:", font=self.fuente_negrita
        ).pack(anchor="w")
        self.caja_resultado_doc = self._crear_caja_texto_con_scroll(tab_doc, self.fuente_mono)

    def _seleccionar_documento(self):
        """Abre un selector de documentos compatibles y actualiza el campo de ruta."""
        tipos = [
            ("Documentos compatibles", "*.pdf *.docx *.pptx"),
            ("Archivos PDF", "*.pdf"),
            ("Documentos Word (.docx)", "*.docx"),
            ("Presentaciones PowerPoint (.pptx)", "*.pptx"),
        ]
        archivo = filedialog.askopenfilename(title="Seleccionar documento", filetypes=tipos)
        if archivo:
            self.entrada_archivo_doc.delete(0, tk.END)
            self.entrada_archivo_doc.insert(0, archivo)

    def _iniciar_extraccion_doc(self):
        """Valida existencia, extensión y permisos de escritura antes de lanzar la extracción."""
        ruta_doc = self.entrada_archivo_doc.get().strip()
        guardar_img = self.var_extraer_img.get()
        usar_ocr = self.var_usar_ocr.get()

        if not ruta_doc:
            messagebox.showwarning("Campo vacío", "Selecciona un documento.")
            return
        if not os.path.isfile(ruta_doc):
            messagebox.showwarning(
                "Archivo inválido",
                "Selecciona un archivo PDF, Word (.docx) o PowerPoint (.pptx) existente.",
            )
            return

        _, ext = os.path.splitext(ruta_doc)
        if ext.lower() not in EXTENSIONES_DOCUMENTO_SOPORTADAS:
            messagebox.showerror(
                "Formato no soportado",
                "Extensión no soportada. Usa .pdf, .docx o .pptx.",
            )
            return

        carpeta_salida = os.path.dirname(os.path.abspath(ruta_doc))
        if not verificar_permisos_escritura(carpeta_salida):
            messagebox.showerror(
                "Sin permisos de escritura",
                f"No hay permisos para guardar el resultado en:\n{carpeta_salida}",
            )
            return

        self.boton_extraer_doc.config(state=tk.DISABLED)
        self.caja_resultado_doc.delete("1.0", tk.END)
        msg_estado = (
            "Extrayendo contenido y ejecutando OCR en imágenes..."
            if usar_ocr
            else "Extrayendo contenido del documento..."
        )
        self.etiqueta_estado_doc.config(text=msg_estado, foreground=COLOR_PROGRESO)
        self._registrar_inicio_tarea()

        threading.Thread(
            target=self._proceso_extraccion_doc,
            args=(ruta_doc, guardar_img, usar_ocr),
            daemon=True,
        ).start()

    def _proceso_extraccion_doc(self, ruta_doc, guardar_img, usar_ocr):
        """Ejecuta la extracción documental en segundo plano y notifica el resultado a la interfaz."""
        try:
            contenido, ruta_salida, advertencias = extraer_documento(
                ruta_doc, guardar_img, usar_ocr
            )
            self._notificar_ui(
                lambda: self._finalizar_exito_doc(contenido, ruta_salida, advertencias)
            )
        except Exception as exc:
            mensaje_error = str(exc)
            self._notificar_ui(lambda msg=mensaje_error: self._finalizar_error_doc(msg))
        finally:
            self._registrar_fin_tarea()

    def _finalizar_exito_doc(self, contenido, ruta_salida, advertencias):
        """Muestra la vista previa, reactiva el botón y reporta éxito o advertencias de extracción."""
        self.boton_extraer_doc.config(state=tk.NORMAL)
        self._insertar_vista_previa(self.caja_resultado_doc, contenido)
        if advertencias:
            resumen_adv = "\n".join(advertencias[:5])
            if len(advertencias) > 5:
                resumen_adv += f"\n... y {len(advertencias) - 5} advertencias adicionales."
            self.etiqueta_estado_doc.config(
                text=f"Guardado con {len(advertencias)} advertencia(s): {os.path.basename(ruta_salida)}.",
                foreground=COLOR_PROGRESO,
            )
            messagebox.showwarning(
                "Extracción completada con advertencias",
                f"Archivo guardado en:\n{ruta_salida}\n\nAdvertencias:\n{resumen_adv}",
            )
        else:
            self.etiqueta_estado_doc.config(
                text=f"Guardado en: {os.path.basename(ruta_salida)}.",
                foreground=COLOR_EXITO,
            )
            messagebox.showinfo(
                "Extracción completada", f"Archivo guardado en:\n{ruta_salida}"
            )

    def _finalizar_error_doc(self, mensaje_error):
        """Restaura el botón de extracción y muestra el error al usuario."""
        self.boton_extraer_doc.config(state=tk.NORMAL)
        self.etiqueta_estado_doc.config(
            text="Error al procesar el documento.", foreground=COLOR_ERROR
        )
        messagebox.showerror(
            "Error de extracción", f"No se pudo extraer el documento:\n{mensaje_error}"
        )


def main():
    """Punto de entrada principal de la interfaz gráfica."""
    ventana = tk.Tk()
    AplicacionSuite(ventana)
    ventana.mainloop()


if __name__ == "__main__":
    main()
