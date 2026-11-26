"""Pruebas unitarias para la lógica de utilidades, extracción y transcripción de main.py."""

import io
import os
import sys
import tempfile
import unittest
from types import SimpleNamespace
from unittest import mock

import docx
from docx.opc.constants import RELATIONSHIP_TYPE
import pymupdf
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

import main


def _crear_png_bytes(color=(255, 0, 0), tamano=(20, 20)):
    """Genera una imagen PNG mínima en memoria para pruebas de extracción y OCR."""
    buf = io.BytesIO()
    Image.new("RGB", tamano, color=color).save(buf, format="PNG")
    return buf.getvalue()


class TestUtilidadesYValidacion(unittest.TestCase):
    """Pruebas de validación de tiempos, opciones de yt_dlp, rutas y sanitización."""

    def test_convertir_a_segundos_validos(self):
        """Comprueba la conversión correcta de cadenas vacías, SS, MM:SS y HH:MM:SS."""
        self.assertIsNone(main.convertir_a_segundos(""))
        self.assertIsNone(main.convertir_a_segundos("   "))
        self.assertEqual(main.convertir_a_segundos("0"), 0)
        self.assertEqual(main.convertir_a_segundos("45"), 45)
        self.assertEqual(main.convertir_a_segundos("02:30"), 150)
        self.assertEqual(main.convertir_a_segundos("01:02:03"), 3723)

    def test_convertir_a_segundos_invalidos(self):
        """Verifica que formatos negativos, fuera de rango o no numéricos lancen ValueError."""
        casos_invalidos = [
            "-10",
            "01:-30",
            "-01:30",
            "01:60",
            "01:90",
            "01:60:00",
            "01:00:60",
            "abc",
            "1:2:3:4",
            "12.5",
        ]
        for caso in casos_invalidos:
            with self.subTest(caso=caso):
                with self.assertRaises(ValueError):
                    main.convertir_a_segundos(caso)

    def test_formatear_segundo(self):
        """Verifica el formato MM:SS, HH:MM:SS y el rechazo de segundos negativos."""
        self.assertEqual(main.formatear_segundo(0), "00:00")
        self.assertEqual(main.formatear_segundo(59.9), "00:59")
        self.assertEqual(main.formatear_segundo(65), "01:05")
        self.assertEqual(main.formatear_segundo(3661), "01:01:01")
        with self.assertRaises(ValueError):
            main.formatear_segundo(-1)

    def test_construir_opciones_yt_dlp_sin_recorte_por_defecto(self):
        """Asegura que sin recorte no se inyecten download_ranges ni extractor_args fijos."""
        opts_vacio = main.construir_opciones_yt_dlp(None, None, "192", "C:/tmp")
        self.assertNotIn("download_ranges", opts_vacio)
        self.assertNotIn("force_keyframes_at_cuts", opts_vacio)
        self.assertNotIn("extractor_args", opts_vacio)

        opts_cero = main.construir_opciones_yt_dlp(0, None, "192", "C:/tmp")
        self.assertNotIn("download_ranges", opts_cero)
        self.assertNotIn("force_keyframes_at_cuts", opts_cero)

    def test_construir_opciones_yt_dlp_con_recorte(self):
        """Asegura que al indicar inicio mayor a 0 o fin se active download_ranges."""
        opts_inicio = main.construir_opciones_yt_dlp(10, None, "192", "C:/tmp")
        self.assertIn("download_ranges", opts_inicio)
        self.assertTrue(opts_inicio["force_keyframes_at_cuts"])

        opts_fin = main.construir_opciones_yt_dlp(None, 30, "192", "C:/tmp")
        self.assertIn("download_ranges", opts_fin)
        self.assertTrue(opts_fin["force_keyframes_at_cuts"])

    def test_ruta_descargas_y_permisos_escritura(self):
        """Verifica la ruta portable de descargas y la comprobación de permisos de escritura."""
        ruta_predet = main.obtener_carpeta_descargas_predeterminada()
        self.assertEqual(ruta_predet, os.path.join(os.path.expanduser("~"), "Downloads"))

        with tempfile.TemporaryDirectory() as tmpdir:
            self.assertTrue(main.verificar_permisos_escritura(tmpdir))
            self.assertFalse(main.verificar_permisos_escritura(os.path.join(tmpdir, "no_existe")))
            self.assertFalse(main.verificar_permisos_escritura(""))

    def test_sanitizar_nombre_archivo(self):
        """Comprueba la eliminación de rutas relativas y caracteres inválidos en nombres."""
        self.assertEqual(main.sanitizar_nombre_archivo("../../malo:nombre*.png"), "malo_nombre_.png")
        self.assertEqual(main.sanitizar_nombre_archivo(""), "imagen")


class TestMarkdownYOCR(unittest.TestCase):
    """Pruebas de formateo de tablas Markdown, ejecución OCR y manejo separado de errores."""

    def test_convertir_tabla_a_md(self):
        """Verifica escape de barras verticales, celdas nulas y relleno de filas irregulares."""
        self.assertEqual(main.convertir_tabla_a_md([]), "")
        self.assertEqual(main.convertir_tabla_a_md([[]]), "")

        filas = [
            ["Col | 1", "Col\r\n2"],
            [None, "A|B", "Extra"],
        ]
        md = main.convertir_tabla_a_md(filas)
        lineas = md.strip().splitlines()
        self.assertEqual(lineas[0], r"| Col \| 1 | Col 2 |  |")
        self.assertEqual(lineas[1], "| --- | --- | --- |")
        self.assertEqual(lineas[2], r"|  | A\|B | Extra |")

    def test_aplicar_ocr_bytes_exito_y_error(self):
        """Verifica que aplicar_ocr_bytes devuelva (texto, None) en éxito y ('', error) en fallo."""
        ocr_ok = mock.MagicMock(return_value=([[[], "Linea 1", 0.99], [[], "Linea 2", 0.98]], None))
        with mock.patch.object(main, "obtener_motor_ocr", return_value=ocr_ok):
            texto, err = main.aplicar_ocr_bytes(b"img")
        self.assertEqual(texto, "Linea 1\nLinea 2")
        self.assertIsNone(err)

        ocr_fail = mock.MagicMock(side_effect=RuntimeError("fallo onnx"))
        with mock.patch.object(main, "obtener_motor_ocr", return_value=ocr_fail):
            texto, err = main.aplicar_ocr_bytes(b"img")
        self.assertEqual(texto, "")
        self.assertIn("fallo onnx", err)

    def test_procesar_imagen_extraida_cita_multilinea_y_error_separado(self):
        """Comprueba el prefijo '> ' multilínea, guardado diferido y separación de errores OCR."""
        with tempfile.TemporaryDirectory() as tmpdir:
            carpeta_img = os.path.join(tmpdir, "imgs")
            advertencias = []

            with mock.patch.object(main, "aplicar_ocr_bytes", return_value=("L1\nL2", None)):
                info, nombre, _, _ = main.procesar_imagen_extraida(
                    img_bytes=b"pngdata",
                    encabezado_md="> **[Imagen #1]**",
                    nombre_archivo="foto.png",
                    guardar_img=True,
                    carpeta_img=carpeta_img,
                    usar_ocr=True,
                    advertencias=advertencias,
                )
            self.assertEqual(nombre, "foto.png")
            self.assertTrue(os.path.isfile(os.path.join(carpeta_img, "foto.png")))
            self.assertIn("> **Texto extraído por OCR:**\n> L1\n> L2", info)
            self.assertEqual(advertencias, [])

            with mock.patch.object(main, "aplicar_ocr_bytes", return_value=("", "imagen corrupta")):
                info_err, _, _, _ = main.procesar_imagen_extraida(
                    img_bytes=b"bad",
                    encabezado_md="> **[Imagen #2]**",
                    nombre_archivo="foto2.png",
                    guardar_img=False,
                    carpeta_img=carpeta_img,
                    usar_ocr=True,
                    advertencias=advertencias,
                )
            self.assertNotIn("Texto extraído por OCR", info_err)
            self.assertIn("OCR no disponible", info_err)
            self.assertEqual(len(advertencias), 1)
            self.assertIn("imagen corrupta", advertencias[0])


class TestExtractoresDocumentos(unittest.TestCase):
    """Pruebas de extracción para PDF, DOCX, PPTX y validación de salida documental."""

    def test_extraer_pdf_deduplica_xref_cierra_descriptor_y_detecta_cifrado(self):
        """Verifica deduplicación de xref, cierre limpio del PDF y rechazo de archivos cifrados."""
        png = _crear_png_bytes()
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta_pdf = os.path.join(tmpdir, "prueba.pdf")
            doc = pymupdf.open()
            p1 = doc.new_page()
            p1.insert_text((72, 72), "Hola PDF pagina 1")
            xref = p1.insert_image(pymupdf.Rect(72, 100, 140, 168), stream=png)
            p2 = doc.new_page()
            p2.insert_text((72, 72), "Hola PDF pagina 2")
            p2.insert_image(pymupdf.Rect(72, 100, 140, 168), xref=xref)
            doc.save(ruta_pdf)
            doc.close()

            carpeta_img = os.path.join(tmpdir, "prueba_imagenes")
            advertencias = []
            with mock.patch.object(main, "aplicar_ocr_bytes", return_value=("Texto OCR", None)) as mock_ocr:
                md = main.extraer_pdf(ruta_pdf, True, carpeta_img, True, advertencias)

            self.assertIn("Hola PDF pagina 1", md)
            self.assertIn("Hola PDF pagina 2", md)
            self.assertEqual(mock_ocr.call_count, 1)
            self.assertEqual(len(os.listdir(carpeta_img)), 1)
            # Verifica que el descriptor fue cerrado (en Windows fallaría si sigue abierto)
            os.remove(ruta_pdf)

            ruta_cifrado = os.path.join(tmpdir, "cifrado.pdf")
            doc_enc = pymupdf.open()
            doc_enc.new_page().insert_text((72, 72), "Secreto")
            doc_enc.save(ruta_cifrado, encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="1234", owner_pw="5678")
            doc_enc.close()
            with self.assertRaises(ValueError):
                main.extraer_pdf(ruta_cifrado, False, carpeta_img, False)

    def test_extraer_docx_con_enlace_externo_image_y_estilo_nulo(self):
        """Comprueba tolerancia ante enlaces externos con 'image', párrafos sin estilo e imágenes."""
        png = _crear_png_bytes()
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta_docx = os.path.join(tmpdir, "prueba.docx")
            d = docx.Document()
            d.add_heading("Sección Uno", level=1)
            p_sin_estilo = d.add_paragraph("Párrafo normal")
            p_sin_estilo.style = None

            # Agregar relación externa cuya URL contiene la palabra "image"
            d.part.relate_to(
                "https://example.com/image/externo.png",
                RELATIONSHIP_TYPE.HYPERLINK,
                is_external=True,
            )

            tabla = d.add_table(rows=2, cols=2)
            tabla.cell(0, 0).text = "Clave"
            tabla.cell(0, 1).text = "Valor"
            tabla.cell(1, 0).text = "A|B"
            tabla.cell(1, 1).text = "123"

            d.add_picture(io.BytesIO(png), width=docx.shared.Inches(1))
            d.save(ruta_docx)

            carpeta_img = os.path.join(tmpdir, "docx_imgs")
            advertencias = []
            with mock.patch.object(main, "aplicar_ocr_bytes", return_value=("OCR Word", None)):
                md = main.extraer_docx(ruta_docx, True, carpeta_img, True, advertencias)

            self.assertIn("## Sección Uno", md)
            self.assertIn("Párrafo normal", md)
            self.assertIn(r"| A\|B | 123 |", md)
            self.assertIn("> OCR Word", md)
            self.assertEqual(len(os.listdir(carpeta_img)), 1)
            self.assertEqual(advertencias, [])

    def test_extraer_pptx_con_imagen_placeholder_y_notas(self):
        """Verifica extracción de texto, notas, imágenes directas y marcadores de posición en PPTX."""
        png = _crear_png_bytes()
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta_pptx = os.path.join(tmpdir, "prueba.pptx")
            prs = Presentation()
            slide_blank = prs.slides.add_slide(prs.slide_layouts[6])
            tx = slide_blank.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            tx.text_frame.text = "Texto de diapositiva"
            slide_blank.shapes.add_picture(io.BytesIO(png), Inches(1), Inches(2), width=Inches(1))
            slide_blank.notes_slide.notes_text_frame.text = "Nota del presentador"

            for layout in prs.slide_layouts:
                pic_phs = [ph for ph in layout.placeholders if ph.placeholder_format.type == 18]
                if pic_phs:
                    slide_ph = prs.slides.add_slide(layout)
                    for ph in slide_ph.placeholders:
                        if ph.placeholder_format.type == 18:
                            ph.insert_picture(io.BytesIO(png))
                            break
                    break

            prs.save(ruta_pptx)

            carpeta_img = os.path.join(tmpdir, "pptx_imgs")
            advertencias = []
            with mock.patch.object(main, "aplicar_ocr_bytes", return_value=("OCR PPT", None)):
                md = main.extraer_pptx(ruta_pptx, True, carpeta_img, True, advertencias)

            self.assertIn("- Texto de diapositiva", md)
            self.assertIn("**Notas del orador:** Nota del presentador", md)
            self.assertIn("> OCR PPT", md)
            self.assertGreaterEqual(len(os.listdir(carpeta_img)), 1)

    def test_extraer_documento_no_crea_carpeta_vacia_ni_acepta_extension_invalida(self):
        """Asegura rechazo de extensiones no soportadas y no creación de carpetas vacías."""
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta_invalida = os.path.join(tmpdir, "archivo.txt")
            with open(ruta_invalida, "w", encoding="utf-8") as f:
                f.write("hola")
            with self.assertRaises(ValueError):
                main.extraer_documento(ruta_invalida, True, False)
            self.assertFalse(os.path.exists(os.path.join(tmpdir, "archivo_imagenes")))

            ruta_pdf_sin_img = os.path.join(tmpdir, "solo_texto.pdf")
            doc = pymupdf.open()
            doc.new_page().insert_text((72, 72), "Solo texto")
            doc.save(ruta_pdf_sin_img)
            doc.close()

            contenido, ruta_salida, advertencias = main.extraer_documento(
                ruta_pdf_sin_img, guardar_img=True, usar_ocr=False
            )
            self.assertTrue(ruta_salida.endswith("_extraido.md"))
            self.assertTrue(os.path.isfile(ruta_salida))
            self.assertIn("Solo texto", contenido)
            self.assertFalse(os.path.exists(os.path.join(tmpdir, "solo_texto_imagenes")))
            self.assertEqual(advertencias, [])


class TestTranscripcionYCacheWhisper(unittest.TestCase):
    """Pruebas de gestión de memoria de caché Whisper y transcripción de audio."""

    def tearDown(self):
        """Limpia la caché global de Whisper al terminar cada prueba."""
        main.modelo_whisper_cache["nombre"] = None
        main.modelo_whisper_cache["instancia"] = None

    def test_obtener_modelo_whisper_libera_anterior_antes_de_cargar_nuevo(self):
        """Comprueba que cambiar de modelo libera primero la instancia previa en memoria."""
        estados_en_creacion = []

        def fake_whisper(nombre, device="cpu", compute_type="int8"):
            """Simula la instanciación de WhisperModel registrando el estado previo de caché."""
            estados_en_creacion.append((nombre, main.modelo_whisper_cache["instancia"]))
            return SimpleNamespace(nombre=nombre)

        fake_mod = SimpleNamespace(WhisperModel=fake_whisper)
        with mock.patch.dict(sys.modules, {"faster_whisper": fake_mod}):
            m1 = main.obtener_modelo_whisper("small")
            m2 = main.obtener_modelo_whisper("medium")

        self.assertEqual(m1.nombre, "small")
        self.assertEqual(m2.nombre, "medium")
        self.assertEqual(estados_en_creacion, [("small", None), ("medium", None)])

    def test_transcribir_audio_filtra_vacios_y_advierte_silencio(self):
        """Verifica omisión de segmentos vacíos y advertencia explícita cuando no hay voz."""
        with tempfile.TemporaryDirectory() as tmpdir:
            ruta_audio = os.path.join(tmpdir, "nota.opus")
            with open(ruta_audio, "wb") as f:
                f.write(b"fake")

            seg_ok = SimpleNamespace(start=0.0, end=2.5, text="  Hola mundo  ")
            seg_vacio = SimpleNamespace(start=2.5, end=4.0, text="   ")
            modelo_ok = mock.MagicMock()
            modelo_ok.transcribe.return_value = ([seg_ok, seg_vacio], None)

            with mock.patch.object(main, "obtener_modelo_whisper", return_value=modelo_ok):
                texto, ruta_txt, adv = main.transcribir_audio(ruta_audio, "small", "es", True)
            self.assertEqual(texto, "[00:00 -> 00:02] Hola mundo")
            self.assertTrue(os.path.isfile(ruta_txt))
            self.assertIsNone(adv)

            modelo_silencio = mock.MagicMock()
            modelo_silencio.transcribe.return_value = ([seg_vacio], None)
            with mock.patch.object(main, "obtener_modelo_whisper", return_value=modelo_silencio):
                texto_vacio, _, adv_silencio = main.transcribir_audio(
                    ruta_audio, "small", "es", False
                )
            self.assertEqual(texto_vacio, "")
            self.assertIsNotNone(adv_silencio)


if __name__ == "__main__":
    unittest.main()
