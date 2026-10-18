"""
Conversor de documentos (PDF, Word, PowerPoint) -> Markdown usando docling,
con aceleración GPU/CPU.

Resistencia añadida frente al script original:
  - Backend forzado a pypdfium2 para PDFs (en vez del backend nativo
    "docling-parse"): se diagnosticó experimentalmente que docling-parse
    provoca un crash nativo (STATUS_HEAP_CORRUPTION / 0xC0000374, y en
    otras variantes STATUS_STACK_BUFFER_OVERRUN / 0xC0000409) al procesar
    ciertas imágenes CMYK de alta resolución con perfil ICC embebido --
    reproducido de forma aislada y reproducible con una portada con imagen
    CMYK 2186x2900 y perfil "Coated FOGRA39", tanto con OCR+GPU como sin
    OCR forzando CPU. pypdfium2 (usado también internamente por docling
    para el renderizado) procesa la misma imagen sin problema, así que se
    fuerza como backend.
    NOTA: la hipótesis inicial apuntaba a imágenes con canal alfa/transparencia
    (RGBA/LA o /SMask); se descartó por diagnóstico -- las imágenes del PDF
    de prueba eran CMYK planas sin alfa (0 con transparencia detectada) y
    el crash persistía. La función de saneado de alfa se conserva como
    salvaguarda adicional, pero el fix que resuelve el crash confirmado
    es el cambio de backend.
  - Saneado automático de imágenes con canal alfa (RGBA/LA/paleta con
    transparencia) en PDFs, como salvaguarda adicional: se aplanan a RGB
    antes de procesar, por si algún PDF distinto sí dispara el problema
    por esa vía.
  - Aislamiento por PROCESO APARTE a nivel de DOCUMENTO COMPLETO (no por
    trozo): todos los trozos de un mismo documento se procesan de forma
    secuencial dentro de UN ÚNICO proceso hijo, que inicializa el
    converter (y por tanto el contexto CUDA) UNA SOLA VEZ para todo el
    documento. Si algo truena a nivel nativo (C++/CUDA), solo muere ese
    proceso hijo -- el resto del lote sigue. Un try/except de Python
    NUNCA puede atrapar ese tipo de crash; aislarlo en un proceso es la
    única forma real de sobrevivirlo.
    NOTA (diagnóstico de rendimiento): la versión anterior aislaba cada
    trozo en su PROPIO proceso, lo que forzaba reinicializar el contexto
    CUDA una vez por trozo. Bajo el driver WDDM de Windows esa
    reinicialización repetida degrada progresivamente el arranque de
    cada proceso hijo (confirmado: 10.4s en un solo proceso vs. >360s
    troceando el mismo documento). Mover el aislamiento a nivel de
    documento completo elimina ese ciclo sin perder la protección del
    lote frente a un documento problemático. Como mitigación adicional,
    el resultado de cada trozo se envía al proceso principal en cuanto
    está listo (no acumulado al final): si el proceso hijo truena a
    mitad de documento, los trozos ya completados antes del crash
    normalmente se conservan.
  - Timeout TOTAL POR DOCUMENTO (ya no por trozo): se calcula como
    timeout_por_trozo × número de trozos del documento, y se aplica una
    sola vez al proceso hijo completo.
  - Umbral de partición elevado y DESACOPLADO del tamaño de trozo (antes
    era tamaño_trozo + 3, ~15 páginas): por defecto, documentos de hasta
    60 páginas NUNCA se parten en trozos. Se diagnosticó que el troceo en
    sí -- vía el aislamiento por proceso a nivel de trozo -- era la causa
    de la degradación de rendimiento; con el aislamiento ahora a nivel de
    documento, partir de más ya no aporta protección extra y solo tiene
    sentido para documentos genuinamente grandes (para limitar memoria
    acumulada del backend en C++ dentro de una sola llamada a convert()).
  - Validación rápida de cada documento (corrupto / con contraseña /
    vacío / formato no soportado) antes de intentar procesarlo.
  - Verificación de que el ejecutable de Tesseract exista antes de
    activar OCR, para no fallar a mitad de proceso con un error críptico.
  - Manejo de dependencias faltantes con mensajes claros en vez de un
    traceback crudo al arrancar.
  - Detección de tablas (table structure) OPCIONAL: se pregunta al usuario,
    ya que es una de las partes más lentas del pipeline y no siempre hace
    falta.
  - Tamaño de trozo (páginas por trozo, para PDFs largos) CONFIGURABLE por
    el usuario en vez de un valor fijo.
  - Soporte para Word (.docx) y PowerPoint (.pptx) además de PDF. Los PPTX
    con diagramas SmartArt o tablas se leen directamente del XML del PowerPoint
    para conservar el texto y la estructura que docling omite; los demás
    siguen usando docling. OCR y detección visual de tablas solo afectan a PDF.
    Solo se admiten .docx/.pptx modernos (Office Open XML), no .doc/.ppt.

  - Corrección automática de orientación de páginas en PDFs (ver
    corregir_orientacion_paginas): detecta páginas cuyo contenido está
    girado 90/180/270° (típico de presentaciones exportadas "de lado") a
    partir de la dirección real del texto y les aplica la rotación
    inversa en una copia temporal, ANTES de sanear/partir/convertir. Solo
    toca páginas con capa de texto y orientación claramente anómala; si no
    hay nada que corregir (o falla algo), se usa el PDF original sin
    cambios. Se puede desactivar con CORREGIR_ORIENTACION = False.

Interfaces:
    python docling_pdf2md.py            -> modo consola (igual que antes)
    python docling_pdf2md.py --gui      -> interfaz gráfica simple (Tkinter)

Ambas usan exactamente la misma lógica de conversión (procesar_lote).
"""

import gc
import io
import math
import multiprocessing as mp
import shutil
import sys
import tempfile
import time
import traceback
from datetime import datetime
from pathlib import Path
from pptx_smartart_a_markdown import convert_file as extraer_smartart_pptx

# ---------------------------------------------------------------------------
# Dependencias principales. Si fallan, no queremos un traceback crudo al
# arrancar -- se reporta de forma clara en main()/lanzar_gui().
# ---------------------------------------------------------------------------
DEPENDENCIAS_CORE_OK = False
_ERROR_DEPENDENCIAS = None
PYPDF_DISPONIBLE = False
SANEADO_DISPONIBLE = False
_DEPENDENCIAS_CARGADAS = False


def cargar_dependencias():
    """Carga Docling y sus dependencias bajo demanda.

    En modo GUI se llama desde el hilo de conversión, de modo que la ventana
    aparece enseguida aunque PyTorch/Docling tarden en inicializarse.
    """
    global torch, DocumentConverter, PdfFormatOption, PdfPipelineOptions
    global TesseractCliOcrOptions, AcceleratorOptions, AcceleratorDevice
    global InputFormat, PyPdfiumDocumentBackend, PdfReader, PdfWriter
    global pymupdf, Image, DEPENDENCIAS_CORE_OK, _ERROR_DEPENDENCIAS
    global PYPDF_DISPONIBLE, SANEADO_DISPONIBLE, _DEPENDENCIAS_CARGADAS

    if _DEPENDENCIAS_CARGADAS:
        return DEPENDENCIAS_CORE_OK
    _DEPENDENCIAS_CARGADAS = True

    try:
        import torch
        from docling.document_converter import DocumentConverter, PdfFormatOption
        from docling.datamodel.pipeline_options import (
            PdfPipelineOptions,
            TesseractCliOcrOptions,
            AcceleratorOptions,
            AcceleratorDevice,
        )
        from docling.datamodel.base_models import InputFormat
        from docling.backend.pypdfium2_backend import PyPdfiumDocumentBackend
        DEPENDENCIAS_CORE_OK = True
    except Exception as e:  # noqa: BLE001 - mostrar fallo claro en main/GUI
        _ERROR_DEPENDENCIAS = e

    try:
        from pypdf import PdfReader, PdfWriter
        PYPDF_DISPONIBLE = True
    except ImportError:
        pass

    try:
        import pymupdf
        from PIL import Image
        SANEADO_DISPONIBLE = True
    except ImportError:
        pass

    return DEPENDENCIAS_CORE_OK


# Formatos de documento soportados. Solo formatos modernos basados en XML
# (Office Open XML); los binarios antiguos .doc/.ppt no son compatibles
# con docling.
EXTENSIONES_SOPORTADAS = (".pdf", ".docx", ".pptx")

# Si es True, los PDFs se revisan antes de procesar y las páginas con el
# contenido girado se enderezan automáticamente (ver
# corregir_orientacion_paginas). Ponlo en False para desactivarlo.
CORREGIR_ORIENTACION = True

# Tamaño de trozo (en páginas) para PDFs que SÍ superan el umbral de
# partición (ver UMBRAL_PARTICION_POR_DEFECTO más abajo). Los trozos se
# parten físicamente antes de pasar por docling, para evitar la
# acumulación de memoria del backend en C++ dentro de una sola llamada a
# convert(). 12-15 es un buen punto de partida para 4GB de VRAM.
# NOTA: ya NO determina el umbral de partición (antes era tamaño_trozo + 3);
# ese umbral ahora es independiente y bastante más alto -- ver abajo.
TAMANO_TROZO_POR_DEFECTO = 12

# Umbral de partición por defecto (en páginas). Documentos con este número
# de páginas o menos NUNCA se parten en trozos, sin importar el tamaño de
# trozo configurado -- se procesan enteros en una sola llamada a convert().
# Se elevó deliberadamente (antes era tamaño_trozo + 3, ~15 páginas): se
# diagnosticó que el troceo en sí -- vía el aislamiento por proceso a nivel
# de trozo que existía antes -- era la causa de una degradación severa de
# rendimiento en Windows (reinicialización repetida del contexto CUDA, ver
# convertir_documento_aislado). Con el aislamiento ahora a nivel de
# documento completo, partir de más ya no aporta protección adicional --
# solo tiene sentido para documentos genuinamente grandes, donde sí
# conviene limitar la memoria acumulada del backend en C++.
UMBRAL_PARTICION_POR_DEFECTO = 60

# Tiempo (segundos) presupuestado POR TROZO para calcular el timeout TOTAL
# de un documento: timeout_total = este valor × número de trozos del
# documento (mínimo 1 trozo, incluso si el documento no se parte). Con el
# aislamiento por proceso ahora a nivel de documento (no de trozo), el
# timeout ya no se aplica trozo a trozo -- se aplica una sola vez al
# proceso hijo completo; este valor es solo la unidad usada para
# presupuestar ese total. Ajustar si tus documentos son muy pesados.
TIMEOUT_POR_TROZO_POR_DEFECTO = 120  # 2 minutos

RUTA_TESSERACT_POR_DEFECTO = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


def detectar_acelerador(log=print):
    """Detecta automáticamente si CUDA está disponible"""
    if not DEPENDENCIAS_CORE_OK:
        return None
    if torch.cuda.is_available():
        gpu_name = torch.cuda.get_device_name(0)
        gpu_memory = torch.cuda.get_device_properties(0).total_memory / (1024**3)
        log(f"\n✓ CUDA disponible: {gpu_name} ({gpu_memory:.1f} GB)")
        return AcceleratorDevice.CUDA
    else:
        log("\n⚠ CUDA no disponible, usando CPU")
        return AcceleratorDevice.CPU


def pedir_ruta(mensaje):
    return Path(input(mensaje).strip().strip('"'))


def pedir_ocr():
    print("\n¿Aplicar OCR?")
    print("  [1] Sí  — para PDFs escaneados o con imágenes")
    print("  [2] No  — para PDFs con texto nativo")
    while True:
        opcion = input("Elige 1 o 2: ").strip()
        if opcion in ("1", "2"):
            return opcion == "1"
        print("  Opción no válida, escribe 1 o 2.")


def pedir_tablas():
    print("\n¿Detectar estructura de tablas?")
    print("  [1] Sí — más preciso con tablas, pero más lento (corre en todas las páginas)")
    print("  [2] No — más rápido, recomendado si el documento no tiene tablas importantes")
    while True:
        opcion = input("Elige 1 o 2: ").strip()
        if opcion in ("1", "2"):
            return opcion == "1"
        print("  Opción no válida, escribe 1 o 2.")


def pedir_tamano_trozo():
    print(f"\n¿Tamaño de trozo para documentos largos? (páginas por trozo)")
    entrada = input(f"Enter para usar el valor por defecto ({TAMANO_TROZO_POR_DEFECTO}): ").strip()
    if not entrada:
        return TAMANO_TROZO_POR_DEFECTO
    try:
        valor = int(entrada)
        if valor > 0:
            return valor
    except ValueError:
        pass
    print(f"  Valor no válido, se usa el valor por defecto ({TAMANO_TROZO_POR_DEFECTO}).")
    return TAMANO_TROZO_POR_DEFECTO


def pedir_umbral_particion(tamano_trozo: int):
    print(f"\n¿Umbral de partición? (páginas; documentos con este número de páginas")
    print(f"  o menos NUNCA se parten en trozos, sin importar el tamaño de trozo)")
    entrada = input(f"Enter para usar el valor por defecto ({UMBRAL_PARTICION_POR_DEFECTO}): ").strip()
    if not entrada:
        valor = UMBRAL_PARTICION_POR_DEFECTO
    else:
        try:
            valor = int(entrada)
            if valor <= 0:
                raise ValueError
        except ValueError:
            print(f"  Valor no válido, se usa el valor por defecto ({UMBRAL_PARTICION_POR_DEFECTO}).")
            valor = UMBRAL_PARTICION_POR_DEFECTO
    # El umbral nunca debe ser menor que el tamaño de trozo: si lo fuera,
    # cualquier documento que lo supere generaría trozos casi tan grandes
    # como él mismo, partiendo sin necesidad real.
    return max(valor, tamano_trozo)


def pedir_timeout():
    print(f"\n¿Segundos presupuestados por trozo? (el timeout total del documento")
    print(f"  se calcula como este valor × número de trozos que tenga)")
    entrada = input(f"Enter para usar el valor por defecto ({TIMEOUT_POR_TROZO_POR_DEFECTO}): ").strip()
    if not entrada:
        return TIMEOUT_POR_TROZO_POR_DEFECTO
    try:
        valor = int(entrada)
        if valor > 0:
            return valor
    except ValueError:
        pass
    print(f"  Valor no válido, se usa el valor por defecto ({TIMEOUT_POR_TROZO_POR_DEFECTO}).")
    return TIMEOUT_POR_TROZO_POR_DEFECTO


def verificar_tesseract(log=print):
    """Comprueba que el ejecutable de Tesseract exista antes de activar OCR,
    para avisar con claridad en vez de fallar a mitad del procesamiento con
    un error críptico del backend de OCR. Devuelve la ruta a usar (puede ser
    la del sistema si la de por defecto no existe)."""
    if Path(RUTA_TESSERACT_POR_DEFECTO).exists():
        return RUTA_TESSERACT_POR_DEFECTO

    alternativa = shutil.which("tesseract")
    if alternativa:
        log(f"⚠ No se encontró Tesseract en la ruta por defecto, usando: {alternativa}")
        return alternativa

    log(f"⚠ Tesseract no encontrado ni en '{RUTA_TESSERACT_POR_DEFECTO}' ni en el PATH del sistema.")
    log("  El OCR probablemente fallará. Instálalo desde https://github.com/UB-Mannheim/tesseract/wiki")
    return RUTA_TESSERACT_POR_DEFECTO


def build_pipeline(usar_ocr: bool, usar_tablas: bool, device, log=print) -> "PdfPipelineOptions":
    """Construye las opciones del pipeline de PDF. Nota: estas opciones solo
    afectan al formato PDF -- Word y PowerPoint se procesan con el backend
    nativo de docling para esos formatos (parseo de XML), que no usa OCR ni
    detección de tablas por visión."""
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_table_structure = usar_tablas
    if usar_tablas:
        pipeline_options.table_structure_options.do_cell_matching = True

    num_threads = 10 if device == AcceleratorDevice.CUDA else 6

    pipeline_options.accelerator_options = AcceleratorOptions(
        device=device,
        num_threads=num_threads,
    )

    if usar_ocr:
        ruta_tesseract = verificar_tesseract(log=log)
        pipeline_options.do_ocr = True
        pipeline_options.ocr_options = TesseractCliOcrOptions(
            lang=["spa", "eng"],
            tesseract_cmd=ruta_tesseract,
        )
    else:
        pipeline_options.do_ocr = False

    return pipeline_options


def crear_converter(pipeline_options) -> "DocumentConverter":
    # Se fuerza el backend pypdfium2 en vez del backend nativo de docling
    # (docling-parse) SOLO para PDF. Se confirmó experimentalmente que
    # docling-parse provoca un crash nativo (STATUS_HEAP_CORRUPTION /
    # 0xC0000374) al procesar ciertas imágenes CMYK de alta resolución con
    # perfil ICC embebido (reproducido con una portada con imagen CMYK
    # 2186x2900, perfil "Coated FOGRA39"). pypdfium2 procesa la misma
    # imagen sin problema, así que se usa como backend por defecto.
    #
    # Word (.docx) y PowerPoint (.pptx) no se configuran aquí: docling usa
    # sus backends nativos por defecto para esos formatos (no involucran
    # docling-parse ni pypdfium2, así que no aplica el mismo bug).
    return DocumentConverter(
        format_options={
            InputFormat.PDF: PdfFormatOption(
                pipeline_options=pipeline_options,
                backend=PyPdfiumDocumentBackend,
            )
        }
    )


def liberar_memoria(device):
    gc.collect()
    if DEPENDENCIAS_CORE_OK and device == AcceleratorDevice.CUDA:
        torch.cuda.empty_cache()
        torch.cuda.synchronize()


def buscar_documentos(ruta: Path):
    """Busca documentos soportados (.pdf, .docx, .pptx) en una ruta.
    Si `ruta` es un archivo, se devuelve tal cual (la validación de formato
    ocurre después, en validar_documento). Si es una carpeta, se buscan
    recursivamente todos los formatos soportados."""
    if ruta.is_file():
        return [ruta]
    if ruta.is_dir():
        archivos = []
        for ext in EXTENSIONES_SOPORTADAS:
            archivos.extend(ruta.rglob(f"*{ext}"))
        return sorted(set(archivos))
    return []


def validar_documento(path: Path):
    """Verificación rápida antes de procesar: detecta formato no soportado,
    archivos vacíos/inexistentes y, para PDFs, además corrupción o
    protección con contraseña. Devuelve (ok: bool, motivo: str)."""
    if path.suffix.lower() not in EXTENSIONES_SOPORTADAS:
        return False, f"Formato no soportado ({path.suffix or 'sin extensión'}); usa .pdf, .docx o .pptx"

    if not path.exists() or path.stat().st_size == 0:
        return False, "Archivo vacío o inexistente"

    if path.suffix.lower() == ".pdf" and PYPDF_DISPONIBLE:
        try:
            reader = PdfReader(str(path))
            if reader.is_encrypted:
                try:
                    reader.decrypt("")
                except Exception:
                    return False, "PDF protegido con contraseña"
            if len(reader.pages) == 0:
                return False, "PDF sin páginas"
        except Exception as e:
            return False, f"PDF corrupto o ilegible: {str(e)[:120]}"

    return True, ""


def corregir_orientacion_paginas(pdf_path: Path, tmp_dir: Path, log=print):
    """Endereza páginas de un PDF cuyo contenido está girado (90/180/270°).

    Método: para cada página con capa de texto se mide la dirección de las
    líneas de texto (ponderada por nº de caracteres). pymupdf entrega esa
    dirección SIN aplicar el atributo /Rotate de la página, así que la
    rotación efectiva visible es (ángulo del texto - /Rotate). Si no es 0,
    se fija /Rotate de modo que el texto quede horizontal y legible.

    Criterios conservadores para no estropear PDFs correctos:
      - páginas sin texto suficiente (<20 caracteres) o sin una dirección
        claramente dominante (>=60 %) se dejan tal cual;
      - solo se corrige en múltiplos de 90°;
      - si no hay ninguna página que corregir, devuelve la ruta original
        (no se crea copia); ante cualquier error, también la original.

    Devuelve la ruta al PDF corregido (copia en tmp_dir) o la original."""
    if not CORREGIR_ORIENTACION or not SANEADO_DISPONIBLE:
        return pdf_path

    try:
        doc = pymupdf.open(str(pdf_path))
        corregidas = []

        for num_pagina, page in enumerate(doc, start=1):
            pesos = {0: 0, 90: 0, 180: 0, 270: 0}
            for bloque in page.get_text("dict").get("blocks", []):
                for linea in bloque.get("lines", []):
                    n = sum(len(sp.get("text", "").strip()) for sp in linea.get("spans", []))
                    if n == 0:
                        continue
                    dx, dy = linea["dir"]
                    # Ángulo antihorario visual del texto (eje y hacia abajo)
                    ang = round(math.degrees(math.atan2(-dy, dx)) / 90) * 90 % 360
                    pesos[ang] += n

            total = sum(pesos.values())
            if total < 20:
                continue
            ang_dominante = max(pesos, key=pesos.get)
            if pesos[ang_dominante] / total < 0.6:
                continue

            # Corrección horaria necesaria = (ángulo del texto - /Rotate)
            correccion = (ang_dominante - page.rotation) % 360
            if correccion != 0:
                page.set_rotation((page.rotation + correccion) % 360)
                corregidas.append((num_pagina, correccion))

        if not corregidas:
            doc.close()
            return pdf_path

        orientado_path = tmp_dir / f"{pdf_path.stem}__orientado.pdf"
        doc.save(str(orientado_path), garbage=3, deflate=True)
        doc.close()

        resumen = ", ".join(f"p.{n} (+{g}°)" for n, g in corregidas[:8])
        if len(corregidas) > 8:
            resumen += f", ... ({len(corregidas)} en total)"
        log(f"  (orientación corregida en {len(corregidas)} página(s): {resumen})")
        return orientado_path

    except Exception:
        log(f"  ⚠ No se pudo revisar/corregir la orientación de {pdf_path.name}, se usa el original")
        return pdf_path


def sanear_imagenes_alfa(pdf_path: Path, tmp_dir: Path, log=print):
    """Detecta imágenes con canal alfa real (incluyendo las que usan una
    máscara de transparencia /SMask como objeto PDF separado, no solo las
    que traen el alfa embebido en los propios píxeles) y las aplana,
    quitando ese canal, en una copia del archivo. Solo aplica a PDF.

    Motivo: salvaguarda adicional -- se confirmó que ciertas imágenes con
    transparencia pueden provocar crashes nativos similares al que motivó
    forzar el backend pypdfium2 (ver crear_converter). Aunque el caso
    diagnosticado no involucraba transparencia, se conserva este saneado
    por si algún otro documento sí la dispara por esa vía.

    Usamos pymupdf.Pixmap en vez de extraer el stream crudo de la imagen:
    extract_image() no combina una /SMask externa, así que una imagen con
    transparencia real vía máscara separada pasaría desapercibida si solo
    miráramos el modo de color del stream crudo. Pixmap sí la compone.

    Devuelve la ruta al PDF saneado, o la original si no había nada que
    hacer o si pymupdf no está disponible."""
    if not SANEADO_DISPONIBLE:
        return pdf_path

    try:
        doc = pymupdf.open(str(pdf_path))
        hubo_cambios = False
        total_imagenes = 0
        con_alfa = 0
        fallidas = 0

        for num_pagina, page in enumerate(doc, start=1):
            for img in page.get_images(full=True):
                xref = img[0]
                total_imagenes += 1
                try:
                    pix = pymupdf.Pixmap(doc, xref)

                    if not pix.alpha:
                        continue

                    con_alfa += 1

                    # Si además es CMYK u otro espacio de >3 canales, pasarlo
                    # a RGB antes de quitar el alfa.
                    if pix.colorspace and pix.colorspace.n >= 4:
                        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)

                    # Quita el canal alfa. No compone pixel a pixel sobre un
                    # fondo blanco -- para nuestro objetivo (evitar el crash)
                    # basta con eliminar el canal; el contenido visible no
                    # cambia de forma perceptible.
                    pix_plano = pymupdf.Pixmap(pix, 0)
                    page.replace_image(xref, pixmap=pix_plano)
                    hubo_cambios = True
                except Exception:
                    fallidas += 1
                    log(f"    ⚠ No se pudo sanear imagen xref={xref} (página {num_pagina}):\n"
                        f"{traceback.format_exc(limit=2)}")
                    continue

        log(f"    (saneado: {total_imagenes} imagen(es) revisada(s), "
            f"{con_alfa} con transparencia, {fallidas} fallida(s))")

        if not hubo_cambios:
            doc.close()
            return pdf_path

        saneado_path = tmp_dir / f"{pdf_path.stem}__saneado.pdf"
        doc.save(str(saneado_path), garbage=3, deflate=True)
        doc.close()
        return saneado_path

    except Exception:
        log(f"  ⚠ No se pudo sanear imágenes de {pdf_path.name}, se usa el original")
        return pdf_path


def partir_pdf(pdf_path: Path, tmp_dir: Path, tamano_trozo: int, umbral_particion: int):
    """Parte un PDF en archivos físicos más pequeños usando pypdf.
    Devuelve lista de (ruta_trozo, pagina_inicio, pagina_fin) en 1-indexado,
    para poder reportar errores con el número de página real del PDF original."""
    reader = PdfReader(str(pdf_path))
    total_paginas = len(reader.pages)

    if total_paginas <= umbral_particion:
        return [(pdf_path, 1, total_paginas)]

    trozos = []
    for inicio in range(0, total_paginas, tamano_trozo):
        fin = min(inicio + tamano_trozo, total_paginas)
        writer = PdfWriter()
        for p in range(inicio, fin):
            writer.add_page(reader.pages[p])
        trozo_path = tmp_dir / f"{pdf_path.stem}__p{inicio + 1}-{fin}.pdf"
        with open(trozo_path, "wb") as f:
            writer.write(f)
        trozos.append((trozo_path, inicio + 1, fin))

    return trozos


def _worker_convertir_documento(trozos_info, pipeline_options, device, queue):
    """Se ejecuta en UN ÚNICO PROCESO HIJO para el documento COMPLETO
    (todos sus trozos, si los tiene). El converter -- y por tanto el
    contexto CUDA -- se inicializa UNA SOLA VEZ aquí y se reutiliza para
    procesar cada trozo secuencialmente. Esto es lo que elimina el ciclo
    de reinicialización repetida de CUDA que causaba la degradación de
    rendimiento cuando cada trozo tenía su propio proceso.

    Si algo truena a nivel nativo (C++/CUDA), solo termina este proceso --
    nunca la aplicación principal ni el resto del lote. La contrapartida
    (documentada en el diagnóstico) es que se pierde granularidad de
    recuperación DENTRO de un mismo documento: un crash nativo genuino a
    mitad de documento se lleva por delante los trozos aún no procesados.
    Como mitigación, cada trozo se envía por la cola en cuanto termina
    (no se acumulan hasta el final), así que el proceso padre normalmente
    conserva los trozos completados antes del crash.

    trozos_info: lista de (ruta_trozo_str, pagina_inicio, pagina_fin).
    """
    try:
        if not cargar_dependencias():
            raise RuntimeError(f"No se pudieron cargar PyTorch/Docling: {_ERROR_DEPENDENCIAS}")
        converter = crear_converter(pipeline_options)
    except Exception:
        queue.put(("fatal", traceback.format_exc()))
        return

    for trozo_path_str, pag_inicio, pag_fin in trozos_info:
        try:
            result = converter.convert(trozo_path_str)
            page_errors = getattr(result, "errors", None)
            advertencias = [str(pe) for pe in page_errors] if page_errors else []
            markdown = result.document.export_to_markdown(traverse_pictures=True)
            queue.put(("chunk_ok", pag_inicio, pag_fin, markdown, advertencias))
        except Exception:
            # Fallo de Python (no un crash nativo) en un trozo puntual --
            # se reporta y se sigue con el resto de trozos del documento.
            queue.put(("chunk_error", pag_inicio, pag_fin, traceback.format_exc()))

    queue.put(("done",))

    try:
        liberar_memoria(device)
    except Exception:
        pass


def convertir_documento_aislado(trozos_info, pipeline_options, device, log_file, archivo: Path, timeout_total: int, log=print):
    """Lanza la conversión de TODOS los trozos de un documento en UN SOLO
    proceso hijo (aislamiento por documento, no por trozo -- ver
    _worker_convertir_documento). El timeout se aplica una sola vez, al
    proceso completo, no trozo a trozo.

    Devuelve (markdown_completo_o_None, lista_de_advertencias)."""
    queue = mp.Queue()
    proceso = mp.Process(
        target=_worker_convertir_documento,
        args=(trozos_info, pipeline_options, device, queue),
    )
    proceso.start()

    hay_varios_trozos = len(trozos_info) > 1
    partes_md = {}
    advertencias = []
    completado = False
    fatal_tb = None

    tiempo_inicio = time.monotonic()
    while True:
        tiempo_restante = timeout_total - (time.monotonic() - tiempo_inicio)
        if tiempo_restante <= 0:
            break
        try:
            mensaje = queue.get(timeout=tiempo_restante)
        except Exception:
            break  # nada nuevo en la cola dentro del tiempo restante

        tipo = mensaje[0]
        if tipo == "chunk_ok":
            _, pag_inicio, pag_fin, markdown, advertencias_trozo = mensaje
            partes_md[pag_inicio] = markdown
            advertencias.extend(advertencias_trozo)
            if hay_varios_trozos:
                log(f"    → completado [páginas {pag_inicio}-{pag_fin}]")
        elif tipo == "chunk_error":
            _, pag_inicio, pag_fin, tb = mensaje
            rango = f"páginas {pag_inicio}-{pag_fin}" if pag_fin else "documento completo"
            log(f"    ✗ falló [{rango}]")
            log_file.write(f"\n[{archivo}] FALLÓ el trozo [{rango}]\n{tb}\n")
            log_file.flush()
            ultima_linea = tb.strip().splitlines()[-1] if tb.strip() else "error desconocido"
            advertencias.append(f"[{rango}] {ultima_linea[:150]}")
            partes_md[pag_inicio] = None
        elif tipo == "fatal":
            fatal_tb = mensaje[1]
            break
        elif tipo == "done":
            completado = True
            break

    fue_timeout = False
    if not completado and fatal_tb is None and proceso.is_alive():
        fue_timeout = True
        proceso.terminate()
        proceso.join()
        mensaje = (f"TIMEOUT tras {timeout_total}s, documento cancelado "
                   f"({len(partes_md)}/{len(trozos_info)} trozo(s) recuperados antes del timeout)")
        log(f"    ⏱ {mensaje}")
        log_file.write(f"\n[{archivo}] {mensaje}\n")
        log_file.flush()
        advertencias.append(f"TIMEOUT tras {timeout_total}s, procesamiento cancelado")
    else:
        proceso.join(10)
        if proceso.is_alive():
            proceso.terminate()
            proceso.join()

    if fatal_tb:
        log_file.write(f"\n[{archivo}] NO SE PUDO INICIALIZAR EL CONVERTER\n{fatal_tb}\n")
        log_file.flush()
        ultima_linea = fatal_tb.strip().splitlines()[-1] if fatal_tb.strip() else "error desconocido"
        return None, [f"No se pudo inicializar el converter: {ultima_linea[:150]}"]

    # El chequeo de exitcode busca CRASHES NATIVOS genuinos (p.ej. el
    # proceso murió solo por STATUS_HEAP_CORRUPTION). Si fuimos NOSOTROS
    # quienes matamos el proceso por timeout, el exitcode también sale
    # distinto de 0 (por la señal de terminate()) pero eso ya se reportó
    # arriba como timeout -- reportarlo aquí también sería un falso "crash".
    if not fue_timeout and proceso.exitcode not in (0, None):
        mensaje = (f"Proceso hijo terminó con código {proceso.exitcode} (probable crash nativo), "
                   f"{len(partes_md)}/{len(trozos_info)} trozo(s) recuperados antes del crash")
        log(f"    ✗ {mensaje}")
        log_file.write(f"\n[{archivo}] {mensaje}\n")
        log_file.flush()
        advertencias.append(f"CRASH NATIVO (código {proceso.exitcode}) tras completar "
                             f"{len(partes_md)}/{len(trozos_info)} trozo(s)")

    if not partes_md:
        return None, advertencias

    partes_ordenadas = []
    for _, pag_inicio, pag_fin in trozos_info:
        markdown_trozo = partes_md.get(pag_inicio)
        if markdown_trozo is None:
            rango = f"páginas {pag_inicio}-{pag_fin}" if pag_fin else "documento completo"
            partes_ordenadas.append(f"\n\n> ⚠ **[Error: no se pudo convertir {rango}]**\n\n")
        else:
            partes_ordenadas.append(markdown_trozo)

    return "\n\n".join(partes_ordenadas), advertencias


def convertir_documento(archivo: Path, pipeline_options, device, log_file, tamano_trozo: int, umbral_particion: int, timeout_por_trozo: int, log=print):
    """Convierte un documento (PDF, DOCX o PPTX) y devuelve
    (markdown_completo_o_None, advertencias).

    Para PDF: sanea imágenes problemáticas y lo parte en trozos si es
    genuinamente largo (por encima de umbral_particion). Para PPTX con
    SmartArt: recupera el texto y la jerarquía del OOXML, que Docling omite,
    además de las tablas nativas. Los DOCX y PPTX sin esos objetos siguen
    la conversión aislada de Docling."""
    if archivo.suffix.lower() == ".pptx":
        try:
            markdown_smartart, detalles = extraer_smartart_pptx(archivo)
        except Exception:
            log_file.write(
                f"\n[{archivo}] NO SE PUDO EXTRAER SMARTART DEL PPTX\n"
                f"{traceback.format_exc()}\n"
            )
            log_file.flush()
            log("  ✗ No se pudo comprobar el contenido SmartArt (ver errores.log)")
            return None, ["No se pudo extraer o verificar el SmartArt del PPTX"]

        if detalles["smartart_diagrams"] or detalles["tables"]:
            if detalles["unsupported_graphics"]:
                mensaje = (
                    f"PPTX contiene {detalles['unsupported_graphics']} objeto(s) gráfico(s) "
                    "no compatibles con la extracción directa"
                )
                log_file.write(f"\n[{archivo}] {mensaje}\n")
                log_file.flush()
                log(f"  ✗ {mensaje} (ver errores.log)")
                return None, [mensaje]
            log(
                f"  (SmartArt: {detalles['smartart_diagrams']} diagrama(s), "
                f"{detalles['emitted_nodes']}/{detalles['smartart_nodes']} nodos; "
                f"tablas: {detalles['tables']}, celdas: {detalles['table_cells']})"
            )
            return markdown_smartart, []

    if archivo.suffix.lower() != ".pdf":
        trozos_info = [(str(archivo), 1, None)]
        timeout_total = timeout_por_trozo
        return convertir_documento_aislado(
            trozos_info, pipeline_options, device, log_file, archivo, timeout_total, log=log
        )

    with tempfile.TemporaryDirectory(prefix="docling_split_") as tmp:
        tmp_dir = Path(tmp)

        # Paso previo: enderezar páginas giradas (no hace nada si no hace falta)
        pdf_base = corregir_orientacion_paginas(archivo, tmp_dir, log=log)

        pdf_saneado = sanear_imagenes_alfa(pdf_base, tmp_dir, log=log) if SANEADO_DISPONIBLE else pdf_base
        if pdf_saneado != pdf_base:
            log("  (imágenes con transparencia saneadas)")

        try:
            trozos = (
                partir_pdf(pdf_saneado, tmp_dir, tamano_trozo, umbral_particion)
                if PYPDF_DISPONIBLE
                else [(pdf_saneado, 1, None)]
            )
        except Exception:
            log_file.write(f"\n[{archivo}] NO SE PUDO PARTIR EL PDF, se procesa entero\n{traceback.format_exc()}\n")
            trozos = [(pdf_saneado, 1, None)]

        if len(trozos) > 1:
            log(f"  (partido en {len(trozos)} trozos de ~{tamano_trozo} pág., "
                f"procesados en un único proceso aislado por documento)")

        # trozos_info se pasa al proceso hijo -- las rutas deben ser str,
        # no Path, para no depender de pickling de objetos Path entre
        # procesos en todas las plataformas.
        trozos_info = [(str(trozo_path), pag_inicio, pag_fin) for trozo_path, pag_inicio, pag_fin in trozos]

        # El timeout se aplica una sola vez, a nivel de documento completo
        # (ver convertir_documento_aislado), así que se presupuesta como
        # timeout_por_trozo × número de trozos del documento.
        timeout_total = timeout_por_trozo * max(1, len(trozos_info))

        return convertir_documento_aislado(
            trozos_info, pipeline_options, device, log_file, archivo, timeout_total, log=log
        )


def procesar_lote(archivos, output_dir: Path, usar_ocr: bool, usar_tablas: bool,
                   tamano_trozo: int, umbral_particion: int, timeout: int, device, log=print, on_progreso=None):
    """Procesa una lista de documentos (PDF/DOCX/PPTX) y escribe los .md
    correspondientes en output_dir. `log(mensaje)` recibe líneas de progreso
    (por defecto, print -> consola). `on_progreso(indice, total, nombre_archivo)`
    es opcional, pensado para actualizar una barra de progreso en la GUI.

    Devuelve un dict con el resumen de la corrida."""
    output_dir.mkdir(parents=True, exist_ok=True)

    pipeline_options = build_pipeline(usar_ocr, usar_tablas, device, log=log)
    modo_ocr = "con OCR" if usar_ocr else "sin OCR"
    modo_tablas = "con detección de tablas" if usar_tablas else "sin detección de tablas"
    modo_gpu = "GPU + CPU" if device == AcceleratorDevice.CUDA else "CPU"

    total = len(archivos)
    log(f"\n▶ Procesando {total} archivo(s) {modo_ocr}, {modo_tablas}, uno por uno (secuencial)")
    log(f"  PDFs de más de {umbral_particion} páginas se parten en trozos de {tamano_trozo}")
    log(f"  (cada documento se procesa en un único proceso aislado, timeout total = {timeout}s × nº de trozos)")
    if any(archivo.suffix.lower() == ".pptx" for archivo in archivos):
        log("  PPTX con SmartArt/tablas: extracción directa; OCR y detección visual de tablas solo afectan a PDFs")
    log(f"  Usando: {modo_gpu} ⚡\n")

    exitosos = 0
    con_advertencias = 0
    errores = 0
    fallidos = []
    log_path = output_dir / "errores.log"

    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(f"\n{'=' * 70}\nCorrida: {datetime.now().isoformat()}\n{'=' * 70}\n")

        for i, archivo in enumerate(archivos, start=1):
            if on_progreso:
                on_progreso(i, total, archivo.name)
            log(f"[{i}/{total}] {archivo.name} ...")

            valido, motivo = validar_documento(archivo)
            if not valido:
                log(f"  ✗ Omitido: {motivo}")
                log_file.write(f"\n[{archivo}] OMITIDO: {motivo}\n")
                errores += 1
                fallidos.append(f"{archivo} ({motivo})")
                continue

            markdown, advertencias = convertir_documento(
                archivo, pipeline_options, device, log_file, tamano_trozo, umbral_particion, timeout, log=log
            )

            if markdown is None:
                log("  ✗ Falló completamente")
                errores += 1
                fallidos.append(str(archivo))
                continue

            output_file = output_dir / f"{archivo.stem}.md"
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(markdown)

            if advertencias:
                con_advertencias += 1
                log_file.write(f"\n[{archivo}] CONVERTIDO CON {len(advertencias)} ADVERTENCIA(S):\n")
                for a in advertencias:
                    log_file.write(f"    - {a}\n")
                log_file.flush()
                log(f"  ✓ Guardado (con {len(advertencias)} advertencia(s), ver log)")
            else:
                log("  ✓ Guardado")

            exitosos += 1

    return {
        "total": total,
        "exitosos": exitosos,
        "con_advertencias": con_advertencias,
        "errores": errores,
        "fallidos": fallidos,
        "log_path": log_path,
    }


def _reportar_dependencias_faltantes(log=print):
    log(f"\n✗ No se pudieron cargar las dependencias principales (torch/docling): {_ERROR_DEPENDENCIAS}")
    log("  Revisa que el entorno virtual esté activado y que la instalación se haya completado.")
    if not PYPDF_DISPONIBLE:
        log("\n⚠ pypdf no está instalado: no se podrán partir PDFs largos ni validar archivos.")
        log("  Instálalo con:  pip install pypdf")
    if not SANEADO_DISPONIBLE:
        log("\n⚠ pymupdf/Pillow no disponibles: no se sanearán imágenes con transparencia.")
        log("  Instálalos con:  pip install pymupdf pillow")


# ---------------------------------------------------------------------------
# Modo consola (comportamiento original, preservado)
# ---------------------------------------------------------------------------

def main():
    print("═" * 60)
    print("   Conversor de documentos → Markdown (GPU + CPU)")
    print("   Soporta: PDF, Word (.docx), PowerPoint (.pptx)")
    print("═" * 60)

    print("Cargando PyTorch y Docling...")
    cargar_dependencias()
    if not DEPENDENCIAS_CORE_OK:
        _reportar_dependencias_faltantes()
        input("\nPresiona Enter para salir...")
        return

    if not PYPDF_DISPONIBLE:
        print("\n⚠ pypdf no está instalado: no se podrán partir PDFs largos.")
        print("  Instálalo con:  pip install pypdf")

    if not SANEADO_DISPONIBLE:
        print("\n⚠ pymupdf/Pillow no disponibles: no se sanearán imágenes con transparencia")
        print("  (algunos PDFs con portadas/imágenes RGBA podrían fallar). Instala con:")
        print("  pip install pymupdf pillow")

    device = detectar_acelerador()

    input_path = pedir_ruta("\nRuta del documento o carpeta: ")
    output_dir = pedir_ruta("Ruta de salida:               ")

    usar_ocr = pedir_ocr()
    usar_tablas = pedir_tablas()
    tamano_trozo = pedir_tamano_trozo()
    umbral_particion = pedir_umbral_particion(tamano_trozo)
    timeout = pedir_timeout()

    if not input_path.exists():
        print("\nLa ruta indicada no existe.")
        input("\nPresiona Enter para salir...")
        return

    archivos = buscar_documentos(input_path)

    if not archivos:
        print("\nNo se encontraron documentos compatibles (.pdf, .docx, .pptx).")
        input("\nPresiona Enter para salir...")
        return

    resumen = procesar_lote(
        archivos, output_dir, usar_ocr, usar_tablas, tamano_trozo, umbral_particion, timeout, device, log=print
    )

    print(f"\n{'═' * 60}")
    print(f"Resumen: {resumen['exitosos']} exitosos ({resumen['con_advertencias']} con advertencias), "
          f"{resumen['errores']} errores totales")
    if resumen["fallidos"]:
        print("\nArchivos que fallaron por completo:")
        for f in resumen["fallidos"]:
            print(f"  - {f}")
    if resumen["con_advertencias"] or resumen["fallidos"]:
        print(f"\nDetalle completo en: {resumen['log_path']}")
    print(f"{'═' * 60}")
    input("\nPresiona Enter para salir...")


# ---------------------------------------------------------------------------
# Interfaz gráfica simple (Tkinter) -- reusa procesar_lote tal cual
# ---------------------------------------------------------------------------

def lanzar_gui():
    import tkinter as tk
    import queue
    import threading
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("Conversor de documentos → Markdown (docling)")
    root.geometry("680x600")
    root.minsize(560, 460)

    ruta_var = tk.StringVar()
    salida_var = tk.StringVar()
    ocr_var = tk.BooleanVar(value=False)
    tablas_var = tk.BooleanVar(value=True)
    trozo_var = tk.StringVar(value=str(TAMANO_TROZO_POR_DEFECTO))
    umbral_var = tk.StringVar(value=str(UMBRAL_PARTICION_POR_DEFECTO))
    timeout_var = tk.StringVar(value=str(TIMEOUT_POR_TROZO_POR_DEFECTO))

    cola_log = queue.Queue()
    estado = {"corriendo": False}

    def log(mensaje):
        cola_log.put(str(mensaje))

    def elegir_archivo():
        ruta = filedialog.askopenfilename(
            filetypes=[
                ("Documentos soportados", "*.pdf *.docx *.pptx"),
                ("PDF", "*.pdf"),
                ("Word", "*.docx"),
                ("PowerPoint", "*.pptx"),
                ("Todos", "*.*"),
            ]
        )
        if ruta:
            ruta_var.set(ruta)

    def elegir_carpeta_entrada():
        ruta = filedialog.askdirectory()
        if ruta:
            ruta_var.set(ruta)

    def elegir_salida():
        ruta = filedialog.askdirectory()
        if ruta:
            salida_var.set(ruta)

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="Documento o carpeta de entrada (PDF, .docx, .pptx):").grid(
        row=0, column=0, columnspan=4, sticky="w"
    )
    ttk.Entry(frm, textvariable=ruta_var).grid(row=1, column=0, columnspan=2, sticky="we", padx=(0, 6))
    ttk.Button(frm, text="Archivo...", command=elegir_archivo).grid(row=1, column=2, padx=2)
    ttk.Button(frm, text="Carpeta...", command=elegir_carpeta_entrada).grid(row=1, column=3, padx=2)

    ttk.Label(frm, text="Carpeta de salida:").grid(row=2, column=0, columnspan=4, sticky="w", pady=(10, 0))
    ttk.Entry(frm, textvariable=salida_var).grid(row=3, column=0, columnspan=2, sticky="we", padx=(0, 6))
    ttk.Button(frm, text="Elegir...", command=elegir_salida).grid(row=3, column=2, padx=2)

    ttk.Checkbutton(frm, text="Aplicar OCR (solo PDF escaneados o con imágenes)", variable=ocr_var).grid(
        row=4, column=0, columnspan=4, sticky="w", pady=(10, 0)
    )
    ttk.Checkbutton(
        frm, text="Detectar estructura de tablas (más lento, solo aplica a PDF)", variable=tablas_var
    ).grid(row=5, column=0, columnspan=4, sticky="w")

    fila_umbral = ttk.Frame(frm)
    fila_umbral.grid(row=6, column=0, columnspan=4, sticky="w", pady=(6, 0))
    ttk.Label(fila_umbral, text="Umbral de partición (páginas; por debajo, nunca se parte):").pack(side="left")
    ttk.Entry(fila_umbral, textvariable=umbral_var, width=6).pack(side="left", padx=(6, 0))

    fila_trozo = ttk.Frame(frm)
    fila_trozo.grid(row=7, column=0, columnspan=4, sticky="w", pady=(4, 0))
    ttk.Label(fila_trozo, text="Tamaño de trozo si sí se parte (páginas):").pack(side="left")
    ttk.Entry(fila_trozo, textvariable=trozo_var, width=6).pack(side="left", padx=(6, 0))

    fila_timeout = ttk.Frame(frm)
    fila_timeout.grid(row=8, column=0, columnspan=4, sticky="w", pady=(4, 0))
    ttk.Label(fila_timeout, text="Segundos presupuestados por trozo (timeout total = esto × nº trozos):").pack(side="left")
    ttk.Entry(fila_timeout, textvariable=timeout_var, width=6).pack(side="left", padx=(6, 0))

    barra = ttk.Progressbar(frm, mode="determinate")
    barra.grid(row=9, column=0, columnspan=4, sticky="we", pady=(10, 4))

    texto_log = tk.Text(frm, height=18, state="disabled", bg="#111111", fg="#33ff33", font=("Consolas", 9))
    texto_log.grid(row=10, column=0, columnspan=4, sticky="nsew", pady=(4, 10))

    scroll = ttk.Scrollbar(frm, command=texto_log.yview)
    texto_log.configure(yscrollcommand=scroll.set)
    scroll.grid(row=10, column=4, sticky="ns")

    btn_iniciar = ttk.Button(frm, text="Iniciar conversión")
    btn_iniciar.grid(row=11, column=0, columnspan=4, sticky="we")

    for i in range(4):
        frm.columnconfigure(i, weight=1)
    frm.rowconfigure(10, weight=1)

    def escribir_log(mensaje):
        texto_log.configure(state="normal")
        texto_log.insert("end", mensaje + "\n")
        texto_log.see("end")
        texto_log.configure(state="disabled")

    def actualizar_progreso(i, total, _nombre):
        barra["maximum"] = max(total, 1)
        barra["value"] = i

    def poll_log():
        try:
            while True:
                escribir_log(cola_log.get_nowait())
        except queue.Empty:
            pass
        root.after(150, poll_log)

    def finalizar_ui():
        estado["corriendo"] = False
        btn_iniciar.configure(state="normal", text="Iniciar conversión")

    def hilo_conversion(ruta_str, salida_str, usar_ocr, usar_tablas, tamano_trozo, umbral_particion, timeout):
        try:
            log("Cargando PyTorch y Docling...")
            cargar_dependencias()
            if not DEPENDENCIAS_CORE_OK:
                _reportar_dependencias_faltantes(log=log)
                return

            input_path = Path(ruta_str)
            output_dir = Path(salida_str)

            if not input_path.exists():
                log("✗ La ruta indicada no existe.")
                return

            archivos = buscar_documentos(input_path)

            if not archivos:
                log("✗ No se encontraron documentos compatibles (.pdf, .docx, .pptx) en la ruta indicada.")
                return

            device = detectar_acelerador(log=log)

            resumen = procesar_lote(
                archivos, output_dir, usar_ocr, usar_tablas, tamano_trozo, umbral_particion, timeout, device,
                log=log,
                on_progreso=lambda i, total, nombre: root.after(0, actualizar_progreso, i, total, nombre),
            )

            log("")
            log(f"Resumen: {resumen['exitosos']} exitosos ({resumen['con_advertencias']} con advertencias), "
                f"{resumen['errores']} errores totales")
            if resumen["fallidos"]:
                log("Archivos que fallaron por completo:")
                for f in resumen["fallidos"]:
                    log(f"  - {f}")
            if resumen["con_advertencias"] or resumen["fallidos"]:
                log(f"Detalle completo en: {resumen['log_path']}")

        except Exception:
            log("✗ Error inesperado:")
            log(traceback.format_exc())
        finally:
            root.after(0, finalizar_ui)

    def iniciar():
        if estado["corriendo"]:
            return
        ruta_str = ruta_var.get().strip().strip('"')
        salida_str = salida_var.get().strip().strip('"')

        if not ruta_str:
            messagebox.showwarning("Falta información", "Elige un documento o carpeta de entrada.")
            return
        if not salida_str:
            messagebox.showwarning("Falta información", "Elige una carpeta de salida.")
            return

        try:
            tamano_trozo = int(trozo_var.get().strip())
            if tamano_trozo <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Valor no válido",
                f"El tamaño de trozo debe ser un número entero positivo. Se usará {TAMANO_TROZO_POR_DEFECTO}.",
            )
            tamano_trozo = TAMANO_TROZO_POR_DEFECTO
            trozo_var.set(str(tamano_trozo))

        try:
            umbral_particion = int(umbral_var.get().strip())
            if umbral_particion <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Valor no válido",
                f"El umbral de partición debe ser un número entero positivo. Se usará {UMBRAL_PARTICION_POR_DEFECTO}.",
            )
            umbral_particion = UMBRAL_PARTICION_POR_DEFECTO
            umbral_var.set(str(umbral_particion))
        # El umbral nunca debe ser menor que el tamaño de trozo (ver
        # pedir_umbral_particion): si lo fuera, cualquier documento que lo
        # supere generaría trozos casi tan grandes como él mismo.
        if umbral_particion < tamano_trozo:
            umbral_particion = tamano_trozo
            umbral_var.set(str(umbral_particion))

        try:
            timeout = int(timeout_var.get().strip())
            if timeout <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Valor no válido",
                f"El timeout debe ser un número entero positivo. Se usará {TIMEOUT_POR_TROZO_POR_DEFECTO}.",
            )
            timeout = TIMEOUT_POR_TROZO_POR_DEFECTO
            timeout_var.set(str(timeout))

        estado["corriendo"] = True
        btn_iniciar.configure(state="disabled", text="Procesando...")
        texto_log.configure(state="normal")
        texto_log.delete("1.0", "end")
        texto_log.configure(state="disabled")
        barra["value"] = 0

        threading.Thread(
            target=hilo_conversion,
            args=(ruta_str, salida_str, ocr_var.get(), tablas_var.get(), tamano_trozo, umbral_particion, timeout),
            daemon=True,
        ).start()

    btn_iniciar.configure(command=iniciar)

    root.after(150, poll_log)
    root.mainloop()


if __name__ == "__main__":
    mp.freeze_support()
    try:
        if "--gui" in sys.argv:
            lanzar_gui()
        else:
            main()
    except KeyboardInterrupt:
        print("\n\nInterrumpido por el usuario.")
