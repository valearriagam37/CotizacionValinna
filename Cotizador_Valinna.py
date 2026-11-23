# Cotizador de Valinna Baby Zone: arma la cotización y la entrega en PDF
# Se ejecuta desde esta carpeta con: python -m streamlit run Cotizador_Valinna.py
#	cd "C:\Users\vale1\OneDrive\Cursos\Analisis de datos con python IA_TecMTY\Proyecto Pág Valinna"
# Importa smtplib para enviar la cotización por correo
import smtplib
# Importa b64encode para incrustar el logo en la portada
from base64 import b64encode
# Importa escape para mostrar con seguridad los textos del catálogo
from html import escape
# Importa BytesIO para manejar el PDF y las imágenes en memoria
from io import BytesIO
# Importa la fecha para validar el evento y calcular la vigencia
from datetime import date, datetime, time, timedelta
# Importa EmailMessage para armar el correo con el PDF adjunto
from email.message import EmailMessage
# Importa lru_cache para recortar el logo una sola vez
from functools import lru_cache
# Importa sha1 para generar el folio de la cotización
from hashlib import sha1
# Importa hmac para comparar la contraseña de administración de forma segura
import hmac
# Importa json para adjuntar y leer el archivo del pedido
import json
# Importa logging para registrar fallas técnicas sin mostrárselas al cliente
import logging
# Importa Path para construir rutas relativas a este archivo
from pathlib import Path
# Importa re para limpiar el celular y el nombre del archivo
import re

# Importa pandas para leer el catálogo de Excel
import pandas as pd
# Importa pypdfium2 para mostrar el PDF en pantalla
import pypdfium2 as pdfium
# Importa streamlit para construir la página
import streamlit as st
# Importa el error que lanza streamlit cuando no existe el archivo de secretos
from streamlit.errors import StreamlitSecretNotFoundError
# Importa utilidades de Pillow para recortar el logo y las fotos
from PIL import Image, ImageChops, ImageOps
# Importa Color para definir los colores del PDF
from reportlab.lib.colors import Color
# Importa el tamaño de página carta, igual al del formato original
from reportlab.lib.pagesizes import letter
# Importa simpleSplit para partir textos largos en varias líneas
from reportlab.lib.utils import ImageReader, simpleSplit
# Importa el lienzo donde se dibuja el PDF
from reportlab.pdfgen import canvas


# Guarda la carpeta donde vive este archivo
CARPETA = Path(__file__).parent
# Define la ruta del catálogo de Excel
RUTA_CATALOGO = CARPETA / "Catalogo_Valinna.xlsx"
# Define la carpeta de las imágenes de paquetes y artículos
CARPETA_IMAGENES = CARPETA / "imagenes"
# Define la ruta del logo
RUTA_LOGO = CARPETA / "Logo Valinna.png"
# Define la ruta de la hoja de estilos de la página
RUTA_ESTILOS = CARPETA / "estilos.css"
# Define cuántos datos obligatorios debe completar el cliente
TOTAL_OBLIGATORIOS = 11
# Define las columnas obligatorias del catálogo
COLUMNAS_CATALOGO = ["Tipo", "Nombre", "Descripcion", "Precio", "Imagen"]

# Define el título del PDF
TITULO_DOCUMENTO = "Cotización"
# Crea el registro de fallas técnicas de la aplicación
registro = logging.getLogger("cotizador_valinna")
# Define el renglón del flete, cuyo costo se entrega después por mensaje
RENGLON_FLETE = {"descripcion": "Servicio de Flete", "cantidad": 1, "precio": None, "total": None}
# Define la aclaración de que los importes no incluyen el flete
NOTA_FLETE = "*Estos importes no incluyen el flete. Su costo se confirmará por mensaje y la reservación será el 50% del total, incluyendo el flete."
# Define el correo de Valinna, que envía y recibe las solicitudes
CORREO_VALINNA = "valinnababyzone@gmail.com"
# Define el servidor de correo de Gmail
SERVIDOR_CORREO = "smtp.gmail.com"
# Define el puerto cifrado del servidor de correo
PUERTO_CORREO = 465
# Define la leyenda de aceptación que cierra el PDF y la página
LEYENDA_ACEPTACION = (
    # Primera parte de la leyenda: la aceptación
    "Acepto la presente cotización y solicito que Valinna me contacte para "
    # Segunda parte de la leyenda: la confirmación pendiente
    "confirmarme la disponibilidad de la fecha, el costo del flete y los siguientes pasos."
# Cierra la leyenda
)
# Define el teléfono de contacto de Valinna
TELEFONO_VALINNA = "81 2583 1055"
# Define la invitación del PDF cuando aún no se solicita el servicio
INVITACION_PDF = f"Da clic en Solicitar servicio en nuestro cotizador y te estaremos contactando al WhatsApp que nos dejaste. Si lo prefieres, también puedes escribirnos al {TELEFONO_VALINNA}."
# Define los términos y condiciones tomados del formato
TERMINOS = [
    # Término sobre la vigencia
    "*Esta cotización cuenta con una validez de 10 días.",
    # Término sobre el IVA
    "*Esta nota de compra no incluye IVA. En caso de requerir factura, se cobra un 16% de IVA.",
    # Término sobre la liquidación
    "*Esta nota de compra se deberá liquidar un día antes del evento.",
    # Término sobre la reserva de la fecha
    "*Valinna no se hace responsable por la disponibilidad o reserva de la fecha del evento "
    # Continuación del término sobre el anticipo
    "hasta que se haya realizado el pago correspondiente al anticipo. La fecha quedará "
    # Continuación del término sobre la confirmación
    "oficialmente apartada una vez confirmado el pago del 50% del total del servicio, incluyendo el flete.",
# Cierra los términos
]

# Define el ancho de la página en puntos
ANCHO_PAGINA = letter[0]
# Define el alto de la página en puntos
ALTO_PAGINA = letter[1]
# Define el morado oscuro de la marca para encabezados
CIRUELA = Color(0.357, 0.129, 0.412)
# Define el morado medio para degradados y títulos
ORQUIDEA = Color(0.557, 0.192, 0.573)
# Define el fondo claro de tarjetas y renglones
LAVANDA = Color(0.976, 0.953, 0.984)
# Define el borde suave de las tarjetas
BORDE = Color(0.898, 0.812, 0.91)
# Define el color principal de los textos
TINTA = Color(0.16, 0.09, 0.20)
# Define el blanco de los textos sobre fondos oscuros
BLANCO = Color(1, 1, 1)
# Define los cuatro colores de los cubos del logo
COLORES_MARCA = [Color(0.969, 0.878, 0.42), Color(1, 0.769, 0.769), Color(0.839, 0.733, 1), Color(0.663, 0.898, 0.984)]
# Define el centro de la columna de cantidad
COLUMNA_CANTIDAD = 300
# Define el centro de la columna de precio
COLUMNA_PRECIO = 410
# Define el centro de la columna de total
COLUMNA_TOTAL = 520
# Define hasta dónde pueden bajar los renglones de la tabla
LIMITE_RENGLONES = 474
# Define cuántas horas antes del evento inicia el montaje
HORAS_MONTAJE = 1
# Define cuántas horas dura la renta a partir del inicio del evento
HORAS_RENTA = 4
# Define la hora más temprana para iniciar un evento
PRIMER_EVENTO = time(8, 0)
# Define la hora límite para el desmontaje del último evento
LIMITE_DESMONTAJE = time(21, 0)
# Define cada cuántos minutos se ofrece un horario de inicio
MINUTOS_ENTRE_HORARIOS = 30
# Define los días de vigencia que indican los términos
DIAS_VIGENCIA = 10
# Define los nombres de los meses en español
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
# Define los nombres de los días en español, iniciando en lunes
DIAS = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]

# Define el título del PDF que arma el personal de Valinna
TITULO_FINAL = "Cotización final"
# Define la tasa de IVA cuando el cliente requiere factura
TASA_IVA = 0.16
# Define la aclaración del plan de pago en la cotización final
NOTA_FINAL = "*Estos importes ya incluyen el flete y los ajustes acordados. La reservación es el 50% del total."
# Define el texto del recuadro final cuando el personal no escribe una nota
TEXTO_FINAL = f"Esta es tu cotización final. Para apartar la fecha realiza el pago de reservación y envíanos tu comprobante por WhatsApp al {TELEFONO_VALINNA}."
# Define el término que sustituye al del IVA cuando sí se factura
TERMINO_CON_IVA = "*Esta nota de compra incluye el 16% de IVA para su facturación."
# Define las columnas de la tabla de conceptos del panel de administración
COLUMNAS_PEDIDO = ["Concepto", "Cantidad", "Precio unitario"]
# Define las columnas de la tabla de promociones
COLUMNAS_PROMOCION = ["Descripción", "Tipo", "Valor"]
# Define los tipos de promoción disponibles
TIPOS_PROMOCION = ["Porcentaje (%)", "Monto fijo ($)"]
# Define cuántos caracteres caben en la nota del recuadro final del PDF
LARGO_NOTA = 180
# Define la carpeta donde se guardan los pedidos del panel de administración
CARPETA_PEDIDOS = CARPETA / "pedidos_guardados"
# Define los tipos de pago que se pueden registrar
TIPOS_PAGO = ["Anticipo", "Abono", "Liquidación"]
# Define las formas de pago aceptadas
FORMAS_PAGO = ["Transferencia", "Efectivo", "Tarjeta", "Depósito"]
# Define el nombre de quien firma los recibos
RESPONSABLE_RECIBOS = "Valeria Arriaga"
# Define los números del 0 al 29 escritos con letra
UNIDADES_LETRA = ["", "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve", "veinte", "veintiún", "veintidós", "veintitrés", "veinticuatro", "veinticinco", "veintiséis", "veintisiete", "veintiocho", "veintinueve"]
# Define las decenas a partir del treinta escritas con letra
DECENAS_LETRA = ["", "", "", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"]
# Define las centenas escritas con letra
CENTENAS_LETRA = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"]


# Da formato de moneda a un número
def moneda(valor):
    # Devuelve el aviso si el importe aún no se conoce
    if valor is None:
        # Indica que el importe sigue pendiente
        return "Por confirmar"
    # Convierte el cero negativo en cero para no mostrar "-$0.00"
    valor = valor + 0.0
    # Coloca el signo negativo antes del signo de pesos en los descuentos
    return f"-${abs(valor):,.2f}" if valor < 0 else f"${valor:,.2f}"


# Escribe una fecha como "13 de octubre 2025"
def fecha_larga(fecha):
    # Arma el texto con el nombre del mes en español
    return f"{fecha.day} de {MESES[fecha.month - 1]} {fecha.year}"


# Escribe una hora como "1:30 pm"
def hora_corta(hora):
    # Convierte la hora de 24 a 12 horas
    hora_12 = hora.hour % 12 or 12
    # Elige el sufijo según la mitad del día
    sufijo = "am" if hora.hour < 12 else "pm"
    # Devuelve la hora ya formateada
    return f"{hora_12}:{hora.minute:02d} {sufijo}"


# Calcula los horarios de montaje y desmontaje a partir del inicio del evento
def horarios_evento(fecha_evento, hora_evento):
    # Une la fecha y la hora del evento
    inicio = datetime.combine(fecha_evento, hora_evento)
    # Calcula el montaje antes del inicio del evento
    montaje = inicio - timedelta(hours=HORAS_MONTAJE)
    # Calcula el desmontaje al terminar las horas de renta
    desmontaje = inicio + timedelta(hours=HORAS_RENTA)
    # Devuelve los horarios como texto
    return {"montaje": hora_corta(montaje.time()), "desmontaje": hora_corta(desmontaje.time())}


# Arma la lista de horas de inicio cuyo desmontaje no rebasa el límite
def horarios_disponibles():
    # Toma cualquier fecha para poder sumar tiempo a las horas
    dia = date.today()
    # Calcula la última hora de inicio permitida
    ultimo = datetime.combine(dia, LIMITE_DESMONTAJE) - timedelta(hours=HORAS_RENTA)
    # Inicia en el primer horario del día
    actual = datetime.combine(dia, PRIMER_EVENTO)
    # Prepara la lista de horarios
    horarios = []
    # Avanza hasta llegar a la última hora permitida
    while actual <= ultimo:
        # Agrega el horario a la lista
        horarios.append(actual.time())
        # Avanza al siguiente horario
        actual += timedelta(minutes=MINUTOS_ENTRE_HORARIOS)
    # Devuelve los horarios disponibles
    return horarios


# Recorta los márgenes blancos del logo
@lru_cache(maxsize=1)
def logo_recortado():
    # Regresa None si el logo no está en la carpeta
    if not RUTA_LOGO.is_file():
        # Sin logo no hay nada que recortar
        return None
    # Abre el logo con su canal de transparencia
    logo = Image.open(RUTA_LOGO).convert("RGBA")
    # Crea un fondo blanco del mismo tamaño
    fondo = Image.new("RGBA", logo.size, "white")
    # Coloca el logo sobre el fondo blanco
    fondo.alpha_composite(logo)
    # Quita la transparencia ya resuelta
    plano = fondo.convert("RGB")
    # Calcula cuánto difiere cada píxel del blanco
    diferencia = ImageChops.difference(plano, Image.new("RGB", plano.size, "white"))
    # Ignora los píxeles casi blancos y ubica el dibujo
    caja = diferencia.convert("L").point(lambda valor: 255 if valor > 15 else 0).getbbox()
    # Prepara un búfer para guardar el resultado
    bufer = BytesIO()
    # Guarda el logo recortado como PNG
    plano.crop(caja).save(bufer, format="PNG")
    # Devuelve los bytes de la imagen
    return bufer.getvalue()


# Lee y valida el catálogo; la marca de tiempo refresca la caché si cambia el Excel
@st.cache_data
def cargar_catalogo(marca_tiempo):
    # Lee la primera hoja del Excel
    catalogo = pd.read_excel(RUTA_CATALOGO)
    # Quita espacios sobrantes en los encabezados
    catalogo.columns = catalogo.columns.astype(str).str.strip()
    # Busca las columnas que no vengan en el archivo
    faltantes = [columna for columna in COLUMNAS_CATALOGO if columna not in catalogo.columns]
    # Detiene la lectura si falta alguna columna
    if faltantes:
        # Informa cuáles columnas agregar
        raise ValueError(f"Al catálogo le faltan las columnas: {', '.join(faltantes)}.")
    # Descarta las filas totalmente vacías
    catalogo = catalogo[COLUMNAS_CATALOGO].dropna(how="all")
    # Recorre las columnas de texto
    for columna in ["Tipo", "Nombre", "Descripcion", "Imagen"]:
        # Convierte vacíos a texto vacío y limpia espacios
        catalogo[columna] = catalogo[columna].map(lambda valor: "" if pd.isna(valor) else str(valor).strip())
    # Unifica mayúsculas en el tipo
    catalogo["Tipo"] = catalogo["Tipo"].str.capitalize()
    # Convierte el precio a número y marca como vacío lo inválido
    catalogo["Precio"] = pd.to_numeric(catalogo["Precio"], errors="coerce")
    # Prepara la lista de problemas encontrados
    errores = []
    # Revisa cada fila del catálogo
    for indice, fila in catalogo.iterrows():
        # Calcula el renglón real en Excel, contando el encabezado
        renglon = indice + 2
        # Valida que el tipo sea uno de los permitidos
        if fila["Tipo"] not in ("Paquete", "Adicional"):
            # Registra el tipo inválido
            errores.append(f"Renglón {renglon}: Tipo debe ser Paquete o Adicional.")
        # Valida que exista el nombre
        if not fila["Nombre"]:
            # Registra el nombre vacío
            errores.append(f"Renglón {renglon}: falta el Nombre.")
        # Valida que el precio sea un número positivo o cero
        if pd.isna(fila["Precio"]) or fila["Precio"] < 0:
            # Registra el precio inválido
            errores.append(f"Renglón {renglon}: Precio debe ser un número mayor o igual a 0.")
    # Busca nombres repetidos, porque identifican a cada opción
    repetidos = catalogo.loc[catalogo["Nombre"].duplicated() & (catalogo["Nombre"] != ""), "Nombre"].unique()
    # Revisa si hubo repetidos
    if len(repetidos):
        # Registra los nombres repetidos
        errores.append(f"Nombres repetidos: {', '.join(repetidos)}.")
    # Revisa que exista al menos un paquete
    if not (catalogo["Tipo"] == "Paquete").any():
        # Registra la ausencia de paquetes
        errores.append("El catálogo necesita al menos un Paquete.")
    # Detiene la lectura si hubo cualquier problema
    if errores:
        # Reporta todos los problemas juntos
        raise ValueError(" ".join(errores))
    # Devuelve el catálogo limpio
    return catalogo


# Obtiene una credencial del archivo de secretos
def secreto(clave):
    # Intenta leer la credencial
    try:
        # Devuelve el valor guardado
        return st.secrets[clave]
    # Atiende la falta de la credencial o del archivo completo
    except (KeyError, FileNotFoundError, StreamlitSecretNotFoundError) as error:
        # Reporta cuál credencial falta
        raise ValueError(f"Falta configurar el secreto {clave}.") from error


# Suma los importes de la cotización, sin contar el flete por confirmar
def total_cotizacion(renglones):
    # Devuelve la suma de los renglones que ya tienen importe
    return sum(renglon["total"] for renglon in renglones if renglon["total"] is not None)


# Arma el archivo del pedido que el personal carga en el panel de administración
def pedido_json(datos, renglones, folio):
    # Reúne solo los conceptos con precio, porque el flete lo define el personal
    conceptos = [{"descripcion": r["descripcion"], "cantidad": int(r["cantidad"]), "precio": float(r["precio"])} for r in renglones if r["precio"] is not None]
    # Reúne los datos del cliente, del evento y los conceptos
    pedido = {
        # Guarda el folio para conservarlo en la cotización final
        "folio": folio,
        # Guarda el nombre del cliente
        "nombre": datos["nombre"],
        # Guarda el celular del cliente
        "celular": datos["celular"],
        # Guarda la dirección del evento
        "direccion": datos["direccion"],
        # Guarda la fecha del evento en formato AAAA-MM-DD
        "fecha_evento": datos["fecha_evento"].isoformat(),
        # Guarda la hora del evento en formato HH:MM
        "hora_evento": datos["hora_evento"].strftime("%H:%M"),
        # Guarda los conceptos elegidos por el cliente
        "conceptos": conceptos,
    # Cierra los datos del pedido
    }
    # Devuelve el pedido como texto JSON legible
    return json.dumps(pedido, ensure_ascii=False, indent=2).encode("utf-8")


# Envía por correo a Valinna el PDF de la cotización solicitada
def enviar_correo(pdf, nombre_archivo, datos, renglones, total):
    # Limpia los datos igual que en el PDF para que coincida el folio
    limpios = limpiar_datos(datos)
    # Calcula el folio de la cotización
    folio = folio_cotizacion(limpios, renglones)
    # Arma una línea por cada concepto cotizado
    conceptos = "\n".join(f"- {r['descripcion']} x {r['cantidad']}: {moneda(r['total'])}" for r in renglones)
    # Calcula los horarios de montaje y desmontaje
    horarios = horarios_evento(limpios["fecha_evento"], limpios["hora_evento"])
    # Crea el mensaje de correo
    mensaje = EmailMessage()
    # Indica el asunto con el folio y el cliente
    mensaje["Subject"] = f"Solicitud de servicio {folio} - {limpios['nombre']}"
    # Indica quién envía el correo
    mensaje["From"] = CORREO_VALINNA
    # Indica quién recibe el correo
    mensaje["To"] = CORREO_VALINNA
    # Escribe el cuerpo con lo necesario para calcular el flete y contactar al cliente
    mensaje.set_content(
        # Agrega el encabezado del aviso
        f"Un cliente solicitó el servicio. La cotización va adjunta.\n\n"
        # Agrega el folio
        f"Folio: {folio}\n"
        # Agrega el nombre del cliente
        f"Cliente: {limpios['nombre']}\n"
        # Agrega el celular del cliente
        f"WhatsApp: {limpios['celular']}\n"
        # Agrega la dirección del evento
        f"Dirección: {limpios['direccion']}\n"
        # Agrega la fecha y la hora del evento
        f"Evento: {fecha_larga(limpios['fecha_evento'])}, {hora_corta(limpios['hora_evento'])}\n"
        # Agrega los horarios de montaje y desmontaje
        f"Montaje: {horarios['montaje']} | Renta: {HORAS_RENTA} horas | Desmontaje: {horarios['desmontaje']}\n\n"
        # Agrega los conceptos cotizados
        f"{conceptos}\n\n"
        # Agrega el total sin flete
        f"Total sin flete: {moneda(total)}\n\n"
        # Explica cómo usar el archivo del pedido
        f"Para armar la cotización final, carga el archivo Pedido_{folio}.json en el panel de administración del cotizador (?admin=1).\n"
    # Cierra el cuerpo del correo
    )
    # Adjunta el PDF de la cotización
    mensaje.add_attachment(pdf, maintype="application", subtype="pdf", filename=nombre_archivo)
    # Adjunta el archivo del pedido para cargarlo en el panel de administración
    mensaje.add_attachment(pedido_json(limpios, renglones, folio), maintype="application", subtype="json", filename=f"Pedido_{folio}.json")
    # Obtiene la contraseña de aplicación sin los espacios con que la muestra Google
    contrasena = secreto("GMAIL_CONTRASENA_APLICACION").replace(" ", "")
    # Abre una conexión cifrada con Gmail
    with smtplib.SMTP_SSL(SERVIDOR_CORREO, PUERTO_CORREO, timeout=30) as servidor:
        # Inicia sesión con la cuenta de Valinna
        servidor.login(CORREO_VALINNA, contrasena)
        # Envía el mensaje
        servidor.send_message(mensaje)


# Dibuja un texto reduciendo la fuente si no cabe en el ancho indicado
def texto_ajustado(lienzo, x, arriba, texto, fuente, tamano, ancho_maximo):
    # Reduce el tamaño mientras el texto sea más ancho que lo permitido
    while lienzo.stringWidth(texto, fuente, tamano) > ancho_maximo and tamano > 6:
        # Baja medio punto en cada intento
        tamano -= 0.5
    # Aplica la fuente final
    lienzo.setFont(fuente, tamano)
    # Escribe el texto midiendo desde la parte superior de la página
    lienzo.drawString(x, ALTO_PAGINA - arriba, texto)


# Rellena un rectángulo con un degradado horizontal
def rectangulo_degradado(lienzo, x, arriba, ancho, alto, radio, color_inicial, color_final):
    # Guarda el estado para limitar el degradado a este rectángulo
    lienzo.saveState()
    # Crea la figura del rectángulo
    ruta = lienzo.beginPath()
    # Traza el rectángulo con sus esquinas redondeadas
    ruta.roundRect(x, ALTO_PAGINA - arriba - alto, ancho, alto, radio)
    # Recorta el área de dibujo a la figura
    lienzo.clipPath(ruta, stroke=0, fill=0)
    # Pinta el degradado de izquierda a derecha
    lienzo.linearGradient(x, 0, x + ancho, 0, (color_inicial, color_final), extend=False)
    # Recupera el estado anterior
    lienzo.restoreState()


# Dibuja la franja con los cuatro colores del logo
def franja_colores(lienzo, arriba):
    # Calcula el ancho de cada color
    ancho = ANCHO_PAGINA / len(COLORES_MARCA)
    # Recorre los colores de la marca
    for posicion, color in enumerate(COLORES_MARCA):
        # Selecciona el color
        lienzo.setFillColor(color)
        # Dibuja el tramo de la franja
        lienzo.rect(posicion * ancho, ALTO_PAGINA - arriba - 6, ancho, 6, stroke=0, fill=1)


# Dibuja el pie con los datos de contacto en cada página
def dibujar_pie(lienzo):
    # Selecciona el color del pie
    lienzo.setFillColor(CIRUELA)
    # Dibuja la banda inferior
    lienzo.rect(0, 0, ANCHO_PAGINA, 22, stroke=0, fill=1)
    # Selecciona el blanco para el texto
    lienzo.setFillColor(BLANCO)
    # Selecciona la fuente del pie
    lienzo.setFont("Helvetica", 8.5)
    # Escribe los datos de contacto
    lienzo.drawCentredString(ANCHO_PAGINA / 2, 8, f"Valinna Baby Zone   |   Monterrey, Nuevo León   |   Tel. {TELEFONO_VALINNA}")
    # Escribe el número de página
    lienzo.drawRightString(576, 8, f"Página {lienzo.getPageNumber()}")


# Dibuja una tarjeta clara con su título
def tarjeta_pdf(lienzo, x, arriba, ancho, alto, titulo):
    # Selecciona el fondo de la tarjeta
    lienzo.setFillColor(LAVANDA)
    # Selecciona el borde de la tarjeta
    lienzo.setStrokeColor(BORDE)
    # Define el grosor del borde
    lienzo.setLineWidth(1)
    # Dibuja la tarjeta
    lienzo.roundRect(x, ALTO_PAGINA - arriba - alto, ancho, alto, 10, stroke=1, fill=1)
    # Selecciona el color del título
    lienzo.setFillColor(ORQUIDEA)
    # Selecciona la fuente del título
    lienzo.setFont("Helvetica-Bold", 8.5)
    # Escribe el título de la tarjeta
    lienzo.drawString(x + 12, ALTO_PAGINA - arriba - 16, titulo)


# Dibuja el encabezado de la tabla a la altura indicada
def dibujar_encabezado_tabla(lienzo, arriba):
    # Dibuja la barra del encabezado
    rectangulo_degradado(lienzo, 36, arriba, 540, 26, 13, CIRUELA, ORQUIDEA)
    # Selecciona el blanco para los títulos
    lienzo.setFillColor(BLANCO)
    # Selecciona la fuente de los títulos
    lienzo.setFont("Helvetica-Bold", 10)
    # Escribe el título de la descripción
    lienzo.drawString(52, ALTO_PAGINA - arriba - 17, "Descripción")
    # Escribe el título de la cantidad
    lienzo.drawCentredString(COLUMNA_CANTIDAD, ALTO_PAGINA - arriba - 17, "Cantidad")
    # Escribe el título del precio
    lienzo.drawCentredString(COLUMNA_PRECIO, ALTO_PAGINA - arriba - 17, "Precio unitario (MXN)")
    # Escribe el título del total
    lienzo.drawCentredString(COLUMNA_TOTAL, ALTO_PAGINA - arriba - 17, "Total")


# Dibuja el encabezado principal y los datos del cliente y del evento
def dibujar_cabecera(lienzo, datos, folio, titulo, lineas=None):
    # Dibuja la banda superior
    rectangulo_degradado(lienzo, 0, 0, ANCHO_PAGINA, 104, 0, CIRUELA, ORQUIDEA)
    # Dibuja la franja con los colores del logo
    franja_colores(lienzo, 104)
    # Selecciona el blanco para el recuadro del logo
    lienzo.setFillColor(BLANCO)
    # Dibuja el recuadro del logo
    lienzo.roundRect(36, ALTO_PAGINA - 94, 116, 84, 14, stroke=0, fill=1)
    # Obtiene el logo recortado
    logo = logo_recortado()
    # Coloca el logo solo si existe
    if logo:
        # Dibuja el logo conservando su proporción
        lienzo.drawImage(ImageReader(BytesIO(logo)), 42, ALTO_PAGINA - 90, width=104, height=76, preserveAspectRatio=True, anchor="c", mask="auto")
    # Reduce la fuente del título si no cabe junto al logo
    tamano_titulo = 30 if lienzo.stringWidth(titulo.upper(), "Helvetica-Bold", 30) < 400 else 24
    # Selecciona la fuente del título
    lienzo.setFont("Helvetica-Bold", tamano_titulo)
    # Escribe el título del documento
    lienzo.drawRightString(576, ALTO_PAGINA - 46, titulo.upper())
    # Selecciona la fuente de los datos del documento
    lienzo.setFont("Helvetica", 9.5)
    # Usa los datos de la cotización cuando no se indican otros, como en los recibos
    if lineas is None:
        # Calcula hasta cuándo es válida la cotización
        vigencia = datos["fecha_cotizacion"] + timedelta(days=DIAS_VIGENCIA)
        # Reúne los datos del documento con su posición
        lineas = ((64, f"Folio: {folio}"), (78, f"Emisión: {fecha_larga(datos['fecha_cotizacion'])}"), (92, f"Vigencia: hasta {fecha_larga(vigencia)}"))
    # Recorre los datos del documento
    for arriba, linea in lineas:
        # Escribe el dato alineado a la derecha
        lienzo.drawRightString(576, ALTO_PAGINA - arriba, linea)
    # Dibuja la tarjeta del cliente
    tarjeta_pdf(lienzo, 36, 122, 200, 96, "CLIENTE")
    # Selecciona el color del texto principal
    lienzo.setFillColor(TINTA)
    # Inicia la posición del nombre
    arriba = 154
    # Parte el nombre para que quepa en la tarjeta
    for linea in simpleSplit(datos["nombre"], "Helvetica-Bold", 11, 176):
        # Escribe la línea del nombre
        texto_ajustado(lienzo, 48, arriba, linea, "Helvetica-Bold", 11, 176)
        # Baja a la siguiente línea
        arriba += 13
    # Selecciona la fuente del celular
    lienzo.setFont("Helvetica", 10)
    # Escribe el celular de contacto
    lienzo.drawString(48, ALTO_PAGINA - 206, f"WhatsApp: {datos['celular']}")
    # Dibuja la tarjeta del evento
    tarjeta_pdf(lienzo, 248, 122, 328, 96, "EVENTO")
    # Selecciona el color del texto principal
    lienzo.setFillColor(TINTA)
    # Selecciona la fuente de los datos del evento
    lienzo.setFont("Helvetica", 9.5)
    # Inicia la posición de los datos del evento
    arriba = 152
    # Parte la dirección para que quepa en la tarjeta
    for linea in simpleSplit(datos["direccion"], "Helvetica", 9.5, 304):
        # Escribe la línea de la dirección
        lienzo.drawString(260, ALTO_PAGINA - arriba, linea)
        # Baja a la siguiente línea
        arriba += 12
    # Arma la línea con la fecha y la hora del evento
    fecha_y_hora = f"{DIAS[datos['fecha_evento'].weekday()].capitalize()} {fecha_larga(datos['fecha_evento'])}   |   {hora_corta(datos['hora_evento'])}"
    # Selecciona la fuente de la fecha y la hora
    lienzo.setFont("Helvetica-Bold", 9.5)
    # Escribe la fecha y la hora
    lienzo.drawString(260, ALTO_PAGINA - arriba, fecha_y_hora)
    # Calcula los horarios de montaje y desmontaje
    horarios = horarios_evento(datos["fecha_evento"], datos["hora_evento"])
    # Escribe los horarios sin salirse de la tarjeta
    texto_ajustado(lienzo, 260, arriba + 12, f"Montaje: {horarios['montaje']}   |   Renta: {HORAS_RENTA} horas   |   Desmontaje: {horarios['desmontaje']}", "Helvetica", 9.5, 304)


# Agrega una página de continuación para la tabla
def nueva_pagina_tabla(lienzo, folio, titulo):
    # Dibuja el pie de la página actual
    dibujar_pie(lienzo)
    # Cierra la página actual
    lienzo.showPage()
    # Dibuja la banda superior reducida
    rectangulo_degradado(lienzo, 0, 0, ANCHO_PAGINA, 40, 0, CIRUELA, ORQUIDEA)
    # Dibuja la franja con los colores del logo
    franja_colores(lienzo, 40)
    # Selecciona el blanco para el texto
    lienzo.setFillColor(BLANCO)
    # Selecciona la fuente del texto
    lienzo.setFont("Helvetica-Bold", 11)
    # Escribe el título y el folio
    lienzo.drawString(36, ALTO_PAGINA - 26, f"{titulo}  |  Folio: {folio}")
    # Repite el encabezado de la tabla
    dibujar_encabezado_tabla(lienzo, 60)


# Dibuja los renglones de la tabla, agregando páginas si hacen falta
def dibujar_renglones(lienzo, renglones, folio, titulo):
    # Inicia debajo del encabezado de la primera página
    arriba = 258
    # Recorre cada concepto cotizado
    for numero, renglon in enumerate(renglones):
        # Parte la descripción para que no invada la columna de cantidad
        lineas = simpleSplit(texto_imprimible(renglon["descripcion"]), "Helvetica-Bold", 10, 200)
        # Calcula el alto que ocupa el concepto
        alto = max(24, 13 * len(lineas) + 11)
        # Revisa si el concepto ya no cabe en la página
        if arriba + alto > LIMITE_RENGLONES:
            # Continúa la tabla en otra página
            nueva_pagina_tabla(lienzo, folio, titulo)
            # Reinicia la posición debajo del encabezado
            arriba = 90
        # Sombrea uno de cada dos renglones para facilitar la lectura
        if numero % 2 == 0:
            # Selecciona el fondo del renglón
            lienzo.setFillColor(LAVANDA)
            # Dibuja el fondo del renglón
            lienzo.roundRect(36, ALTO_PAGINA - arriba - alto, 540, alto, 6, stroke=0, fill=1)
        # Selecciona el color del texto
        lienzo.setFillColor(TINTA)
        # Selecciona la fuente de la descripción
        lienzo.setFont("Helvetica-Bold", 10)
        # Recorre cada línea de la descripción
        for posicion, linea in enumerate(lineas):
            # Escribe la línea de la descripción
            lienzo.drawString(52, ALTO_PAGINA - arriba - 16 - 13 * posicion, linea)
        # Selecciona la fuente de los importes
        lienzo.setFont("Helvetica", 10)
        # Escribe la cantidad
        lienzo.drawCentredString(COLUMNA_CANTIDAD, ALTO_PAGINA - arriba - 16, str(renglon["cantidad"]))
        # Resalta los importes que siguen por confirmar
        if renglon["total"] is None:
            # Selecciona el color del aviso
            lienzo.setFillColor(ORQUIDEA)
            # Selecciona la fuente del aviso
            lienzo.setFont("Helvetica-Oblique", 10)
        # Escribe el precio unitario
        lienzo.drawCentredString(COLUMNA_PRECIO, ALTO_PAGINA - arriba - 16, moneda(renglon["precio"]))
        # Escribe el total del renglón
        lienzo.drawCentredString(COLUMNA_TOTAL, ALTO_PAGINA - arriba - 16, moneda(renglon["total"]))
        # Avanza al siguiente renglón
        arriba += alto


# Dibuja el plan de pago y el total; final trae los ajustes de la cotización final
def dibujar_totales(lienzo, total, final=None):
    # Usa el anticipo ya recibido cuando el personal lo registró
    if final and final["anticipo"] > 0:
        # Toma el anticipo recibido
        anticipo = final["anticipo"]
        # Nombra el importe como pagos ya recibidos
        etiqueta_anticipo = "Pagos recibidos"
    # Calcula la reservación cuando aún no hay anticipo
    else:
        # Calcula el anticipo del 50% que piden los términos
        anticipo = round(total / 2, 2)
        # Nombra el pago como reservación
        etiqueta_anticipo = "Pago de Reservación (50%)"
    # Dibuja la tarjeta del plan de pago
    tarjeta_pdf(lienzo, 36, 484, 312, 80, "PLAN DE PAGO")
    # Selecciona el color del texto principal
    lienzo.setFillColor(TINTA)
    # Selecciona la fuente de los pagos
    lienzo.setFont("Helvetica-Bold", 10.5)
    # Reúne los pagos con su posición
    pagos = ((516, etiqueta_anticipo, anticipo), (531, "Saldo pendiente", total - anticipo))
    # Recorre los pagos
    for arriba, etiqueta, importe in pagos:
        # Escribe el concepto del pago
        lienzo.drawString(48, ALTO_PAGINA - arriba, etiqueta)
        # Escribe el importe del pago
        lienzo.drawRightString(336, ALTO_PAGINA - arriba, moneda(importe))
    # Selecciona la fuente de la aclaración del flete
    lienzo.setFont("Helvetica", 7.5)
    # Inicia la posición de la aclaración
    arriba = 545
    # Parte la aclaración para respetar el ancho de la tarjeta
    for linea in simpleSplit(NOTA_FINAL if final else NOTA_FLETE, "Helvetica", 7.5, 288):
        # Escribe la línea de la aclaración
        lienzo.drawString(48, ALTO_PAGINA - arriba, linea)
        # Baja a la siguiente línea
        arriba += 9.5
    # Dibuja el recuadro del total
    rectangulo_degradado(lienzo, 360, 484, 216, 80, 12, CIRUELA, ORQUIDEA)
    # Selecciona el blanco para el texto del total
    lienzo.setFillColor(BLANCO)
    # Selecciona la fuente de la etiqueta
    lienzo.setFont("Helvetica-Bold", 9.5)
    # Escribe la etiqueta del total
    lienzo.drawCentredString(468, ALTO_PAGINA - 505, "TOTAL" if final else "TOTAL (SIN FLETE)")
    # Calcula el tamaño de letra para que el total quepa en el recuadro
    tamano = 26 if lienzo.stringWidth(moneda(total), "Helvetica-Bold", 26) < 196 else 18
    # Selecciona la fuente del importe
    lienzo.setFont("Helvetica-Bold", tamano)
    # Escribe el importe del total
    lienzo.drawCentredString(468, ALTO_PAGINA - 534, moneda(total))
    # Selecciona la fuente de la moneda
    lienzo.setFont("Helvetica", 8.5)
    # Escribe la moneda
    lienzo.drawCentredString(468, ALTO_PAGINA - 551, "Pesos mexicanos")


# Dibuja los términos y los datos de contacto
def dibujar_terminos(lienzo, terminos=TERMINOS):
    # Selecciona el fondo de la banda
    lienzo.setFillColor(LAVANDA)
    # Dibuja la banda a todo lo ancho
    lienzo.rect(0, ALTO_PAGINA - 692, ANCHO_PAGINA, 118, stroke=0, fill=1)
    # Selecciona el color de los títulos
    lienzo.setFillColor(CIRUELA)
    # Selecciona la fuente de los títulos
    lienzo.setFont("Helvetica-Bold", 11)
    # Escribe el título de los términos
    lienzo.drawString(36, ALTO_PAGINA - 592, "Términos y condiciones")
    # Escribe el título de los datos de contacto
    lienzo.drawString(420, ALTO_PAGINA - 592, "Nuestros datos")
    # Selecciona el color del texto
    lienzo.setFillColor(TINTA)
    # Selecciona la fuente de los términos
    lienzo.setFont("Helvetica", 8.5)
    # Inicia la posición de los términos
    arriba = 606
    # Recorre cada término
    for termino in terminos:
        # Parte el término para respetar el ancho de la columna
        for linea in simpleSplit(termino, "Helvetica", 8.5, 360):
            # Escribe la línea del término
            lienzo.drawString(36, ALTO_PAGINA - arriba, linea)
            # Baja a la siguiente línea
            arriba += 11
    # Selecciona la fuente de los datos de contacto
    lienzo.setFont("Helvetica", 10)
    # Recorre los datos de contacto con su posición
    for arriba, linea in zip((608, 623, 638), ("Valeria Arriaga", TELEFONO_VALINNA, "Monterrey, Nuevo León.")):
        # Escribe el dato de contacto
        lienzo.drawString(420, ALTO_PAGINA - arriba, linea)


# Dibuja el recuadro final según si el cliente ya solicitó el servicio
def dibujar_recuadro_final(lienzo, solicitada, final=None):
    # Usa el título y la nota de la cotización final cuando aplica
    if final:
        # Elige el título del recuadro final, que los recibos pueden cambiar
        titulo = final.get("titulo", "COTIZACIÓN FINAL")
        # Usa la nota del personal o el texto predeterminado
        texto = texto_imprimible(final["notas"]) or TEXTO_FINAL
    # Usa los textos de la cotización del cliente
    else:
        # Elige el título del recuadro
        titulo = "SERVICIO SOLICITADO" if solicitada else "¿TE INTERESA ESTE SERVICIO?"
        # Elige el texto del recuadro
        texto = LEYENDA_ACEPTACION if solicitada else INVITACION_PDF
    # Selecciona el contorno del recuadro
    lienzo.setStrokeColor(ORQUIDEA)
    # Define el grosor del contorno
    lienzo.setLineWidth(1.5)
    # Dibuja el recuadro
    lienzo.roundRect(36, ALTO_PAGINA - 758, 540, 54, 10, stroke=1, fill=0)
    # Selecciona el color del título
    lienzo.setFillColor(ORQUIDEA)
    # Selecciona la fuente del título
    lienzo.setFont("Helvetica-Bold", 9.5)
    # Escribe el título centrado
    lienzo.drawCentredString(ANCHO_PAGINA / 2, ALTO_PAGINA - 719, titulo)
    # Selecciona el color del texto
    lienzo.setFillColor(TINTA)
    # Selecciona la fuente del texto
    lienzo.setFont("Helvetica", 9.5)
    # Inicia la posición del texto
    arriba = 733
    # Parte el texto para que quepa en el recuadro
    for linea in simpleSplit(texto, "Helvetica", 9.5, 510):
        # Escribe la línea centrada
        lienzo.drawCentredString(ANCHO_PAGINA / 2, ALTO_PAGINA - arriba, linea)
        # Baja a la siguiente línea
        arriba += 12


# Quita los caracteres que las fuentes del PDF no pueden dibujar, como los emojis
def texto_imprimible(texto):
    # Conserva solo lo representable y limpia los espacios sobrantes
    return texto.encode("cp1252", "ignore").decode("cp1252").strip()


# Limpia el nombre y la dirección capturados por el cliente
def limpiar_datos(datos):
    # Devuelve una copia de los datos con sus textos imprimibles
    return {**datos, "nombre": texto_imprimible(datos["nombre"]), "direccion": texto_imprimible(datos["direccion"])}


# Genera un folio estable a partir del contenido de la cotización
def folio_cotizacion(datos, renglones):
    # Calcula una huella del contenido de la cotización
    huella = sha1(repr((datos, renglones)).encode("utf-8")).hexdigest()
    # Arma el folio con la fecha y parte de la huella
    return f"VAL-{datos['fecha_cotizacion']:%y%m%d}-{huella[:4].upper()}"


# Genera el PDF completo y devuelve sus bytes
def generar_pdf(datos, renglones, solicitada, final=None):
    # Limpia los datos capturados por el cliente
    datos = limpiar_datos(datos)
    # Conserva el folio original del pedido o calcula uno nuevo
    folio = (final or {}).get("folio") or folio_cotizacion(datos, renglones)
    # Elige el título según si es la cotización final
    titulo = TITULO_FINAL if final else TITULO_DOCUMENTO
    # Cambia el término del IVA cuando la cotización final sí lo incluye
    terminos = [TERMINO_CON_IVA if "IVA" in termino else termino for termino in TERMINOS] if final and final["con_iva"] else TERMINOS
    # Crea el búfer que recibirá el PDF
    bufer = BytesIO()
    # Crea el lienzo en tamaño carta con un resultado reproducible
    lienzo = canvas.Canvas(bufer, pagesize=letter, invariant=1)
    # Asigna el título interno del archivo
    lienzo.setTitle(f"{titulo} Valinna - {datos['nombre']}")
    # Dibuja el encabezado y los datos del cliente y del evento
    dibujar_cabecera(lienzo, datos, folio, titulo)
    # Dibuja el encabezado de la tabla
    dibujar_encabezado_tabla(lienzo, 230)
    # Dibuja los conceptos cotizados
    dibujar_renglones(lienzo, renglones, folio, titulo)
    # Dibuja el plan de pago y el total
    dibujar_totales(lienzo, total_cotizacion(renglones), final)
    # Dibuja los términos y los datos de contacto
    dibujar_terminos(lienzo, terminos)
    # Dibuja el recuadro final
    dibujar_recuadro_final(lienzo, solicitada, final)
    # Dibuja el pie de la última página
    dibujar_pie(lienzo)
    # Guarda el documento
    lienzo.save()
    # Devuelve el contenido del PDF
    return bufer.getvalue()


# Escribe con letra un número entero del 1 al 999
def letras_menor_mil(numero):
    # Atiende el caso especial del cien exacto
    if numero == 100:
        # Devuelve "cien" en lugar de "ciento"
        return "cien"
    # Separa las centenas del resto
    centena, resto = divmod(numero, 100)
    # Inicia con la centena, si la hay
    partes = [CENTENAS_LETRA[centena]] if centena else []
    # Usa la lista directa para los números menores a 30
    if 0 < resto < 30:
        # Agrega el número escrito
        partes.append(UNIDADES_LETRA[resto])
    # Arma las decenas con "y" para los números de 30 en adelante
    elif resto >= 30:
        # Separa la decena de la unidad
        decena, unidad = divmod(resto, 10)
        # Agrega la decena y, si existe, la unidad
        partes.append(DECENAS_LETRA[decena] + (f" y {UNIDADES_LETRA[unidad]}" if unidad else ""))
    # Une las partes con espacios
    return " ".join(partes)


# Escribe con letra un número entero de pesos
def numero_letras(numero):
    # Atiende el cero
    if numero == 0:
        # Devuelve el cero escrito
        return "cero"
    # Separa los millones del resto
    millones, resto = divmod(numero, 1_000_000)
    # Separa los miles de las unidades
    miles, unidades = divmod(resto, 1000)
    # Prepara las partes del texto
    partes = []
    # Escribe los millones, si los hay
    if millones:
        # Usa el singular para un millón
        partes.append("un millón" if millones == 1 else f"{letras_menor_mil(millones)} millones")
    # Escribe los miles, si los hay
    if miles:
        # Usa "mil" solo para mil exacto
        partes.append("mil" if miles == 1 else f"{letras_menor_mil(miles)} mil")
    # Escribe las unidades, si las hay
    if unidades:
        # Agrega las unidades escritas
        partes.append(letras_menor_mil(unidades))
    # Une las partes con espacios
    return " ".join(partes)


# Escribe un importe con letra como "(DOS MIL PESOS 00/100 M.N.)"
def importe_letras(valor):
    # Convierte el importe en centavos para evitar errores de redondeo
    pesos, centavos = divmod(round(valor * 100), 100)
    # Usa el singular solo para un peso
    palabra = "peso" if pesos == 1 else "pesos"
    # Agrega "de" a los millones exactos, como en "un millón de pesos"
    if pesos and pesos % 1_000_000 == 0:
        # Ajusta la palabra de la moneda
        palabra = "de pesos"
    # Devuelve el importe en mayúsculas con los centavos
    return f"({numero_letras(pesos)} {palabra} {centavos:02d}/100 M.N.)".upper()


# Genera el PDF del recibo de un pago y devuelve sus bytes
def generar_recibo(datos, pago, anteriores, total):
    # Limpia los datos del cliente para el PDF
    datos = limpiar_datos(datos)
    # Usa el folio del pedido, o lo toma del número de recibo
    folio = datos["folio"] or pago["recibo"].rsplit("-R", 1)[0]
    # Toma el importe de este pago
    monto = float(pago["monto"])
    # Usa el total vigente cuando se registró el pago, para que el recibo no cambie después
    total = float(pago.get("total", total))
    # Calcula el saldo que queda después de este pago
    saldo = round(total - anteriores - monto, 2)
    # Revisa si con este pago queda liquidado el servicio
    liquidado = saldo <= 0.005
    # Elige el título del recibo
    titulo = "Recibo de liquidación" if liquidado else "Recibo de pago"
    # Convierte la fecha del pago
    fecha_pago = date.fromisoformat(pago["fecha"])
    # Calcula la fecha límite para liquidar, un día antes del evento
    limite = datos["fecha_evento"] - timedelta(days=1)
    # Crea el búfer que recibirá el PDF
    bufer = BytesIO()
    # Crea el lienzo en tamaño carta con un resultado reproducible
    lienzo = canvas.Canvas(bufer, pagesize=letter, invariant=1)
    # Asigna el título interno del archivo
    lienzo.setTitle(f"{titulo} {pago['recibo']} - {datos['nombre']}")
    # Reúne los datos del recibo con su posición
    lineas = ((64, f"Recibo: {pago['recibo']}"), (78, f"Fecha de pago: {fecha_larga(fecha_pago)}"), (92, f"Cotización: {folio}"))
    # Dibuja el encabezado y los datos del cliente y del evento
    dibujar_cabecera(lienzo, datos, folio, titulo, lineas)
    # Selecciona el color del texto principal
    lienzo.setFillColor(TINTA)
    # Escribe a quién se le recibe el pago
    texto_ajustado(lienzo, 36, 250, f"Recibimos de {datos['nombre']} la cantidad de:", "Helvetica", 11, 540)
    # Dibuja el recuadro del importe
    rectangulo_degradado(lienzo, 36, 262, 540, 74, 14, CIRUELA, ORQUIDEA)
    # Selecciona el blanco para el importe
    lienzo.setFillColor(BLANCO)
    # Selecciona la fuente del importe
    lienzo.setFont("Helvetica-Bold", 28)
    # Escribe el importe centrado
    lienzo.drawCentredString(ANCHO_PAGINA / 2, ALTO_PAGINA - 300, moneda(monto))
    # Escribe el importe con letra
    letras = importe_letras(monto)
    # Inicia el tamaño de la letra del importe escrito
    tamano = 9.5
    # Reduce la letra si el importe escrito no cabe
    while lienzo.stringWidth(letras, "Helvetica", tamano) > 510 and tamano > 6:
        # Baja medio punto en cada intento
        tamano -= 0.5
    # Selecciona la fuente del importe escrito
    lienzo.setFont("Helvetica", tamano)
    # Escribe el importe con letra centrado
    lienzo.drawCentredString(ANCHO_PAGINA / 2, ALTO_PAGINA - 322, letras)
    # Dibuja la barra del encabezado de la tabla
    rectangulo_degradado(lienzo, 36, 356, 540, 26, 13, CIRUELA, ORQUIDEA)
    # Selecciona la fuente de los títulos de la tabla
    lienzo.setFont("Helvetica-Bold", 10)
    # Escribe el título del concepto
    lienzo.drawString(52, ALTO_PAGINA - 373, "Concepto")
    # Escribe el título de la forma de pago
    lienzo.drawCentredString(COLUMNA_CANTIDAD, ALTO_PAGINA - 373, "Forma de pago")
    # Escribe el título de la referencia
    lienzo.drawCentredString(COLUMNA_PRECIO, ALTO_PAGINA - 373, "Referencia")
    # Escribe el título del importe
    lienzo.drawCentredString(COLUMNA_TOTAL, ALTO_PAGINA - 373, "Importe")
    # Selecciona el fondo del renglón
    lienzo.setFillColor(LAVANDA)
    # Dibuja el fondo del renglón
    lienzo.roundRect(36, ALTO_PAGINA - 382 - 34, 540, 34, 6, stroke=0, fill=1)
    # Selecciona el color del texto
    lienzo.setFillColor(TINTA)
    # Selecciona la fuente del concepto
    lienzo.setFont("Helvetica-Bold", 10)
    # Parte el concepto en máximo dos líneas
    for posicion, linea in enumerate(simpleSplit(f"{pago['tipo']} - Cotización {folio}", "Helvetica-Bold", 10, 200)[:2]):
        # Escribe la línea del concepto
        lienzo.drawString(52, ALTO_PAGINA - 397 - 13 * posicion, linea)
    # Selecciona la fuente de los datos del pago
    lienzo.setFont("Helvetica", 10)
    # Escribe la forma de pago
    lienzo.drawCentredString(COLUMNA_CANTIDAD, ALTO_PAGINA - 397, pago["forma"])
    # Toma solo la primera línea de la referencia para que no invada otra columna
    referencia = (simpleSplit(texto_imprimible(pago.get("referencia", "")), "Helvetica", 10, 100) or ["-"])[0]
    # Escribe la referencia
    lienzo.drawCentredString(COLUMNA_PRECIO, ALTO_PAGINA - 397, referencia)
    # Selecciona la fuente del importe
    lienzo.setFont("Helvetica-Bold", 10)
    # Escribe el importe del pago
    lienzo.drawCentredString(COLUMNA_TOTAL, ALTO_PAGINA - 397, moneda(monto))
    # Dibuja la tarjeta del estado de cuenta
    tarjeta_pdf(lienzo, 36, 436, 312, 112, "ESTADO DE CUENTA")
    # Selecciona el color del texto
    lienzo.setFillColor(TINTA)
    # Selecciona la fuente de los importes
    lienzo.setFont("Helvetica", 10.5)
    # Reúne los importes del estado de cuenta con su posición
    cuentas = ((466, "Total del servicio", total), (482, "Pagos anteriores", anteriores), (498, "Este pago", monto))
    # Recorre los importes
    for arriba, etiqueta, importe in cuentas:
        # Escribe el concepto
        lienzo.drawString(48, ALTO_PAGINA - arriba, etiqueta)
        # Escribe el importe
        lienzo.drawRightString(336, ALTO_PAGINA - arriba, moneda(importe))
    # Selecciona el color de la línea divisoria
    lienzo.setStrokeColor(BORDE)
    # Dibuja la línea antes del saldo
    lienzo.line(48, ALTO_PAGINA - 507, 336, ALTO_PAGINA - 507)
    # Selecciona la fuente del saldo
    lienzo.setFont("Helvetica-Bold", 12)
    # Escribe la etiqueta del saldo
    lienzo.drawString(48, ALTO_PAGINA - 527, "Saldo pendiente")
    # Escribe el saldo, sin mostrar negativos
    lienzo.drawRightString(336, ALTO_PAGINA - 527, moneda(max(saldo, 0)))
    # Dibuja el sello de liquidado cuando ya no hay saldo
    if liquidado:
        # Dibuja el recuadro morado
        rectangulo_degradado(lienzo, 360, 436, 216, 112, 12, CIRUELA, ORQUIDEA)
        # Selecciona el blanco para el texto
        lienzo.setFillColor(BLANCO)
        # Selecciona la fuente del sello
        lienzo.setFont("Helvetica-Bold", 26)
        # Escribe el sello
        lienzo.drawCentredString(468, ALTO_PAGINA - 494, "LIQUIDADO")
        # Selecciona la fuente del agradecimiento
        lienzo.setFont("Helvetica", 10)
        # Escribe el agradecimiento
        lienzo.drawCentredString(468, ALTO_PAGINA - 516, "¡Gracias por tu pago!")
    # Dibuja el saldo pendiente y su fecha límite
    else:
        # Selecciona el fondo del recuadro
        lienzo.setFillColor(LAVANDA)
        # Selecciona el contorno del recuadro
        lienzo.setStrokeColor(ORQUIDEA)
        # Define el grosor del contorno
        lienzo.setLineWidth(1.5)
        # Dibuja el recuadro
        lienzo.roundRect(360, ALTO_PAGINA - 436 - 112, 216, 112, 12, stroke=1, fill=1)
        # Selecciona el color del título
        lienzo.setFillColor(ORQUIDEA)
        # Selecciona la fuente del título
        lienzo.setFont("Helvetica-Bold", 9.5)
        # Escribe el título
        lienzo.drawCentredString(468, ALTO_PAGINA - 462, "SALDO PENDIENTE")
        # Selecciona el color del importe
        lienzo.setFillColor(TINTA)
        # Selecciona la fuente del importe
        lienzo.setFont("Helvetica-Bold", 22)
        # Escribe el saldo
        lienzo.drawCentredString(468, ALTO_PAGINA - 492, moneda(saldo))
        # Selecciona la fuente de la fecha límite
        lienzo.setFont("Helvetica", 8.5)
        # Escribe la indicación de la fecha límite
        lienzo.drawCentredString(468, ALTO_PAGINA - 514, "Liquidar a más tardar el")
        # Selecciona la fuente de la fecha
        lienzo.setFont("Helvetica-Bold", 9.5)
        # Escribe la fecha límite
        lienzo.drawCentredString(468, ALTO_PAGINA - 528, fecha_larga(limite))
    # Selecciona el color de la línea de firma
    lienzo.setStrokeColor(TINTA)
    # Define el grosor de la línea de firma
    lienzo.setLineWidth(0.8)
    # Dibuja la línea de firma
    lienzo.line(340, ALTO_PAGINA - 640, 560, ALTO_PAGINA - 640)
    # Selecciona el color del texto
    lienzo.setFillColor(TINTA)
    # Selecciona la fuente de la firma
    lienzo.setFont("Helvetica", 9.5)
    # Escribe quién recibe el pago
    lienzo.drawCentredString(450, ALTO_PAGINA - 654, f"Recibió: {RESPONSABLE_RECIBOS}")
    # Escribe el nombre del negocio debajo de la firma
    lienzo.drawCentredString(450, ALTO_PAGINA - 667, "Valinna Baby Zone")
    # Selecciona la fuente de la aclaración fiscal
    lienzo.setFont("Helvetica-Oblique", 8.5)
    # Aclara que el recibo no es una factura
    lienzo.drawString(36, ALTO_PAGINA - 654, "Este recibo no es un comprobante fiscal.")
    # Elige el mensaje del recuadro final según el estado del pago
    if liquidado:
        # Agradece la liquidación total
        mensaje = "Tu servicio quedó liquidado en su totalidad. ¡Gracias por confiar en Valinna Baby Zone para tu evento!"
    # Atiende el anticipo que aparta la fecha
    elif pago["tipo"] == "Anticipo":
        # Confirma la fecha y recuerda el saldo
        mensaje = f"Con este pago queda apartada la fecha de tu evento. El saldo de {moneda(saldo)} se deberá liquidar a más tardar el {fecha_larga(limite)}."
    # Atiende los abonos
    else:
        # Agradece y recuerda el saldo
        mensaje = f"Gracias por tu pago. El saldo de {moneda(saldo)} se deberá liquidar a más tardar el {fecha_larga(limite)}."
    # Dibuja el recuadro final con el mensaje
    dibujar_recuadro_final(lienzo, True, {"titulo": "GRACIAS POR TU PAGO", "notas": mensaje})
    # Dibuja el pie de la página
    dibujar_pie(lienzo)
    # Guarda el documento
    lienzo.save()
    # Devuelve el contenido del PDF
    return bufer.getvalue()


# Convierte el PDF en imágenes para mostrarlo en pantalla
@st.cache_data(max_entries=20)
def paginas_pdf(pdf):
    # Abre el PDF en memoria
    documento = pdfium.PdfDocument(pdf)
    # Prepara la lista de imágenes
    paginas = []
    # Recorre cada página del documento
    for pagina in documento:
        # Prepara un búfer para la imagen
        bufer = BytesIO()
        # Dibuja la página en alta resolución y la guarda como PNG
        pagina.render(scale=2).to_pil().save(bufer, format="PNG")
        # Agrega la imagen a la lista
        paginas.append(bufer.getvalue())
    # Cierra el documento
    documento.close()
    # Devuelve las imágenes de las páginas
    return paginas


# Prepara la foto de una tarjeta en formato 4:3 para que todas midan igual
@st.cache_data
def imagen_tarjeta(ruta_texto, marca_tiempo):
    # Abre la foto respetando la orientación con la que se tomó
    foto = ImageOps.exif_transpose(Image.open(ruta_texto)).convert("RGB")
    # Recorta la foto desde el centro al tamaño de la tarjeta
    recorte = ImageOps.fit(foto, (800, 600))
    # Prepara un búfer para guardar el resultado
    bufer = BytesIO()
    # Guarda la foto como JPEG liviano
    recorte.save(bufer, format="JPEG", quality=85)
    # Devuelve los bytes de la foto
    return bufer.getvalue()


# Prepara el logo en formato 4:3, sin recortarlo, para las tarjetas sin foto
@lru_cache(maxsize=1)
def logo_tarjeta():
    # Centra el logo sobre un lienzo blanco del tamaño de la tarjeta
    lienzo = ImageOps.pad(Image.open(BytesIO(logo_recortado())), (800, 600), color="white", centering=(0.5, 0.5))
    # Reduce la escala del logo para dejarle aire alrededor
    marco = Image.new("RGB", (800, 600), "white")
    # Coloca el logo reducido al centro del marco
    marco.paste(lienzo.resize((560, 420)), (120, 90))
    # Prepara un búfer para guardar el resultado
    bufer = BytesIO()
    # Guarda el resultado como PNG
    marco.save(bufer, format="PNG")
    # Devuelve los bytes de la imagen
    return bufer.getvalue()


# Muestra la imagen de una opción del catálogo
def mostrar_imagen(nombre_archivo):
    # Construye la ruta de la imagen indicada en el catálogo
    ruta = CARPETA_IMAGENES / nombre_archivo
    # Usa la imagen si el catálogo la indica y existe
    if nombre_archivo and ruta.is_file():
        # Intenta preparar la foto
        try:
            # Muestra la imagen del producto
            st.image(imagen_tarjeta(str(ruta), ruta.stat().st_mtime), width="stretch")
            # Termina porque la foto ya se mostró
            return
        # Atiende una foto dañada o con un formato no válido
        except OSError as error:
            # Registra el archivo que falló
            registro.warning("No se pudo abrir la imagen %s: %s", nombre_archivo, error)
    # Usa el logo como respaldo cuando no hay imagen
    if logo_recortado():
        # Muestra el logo en lugar de la imagen faltante
        st.image(logo_tarjeta(), width="stretch")


# Aplica el diseño de la página
def aplicar_estilos():
    # Revisa que exista la hoja de estilos
    if RUTA_ESTILOS.is_file():
        # Inserta los estilos en la página
        st.markdown(f"<style>{RUTA_ESTILOS.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


# Muestra la portada con el logo y el mensaje principal
def mostrar_portada():
    # Obtiene el logo recortado
    logo = logo_recortado()
    # Convierte el logo en una imagen incrustada, si existe
    imagen = f'<div class="portada-logo"><img src="data:image/png;base64,{b64encode(logo).decode()}" alt="Valinna Baby Zone"></div>' if logo else ""
    # Arma los tres pasos que resumen el proceso
    pasos = '<div class="portada-pasos"><span>1. Elige o personaliza</span><span>2. Cuéntanos de tu evento</span><span>3. Solicita tu servicio</span></div>'
    # Arma el mensaje principal
    texto = f'<div class="portada-texto"><h1>Cotiza tu evento en minutos</h1><p>Arma tu paquete, agrega lo que quieras y obtén tu cotización en PDF al instante.</p>{pasos}</div>'
    # Muestra la portada completa
    st.markdown(f'<div class="portada">{imagen}{texto}</div>', unsafe_allow_html=True)


# Muestra el título numerado de un paso
def titulo_paso(numero, titulo, nota=""):
    # Arma la nota aclaratoria solo si existe
    aclaracion = f'<span class="paso-nota">{escape(nota)}</span>' if nota else ""
    # Muestra el número y el título del paso
    st.markdown(f'<div class="paso"><span class="paso-numero paso-{numero}">{numero}</span><span class="paso-titulo">{escape(titulo)}</span>{aclaracion}</div>', unsafe_allow_html=True)


# Muestra el nombre, la descripción y el precio de una tarjeta
def info_tarjeta(opcion):
    # Escribe "Sin costo" cuando el precio es cero
    precio = "Sin costo" if opcion["Precio"] == 0 else moneda(opcion["Precio"])
    # Arma la descripción solo si existe
    descripcion = f'<div class="tarjeta-descripcion">{escape(opcion["Descripcion"])}</div>' if opcion["Descripcion"] else ""
    # Muestra la información de la opción
    st.markdown(f'<div class="tarjeta-nombre">{escape(opcion["Nombre"])}</div>{descripcion}<div class="tarjeta-precio">{precio}</div>', unsafe_allow_html=True)


# Dibuja una tarjeta y devuelve la zona donde van sus controles
def cuerpo_tarjeta(opcion, horizontal):
    # Apila imagen e información cuando hay varias opciones por fila
    if not horizontal:
        # Muestra la imagen de la opción
        mostrar_imagen(opcion["Imagen"])
        # Muestra el nombre, la descripción y el precio
        info_tarjeta(opcion)
        # Devuelve la tarjeta misma para colocar los controles
        return st.container()
    # Crea dos columnas para aprovechar el ancho cuando la opción es única
    columna_imagen, columna_info = st.columns(2, vertical_alignment="center")
    # Entra a la columna de la imagen
    with columna_imagen:
        # Muestra la imagen de la opción
        mostrar_imagen(opcion["Imagen"])
    # Entra a la columna de la información
    with columna_info:
        # Muestra el nombre, la descripción y el precio
        info_tarjeta(opcion)
    # Devuelve la columna de la información para colocar los controles
    return columna_info


# Quita un concepto de la cotización desmarcando su casilla
def quitar_concepto(indice):
    # Desmarca la casilla del concepto en su pestaña
    st.session_state[f"concepto_{indice}"] = False


# Muestra las opciones como tarjetas y devuelve los renglones elegidos
def tarjetas_opciones(opciones):
    # Prepara la lista de conceptos elegidos
    renglones = []
    # Avisa si esta pestaña aún no tiene opciones
    if opciones.empty:
        # Muestra el aviso en la pestaña
        st.info("Por ahora no hay opciones disponibles en esta sección.")
        # Regresa la lista vacía
        return renglones
    # Usa toda la fila cuando solo hay una opción
    por_fila = 1 if len(opciones) == 1 else 2
    # Recorre las opciones en grupos por fila
    for inicio in range(0, len(opciones), por_fila):
        # Crea las columnas de la fila
        columnas = st.columns(por_fila)
        # Recorre las opciones de este grupo
        for columna, (indice, opcion) in zip(columnas, opciones.iloc[inicio:inicio + por_fila].iterrows()):
            # Revisa si la opción ya está marcada
            marcada = st.session_state.get(f"concepto_{indice}", False)
            # Crea la tarjeta, resaltada si la opción está incluida
            with columna.container(key=f"tarjeta_{'elegida' if marcada else 'normal'}_{indice}"):
                # Dibuja la tarjeta y entra a la zona de sus controles
                with cuerpo_tarjeta(opcion, por_fila == 1):
                    # Muestra la casilla para incluir la opción
                    incluida = st.checkbox("Agregar a mi cotización", key=f"concepto_{indice}")
                    # Pide la cantidad solo si la opción está incluida
                    if incluida:
                        # Solicita cuántas piezas quiere
                        cantidad = st.number_input("Cantidad", min_value=1, value=1, step=1, key=f"cantidad_{indice}")
                        # Agrega la opción con su importe
                        renglones.append({"indice": indice, "descripcion": opcion["Nombre"], "cantidad": cantidad, "precio": float(opcion["Precio"]), "total": cantidad * float(opcion["Precio"])})
    # Devuelve los conceptos elegidos
    return renglones


# Muestra las dos pestañas del catálogo y devuelve todos los conceptos elegidos
def seccion_catalogo(catalogo):
    # Escribe el título de la sección
    titulo_paso(1, "Arma tu cotización", "Puedes elegir más de uno")
    # Crea las dos pestañas
    pestana_paquetes, pestana_personaliza = st.tabs(["Paquetes armados", "Personaliza tu paquete"])
    # Entra a la pestaña de paquetes
    with pestana_paquetes:
        # Muestra los paquetes y recibe los elegidos
        paquetes = tarjetas_opciones(catalogo[catalogo["Tipo"] == "Paquete"])
    # Entra a la pestaña de personalización
    with pestana_personaliza:
        # Muestra los artículos y recibe los elegidos
        articulos = tarjetas_opciones(catalogo[catalogo["Tipo"] == "Adicional"])
    # Devuelve primero los paquetes y después los artículos
    return paquetes + articulos


# Muestra los campos obligatorios y devuelve lo capturado
def seccion_datos():
    # Escribe el título de la sección
    titulo_paso(2, "Cuéntanos de tu evento", "Todos los campos con * son obligatorios")
    # Crea el panel del formulario
    with st.container(key="panel_datos"):
        # Crea las columnas del nombre y el celular
        columna_nombre, columna_celular = st.columns(2)
        # Pide el nombre completo
        nombre = columna_nombre.text_input("Nombre completo *", max_chars=60).strip()
        # Pide el celular y deja solo sus dígitos
        celular = re.sub(r"\D", "", columna_celular.text_input("Celular de contacto (10 dígitos) *", max_chars=20))
        # Avisa que siguen los campos de la dirección
        st.markdown('<div class="subtitulo-panel">Dirección completa del evento</div>', unsafe_allow_html=True)
        # Crea las columnas de la calle y el número
        columna_calle, columna_numero = st.columns([3, 1])
        # Pide la calle con un largo que quepa en el PDF
        calle = columna_calle.text_input("Calle *", max_chars=50).strip()
        # Pide el número, que admite interior o S/N
        numero = columna_numero.text_input("Número *", max_chars=10).strip()
        # Crea las columnas de la colonia y el municipio
        columna_colonia, columna_municipio = st.columns(2)
        # Pide la colonia
        colonia = columna_colonia.text_input("Colonia *", max_chars=40).strip()
        # Pide el municipio
        municipio = columna_municipio.text_input("Municipio *", max_chars=30).strip()
        # Crea las columnas del estado y el código postal
        columna_estado, columna_cp = st.columns(2)
        # Pide el estado
        estado = columna_estado.text_input("Estado *", max_chars=20).strip()
        # Pide el código postal y deja solo sus dígitos
        codigo_postal = re.sub(r"\D", "", columna_cp.text_input("Código postal (5 dígitos) *", max_chars=5))
        # Avisa que siguen la fecha y la hora
        st.markdown('<div class="subtitulo-panel">Fecha y hora</div>', unsafe_allow_html=True)
        # Crea dos columnas para la fecha y la hora
        columna_fecha, columna_hora = st.columns(2)
        # Pide la fecha sin permitir días pasados
        fecha_evento = columna_fecha.date_input("Fecha del evento *", value=None, min_value=date.today(), format="DD/MM/YYYY")
        # Obtiene las horas de inicio permitidas
        horarios_inicio = horarios_disponibles()
        # Pide la hora mostrando las opciones en formato am/pm
        hora_evento = columna_hora.selectbox("Hora del evento *", horarios_inicio, index=None, format_func=hora_corta, placeholder="Elige la hora", help=f"El último evento inicia a las {hora_corta(horarios_inicio[-1])} para que el desmontaje sea a las {hora_corta(LIMITE_DESMONTAJE)}.")
        # Muestra los horarios calculados cuando ya hay fecha y hora
        if fecha_evento and hora_evento:
            # Calcula los horarios de montaje y desmontaje
            horarios = horarios_evento(fecha_evento, hora_evento)
            # Informa los horarios al cliente
            st.caption(f"Montaje: {horarios['montaje']}  |  Renta: {HORAS_RENTA} horas a partir del inicio del evento  |  Desmontaje: {horarios['desmontaje']}")
    # Agrupa los campos de la dirección con el nombre que se muestra si faltan
    campos_direccion = {"Calle": calle, "Número": numero, "Colonia": colonia, "Municipio": municipio, "Estado": estado}
    # Une la dirección en un solo texto para el PDF
    direccion = f"{calle} {numero}, {colonia}, {municipio}, {estado}, C.P. {codigo_postal}"
    # Devuelve todos los datos capturados
    return {"nombre": nombre, "celular": celular, "campos_direccion": campos_direccion, "codigo_postal": codigo_postal, "direccion": direccion, "fecha_evento": fecha_evento, "hora_evento": hora_evento, "fecha_cotizacion": date.today()}


# Revisa qué datos obligatorios faltan
def datos_faltantes(datos, hay_conceptos):
    # Prepara la lista de pendientes
    faltantes = []
    # Revisa que haya al menos un concepto elegido
    if not hay_conceptos:
        # Registra el concepto pendiente
        faltantes.append("Al menos un paquete o artículo")
    # Revisa el nombre
    if not datos["nombre"]:
        # Registra el nombre pendiente
        faltantes.append("Nombre completo")
    # Revisa que el celular tenga 10 dígitos
    if len(datos["celular"]) != 10:
        # Registra el celular pendiente
        faltantes.append("Celular de contacto de 10 dígitos")
    # Registra los campos de la dirección que sigan vacíos
    faltantes += [campo for campo, valor in datos["campos_direccion"].items() if not valor]
    # Revisa que el código postal tenga 5 dígitos
    if len(datos["codigo_postal"]) != 5:
        # Registra el código postal pendiente
        faltantes.append("Código postal de 5 dígitos")
    # Revisa la fecha
    if datos["fecha_evento"] is None:
        # Registra la fecha pendiente
        faltantes.append("Fecha del evento")
    # Revisa la hora
    if datos["hora_evento"] is None:
        # Registra la hora pendiente
        faltantes.append("Hora del evento")
    # Devuelve los pendientes
    return faltantes


# Muestra los conceptos, con opción de quitarlos, y el total dentro del panel de resumen
def mostrar_importes(conceptos, total):
    # Agrupa las líneas para mantenerlas juntas
    with st.container(key="lista_conceptos"):
        # Recorre los conceptos elegidos
        for concepto in conceptos:
            # Crea una columna para el texto y otra para el botón de quitar
            columna_texto, columna_quitar = st.columns([9, 2], vertical_alignment="center")
            # Muestra el concepto con su cantidad e importe
            columna_texto.markdown(f'<div class="resumen-linea"><span>{escape(str(concepto["descripcion"]))}<small>Cantidad: {concepto["cantidad"]}</small></span><strong>{moneda(concepto["total"])}</strong></div>', unsafe_allow_html=True)
            # Muestra el botón que quita el concepto de la cotización
            columna_quitar.button("✕", key=f"quitar_{concepto['indice']}", help=f"Quitar {concepto['descripcion']}", on_click=quitar_concepto, args=(concepto["indice"],))
        # Crea las columnas del flete, que no se puede quitar
        columna_flete, _ = st.columns([9, 2], vertical_alignment="center")
        # Muestra el flete con su costo por confirmar
        columna_flete.markdown(f'<div class="resumen-linea"><span>{RENGLON_FLETE["descripcion"]}<small>Por mensaje</small></span><strong>{moneda(None)}</strong></div>', unsafe_allow_html=True)
    # Arma el recuadro del total
    recuadro = f'<div class="resumen-total"><span>Total sin flete (MXN)</span><strong>{moneda(total)}</strong></div>'
    # Arma la línea del pago de reservación
    reservacion = f'<div class="resumen-reserva">Pago de Reservación (50%): <strong>{moneda(round(total / 2, 2))}</strong> + 50% del flete</div>'
    # Muestra el total y la reservación
    st.markdown(f"{recuadro}{reservacion}", unsafe_allow_html=True)
    # Aclara el tratamiento del flete y del IVA
    st.caption("El costo del flete se confirmará por mensaje y la reservación será el 50% del total, incluyendo el flete. No incluye IVA. En caso de requerir factura, se cobra un 16% de IVA.")


# Arma el aviso que recibe el cliente al solicitar el servicio
def mensaje_solicitud(celular):
    # Devuelve el aviso con las dos formas de contacto
    return f"¡Recibimos tu solicitud! Nos pondremos en contacto contigo al WhatsApp {celular} para confirmar disponibilidad, costo del flete y siguientes pasos. Si lo prefieres, también puedes escribirnos al {TELEFONO_VALINNA}."


# Muestra en una ventana la cotización completa con sus acciones
@st.dialog("Tu cotización completa", width="large")
def ventana_cotizacion(datos, renglones, nombre_archivo, total, firma):
    # Revisa si esta cotización ya se solicitó
    solicitada = st.session_state.get("firma_solicitada") == firma
    # Genera el PDF, con la leyenda solo si ya se solicitó
    pdf = generar_pdf(datos, renglones, solicitada)
    # Recorre las páginas de la cotización
    for pagina in paginas_pdf(pdf):
        # Muestra la página a todo el ancho de la ventana
        st.image(pagina, width="stretch")
    # Muestra solo la confirmación y la descarga si ya se solicitó
    if solicitada:
        # Recuerda al cliente que será contactado
        st.success(mensaje_solicitud(datos["celular"]))
        # Muestra el botón de descarga sin cerrar la ventana
        st.download_button("Descargar PDF", pdf, file_name=nombre_archivo, mime="application/pdf", on_click="ignore", width="stretch")
        # Termina porque la solicitud ya se envió
        return
    # Invita al cliente a solicitar el servicio
    st.markdown(f'<div class="invitacion"><strong>¿Te interesa el servicio?</strong> Da clic en <strong>Solicitar servicio</strong> y te estaremos contactando al WhatsApp que nos dejaste: <strong>{escape(datos["celular"])}</strong>. Si lo prefieres, también puedes escribirnos al <strong>{TELEFONO_VALINNA}</strong>.</div>', unsafe_allow_html=True)
    # Crea las columnas de las dos acciones
    columna_descarga, columna_solicitud = st.columns(2)
    # Muestra el botón de descarga sin cerrar la ventana
    columna_descarga.download_button("Descargar PDF", pdf, file_name=nombre_archivo, mime="application/pdf", on_click="ignore", width="stretch")
    # Muestra el botón con el que el cliente solicita el servicio
    if columna_solicitud.button("Solicitar servicio", key="solicitar", type="primary", width="stretch"):
        # Intenta enviar el PDF a Valinna
        try:
            # Muestra un aviso mientras termina el envío
            with st.spinner("Enviando tu solicitud a Valinna..."):
                # Envía el PDF con la leyenda de aceptación
                enviar_correo(generar_pdf(datos, renglones, True), nombre_archivo, datos, renglones, total)
        # Atiende fallas de conexión, de configuración o rechazos del envío
        except (smtplib.SMTPException, OSError, ValueError) as error:
            # Registra el detalle técnico para Valinna
            registro.warning("No se pudo enviar la solicitud: %s", error)
            # Avisa al cliente que la solicitud no se realizó
            st.error(f"No pudimos enviar tu solicitud. Intenta de nuevo o escríbenos al {TELEFONO_VALINNA}.")
        # Continúa solo si el envío fue exitoso
        else:
            # Guarda qué cotización quedó solicitada
            st.session_state["firma_solicitada"] = firma
            # Pide celebrar al recargar la página
            st.session_state["celebrar"] = True
            # Recarga la página para cerrar la ventana y mostrar la confirmación
            st.rerun()


# Muestra el panel con el avance, el total y las acciones de la cotización
def seccion_resumen(datos, conceptos):
    # Escribe el título del panel
    st.markdown('<div class="resumen-titulo">Tu cotización</div>', unsafe_allow_html=True)
    # Agrega el flete al final, con su costo por confirmar
    renglones = conceptos + [RENGLON_FLETE]
    # Calcula el total de la cotización sin el flete
    total = total_cotizacion(renglones)
    # Revisa si ya hay conceptos para resumir
    if conceptos:
        # Muestra los conceptos y el total
        mostrar_importes(conceptos, total)
    # Invita a empezar si aún no hay conceptos
    else:
        # Muestra el mensaje inicial
        st.markdown('<div class="resumen-vacio">Elige un paquete o artículo para ver aquí tu total.</div>', unsafe_allow_html=True)
    # Calcula qué datos obligatorios siguen pendientes
    faltantes = datos_faltantes(datos, bool(conceptos))
    # Bloquea el PDF mientras falte algún dato
    if faltantes:
        # Calcula cuántos datos obligatorios ya están listos
        listos = TOTAL_OBLIGATORIOS - len(faltantes)
        # Muestra el avance del cliente
        st.progress(listos / TOTAL_OBLIGATORIOS, text=f"Tu avance: {listos} de {TOTAL_OBLIGATORIOS}")
        # Indica qué completar para continuar
        st.info("Para obtener tu PDF completa: " + ", ".join(faltantes) + ".")
        # Termina aquí mientras haya pendientes
        return
    # Limpia el nombre para usarlo en el archivo
    nombre_limpio = re.sub(r"[^\w]+", "_", datos["nombre"]).strip("_")
    # Arma el nombre del archivo
    nombre_archivo = f"Cotizacion_Valinna_{nombre_limpio}.pdf"
    # Identifica esta cotización para detectar si cambia después de solicitarla
    firma = repr((datos, renglones))
    # Muestra la confirmación si esta cotización ya se solicitó
    if st.session_state.get("firma_solicitada") == firma:
        # Celebra una sola vez, justo después de solicitar
        if st.session_state.pop("celebrar", False):
            # Lanza los globos en la página
            st.balloons()
        # Avisa que lo contactarán al WhatsApp que dejó
        st.success(mensaje_solicitud(datos["celular"]))
    # Muestra el botón que abre la cotización completa
    if st.button("Ver mi cotización", key="ver_cotizacion", type="primary", width="stretch"):
        # Abre la ventana con la cotización
        ventana_cotizacion(datos, renglones, nombre_archivo, total, firma)


# ---------------------------------------------------------------------------
# Panel de administración: solo para el personal que atiende los pedidos
# ---------------------------------------------------------------------------


# Crea una tabla vacía de conceptos con los tipos de dato correctos
def conceptos_vacios():
    # Devuelve la tabla sin renglones
    return pd.DataFrame({"Concepto": pd.Series(dtype="object"), "Cantidad": pd.Series(dtype="float64"), "Precio unitario": pd.Series(dtype="float64")})


# Crea una tabla vacía de promociones con los tipos de dato correctos
def promociones_vacias():
    # Devuelve la tabla sin renglones
    return pd.DataFrame({"Descripción": pd.Series(dtype="object"), "Tipo": pd.Series(dtype="object"), "Valor": pd.Series(dtype="float64")})


# Revisa la contraseña del personal y devuelve si tiene acceso
def acceso_administracion():
    # Permite el paso si ya inició sesión
    if st.session_state.get("admin_autorizado"):
        # Confirma el acceso
        return True
    # Intenta leer la contraseña guardada en los secretos
    try:
        # Obtiene la contraseña de administración
        clave = str(secreto("CONTRASENA_ADMIN"))
    # Atiende la falta del secreto
    except ValueError as error:
        # Avisa qué falta configurar
        st.error(str(error))
        # Niega el acceso
        return False
    # Crea el formulario de acceso
    with st.form("acceso_admin"):
        # Pide la contraseña sin mostrarla
        intento = st.text_input("Contraseña de administración", type="password")
        # Muestra el botón para entrar
        entrar = st.form_submit_button("Entrar", type="primary")
    # Valida la contraseña al enviar el formulario
    if entrar:
        # Compara la contraseña sin filtrar información por el tiempo de respuesta
        if hmac.compare_digest(intento.encode("utf-8"), clave.encode("utf-8")):
            # Guarda que la sesión está autorizada
            st.session_state["admin_autorizado"] = True
            # Recarga para mostrar el panel
            st.rerun()
        # Avisa que la contraseña no coincide
        st.error("Contraseña incorrecta.")
    # Niega el acceso mientras no se valide
    return False


# Cierra la sesión de administración
def cerrar_sesion_admin():
    # Borra la autorización de la sesión
    st.session_state.pop("admin_autorizado", None)


# Lee el archivo JSON del pedido, ya sea el del correo o uno guardado
def leer_pedido(contenido):
    # Convierte el texto JSON y lo prepara para el panel
    return pedido_desde_dict(json.loads(contenido.decode("utf-8")))


# Convierte un pedido en formato JSON en los valores del panel
def pedido_desde_dict(pedido):
    # Convierte los conceptos en la tabla editable
    conceptos = pd.DataFrame([{"Concepto": str(c["descripcion"]), "Cantidad": float(c["cantidad"]), "Precio unitario": float(c["precio"])} for c in pedido["conceptos"]], columns=COLUMNAS_PEDIDO)
    # Convierte las promociones guardadas en la tabla editable
    promociones = pd.DataFrame([{"Descripción": str(p["descripcion"]), "Tipo": p["tipo"] if p["tipo"] in TIPOS_PROMOCION else TIPOS_PROMOCION[0], "Valor": float(p["valor"])} for p in pedido.get("promociones", [])], columns=COLUMNAS_PROMOCION)
    # Convierte los pagos guardados, validando cada campo
    pagos = [{"numero": int(p["numero"]), "recibo": str(p["recibo"]), "fecha": date.fromisoformat(p["fecha"]).isoformat(), "tipo": str(p["tipo"]), "monto": float(p["monto"]), "forma": str(p["forma"]), "referencia": str(p.get("referencia", "")), "total": float(p["total"])} for p in pedido.get("pagos", [])]
    # Devuelve los datos listos para los campos del panel
    return {
        # Entrega las promociones, o una tabla vacía si no hay
        "promociones": promociones if not promociones.empty else promociones_vacias(),
        # Entrega el costo del flete, o cero si el pedido viene del correo
        "flete": float(pedido.get("flete", 0)),
        # Entrega si el cliente requiere factura
        "con_iva": bool(pedido.get("con_iva", False)),
        # Entrega la nota para el cliente
        "notas": str(pedido.get("notas", ""))[:LARGO_NOTA],
        # Entrega los pagos registrados
        "pagos": pagos,
        # Conserva el folio original
        "folio": str(pedido.get("folio", "")),
        # Toma el nombre del cliente
        "nombre": str(pedido["nombre"]),
        # Toma el celular del cliente
        "celular": str(pedido["celular"]),
        # Toma la dirección del evento
        "direccion": str(pedido["direccion"]),
        # Convierte la fecha del evento
        "fecha_evento": date.fromisoformat(pedido["fecha_evento"]),
        # Convierte la hora del evento
        "hora_evento": time.fromisoformat(pedido["hora_evento"]),
        # Entrega los conceptos, o una tabla vacía si no trae ninguno
        "conceptos": conceptos if not conceptos.empty else conceptos_vacios(),
    # Cierra los datos del pedido
    }


# Coloca el pedido leído en los campos del panel
def aplicar_pedido(pedido):
    # Recorre los campos simples del pedido
    for campo in ("folio", "nombre", "celular", "direccion", "fecha_evento", "hora_evento"):
        # Guarda el valor en el campo correspondiente del panel
        st.session_state[f"adm_{campo}"] = pedido[campo]
    # Guarda los conceptos como base de la tabla editable
    st.session_state["adm_conceptos"] = pedido["conceptos"]
    # Guarda las promociones como base de su tabla editable
    st.session_state["adm_promociones"] = pedido["promociones"]
    # Coloca el costo del flete
    st.session_state["adm_flete"] = pedido["flete"]
    # Coloca si el cliente requiere factura
    st.session_state["adm_iva"] = pedido["con_iva"]
    # Coloca la nota para el cliente
    st.session_state["adm_notas"] = pedido["notas"]
    # Coloca los pagos registrados
    st.session_state["adm_pagos"] = pedido["pagos"]
    # Marca que el pedido aún no coincide con uno guardado
    st.session_state["adm_huella_guardada"] = None
    # Cambia la versión de las tablas para que muestren los datos nuevos
    st.session_state["adm_version"] = st.session_state.get("adm_version", 0) + 1
    # Cambia también la versión de la tabla de promociones
    st.session_state["adm_version_promo"] = st.session_state.get("adm_version_promo", 0) + 1


# Deja el panel en blanco para capturar un pedido a mano
def nuevo_pedido():
    # Aplica un pedido vacío a todos los campos
    aplicar_pedido({"folio": "", "nombre": "", "celular": "", "direccion": "", "fecha_evento": None, "hora_evento": None, "conceptos": conceptos_vacios(), "promociones": promociones_vacias(), "flete": 0.0, "con_iva": False, "notas": "", "pagos": []})
    # Olvida el último archivo cargado para poder cargarlo de nuevo
    st.session_state.pop("adm_archivo_cargado", None)


# Prepara los valores iniciales del panel la primera vez que se abre
def iniciar_panel():
    # Revisa si el panel ya se había iniciado con todos sus valores
    if "adm_pagos" not in st.session_state:
        # Deja todos los campos en blanco
        nuevo_pedido()


# Arma la ruta del archivo donde se guarda un pedido
def ruta_pedido(folio):
    # Usa el folio como nombre de archivo, quitando caracteres no válidos
    return CARPETA_PEDIDOS / f"{re.sub(r'[^\w-]+', '_', folio)}.json"


# Lista los pedidos guardados, del más reciente al más antiguo
def pedidos_guardados():
    # Regresa una lista vacía si aún no existe la carpeta
    if not CARPETA_PEDIDOS.is_dir():
        # No hay pedidos guardados
        return []
    # Prepara la lista de pedidos
    lista = []
    # Recorre los archivos ordenados por fecha de modificación
    for ruta in sorted(CARPETA_PEDIDOS.glob("*.json"), key=lambda archivo: archivo.stat().st_mtime, reverse=True):
        # Intenta leer el pedido
        try:
            # Convierte el archivo en diccionario
            pedido = json.loads(ruta.read_text(encoding="utf-8"))
        # Omite archivos dañados
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            # Pasa al siguiente archivo
            continue
        # Suma lo pagado
        pagado = sum(float(pago.get("monto", 0)) for pago in pedido.get("pagos", []))
        # Toma el total guardado
        total = float(pedido.get("total", 0))
        # Describe el estado de pago del pedido
        estado = "Liquidado" if total and pagado >= total - 0.005 else f"Saldo {moneda(total - pagado)}"
        # Agrega el pedido con una etiqueta descriptiva
        lista.append((ruta.name, f"{pedido.get('folio', '')} · {pedido.get('nombre', '')} · evento {pedido.get('fecha_evento', '')} · {estado}"))
    # Devuelve los pedidos encontrados
    return lista


# Abre un pedido guardado en el panel
def abrir_guardado(nombre_archivo):
    # Intenta leer el pedido
    try:
        # Lee y convierte el pedido guardado
        pedido = pedido_desde_dict(json.loads((CARPETA_PEDIDOS / nombre_archivo).read_text(encoding="utf-8")))
    # Atiende archivos dañados o con otro formato
    except (OSError, UnicodeDecodeError, json.JSONDecodeError, KeyError, TypeError, ValueError) as error:
        # Registra el detalle técnico
        registro.warning("No se pudo abrir el pedido %s: %s", nombre_archivo, error)
        # Deja el aviso para mostrarlo en pantalla
        st.session_state["adm_error_abrir"] = "No se pudo abrir el pedido guardado."
        # Termina sin cambiar nada
        return
    # Coloca el pedido en los campos del panel
    aplicar_pedido(pedido)
    # Pide marcar el pedido como guardado al terminar de dibujar el panel
    st.session_state["adm_sincronizar"] = True
    # Confirma la apertura
    st.toast(f"Pedido {pedido['folio']} abierto.")


# Muestra cómo abrir, cargar o empezar un pedido
def seccion_cargar_pedido():
    # Escribe el título de la sección
    titulo_paso(1, "Pedido", "Abre uno guardado, carga el .json del correo o captura a mano")
    # Muestra el error de apertura, si lo hubo
    error = st.session_state.pop("adm_error_abrir", None)
    # Revisa si hay error que mostrar
    if error:
        # Muestra el error
        st.error(error)
    # Obtiene los pedidos guardados con su etiqueta
    etiquetas = dict(pedidos_guardados())
    # Crea las columnas del selector y del botón de abrir
    columna_guardado, columna_abrir = st.columns([3, 1], vertical_alignment="bottom")
    # Pide el pedido guardado que se quiere abrir
    seleccion = columna_guardado.selectbox("Pedidos guardados", list(etiquetas), index=None, format_func=lambda nombre: etiquetas.get(nombre, nombre), placeholder="Elige un pedido guardado" if etiquetas else "Aún no hay pedidos guardados", key="adm_guardado_sel")
    # Muestra el botón que abre el pedido elegido
    columna_abrir.button("Abrir", on_click=abrir_guardado, args=(seleccion,), disabled=seleccion is None, width="stretch")
    # Crea las columnas del cargador y del botón de pedido nuevo
    columna_archivo, columna_nuevo = st.columns([3, 1], vertical_alignment="bottom")
    # Pide el archivo del pedido
    archivo = columna_archivo.file_uploader("Archivo del pedido (Pedido_VAL-....json)", type="json", key="adm_archivo")
    # Muestra el botón para empezar un pedido en blanco
    columna_nuevo.button("Pedido nuevo en blanco", on_click=nuevo_pedido, width="stretch")
    # Termina si no hay archivo
    if archivo is None:
        # No hay nada que cargar
        return
    # Calcula una huella del archivo para no cargarlo dos veces
    huella = sha1(archivo.getvalue()).hexdigest()
    # Termina si este archivo ya se cargó
    if st.session_state.get("adm_archivo_cargado") == huella:
        # Evita borrar los ajustes ya hechos
        return
    # Intenta leer el pedido
    try:
        # Convierte el archivo en datos del panel
        pedido = leer_pedido(archivo.getvalue())
    # Atiende archivos dañados o con otro formato
    except (UnicodeDecodeError, json.JSONDecodeError, KeyError, TypeError, ValueError) as error:
        # Registra el detalle técnico
        registro.warning("No se pudo leer el pedido: %s", error)
        # Avisa al personal
        st.error("El archivo no es un pedido válido del cotizador.")
        # Termina sin cambiar nada
        return
    # Coloca el pedido en los campos del panel
    aplicar_pedido(pedido)
    # Recuerda que este archivo ya se cargó
    st.session_state["adm_archivo_cargado"] = huella
    # Confirma la carga
    st.success(f"Pedido {pedido['folio'] or 'sin folio'} de {pedido['nombre']} cargado.")


# Muestra los datos del cliente y del evento, y los devuelve
def seccion_datos_admin():
    # Escribe el título de la sección
    titulo_paso(2, "Cliente y evento")
    # Crea el panel de los datos
    with st.container(key="panel_datos_admin"):
        # Crea las columnas del folio, el nombre y el celular
        columna_folio, columna_nombre, columna_celular = st.columns([1, 2, 1])
        # Pide el folio, que puede conservarse del pedido original
        folio = columna_folio.text_input("Folio", key="adm_folio", max_chars=20, help="Si lo dejas vacío se genera uno nuevo.").strip()
        # Pide el nombre del cliente
        nombre = columna_nombre.text_input("Nombre completo *", key="adm_nombre", max_chars=60).strip()
        # Pide el celular y deja solo sus dígitos
        celular = re.sub(r"\D", "", columna_celular.text_input("Celular (10 dígitos) *", key="adm_celular", max_chars=20))
        # Pide la dirección completa en un solo campo
        direccion = st.text_input("Dirección completa del evento *", key="adm_direccion", max_chars=200).strip()
        # Crea las columnas de la fecha y la hora
        columna_fecha, columna_hora = st.columns(2)
        # Pide la fecha del evento
        fecha_evento = columna_fecha.date_input("Fecha del evento *", key="adm_fecha_evento", format="DD/MM/YYYY")
        # Pide la hora del evento en intervalos de media hora
        hora_evento = columna_hora.time_input("Hora del evento *", key="adm_hora_evento", step=timedelta(minutes=MINUTOS_ENTRE_HORARIOS))
    # Devuelve los datos con la misma forma que usa el PDF
    return {"folio": folio, "nombre": nombre, "celular": celular, "direccion": direccion, "fecha_evento": fecha_evento, "hora_evento": hora_evento, "fecha_cotizacion": date.today()}


# Muestra la tabla editable de conceptos y devuelve la versión editada
def seccion_conceptos_admin(catalogo):
    # Escribe el título de la sección
    titulo_paso(3, "Conceptos y adecuaciones", "Edita nombre, cantidad o precio; agrega o borra renglones")
    # Toma la versión actual de las tablas
    version = st.session_state["adm_version"]
    # Muestra la tabla editable de conceptos
    editados = st.data_editor(
        # Usa como base los conceptos guardados
        st.session_state["adm_conceptos"],
        # Cambia la clave al cargar otro pedido para refrescar la tabla
        key=f"editor_conceptos_{version}",
        # Permite agregar y borrar renglones
        num_rows="dynamic",
        # Usa todo el ancho disponible
        width="stretch",
        # Oculta el índice de pandas
        hide_index=True,
        # Configura cada columna
        column_config={
            # Configura el nombre del concepto, que puede describir la adecuación
            "Concepto": st.column_config.TextColumn("Concepto", required=True, max_chars=80, help="Ej. Paquete Arte en mesa (sin inflable, +2 sillas)"),
            # Configura la cantidad como número entero
            "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=1, step=1, format="%d", default=1),
            # Configura el precio unitario en pesos
            "Precio unitario": st.column_config.NumberColumn("Precio unitario (MXN)", min_value=0.0, step=50.0, format="$%.2f", default=0.0),
        # Cierra la configuración de columnas
        },
    # Cierra la tabla editable
    )
    # Arma un diccionario con el precio de cada opción del catálogo
    precios = dict(zip(catalogo["Nombre"], catalogo["Precio"]))
    # Crea las columnas para agregar desde el catálogo
    columna_opcion, columna_agregar = st.columns([3, 1], vertical_alignment="bottom")
    # Pide la opción del catálogo que se quiere agregar
    opcion = columna_opcion.selectbox("Agregar del catálogo", list(precios), index=None, placeholder="Elige un paquete o artículo", format_func=lambda nombre: f"{nombre} ({moneda(precios[nombre])})", key=f"adm_opcion_{version}")
    # Agrega la opción al presionar el botón
    if columna_agregar.button("Agregar", disabled=opcion is None, width="stretch"):
        # Crea el renglón con el precio de catálogo
        fila = pd.DataFrame([{"Concepto": opcion, "Cantidad": 1.0, "Precio unitario": float(precios[opcion])}])
        # Une el renglón a la tabla ya editada, sin perder los cambios hechos
        st.session_state["adm_conceptos"] = pd.concat([editados, fila], ignore_index=True)
        # Cambia la versión para mostrar la tabla actualizada
        st.session_state["adm_version"] = version + 1
        # Recarga el panel
        st.rerun()
    # Devuelve la tabla con los cambios del personal
    return editados


# Muestra el flete, las promociones, el IVA y la nota; devuelve los ajustes
def seccion_ajustes_admin():
    # Escribe el título de la sección
    titulo_paso(4, "Flete y promociones")
    # Pide el costo del flete
    flete = st.number_input("Costo del flete (MXN)", key="adm_flete", min_value=0.0, step=50.0, format="%.2f")
    # Suma los pagos registrados, que la cotización final muestra como recibidos
    anticipo = round(sum(pago["monto"] for pago in st.session_state["adm_pagos"]), 2)
    # Explica cómo funcionan las promociones
    st.caption("Promociones y descuentos: el porcentaje se aplica sobre los conceptos (sin flete); el monto fijo se descuenta directo.")
    # Muestra la tabla editable de promociones
    promociones = st.data_editor(
        # Usa como base las promociones guardadas
        st.session_state["adm_promociones"],
        # Usa su propia versión para no perder cambios al agregar conceptos
        key=f"editor_promociones_{st.session_state['adm_version_promo']}",
        # Permite agregar y borrar promociones
        num_rows="dynamic",
        # Usa todo el ancho disponible
        width="stretch",
        # Oculta el índice de pandas
        hide_index=True,
        # Configura cada columna
        column_config={
            # Configura la descripción de la promoción
            "Descripción": st.column_config.TextColumn("Descripción", max_chars=60, default="Promoción", help="Ej. Descuento de temporada"),
            # Configura el tipo de promoción
            "Tipo": st.column_config.SelectboxColumn("Tipo", options=TIPOS_PROMOCION, required=True, default=TIPOS_PROMOCION[0]),
            # Configura el valor de la promoción
            "Valor": st.column_config.NumberColumn("Valor", min_value=0.0, step=5.0, format="%.2f", default=0.0),
        # Cierra la configuración de columnas
        },
    # Cierra la tabla editable
    )
    # Pide si el cliente requiere factura
    con_iva = st.checkbox(f"El cliente requiere factura (agregar IVA {TASA_IVA:.0%})", key="adm_iva")
    # Pide una nota opcional para el cliente
    notas = st.text_area("Nota para el cliente en el PDF (opcional)", key="adm_notas", max_chars=LARGO_NOTA, placeholder=TEXTO_FINAL)
    # Devuelve los ajustes capturados
    return {"flete": flete, "anticipo": anticipo, "promociones": promociones, "con_iva": con_iva, "notas": notas.strip()}


# Convierte un valor de la tabla en texto, tratando los vacíos
def texto_celda(valor):
    # Devuelve texto vacío para celdas vacías y el texto limpio en otro caso
    return "" if pd.isna(valor) else str(valor).strip()


# Convierte la tabla de conceptos en una lista limpia, sin renglones vacíos
def conceptos_limpios(conceptos):
    # Prepara la lista de conceptos
    lista = []
    # Recorre los conceptos editados por el personal
    for fila in conceptos.to_dict("records"):
        # Obtiene el nombre del concepto
        descripcion = texto_celda(fila["Concepto"])
        # Omite los renglones sin concepto
        if not descripcion:
            # Pasa al siguiente renglón
            continue
        # Obtiene la cantidad, con 1 si quedó vacía
        cantidad = 1 if pd.isna(fila["Cantidad"]) else max(1, int(fila["Cantidad"]))
        # Obtiene el precio, con 0 si quedó vacío
        precio = 0.0 if pd.isna(fila["Precio unitario"]) else round(float(fila["Precio unitario"]), 2)
        # Agrega el concepto limpio
        lista.append({"descripcion": descripcion, "cantidad": cantidad, "precio": precio})
    # Devuelve los conceptos limpios
    return lista


# Convierte la tabla de promociones en una lista limpia, solo con las que tienen valor
def promociones_limpias(promociones):
    # Prepara la lista de promociones
    lista = []
    # Recorre las promociones capturadas
    for fila in promociones.to_dict("records"):
        # Omite las promociones sin valor
        if pd.isna(fila["Valor"]) or fila["Valor"] <= 0:
            # Pasa a la siguiente promoción
            continue
        # Usa el porcentaje si el tipo quedó vacío o no es válido
        tipo = fila["Tipo"] if fila["Tipo"] in TIPOS_PROMOCION else TIPOS_PROMOCION[0]
        # Agrega la promoción, con "Promoción" si no tiene descripción
        lista.append({"descripcion": texto_celda(fila["Descripción"]) or "Promoción", "tipo": tipo, "valor": float(fila["Valor"])})
    # Devuelve las promociones limpias
    return lista


# Calcula los renglones y los importes de la cotización final
def calcular_final(conceptos, ajustes):
    # Prepara los renglones del PDF con el importe de cada concepto
    renglones = [{**concepto, "total": round(concepto["cantidad"] * concepto["precio"], 2)} for concepto in conceptos_limpios(conceptos)]
    # Suma los conceptos antes de flete y promociones
    subtotal = total_cotizacion(renglones)
    # Toma el costo del flete
    flete = round(float(ajustes["flete"]), 2)
    # Agrega el flete con su costo ya confirmado
    renglones.append({**RENGLON_FLETE, "precio": flete, "total": flete})
    # Prepara la suma de descuentos
    descuentos = 0.0
    # Prepara la lista de problemas con las promociones
    errores = []
    # Recorre las promociones con valor
    for promocion in promociones_limpias(ajustes["promociones"]):
        # Toma la descripción de la promoción
        descripcion = promocion["descripcion"]
        # Toma el valor capturado
        valor = promocion["valor"]
        # Calcula el descuento por porcentaje
        if promocion["tipo"] == TIPOS_PROMOCION[0]:
            # Revisa que el porcentaje no pase de 100
            if valor > 100:
                # Registra el porcentaje inválido
                errores.append(f"El porcentaje de «{descripcion}» no puede ser mayor a 100.")
            # Calcula el monto sobre los conceptos
            monto = round(subtotal * valor / 100, 2)
            # Muestra el porcentaje en la descripción
            descripcion = f"{descripcion} ({valor:g}%)"
        # Usa el monto fijo tal cual
        else:
            # Toma el monto fijo
            monto = round(valor, 2)
        # Agrega la promoción como renglón negativo
        renglones.append({"descripcion": descripcion, "cantidad": 1, "precio": -monto, "total": -monto})
        # Acumula el descuento
        descuentos += monto
    # Calcula la base después de flete y descuentos
    base = round(subtotal + flete - descuentos, 2)
    # Revisa que los descuentos no dejen el total en negativo
    if base < 0:
        # Registra el problema
        errores.append("Las promociones superan el importe de la cotización.")
    # Calcula el IVA solo si se requiere factura
    iva = round(max(base, 0) * TASA_IVA, 2) if ajustes["con_iva"] else 0.0
    # Agrega el IVA como renglón
    if iva:
        # Agrega el renglón del IVA
        renglones.append({"descripcion": f"IVA ({TASA_IVA:.0%})", "cantidad": 1, "precio": iva, "total": iva})
    # Calcula el total final
    total = total_cotizacion(renglones)
    # Revisa que los pagos registrados no superen el total
    if ajustes["anticipo"] > total + 0.005:
        # Registra el problema
        errores.append("Los pagos registrados superan el total de la cotización.")
    # Devuelve los renglones, los importes y los problemas
    return renglones, {"subtotal": subtotal, "flete": flete, "descuentos": descuentos, "iva": iva, "total": total}, errores


# Revisa los datos obligatorios de la cotización final
def faltantes_admin(datos, renglones):
    # Prepara la lista de pendientes
    faltantes = []
    # Revisa que haya al menos un concepto además del flete
    if not any(renglon["descripcion"] != RENGLON_FLETE["descripcion"] and renglon["total"] >= 0 for renglon in renglones):
        # Registra el concepto pendiente
        faltantes.append("al menos un concepto")
    # Revisa el nombre
    if not datos["nombre"]:
        # Registra el nombre pendiente
        faltantes.append("nombre")
    # Revisa el celular
    if len(datos["celular"]) != 10:
        # Registra el celular pendiente
        faltantes.append("celular de 10 dígitos")
    # Revisa la dirección
    if not datos["direccion"]:
        # Registra la dirección pendiente
        faltantes.append("dirección")
    # Revisa la fecha
    if datos["fecha_evento"] is None:
        # Registra la fecha pendiente
        faltantes.append("fecha del evento")
    # Revisa la hora
    if datos["hora_evento"] is None:
        # Registra la hora pendiente
        faltantes.append("hora del evento")
    # Devuelve los pendientes
    return faltantes


# Muestra en una ventana la cotización final
@st.dialog("Cotización final", width="large")
def ventana_final(pdf, nombre_archivo):
    # Recorre las páginas del PDF
    for pagina in paginas_pdf(pdf):
        # Muestra la página a todo el ancho
        st.image(pagina, width="stretch")
    # Muestra el botón de descarga sin cerrar la ventana
    st.download_button("Descargar PDF final", pdf, file_name=nombre_archivo, mime="application/pdf", on_click="ignore", width="stretch", key="descargar_final_ventana")


# Arma el pedido completo en el formato con que se guarda
def armar_pedido(datos, conceptos, ajustes, total):
    # Devuelve todos los datos necesarios para reabrir el pedido
    return {
        # Guarda el folio
        "folio": datos["folio"],
        # Guarda el nombre del cliente
        "nombre": datos["nombre"],
        # Guarda el celular del cliente
        "celular": datos["celular"],
        # Guarda la dirección del evento
        "direccion": datos["direccion"],
        # Guarda la fecha del evento en formato AAAA-MM-DD
        "fecha_evento": datos["fecha_evento"].isoformat(),
        # Guarda la hora del evento en formato HH:MM
        "hora_evento": datos["hora_evento"].strftime("%H:%M"),
        # Guarda los conceptos limpios
        "conceptos": conceptos_limpios(conceptos),
        # Guarda las promociones con valor
        "promociones": promociones_limpias(ajustes["promociones"]),
        # Guarda el costo del flete
        "flete": round(float(ajustes["flete"]), 2),
        # Guarda si se requiere factura
        "con_iva": bool(ajustes["con_iva"]),
        # Guarda la nota para el cliente
        "notas": ajustes["notas"],
        # Guarda los pagos registrados
        "pagos": st.session_state["adm_pagos"],
        # Guarda el total para mostrar el estado en la lista de pedidos
        "total": total,
    # Cierra el pedido
    }


# Calcula una huella del pedido para saber si cambió desde que se guardó
def huella_pedido(pedido):
    # Devuelve la huella del contenido ordenado
    return sha1(json.dumps(pedido, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()


# Asegura que el pedido tenga folio, generándolo si está vacío
def asegurar_folio(datos, renglones):
    # Genera el folio solo si falta
    if not datos["folio"]:
        # Calcula un folio nuevo a partir del contenido
        datos["folio"] = folio_cotizacion(limpiar_datos(datos), renglones)
        # Lo coloca en el campo del panel
        st.session_state["adm_folio"] = datos["folio"]
    # Devuelve el folio
    return datos["folio"]


# Guarda el pedido en su archivo; se usa desde los botones del panel
def guardar_pedido(datos, conceptos, ajustes, renglones, total):
    # Asegura que exista el folio, que da nombre al archivo
    asegurar_folio(datos, renglones)
    # Arma el pedido completo
    pedido = armar_pedido(datos, conceptos, ajustes, total)
    # Crea la carpeta de pedidos si aún no existe
    CARPETA_PEDIDOS.mkdir(exist_ok=True)
    # Escribe el pedido con la fecha de su última actualización
    ruta_pedido(datos["folio"]).write_text(json.dumps({**pedido, "actualizado": datetime.now().isoformat(timespec="seconds")}, ensure_ascii=False, indent=2), encoding="utf-8")
    # Recuerda la huella de lo guardado
    st.session_state["adm_huella_guardada"] = huella_pedido(pedido)


# Guarda el pedido al presionar el botón Guardar
def guardar_desde_boton(datos, conceptos, ajustes, renglones, total):
    # Intenta guardar el pedido
    try:
        # Escribe el archivo del pedido
        guardar_pedido(datos, conceptos, ajustes, renglones, total)
    # Atiende fallas al escribir el archivo
    except OSError as error:
        # Registra el detalle técnico
        registro.warning("No se pudo guardar el pedido: %s", error)
        # Deja el aviso para mostrarlo en pantalla
        st.session_state["adm_error_guardar"] = "No se pudo guardar el pedido. Revisa que la carpeta no esté bloqueada."
        # Termina sin confirmar
        return
    # Confirma que se guardó
    st.toast(f"Pedido {datos['folio']} guardado.")


# Registra un pago con los datos del formulario y guarda el pedido
def registrar_pago(datos, conceptos, ajustes, renglones, total):
    # Copia los pagos actuales para no modificarlos si algo falla
    pagos = list(st.session_state["adm_pagos"])
    # Suma lo pagado hasta ahora
    pagado = round(sum(pago["monto"] for pago in pagos), 2)
    # Calcula el saldo pendiente
    saldo = round(total - pagado, 2)
    # Toma el tipo de pago elegido
    tipo = st.session_state["pago_tipo"]
    # Toma el monto capturado
    monto = round(float(st.session_state["pago_monto"]), 2)
    # Calcula el monto automático cuando se dejó en cero
    if monto == 0:
        # La liquidación cubre todo el saldo y el anticipo completa el 50%
        monto = saldo if tipo == "Liquidación" else round(max(total / 2 - pagado, 0), 2) if tipo == "Anticipo" else 0.0
    # Revisa que el monto sea positivo
    if monto <= 0:
        # Deja el aviso para mostrarlo en pantalla
        st.session_state["adm_error_pago"] = "Captura un monto mayor a cero."
        # Termina sin registrar
        return
    # Revisa que el monto no supere el saldo
    if monto > saldo + 0.005:
        # Deja el aviso para mostrarlo en pantalla
        st.session_state["adm_error_pago"] = f"El monto supera el saldo pendiente ({moneda(saldo)})."
        # Termina sin registrar
        return
    # Asegura que el pedido tenga folio para numerar el recibo
    folio = asegurar_folio(datos, renglones)
    # Calcula el siguiente número de recibo
    numero = max((pago["numero"] for pago in pagos), default=0) + 1
    # Agrega el pago con todos sus datos
    pagos.append({
        # Guarda el número consecutivo
        "numero": numero,
        # Guarda el folio del recibo
        "recibo": f"{folio}-R{numero}",
        # Guarda la fecha del pago
        "fecha": st.session_state["pago_fecha"].isoformat(),
        # Guarda el tipo de pago
        "tipo": tipo,
        # Guarda el importe
        "monto": monto,
        # Guarda la forma de pago
        "forma": st.session_state["pago_forma"],
        # Guarda la referencia sin caracteres que el PDF no pueda dibujar
        "referencia": texto_imprimible(st.session_state["pago_referencia"]),
        # Guarda el total vigente para que el recibo no cambie después
        "total": total,
    # Cierra el pago
    })
    # Guarda la lista de pagos en la sesión
    st.session_state["adm_pagos"] = pagos
    # Intenta guardar el pedido con el pago nuevo
    try:
        # Escribe el archivo del pedido
        guardar_pedido(datos, conceptos, ajustes, renglones, total)
    # Atiende fallas al escribir el archivo
    except OSError as error:
        # Quita el pago que no se pudo guardar
        st.session_state["adm_pagos"] = pagos[:-1]
        # Registra el detalle técnico
        registro.warning("No se pudo guardar el pago: %s", error)
        # Deja el aviso para mostrarlo en pantalla
        st.session_state["adm_error_pago"] = "No se pudo guardar el pago. Intenta de nuevo."
        # Termina sin confirmar
        return
    # Limpia el monto del formulario
    st.session_state["pago_monto"] = 0.0
    # Limpia la referencia del formulario
    st.session_state["pago_referencia"] = ""
    # Confirma el registro
    st.toast(f"Pago {folio}-R{numero} por {moneda(monto)} registrado.")


# Elimina un pago registrado y guarda el pedido
def eliminar_pago(numero, datos, conceptos, ajustes, renglones, total):
    # Guarda los pagos anteriores por si falla el guardado
    anteriores = st.session_state["adm_pagos"]
    # Quita el pago indicado
    st.session_state["adm_pagos"] = [pago for pago in anteriores if pago["numero"] != numero]
    # Intenta guardar el pedido sin ese pago
    try:
        # Escribe el archivo del pedido
        guardar_pedido(datos, conceptos, ajustes, renglones, total)
    # Atiende fallas al escribir el archivo
    except OSError as error:
        # Recupera los pagos anteriores
        st.session_state["adm_pagos"] = anteriores
        # Registra el detalle técnico
        registro.warning("No se pudo eliminar el pago: %s", error)
        # Deja el aviso para mostrarlo en pantalla
        st.session_state["adm_error_pago"] = "No se pudo eliminar el pago. Intenta de nuevo."
        # Termina sin confirmar
        return
    # Confirma la eliminación
    st.toast("Pago eliminado.")


# Muestra los pagos registrados, sus recibos y el formulario para registrar otro
def seccion_pagos_admin(datos, conceptos, ajustes, renglones, importes, listo):
    # Escribe el título de la sección
    titulo_paso(5, "Pagos y recibos", "Anticipos, abonos y liquidación")
    # Toma los pagos registrados
    pagos = st.session_state["adm_pagos"]
    # Toma el total de la cotización
    total = importes["total"]
    # Suma lo pagado
    pagado = round(sum(pago["monto"] for pago in pagos), 2)
    # Crea tres columnas para los indicadores
    columna_total, columna_pagado, columna_saldo = st.columns(3)
    # Muestra el total
    columna_total.metric("Total", moneda(total))
    # Muestra lo pagado
    columna_pagado.metric("Pagado", moneda(pagado))
    # Muestra el saldo
    columna_saldo.metric("Saldo", moneda(max(total - pagado, 0)))
    # Muestra el error del último pago, si lo hubo
    error = st.session_state.pop("adm_error_pago", None)
    # Revisa si hay error que mostrar
    if error:
        # Muestra el error
        st.error(error)
    # Prepara la suma de pagos anteriores a cada recibo
    anteriores = 0.0
    # Recorre los pagos en el orden en que se registraron
    for pago in pagos:
        # Crea las columnas del pago
        columna_dato, columna_monto, columna_recibo, columna_quitar = st.columns([4, 2, 2, 1], vertical_alignment="center")
        # Muestra el recibo, el tipo, la fecha y la forma de pago
        columna_dato.markdown(f"**{pago['recibo']}** · {pago['tipo']}  \n{fecha_larga(date.fromisoformat(pago['fecha']))} · {pago['forma']}")
        # Muestra el importe
        columna_monto.markdown(f"**{moneda(pago['monto'])}**")
        # Ofrece el recibo solo cuando la cotización está completa
        if listo:
            # Muestra el botón para descargar el recibo
            columna_recibo.download_button("Recibo PDF", generar_recibo(datos, pago, anteriores, total), file_name=f"Recibo_{pago['recibo']}.pdf", mime="application/pdf", on_click="ignore", key=f"recibo_{pago['numero']}", width="stretch")
            # Pide confirmar antes de eliminar el pago
            with columna_quitar.popover("✕", help=f"Eliminar {pago['recibo']}"):
                # Muestra el botón que elimina el pago
                st.button("Eliminar pago", key=f"eliminar_pago_{pago['numero']}", type="primary", on_click=eliminar_pago, args=(pago["numero"], datos, conceptos, ajustes, renglones, total))
        # Acumula el pago para el siguiente recibo
        anteriores += pago["monto"]
    # Avisa que falta completar la cotización antes de registrar pagos
    if not listo:
        # Muestra el aviso
        st.info("Completa la cotización (ver el panel de la derecha) para registrar pagos y descargar recibos.")
        # Termina sin mostrar el formulario
        return
    # Avisa cuando ya no hay saldo
    if pagado >= total - 0.005:
        # Muestra la confirmación de liquidación
        st.success("Este pedido ya está liquidado.")
        # Termina sin mostrar el formulario
        return
    # Crea el formulario del pago
    with st.form("form_pago"):
        # Crea las columnas del tipo y el monto
        columna_tipo, columna_cantidad = st.columns(2)
        # Pide el tipo de pago
        columna_tipo.selectbox("Tipo de pago", TIPOS_PAGO, key="pago_tipo")
        # Pide el monto
        columna_cantidad.number_input("Monto (MXN)", min_value=0.0, step=100.0, format="%.2f", key="pago_monto", help="Déjalo en 0: Anticipo registra el 50% del total y Liquidación registra el saldo completo.")
        # Crea las columnas de la fecha y la forma de pago
        columna_fecha, columna_forma = st.columns(2)
        # Pide la fecha del pago
        columna_fecha.date_input("Fecha del pago", key="pago_fecha", format="DD/MM/YYYY")
        # Pide la forma de pago
        columna_forma.selectbox("Forma de pago", FORMAS_PAGO, key="pago_forma")
        # Pide una referencia opcional
        st.text_input("Referencia (opcional)", key="pago_referencia", max_chars=30, placeholder="Ej. folio de transferencia")
        # Muestra el botón que registra el pago y guarda el pedido
        st.form_submit_button("Registrar pago y guardar", type="primary", width="stretch", on_click=registrar_pago, args=(datos, conceptos, ajustes, renglones, total))


# Muestra el resumen de importes y las acciones de la cotización final
def resumen_admin(datos, conceptos, renglones, importes, ajustes, errores):
    # Escribe el título del panel
    st.markdown('<div class="resumen-titulo">Cotización final</div>', unsafe_allow_html=True)
    # Reúne las líneas del resumen
    lineas = [("Conceptos", importes["subtotal"]), ("Flete", importes["flete"]), ("Promociones", -importes["descuentos"]), (f"IVA ({TASA_IVA:.0%})", importes["iva"])]
    # Recorre las líneas del resumen
    for etiqueta, importe in lineas:
        # Muestra la línea con su importe
        st.markdown(f'<div class="resumen-linea"><span>{etiqueta}</span><strong>{moneda(importe)}</strong></div>', unsafe_allow_html=True)
    # Calcula lo que se cobra como anticipo
    anticipo = ajustes["anticipo"] if ajustes["anticipo"] > 0 else round(importes["total"] / 2, 2)
    # Nombra el anticipo según si ya se recibieron pagos
    etiqueta_anticipo = "Pagos recibidos" if ajustes["anticipo"] > 0 else "Reservación (50%)"
    # Muestra el total final
    st.markdown(f'<div class="resumen-total"><span>Total (MXN)</span><strong>{moneda(importes["total"])}</strong></div>', unsafe_allow_html=True)
    # Muestra el anticipo y el saldo
    st.markdown(f'<div class="resumen-reserva">{etiqueta_anticipo}: <strong>{moneda(anticipo)}</strong> · Saldo: <strong>{moneda(importes["total"] - anticipo)}</strong></div>', unsafe_allow_html=True)
    # Muestra los problemas que impiden generar el PDF
    for error in errores:
        # Muestra el problema
        st.error(error)
    # Toma la petición de marcar como guardado un pedido recién abierto
    sincronizar = st.session_state.pop("adm_sincronizar", False)
    # Calcula qué datos siguen pendientes
    faltantes = faltantes_admin(datos, renglones)
    # Avisa qué completar
    if faltantes:
        # Muestra los pendientes
        st.info("Para generar el PDF y guardar completa: " + ", ".join(faltantes) + ".")
    # Termina si hay problemas o pendientes
    if errores or faltantes:
        # No se genera el PDF todavía
        return
    # Calcula la huella del pedido tal como está en pantalla
    huella_actual = huella_pedido(armar_pedido(datos, conceptos, ajustes, importes["total"]))
    # Marca como guardado el pedido que se acaba de abrir
    if sincronizar:
        # Recuerda la huella del pedido abierto
        st.session_state["adm_huella_guardada"] = huella_actual
    # Revisa si lo que está en pantalla coincide con lo guardado
    guardado = st.session_state.get("adm_huella_guardada") == huella_actual
    # Reúne los ajustes que cambian el PDF
    final = {"folio": datos["folio"], "anticipo": ajustes["anticipo"], "notas": ajustes["notas"], "con_iva": ajustes["con_iva"]}
    # Genera el PDF de la cotización final
    pdf = generar_pdf(datos, renglones, True, final)
    # Limpia el nombre para usarlo en el archivo
    nombre_limpio = re.sub(r"[^\w]+", "_", datos["nombre"]).strip("_")
    # Arma el nombre del archivo
    nombre_archivo = f"Cotizacion_Final_Valinna_{nombre_limpio}.pdf"
    # Muestra el botón para revisar el PDF
    if st.button("Ver cotización final", type="primary", width="stretch"):
        # Abre la ventana con el PDF
        ventana_final(pdf, nombre_archivo)
    # Muestra el botón de descarga directa
    st.download_button("Descargar PDF final", pdf, file_name=nombre_archivo, mime="application/pdf", on_click="ignore", width="stretch")
    # Muestra el botón para guardar el pedido
    st.button("Guardar pedido", key="guardar_pedido", width="stretch", on_click=guardar_desde_boton, args=(datos, conceptos, ajustes, renglones, importes["total"]))
    # Muestra el error al guardar, si lo hubo
    error = st.session_state.pop("adm_error_guardar", None)
    # Revisa si hay error que mostrar
    if error:
        # Muestra el error
        st.error(error)
    # Informa si el pedido está guardado o tiene cambios pendientes
    if guardado:
        # Confirma que no hay cambios pendientes
        st.caption(f"✓ Guardado en {CARPETA_PEDIDOS.name}/{ruta_pedido(datos['folio']).name}")
        # Toma la ruta del archivo guardado
        ruta = ruta_pedido(datos["folio"])
        # Ofrece descargar un respaldo si el archivo existe
        if ruta.is_file():
            # Muestra el botón del respaldo, que se puede volver a cargar en el paso 1
            st.download_button("Descargar respaldo (.json)", ruta.read_bytes(), file_name=ruta.name, mime="application/json", on_click="ignore", width="stretch")
    # Avisa que hay cambios sin guardar
    else:
        # Muestra el aviso
        st.warning("Hay cambios sin guardar.")


# Construye la página del personal que administra los pedidos
def pagina_administracion(catalogo):
    # Escribe el encabezado del panel
    st.markdown('<div class="resumen-titulo">Panel de administración de pedidos</div>', unsafe_allow_html=True)
    # Pide la contraseña y termina si no hay acceso
    if not acceso_administracion():
        # Detiene la página hasta iniciar sesión
        return
    # Muestra el botón para cerrar la sesión
    st.button("Cerrar sesión", on_click=cerrar_sesion_admin)
    # Prepara los valores iniciales del panel
    iniciar_panel()
    # Crea la columna de captura y la del resumen
    izquierda, derecha = st.columns([1.75, 1], gap="large")
    # Entra a la columna de captura
    with izquierda:
        # Carga el pedido recibido, antes de dibujar los campos que llena
        seccion_cargar_pedido()
        # Muestra los datos del cliente y del evento
        datos = seccion_datos_admin()
        # Muestra los conceptos editables
        conceptos = seccion_conceptos_admin(catalogo)
        # Muestra el flete y las promociones
        ajustes = seccion_ajustes_admin()
        # Calcula la cotización final
        renglones, importes, errores = calcular_final(conceptos, ajustes)
        # Revisa si la cotización está completa y sin problemas
        listo = not errores and not faltantes_admin(datos, renglones)
        # Muestra los pagos y los recibos
        seccion_pagos_admin(datos, conceptos, ajustes, renglones, importes, listo)
    # Crea el panel de resumen en la columna derecha
    with derecha.container(key="panel_resumen"):
        # Muestra el resumen y las acciones
        resumen_admin(datos, conceptos, renglones, importes, ajustes, errores)


# Construye la página completa
def main():
    # Configura la pestaña y el ancho de la página
    st.set_page_config(page_title="Cotizador Valinna", page_icon="🎈", layout="wide")
    # Aplica el diseño de la página
    aplicar_estilos()
    # Revisa si se abrió la página del personal con ?admin=1 en la dirección
    es_admin = "admin" in st.query_params
    # Muestra la portada solo a los clientes
    if not es_admin:
        # Muestra la portada
        mostrar_portada()
    # Intenta leer el catálogo
    try:
        # Carga el catálogo usando su fecha de modificación
        catalogo = cargar_catalogo(RUTA_CATALOGO.stat().st_mtime)
    # Atiende el caso de un catálogo que no está en la carpeta
    except FileNotFoundError:
        # Avisa cuál archivo falta
        st.error(f"No se encontró el catálogo {RUTA_CATALOGO.name} en la carpeta del cotizador.")
        # Detiene la página
        st.stop()
    # Atiende el caso de un Excel abierto y bloqueado
    except PermissionError:
        # Pide cerrar el archivo
        st.error(f"No se pudo leer {RUTA_CATALOGO.name}. Ciérralo en Excel y recarga la página.")
        # Detiene la página
        st.stop()
    # Atiende los errores de contenido del catálogo
    except ValueError as error:
        # Muestra qué corregir
        st.error(str(error))
        # Detiene la página
        st.stop()
    # Muestra el panel del personal en lugar del cotizador del cliente
    if es_admin:
        # Construye el panel de administración
        pagina_administracion(catalogo)
        # Termina para no mostrar el cotizador del cliente
        return
    # Crea la columna de captura y la del resumen
    izquierda, derecha = st.columns([1.75, 1], gap="large")
    # Entra a la columna de captura
    with izquierda:
        # Muestra las pestañas del catálogo y recibe los conceptos elegidos
        conceptos = seccion_catalogo(catalogo)
        # Muestra el formulario y recibe los datos
        datos = seccion_datos()
    # Crea el panel de resumen en la columna derecha
    with derecha.container(key="panel_resumen"):
        # Muestra el resumen y las acciones
        seccion_resumen(datos, conceptos)


# Ejecuta la página cuando se corre con streamlit
if __name__ == "__main__":
    # Inicia la aplicación
    main()
