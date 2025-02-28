from django.shortcuts import render, redirect
from .models import camiones_collection, datos_collection
from django.http import JsonResponse, HttpResponse
from .db_connection import db
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import matplotlib.pyplot as plt
import datetime
from docx.shared import Inches

thresholds = {
    "viscocidad": {"monitoreo": 90, "accion": 90},  # Solo se evalúa acción: >90
    "fe": {"monitoreo": 50, "accion": 100},
    "cu": {"monitoreo": 8, "accion": 15},
    "pb": {"monitoreo": 7, "accion": 10},
    "al": {"monitoreo": 8, "accion": 15},
    "si": {"monitoreo": 30, "accion": 80},
    "na": {"monitoreo": None, "accion": 1}
}


def set_cell_shading(cell, fill_color):
    """
    Establece el color de fondo de una celda.
    fill_color: cadena hexadecimal, ej. "FFFF00" para amarillo.
    """
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), fill_color)
    tcPr = cell._tc.get_or_add_tcPr()
    # Eliminar sombreado previo, si existe
    for child in tcPr.findall(qn('w:shd')):
        tcPr.remove(child)
    tcPr.append(shading_elm)

def set_row_shading(row_cells, fill_color):
    """
    Aplica el sombreado solo a la fila completa.
    """
    for cell in row_cells:
        set_cell_shading(cell, fill_color)


def index(request):
    camiones = datos_collection.find({})
    data = {camiones: "camiones"}
    return render (request, "index.html", data)
# Create your views here.
def create_analisis(data: dict):
    return datos_collection.insert_one(data).inserted_id
def lista_camiones(request):
    """Obtiene todos los camiones y los muestra en la lista."""
    camiones = list(camiones_collection.find({}, {"_id": 0, "numero": 1}))  # Solo traemos el campo 'numero'
    return render(request, "lista.html", {"camiones": camiones})

def create_camion(request):
    """Crea un nuevo camión y redirige a la lista."""
    if request.method == "POST":
        numero_camion = request.POST.get("numero_camion")

        if numero_camion:
            # Verificar si el camión ya existe
            if camiones_collection.find_one({"numero": numero_camion}):
                return redirect("lista_camiones")  # Redirigir a la lista de camiones si ya existe

            # Insertar nuevo camión
            camiones_collection.insert_one({"numero": numero_camion})

            return redirect("lista_camiones")  # Redirigir a la lista de camiones tras agregarlo

    return redirect("lista_camiones")   # Si el formulario no es válido, redirigir igual


from django.http import JsonResponse

PREGUNTAS = [
    "Número de camión: ",
    "Número de muestra: ",
    "Motor_1: ",
    "Motor_2: ",
    "Tipo de pauta: ",
    "Horas componente: ",
    "Fecha análisis (dia-mes-año): ",
    "Cambio de aceite: ",
    "Lubricant: ",
    "Horómetro: ",
    "Fecha de la muestra (dia-mes-año): ",
    "Viscosidad: ",
    "Agua %: ",
    "Cfe: ",
    "Fe: ",
    "Cu: ",
    "Pb: ",
    "Al: ",
    "Sn: ",
    "Ag: ",
    "Cr: ",
    "Ni: ",
    "Mo: ",
    "Ti: ",
    "Si: ",
    "Na: ",
    "K: ",
    "B: ",
    "V: ",
    "Mg: ",
    "Ca: ",
    "P: ",
    "Zn: ",
    "Ba: ",
    "Cd: ",
    "Li: ",
    "Mn: ",
    "Sb: "
]

def selecciona_camion(request):
    """
    Vista para seleccionar un camión antes de iniciar el análisis.
    Se guarda el camión seleccionado en la sesión y se inicializa el diccionario de datos.
    """
    camiones_collection = db["camion"]
    if request.method == "POST":
        camion_id = request.POST.get("camion")
        if camion_id:
            # Buscamos en la colección "camion" usando el campo "numero"
            camion = camiones_collection.find_one({"numero": camion_id})
            if camion:
                # Guardamos en la sesión el número del camión
                request.session["selected_camion"] = camion["numero"]
                # Inicializamos el diccionario de análisis con el camión ya seleccionado
                request.session["analisis_data"] = {"Número de camión:": camion["numero"]}
                print("✅ Camión almacenado en sesión:", camion["numero"])
                return redirect('cht')
            else:
                # Si el camión no existe, se muestra un error
                camiones = list(camiones_collection.find({}))
                return render(request, 'selec_cam.html', {
                    'error': 'El camión seleccionado no existe.',
                    'camiones': camiones
                })
    # Método GET: se muestra la lista de camiones para elegir
    camiones = list(camiones_collection.find({}))
    return render(request, 'selec_cam.html', {'camiones': camiones})


def Chatbot(request): 
    if request.method != "POST":
        return JsonResponse({"error": "Método no permitido"}, status=405)

    # Obtener la entrada del usuario y los datos actuales de la sesión
    user_input = request.POST.get("user_input", "").strip()
    session_data = request.session.get("analisis_data", {})

    # Recuperar el número de camión desde la sesión (ya se guardó en selecciona_camion)
    camion_numero = request.session.get("selected_camion")
    print(f"🔍 Debug - Número de camión en sesión: {camion_numero}")
    print(f"📦 Datos actuales en sesión: {session_data}")

    # Si el camión ya está en la sesión, no es necesario volver a preguntar
    # Por lo tanto, usamos la lista de preguntas a partir de la segunda
    if "Número de camión:" in session_data:
        preguntas = PREGUNTAS[1:]
    else:
        preguntas = PREGUNTAS

    # Calcular el avance en la conversación (restamos 1 si ya se tiene el camión)
    avance = len(session_data) - (1 if "Número de camión:" in session_data else 0)

    # Si aún no se han respondido todas las preguntas, guardamos la respuesta actual
    if avance < len(preguntas):
        session_data[preguntas[avance]] = user_input
        request.session["analisis_data"] = session_data

    # Si se han respondido todas las preguntas, se procesa y se guarda el análisis
    if avance + 1 == len(preguntas):
        try:
            analisis_data = {
                "camion": {"numero": camion_numero},  # Se usa el camión seleccionado
                "numero_muestra": session_data.get("Número de muestra: ", ""),
                "motor_1": session_data.get("Motor_1: ", "false").lower() in ["true", "si", "1"],
                "motor_2": session_data.get("Motor_2: ", "false").lower() in ["true", "si", "1"],
                "tipo_pauta": session_data.get("Tipo de pauta: ", ""),
                "horas_componentes": session_data.get("Horas componente: ", ""),
                "fecha_analisis": session_data.get("Fecha análisis (dia-mes-año): ", ""),
                "cambio_aceite": session_data.get("Cambio de aceite: ", ""),
                "lubricant": session_data.get("Lubricant: ", ""),
                "horometro": session_data.get("Horómetro: ", ""),
                "Fecha_muestra": session_data.get("Fecha de la muestra (dia-mes-año): ", ""),
                "componente": session_data.get("Componente:", ""),
                "viscocidad": session_data.get("Viscosidad: ", ""),
                "agua": session_data.get("Agua %: ", ""),
                "cfe": session_data.get("Cfe: ", ""),
                "fe": session_data.get("Fe: ", ""),
                "cu": session_data.get("Cu: ", ""),
                "pb": session_data.get("Pb: ", ""),
                "al": session_data.get("Al: ", ""),
                "sn": session_data.get("Sn: ", ""),
                "ag": session_data.get("Ag: ", ""),
                "cr": session_data.get("Cr: ", ""),
                "ni": session_data.get("Ni: ", ""),
                "mo": session_data.get("Mo: ", ""),
                "ti": session_data.get("Ti: ", ""),
                "si": session_data.get("Si: ", ""),
                "na": session_data.get("Na: ", ""),
                "k": session_data.get("K: ", ""),
                "b": session_data.get("B: ", ""),
                "v": session_data.get("V: ", ""),
                "mg": session_data.get("Mg: ", ""),
                "ca": session_data.get("Ca: ", ""),
                "p": session_data.get("P: ", ""),
                "zn": session_data.get("Zn: ", ""),
                "ba": session_data.get("Ba: ", ""),
                "cd": session_data.get("Cd: ", ""),
                "li": session_data.get("Li: ", ""),
                "mn": session_data.get("Mn: ", ""),
                "sb": session_data.get("Sb: ", "")
            }

            print("🚚 Datos a guardar en MongoDB:", analisis_data)

            def evaluar_condicion(elemento, valor):
                 rangos = {
                     "viscocidad": (90, 90),  # Mismo valor para monitoreo y acción si no hay rango definido
                     "fe": (50, 100),
                     "cu": (8, 15),
                     "pb": (7, 10),
                     "al": (8, 15),
                     "si": (30, 80),
                     "na": (1, 1),  # Solo tiene NORMAL y ACCIÓN REQUERIDA
                 }
                 if elemento in rangos:
                     monitoreo, accion = rangos[elemento]
                     if valor > accion:
                         return "ACCIÓN REQUERIDA"
                     elif valor > monitoreo:
                         return "MONITOREO"
                 return "NORMAL"
             
             # Evaluar todos los elementos convirtiendo a float
            estados = {elem: evaluar_condicion(elem, float(analisis_data.get(elem, 0) or 0)) for elem in ["viscocidad", "fe", "cu", "pb", "al", "si", "na"]}
             
             # Determinar la condición global del camión
            prioridad = {"NORMAL": 0, "MONITOREO": 1, "ACCIÓN REQUERIDA": 2}
            estado_final = "NORMAL"
            for estado in estados.values():
                 if prioridad[estado] > prioridad[estado_final]:
                     estado_final = estado
             
            
            # Insertar el documento en la colección "datos"
            datos_collection.insert_one(analisis_data)
            
            # Reiniciamos la sesión para la conversación, conservando el camión seleccionado
            request.session["analisis_data"] = {"Número de camión:": camion_numero}
            
            mensaje_final = f"La condición del análisis de este camión es: {estado_final}"
            
            return JsonResponse({"message": f"Datos guardados exitosamente 🎉\n\n{mensaje_final}"})
        except Exception as e:
            return JsonResponse({"error": f"Error al guardar: {str(e)}"}, status=500)
    else:
        # Si aún quedan preguntas, se devuelve la siguiente pregunta
        return JsonResponse({"message": preguntas[avance + 1]})


def to_float(val):
    try:
        return float(val)
    except:
        return 0.0


def evaluar_condicion(elemento, valor):

    rangos = {
        "viscocidad": (90, 90),  # Sin rango de monitoreo, solo NORMAL y ACCIÓN REQUERIDA
        "fe": (50, 100),
        "cu": (8, 15),
        "pb": (7, 10),
        "al": (8, 15),
        "si": (30, 80),
        "na": (1, 1),  # Solo tiene NORMAL y ACCIÓN REQUERIDA
    }
    if elemento in rangos:
        monitoreo, accion = rangos[elemento]
        if valor > accion:
            return "ACCIÓN REQUERIDA"
        elif valor > monitoreo:
            return "MONITOREO"
    return "NORMAL"

def generar_informe(request):
    from docx import Document
    from docx.shared import Inches, Pt
    from django.http import HttpResponse
    import datetime
    import matplotlib.pyplot as plt
    from io import BytesIO

    # Obtener todos los camiones de la colección "camion"
    camiones = list(db["camion"].find({}))
    if not camiones:
        return HttpResponse("No hay camiones registrados.", status=400)
    
    # Eliminar duplicados basados en el campo "numero"
    unique_camiones = {camion["numero"]: camion for camion in camiones}
    camiones = list(unique_camiones.values())
    
    doc = Document()
    
    # Encabezado principal
    table_enc = doc.add_table(rows=2, cols=3)
    table_enc.style = 'Table Grid'
    image_path = "C:/Users/Andres Villarroel/chatbotai2/static/images/image.png"
    cell_imagen = table_enc.cell(0, 0)
    paragraph = cell_imagen.paragraphs[0]
    run = paragraph.add_run()
    run.add_picture(image_path, width=Inches(1))
    table_enc.cell(0, 1).text = "Centro de Reparación de Componentes Antofagasta"
    table_enc.cell(0, 2).text = "ANÁLISIS SEMANAL DE ACEITE"
    table_enc.cell(1, 0).merge(table_enc.cell(1, 2))
    doc.add_paragraph("\n")
    
    # Tabla resumen
    table_datos = doc.add_table(rows=1, cols=8)
    table_datos.style = 'Table Grid'
    hdr_cells = table_datos.rows[0].cells
    encabezados = ['Camión', 'Número de muestra', 'Fecha de Análisis', 'MT1 Horas', 'MT2 Horas', 'Condición', 'Observación', 'Recomendaciones']
    for i, txt in enumerate(encabezados):
        hdr_cells[i].text = txt
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(9)
                run.bold = True

    datos_collection = db["datos"]
    critical_camiones = []
    elementos = ["viscocidad", "fe", "cu", "pb", "al", "si", "na"]
    prioridad = {"NORMAL": 0, "MONITOREO": 1, "ACCIÓN REQUERIDA": 2}

    for camion in camiones:
        numero = camion.get("numero")
        analisis_m1 = datos_collection.find_one({"camion.numero": numero, "motor_1": True}, sort=[("numero_muestra", -1)])
        analisis_m2 = datos_collection.find_one({"camion.numero": numero, "motor_2": True}, sort=[("numero_muestra", -1)])
        
        if not analisis_m1 and not analisis_m2:
            continue

        # Procesar cada motor por separado
        estados_m1 = {}
        estados_m2 = {}
        motor_origen_param = {}
        
        for elem in elementos:
            key = elem.lower()
            val1 = to_float(analisis_m1.get(key, 0)) if analisis_m1 else 0
            val2 = to_float(analisis_m2.get(key, 0)) if analisis_m2 else 0
            
            estados_m1[elem] = evaluar_condicion(elem, val1)
            estados_m2[elem] = evaluar_condicion(elem, val2)
            motor_origen_param[elem] = 1 if val1 >= val2 else 2

        # Determinar estado final por motor
        estado_final_m1 = "NORMAL"
        estado_final_m2 = "NORMAL"
        for estado in estados_m1.values():
            if prioridad[estado.upper()] > prioridad[estado_final_m1.upper()]:
                estado_final_m1 = estado
        for estado in estados_m2.values():
            if prioridad[estado.upper()] > prioridad[estado_final_m2.upper()]:
                estado_final_m2 = estado

        # Datos comunes
        fecha_m1 = analisis_m1.get("fecha_analisis", "") if analisis_m1 else ""
        horometro_m1 = str(analisis_m1.get("horometro", "N/A")) if analisis_m1 else "N/A"
        nro_muestra_m1 = analisis_m1.get("numero_muestra", "") if analisis_m1 else ""
        
        fecha_m2 = analisis_m2.get("fecha_analisis", "") if analisis_m2 else ""
        horometro_m2 = str(analisis_m2.get("horometro", "N/A")) if analisis_m2 else "N/A"
        nro_muestra_m2 = analisis_m2.get("numero_muestra", "") if analisis_m2 else ""
        
        fecha_comb = f"{fecha_m1} / {fecha_m2}" if fecha_m1 and fecha_m2 else (fecha_m1 or fecha_m2)
        nro_muestra_comb = f"{nro_muestra_m1} / {nro_muestra_m2}" if nro_muestra_m1 and nro_muestra_m2 else (nro_muestra_m1 or nro_muestra_m2)
        
        # Determinar estado global
        estado_global = estado_final_m1 if prioridad[estado_final_m1.upper()] > prioridad[estado_final_m2.upper()] else estado_final_m2
        
        # Observaciones
        obs_items = []
        for elem in elementos:
            if estados_m1[elem] != "NORMAL":
                obs_items.append(f"Motor 1: {elem.upper()} ({analisis_m1.get(elem.lower(), 0)})")
            if estados_m2[elem] != "NORMAL":
                obs_items.append(f"Motor 2: {elem.upper()} ({analisis_m2.get(elem.lower(), 0)})")
        
        observacion_global = ", ".join(obs_items)
        recomendaciones = "actualizar horometro"
        
        # Agregar fila a la tabla
        row_cells = table_datos.add_row().cells
        row_cells[0].text = str(numero)
        row_cells[1].text = nro_muestra_comb
        row_cells[2].text = fecha_comb
        row_cells[3].text = horometro_m1
        row_cells[4].text = horometro_m2
        row_cells[5].text = estado_global
        row_cells[6].text = observacion_global
        row_cells[7].text = recomendaciones

        if estado_global.upper() == "MONITOREO":
            set_row_shading(row_cells, "FFFF00")
        elif estado_global.upper() == "ACCIÓN REQUERIDA":
            set_row_shading(row_cells, "FF0000")

        if estado_global.upper() != "NORMAL":
            critical_camiones.append({
                "numero": numero,
                "estado_m1": estado_final_m1,
                "estado_m2": estado_final_m2,
                "horometro_m1": horometro_m1,
                "horometro_m2": horometro_m2,
                "params": {
                    elem: {
                        "motor": motor_origen_param[elem],
                        "estado": estados_m1[elem] if motor_origen_param[elem] == 1 else estados_m2[elem]
                    } 
                    for elem in elementos 
                    if estados_m1[elem] != "NORMAL" or estados_m2[elem] != "NORMAL"
                }
            })

    # Sección de gráficos
    motor_info = {
        "2CAM3080": {"MT1": "W09060763", "MT2": "W09040922"},
        "2CAM3082": {"MT1": "W09070380", "MT2": ""},
        "2CAM3083": {"MT1": "W08040386", "MT2": "W11030045"},
        "2CAM3085": {"MT1": "W11070959", "MT2": "WX15030035"},
        "2CAM3086": {"MT1": "W11050868", "MT2": "W11090094"},
        "2CAM3087": {"MT1": "W09040921", "MT2": ""},
        "2CAM3090": {"MT1": "W06010096", "MT2": "W13010323"},
        "2CAM3091": {"MT1": "W11060945", "MT2": "W08070121"},
        "2CAM3092": {"MT1": "W12070264", "MT2": "Sin placa"}
    }

    plt.rcParams.update({
        'font.size': 8,
        'font.family': 'Calibri',
        'axes.titlesize': 9,
        'axes.labelsize': 8
    })

    for camion_data in critical_camiones:
        numero = camion_data["numero"]
        params_data = camion_data["params"]
        
        analisis_records = list(datos_collection.find({"camion.numero": numero})
                            .sort("fecha_analisis", 1).limit(5))
        if not analisis_records:
            continue
            
        fechas = [rec.get("fecha_analisis", "")[:10] for rec in analisis_records[:5]]
        fechas.reverse()

        for param, param_info in params_data.items():
            motor_origen = param_info["motor"]
            estado_motor = camion_data[f"estado_m{motor_origen}"]
            
            doc.add_page_break()
            tabla_info = doc.add_table(rows=4, cols=2)
            tabla_info.style = 'Table Grid'
            
            # Motor 1
            celda = tabla_info.cell(0, 0)
            celda.text = f"Motor 1"
            celda = tabla_info.cell(1, 0)
            celda.text = f"N° Serie: {motor_info.get(numero, {}).get('MT1', 'N/A')}"
            celda = tabla_info.cell(2, 0)
            celda.text = f"Estado: {camion_data['estado_m1']}"
            celda = tabla_info.cell(3, 0)
            horometro = camion_data['horometro_m1'] if camion_data['horometro_m1'] not in ['N/A', 'Sin registro'] else 'Actualizar horometro'
            celda.text = f"Horómetro: {horometro} hrs"

            # Motor 2
            celda = tabla_info.cell(0, 1)
            celda.text = f"Motor 2"
            celda = tabla_info.cell(1, 1)
            celda.text = f"N° Serie: {motor_info.get(numero, {}).get('MT2', 'N/A')}"
            celda = tabla_info.cell(2, 1)
            celda.text = f"Estado: {camion_data['estado_m2']}"
            celda = tabla_info.cell(3, 1)
            horometro = camion_data['horometro_m2'] if camion_data['horometro_m2'] not in ['N/A', 'Sin registro'] else 'Actualizar horometro'
            celda.text = f"Horómetro: {horometro} hrs"

            # Resaltar motor afectado
            motor_celda = tabla_info.cell(0, motor_origen - 1)
            estado_celda = tabla_info.cell(2, motor_origen - 1)
            color = "FFFF00" if estado_motor == "MONITOREO" else "FF0000"
            set_cell_shading(motor_celda, color)
            set_cell_shading(estado_celda, color)

            # Configurar gráfico
            valores = [to_float(rec.get(param.lower(), 0)) for rec in analisis_records[:5]]
            valores.reverse()

            fig, ax = plt.subplots(figsize=(6, 3))
            ax.plot(fechas, valores, marker='o', linewidth=1, label=param.upper())

            accion = thresholds.get(param.lower(), {}).get("accion")
            monitoreo = thresholds.get(param.lower(), {}).get("monitoreo")
            
            if accion is not None:
                ax.axhline(y=accion, color='red', linestyle='--', label="Acción Requerida")
            if monitoreo is not None:
                ax.axhline(y=monitoreo, color='yellow', linestyle='--', label="Monitoreo")

            ax.set_title(f"Tendencia de {param.upper()} (Motor {motor_origen})")
            ax.legend(loc='upper left', prop={'size': 8})
            ax.grid(True, linestyle=':')
            plt.xticks(rotation=45)
            plt.tight_layout()

            img_stream = BytesIO()
            plt.savefig(img_stream, format='png', dpi=150)
            plt.close()
            doc.add_picture(img_stream, width=Inches(5))

    for row in table_datos.rows[1:]:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(8)

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response['Content-Disposition'] = 'attachment; filename="informe_analisis_aceite.docx"'
    doc.save(response)
    return response