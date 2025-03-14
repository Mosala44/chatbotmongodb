for camion_data in critical_camiones:
        numero = camion_data["numero"]
        params_data = camion_data["params"]

        # Obtener registros históricos separados
        records_m1 = list(datos_collection.find({
            "camion.numero": numero, 
            "motor_1": True
        }).sort([("fecha_analisis", -1), ("numero_muestra", -1)]))
        
        records_m2 = list(datos_collection.find({
            "camion.numero": numero, 
            "motor_2": True
        }).sort([("fecha_analisis", -1), ("numero_muestra", -1)]))
    
        # Verificar si el camión tiene parámetros que generen gráficos
        tiene_graficos = any(
            param.lower() in ["cfe", "fe", "si", "cu", "cr", "ni"] and param_info["estado"] in ["MONITOREO", "ACCIÓN REQUERIDA"]
            for param, param_info in params_data.items()
        )
        # Si no existen parámetros críticos para graficar, se omite la generación de la tabla y gráficos
        if not tiene_graficos:
            continue
    
        # Crear página y título del análisis
        doc.add_page_break()
        titulo = doc.add_paragraph()
        titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Centrar texto
        run = titulo.add_run(f"ANÁLISIS DEL CAMIÓN {numero}")
        run.bold = True
        run.font.size = Pt(14)
    
        # Crear tabla de información de motores (solo se crea si hay gráficos a generar)
        tabla_info = doc.add_table(rows=4, cols=2)
        tabla_info.style = 'Table Grid'
    
        # Motor 1
        celda = tabla_info.cell(0, 0)
        celda.text = "Motor 1"
        celda = tabla_info.cell(1, 0)
        celda.text = f"N° Serie: {motor_info.get(numero, {}).get('MT1', 'N/A')}"
        celda = tabla_info.cell(2, 0)
        celda.text = f"Estado: {camion_data['estado_m1']}"
        celda = tabla_info.cell(3, 0)
        horas_componentes_m1 = (camion_data['horas_componentes_m1']
                                 if camion_data['horas_componentes_m1'] not in ['N/A', 'Sin registro', '0', 'no hay registros']
                                 else 'Actualizar horometro')
        celda.text = f"Horómetro: {horas_componentes_m1} hrs"
    
        # Motor 2
        celda = tabla_info.cell(0, 1)
        celda.text = "Motor 2"
        celda = tabla_info.cell(1, 1)
        celda.text = f"N° Serie: {motor_info.get(numero, {}).get('MT2', 'N/A')}"
        celda = tabla_info.cell(2, 1)
        celda.text = f"Estado: {camion_data['estado_m2']}"
        celda = tabla_info.cell(3, 1)
        horas_componentes_m2 = (camion_data['horas_componentes_m2']
                                 if camion_data['horas_componentes_m2'] not in ['N/A', 'Sin registro', '0', 'no hay registros']
                                 else 'Actualizar horometro')
        celda.text = f"Horómetro: {horas_componentes_m2} hrs"
    
        # Resaltar motor afectado (se selecciona el primero que cumpla la condición en los parámetros)
        motor_afectado = None
        for param, param_info in params_data.items():
            if param_info["estado"] in ["MONITOREO", "ACCIÓN REQUERIDA"]:
                motor_afectado = param_info["motor"]
                break
        if motor_afectado:
            motor_celda = tabla_info.cell(0, motor_afectado - 1)
            estado_celda = tabla_info.cell(2, motor_afectado - 1)
            color = "FFFF00" if camion_data[f"estado_m{motor_afectado}"] == "MONITOREO" else "FF0000"
            set_cell_shading(motor_celda, color)
            set_cell_shading(estado_celda, color)
    
        # Iterar sobre cada parámetro que cumpla la condición para generar gráficos
        for param, param_info in params_data.items():
            if param.lower() in ["cfe", "fe", "si", "cu", "cr", "ni"] and param_info["estado"] in ["MONITOREO", "ACCIÓN REQUERIDA"]:
                motor_origen = param_info["motor"]
    
                # Seleccionar registros según el motor
                if motor_origen == 1:
                    registros_motor = records_m1
                else:
                    registros_motor = records_m2
    
                # Eliminar duplicados de fecha (tomar el registro con el número de muestra más reciente)
                registros_unicos = {}
                for rec in registros_motor:
                    fecha = rec.get("fecha_analisis", "")[:10]
                    if fecha not in registros_unicos or int(rec.get("numero_muestra", 0)) > int(registros_unicos[fecha].get("numero_muestra", 0)):
                        registros_unicos[fecha] = rec
    
                # Ordenar cronológicamente (más antiguo primero)
                registros_ordenados = sorted(
                    registros_unicos.values(),
                    key=lambda r: datetime.strptime(r.get("fecha_analisis", "01-01-2000")[:10], "%d-%m-%Y")
                )
    
                # Tomar últimos 5 registros (ahora estarán ordenados correctamente)
                ultimos_5_registros = registros_ordenados[-5:] if len(registros_ordenados) >= 5 else registros_ordenados
    
                # Extraer fechas y valores
                fechas = [rec.get("fecha_analisis", "")[:10] for rec in ultimos_5_registros]
                valores = [to_float(rec.get(param.lower(), 0)) for rec in ultimos_5_registros]
    
                # Generar gráfico para el parámetro (FE o SI)
                fig, ax = plt.subplots(figsize=(6, 3))
                ax.plot(fechas, valores, marker='o', linewidth=1, label=param.upper())
                
                # Líneas de umbral, si existen
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
    
                # Agregar recomendaciones específicas para el parámetro
                if param.lower() == "si":
                    doc.add_paragraph(
                        "En caso de continuar la condición o incrementar los niveles de las partículas silicio, se recomienda realizar las siguientes acciones:",
                        style="Heading3"
                    )
                    recomendaciones_silicio = [
                        "- Realizar micro filtrado y/o cambio de aceite según plan de mantenimiento.",
                        "- Realizar chequeo de tapa de llenado de aceite de cárter.",
                        "- Realizar chequeo de tapa de inspección planetarios.",
                        "- Realizar chequeo de mangueras y abrazaderas de respiradero.",
                        "- Revisar el estado y la fecha del último cambio del filtro de respiradero del MT.",
                        "- Realizar chequeo y registro fotográfico del tapón magnético del Carter.",
                        "- Realizar chequeo y registro fotográfico del piñón solar y los dientes de los planetarios.",
                        "- Se recomienda realizar metrología y/o cambio de piñón solar por horas de operación.",
                        "- Según inspección realizar cambio del piñón solar por horas de operación.",
                        "- Se recomienda programar una medición de Backlash y EndPlay."
                    ]
                    for rec in recomendaciones_silicio:
                        doc.add_paragraph(rec, style="List Bullet")
                elif param.lower() == "fe":
                    doc.add_paragraph(
                        "En caso de continuar la condición o incrementar los niveles de las partículas fierro o ferromagnéticas, se recomienda realizar las siguientes acciones:",
                        style="Heading3"
                    )
                    recomendaciones_fierro = [
                        "- Realizar micro filtrado y/o cambio de aceite según plan de mantenimiento.",
                        "- Realizar chequeo y registro fotográfico del tapón magnético del Carter.",
                        "- Realizar chequeo y registro fotográfico del piñón solar y los dientes de los planetarios.",
                    ]
                    motor_origen = param_info.get("motor", 1)
                    motor = "Motor 1" if motor_origen == 1 else "Motor 2"
                    # Obtener las horas de operación y el estado según el motor
                    horas_operacion = camion_data.get("horas_componentes_m1", None) if motor == "Motor 1" else camion_data.get("horas_componentes_m2", None)
                    estado_fe = camion_data.get("estado_m1", "Normal") if motor == "Motor 1" else camion_data.get("estado_m2", "Normal")
    
                    if horas_operacion is not None and isinstance(horas_operacion, (int, float)):
                        if horas_operacion >= 8000:
                            recomendaciones_fierro.extend([
                                "- Se recomienda realizar metrología y/o cambio de piñón solar por horas de operación.",
                                "- Según inspección realizar cambio del piñón solar por horas de operación.",
                                "- Se recomienda programar una medición de Backlash y EndPlay."
                            ])
                        elif estado_fe in ["Monitoreo", "Acción Requerida"]:
                            recomendaciones_fierro.extend([
                                "- Según inspección realizar cambio del piñón solar por horas de operación.",
                                "- Se recomienda programar una medición de Backlash y EndPlay."
                            ])
                    
                    for rec in recomendaciones_fierro:
                        doc.add_paragraph(rec, style="List Bullet")
                elif param.lower() == "cfe":
                    doc.add_paragraph(
                        "En caso de continuar la condición o incrementar los niveles de las partículas fierro o ferromagnéticas, se recomienda realizar las siguientes acciones:",
                        style="Heading3"
                    )
                    recomendaciones_cfe= [
                        "- Realizar micro filtrado y/o cambio de aceite según plan de mantenimiento.",
                        "- Realizar chequeo y registro fotográfico del tapón magnético del Carter.",
                        "- Realizar chequeo y registro fotográfico del piñón solar y los dientes de los planetarios.",
                        "- Se recomienda programar una medición de Backlash y EndPlay. "
                    ]
                    for rec in recomendaciones_cfe:
                        doc.add_paragraph(rec, style="List Bullet")
    
            
    
    
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