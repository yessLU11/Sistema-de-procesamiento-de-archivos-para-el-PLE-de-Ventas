# processor.py
import pandas as pd
import io
import time

# --- REGLAS DE NEGOCIO ---
PREFIX_RULES = {
    "01": "F",
    "03": "B"
}

def normalizar_tipo_doc(valor):
    """Convierte 1 -> 01, 3 -> 03."""
    try:
        return str(int(valor)).zfill(2)
    except:
        return str(valor)

def normalizar_codigo(tipo_doc, codigo_original):
    """Normaliza el código de establecimiento añadiendo prefijos si no existen."""
    if tipo_doc not in PREFIX_RULES:
        return str(codigo_original)

    prefijo_necesario = PREFIX_RULES[tipo_doc]
    codigo_str = str(codigo_original).strip()

    if codigo_str.startswith(prefijo_necesario):
        return codigo_str
    
    return prefijo_necesario + codigo_str

def procesar_excel(file_stream, progress_callback=None):
    """Orquestador principal: Extrae, Normaliza y Detecta Duplicados."""
    start_time = time.perf_counter()
    try:
        excel_file = pd.ExcelFile(file_stream)
        hojas = excel_file.sheet_names
        total_hojas = len(hojas)
        
        df_limpio = pd.DataFrame()
        lista_errores = []
        
        total_filas = 0
        duplicados_encontrados = 0

        for idx, hoja in enumerate(hojas):
            df = pd.read_excel(excel_file, sheet_name=hoja)
            
            if df.empty:
                lista_errores.append(f"Hoja '{hoja}' está vacía.")
                continue

            # Ajustar índices de fila para auditoría (Excel inicia en 1, cabecera es 1)
            df['_ExcelRow'] = df.index + 2 

            # Normalización de Tipo de Documento
            if 'TipoDoc' in df.columns:
                df['TipoDoc_Norm'] = df['TipoDoc'].apply(normalizar_tipo_doc)
            else:
                lista_errores.append(f"Hoja '{hoja}': Falta columna 'TipoDoc'")
                continue

            # Normalización de Código de Establecimiento
            if 'CodigoEstablecimiento' in df.columns:
                df['CodigoEstablecimiento_Norm'] = df.apply(
                    lambda x: normalizar_codigo(x['TipoDoc_Norm'], x['CodigoEstablecimiento']), axis=1
                )
            else:
                lista_errores.append(f"Hoja '{hoja}': Falta columna 'CodigoEstablecimiento'")

            df['Hoja_Nombre'] = hoja
            df_limpio = pd.concat([df_limpio, df], ignore_index=True)

            if progress_callback:
                progress_callback((idx + 1) / max(total_hojas, 1), f"procesando hoja {idx + 1} de {total_hojas}: {hoja}")

        # Detección de duplicados para documentos 01 y 03
        df_solo_validos = df_limpio[df_limpio['TipoDoc_Norm'].isin(['01', '03'])].copy()

        if 'NumeroCorrelativo' in df_solo_validos.columns:
            df_solo_validos['Clave_Unica'] = (
                df_solo_validos['CodigoEstablecimiento_Norm'].astype(str) + '-' + 
                df_solo_validos['NumeroCorrelativo'].astype(str)
            )

            duplicated_mask = df_solo_validos.duplicated(subset=['Clave_Unica'], keep=False)
            df_duplicados = df_solo_validos[duplicated_mask].copy()
            
            df_trazabilidad = df_duplicados[['Hoja_Nombre', '_ExcelRow', 'TipoDoc_Norm', 'CodigoEstablecimiento_Norm', 'NumeroCorrelativo', 'Clave_Unica']].sort_values(by='Clave_Unica')
            
            duplicados_encontrados = len(df_duplicados)
        else:
            df_trazabilidad = pd.DataFrame()
            lista_errores.append("Falta columna 'NumeroCorrelativo' para generar claves.")

        total_filas = len(df_limpio)
        elapsed = time.perf_counter() - start_time
        if progress_callback:
            progress_callback(1.0, f"finalizado en {elapsed:.2f} segundos")

        return {
            "success": True,
            "hojas_procesadas": hojas,
            "total_registros": total_filas,
            "duplicados_count": duplicados_encontrados,
            "df_duplicados": df_trazabilidad,
            "df_limpio": df_limpio, 
            "errores": lista_errores,
            "elapsed_seconds": elapsed
        }

    except Exception as e:
        return {
            "success": False,
            "message": f"Error fatal en el procesamiento: {str(e)}",
            "errores": []
        }
#Pertenece a 3_📅_Ordenar_Boletas.py
# ============================================================================
# NUEVO MÓDULO: ORDENAR BOLETAS POR DÍA
# ============================================================================
def ordenar_boletas(file_stream, progress_callback=None):
    """
    Procesa un archivo Excel de boletas (PLE Ventas) y genera un resumen con columnas fijas.
    """
    start_time = time.perf_counter()
    try:
        # =========================================================================
        # 1. DEFINICIÓN DE LAS COLUMNAS FIJAS (en el orden exacto del requerimiento)
        # =========================================================================
        COLUMNAS_FIJAS = [
            'Tipo', 'Periodo', 'IDComprobante', 'Serie', 'FechaEmision',
            'FechaVencimiento', 'TipoDoc', 'CodigoEstablecimiento', 'NumeroCorrelativo',
            'campo_10', 'TipoDocCliente', 'NumeroDocCliente', 'RazonSocialCliente',
            'MontoOperacionesExoneradas', 'MontoOperacionesGravadas', 'MontoIGV',
            'CodigoCentroCosto', 'MontoIGVRetenido', 'MontoReversiones',
            'MontoOtrosConceptos', 'MontoPercepciones', 'MontoDetraccion',
            'MontoTotal', 'campo_23', 'campo_24', 'campo_25', 'campo_26',
            'campo_27', 'campo_28', 'campo_29', 'campo_30', 'campo_31',
            'campo_32', 'campo_33', 'campo_34', 'campo_35'
        ]
        # =========================================================================
        # 2. LECTURA Y PROCESAMIENTO HOJA POR HOJA (SALTANDO LAS SIN DATOS VÁLIDOS)
        # =========================================================================
        excel_file = pd.ExcelFile(file_stream, engine='openpyxl')
        hojas = excel_file.sheet_names
        total_hojas = len(hojas)
        
        if progress_callback:
            progress_callback(0.05, f'Leyendo {total_hojas} hojas del archivo Excel...')
        
        # Lista para acumular DataFrames válidos
        dfs_validos = []
        hojas_procesadas = 0
        hojas_saltadas = 0
        
        for idx, hoja_nombre in enumerate(hojas):
            try:
                df_hoja = pd.read_excel(excel_file, sheet_name=hoja_nombre, dtype=str)
                
                if df_hoja.empty:
                    hojas_saltadas += 1
                    continue
                
                # Verificar si tiene las columnas críticas
                if 'TipoDoc' not in df_hoja.columns or 'CodigoEstablecimiento' not in df_hoja.columns:
                    hojas_saltadas += 1
                    continue
                
                # Limpiar basura en esta hoja
                basura_mask = df_hoja.apply(lambda row: row.astype(str).str.contains('RESUMEN DE CONVERSIÓN|Unnamed', case=False, na=False).any(), axis=1)
                df_hoja = df_hoja.loc[~basura_mask].copy()
                
                # Filtrar TipoDoc = '03'
                df_hoja = df_hoja[df_hoja['TipoDoc'] == '03'].copy()
                if df_hoja.empty:
                    hojas_saltadas += 1
                    continue
                
                # Filtrar CodigoEstablecimiento que empiece con 'B'
                df_hoja = df_hoja[df_hoja['CodigoEstablecimiento'].str.startswith('B', na=False)].copy()
                if df_hoja.empty:
                    hojas_saltadas += 1
                    continue
                
                dfs_validos.append(df_hoja)
                hojas_procesadas += 1
                
                if progress_callback:
                    progress_callback(0.05 + 0.05 * ((idx + 1) / total_hojas), f'Hoja {idx+1}/{total_hojas}: "{hoja_nombre}" - {len(df_hoja):,} filas válidas')
                
            except Exception as e:
                hojas_saltadas += 1
                continue
        
        # Verificar si se encontraron datos válidos
        if not dfs_validos:
            return {
                'success': False, 
                'error': f'No se encontraron datos válidos en ninguna de las {total_hojas} hojas. Verifica que existan las columnas "TipoDoc" y "CodigoEstablecimiento" y que haya registros TipoDoc="03" con códigos que empiecen con "B".'
            }
        
        # Concatenar SOLO las hojas válidas
        df = pd.concat(dfs_validos, ignore_index=True)
        
        if progress_callback:
            progress_callback(0.12, f'✅ Hojas procesadas: {hojas_procesadas} | Hojas omitidas: {hojas_saltadas}')

        # =========================================================================
        # 3. LIMPIEZA DE FILAS BASURA (RESUMEN DE CONVERSIÓN, UNNAMED, ETC.)
        # =========================================================================
        if progress_callback:
            progress_callback(0.08, 'Limpiando filas de resumen...')
        basura_mask = df.apply(lambda row: row.astype(str).str.contains('RESUMEN DE CONVERSIÓN|Unnamed', case=False, na=False).any(), axis=1)
        df = df.loc[~basura_mask].copy()
        
        
        # ================================================================
        # 🚀 FILTRO TEMPRANO
        # ================================================================
        
        # Guardar el número original de filas ANTES de filtrar
        original_rows = len(df)
        
        if progress_callback:
            progress_callback(0.10, f'Filtrando datos (total original: {original_rows:,} filas)...')
        
        if 'TipoDoc' not in df.columns:
            return {'#verificar columnas necesarias para filtrarsuccess': False, 'error': 'El archivo no tiene la columna "TipoDoc" necesaria para el filtrado.'}
        if 'CodigoEstablecimiento' not in df.columns:
            return {'success': False, 'error': 'El archivo no tiene la columna "CodigoEstablecimiento" necesaria para el filtrado.'}
        # Filtrar TipoDoc = '03'
        df = df[df['TipoDoc'] == '03'].copy()
        if df.empty:
            return {'success': False, 'error': 'No se encontraron registros con TipoDoc = "03"'}
        
        # Filtrar CodigoEstablecimiento que empiece con 'B'
        df = df[df['CodigoEstablecimiento'].str.startswith('B', na=False)].copy()
        if df.empty:
            return {'success': False, 'error': 'No hay códigos de establecimiento que comiencen con "B"'}
        
        # Calcular reducción
        filas_finales = len(df)
        reduccion = (1 - filas_finales / original_rows) * 100 if original_rows > 0 else 0
        
        if progress_callback:
            progress_callback(0.15, f'Después del filtro: {filas_finales:,} filas (reducción del {reduccion:.1f}%)')

        # =========================================================================
        # 4. VERIFICAR QUE EXISTAN LAS COLUMNAS MÍNIMAS NECESARIAS PARA EL PROCESO
        # =========================================================================
        #aqui no es necesario colocar TipoDoc y CodigoEstablecimiento porque ya se filtro por esas columnas, pero si es necesario verificar que existan las columnas NumeroCorrelativo, MontoOtrosConceptos, IDComprobante y Serie para poder realizar los calculos y generar el nuevo excel con las columnas fijas, si no existen esas columnas se debe retornar un error indicando que el archivo no tiene las columnas necesarias para el procesamiento
        columnas_requeridas = ['FechaEmision', 
                               'NumeroCorrelativo', 'MontoOtrosConceptos', 
                               'IDComprobante', 'Serie']
        missing = [c for c in columnas_requeridas if c not in df.columns]
        if missing:
            return {
                'success': False,
                'error': f"El archivo no tiene las columnas requeridas: {', '.join(missing)}"
            }
        
        # =========================================================================
        # 5. FILTROS: TIPODOC = '03' Y CÓDIGO ESTABLECIMIENTO EMPIEZA CON 'B'
        # =========================================================================
        # df = df[df['TipoDoc'] == '03'].copy()
        #if df.empty:
        #    return {'success': False, 'error': 'No se encontraron registros con TipoDoc = "03"'}
        
        #df = df[df['CodigoEstablecimiento'].str.startswith('B', na=False)].copy()
        #if df.empty:
        #    return {'success': False, 'error': 'No hay códigos de establecimiento que comiencen con "B"'}
        
        # =========================================================================
        # 6. CONVERTIR COLUMNAS NUMÉRICAS
        # =========================================================================
        if progress_callback:
            progress_callback(0.20, 'Convirtiendo columnas numéricas...')

        df['NumeroCorrelativo'] = pd.to_numeric(df['NumeroCorrelativo'], errors='coerce')
        df['MontoOtrosConceptos'] = pd.to_numeric(df['MontoOtrosConceptos'], errors='coerce').fillna(0)
        
        # =========================================================================
        # 7. AGRUPAR POR FECHAEMISION Y CODIGOESTABLECIMIENTO
        # =========================================================================
        if progress_callback:
            progress_callback(0.25, 'Agrupando por FechaEmision y CodigoEstablecimiento...')
            
        #ordenar: primero por FechaEmision ascendente, luego por CodigoEstablecimiento ascendente y finalmente por NumeroCorrelativo ascendente para asegurar que el primer correlativo del grupo sea el menor y el último correlativo del grupo sea el mayor
        df_sorted = df.sort_values(['FechaEmision', 'CodigoEstablecimiento', 'NumeroCorrelativo'])
        grupos = df_sorted.groupby(['FechaEmision', 'CodigoEstablecimiento'])
        total_grupos = grupos.ngroups if hasattr(grupos, 'ngroups') else 1
        
        output_rows = []
        id_counter = 1
        serie_counter = 1
        if progress_callback:
            progress_callback(0.30, f'Generando {total_grupos} grupos...')
        for idx, ((fecha, establecimiento), grupo) in enumerate(grupos):
            primera = grupo.iloc[0]
            primer_correlativo = grupo['NumeroCorrelativo'].min()
            ultimo_correlativo = int(grupo['NumeroCorrelativo'].max())
            suma_total = grupo['MontoOtrosConceptos'].sum()
            
            # Generar IDComprobante
            # el id debe ser tal cual esta en el excel original, solo debe tomar los 7 numeros que por ejemplo el id en el excel ingresado es 123-2981234 entonces el id en el excel a generar es 123-2981
            id_original = str(primera['IDComprobante'])
            if '-' in id_original:
                id_parts = id_original.split('-')
                nuevo_id = f"{id_parts[0]}-{id_parts[1][:4]}"# si tengo 123-7354829 entonces el nuevo id es 123-7354
            else:
                nuevo_id = id_original[:7]  # Si no tiene guion, tomar los primeros 7 caracteres

            # Generar Serie
            nueva_serie = f"M123"  # {serie_counter:04d}"#ejemplo de lo que hace serie_counter: M1230001, M1230002, etc.
            serie_counter += 1
            
            # Construir diccionario respetando el orden fijo de columnas
            row_dict = {}
            for col in COLUMNAS_FIJAS:
                if col == 'Tipo':
                    # Esta columna se reescribirá después con el número correlativo global
                    # Por ahora la dejamos como placeholder
                    row_dict[col] = " "
                elif col == 'IDComprobante':
                    row_dict[col] = nuevo_id
                elif col == 'Serie':
                    row_dict[col] = nueva_serie
                elif col == 'NumeroCorrelativo': # aqui debe ir el 1er correlativo del grupo, no el ultimo ejempplo para la fecha 03/05/2026 su numerocorrelativo es 41834397 y en la columna campo_10 debe ir el último número correlativo del grupo	41835032
                    row_dict[col] = primer_correlativo
                elif col == 'campo_10':
                    row_dict[col] = ultimo_correlativo #indica el último número correlativo del grupo
                elif col == 'MontoOtrosConceptos':
                    row_dict[col] = suma_total
                elif col == 'campo_25':
                    row_dict[col] = suma_total
                elif col == 'campo_35':
                    row_dict[col] = 1
                elif col in ['campo_23', 'campo_24']:
                    row_dict[col] = ''   # K, L, M vacías
                elif col == 'NumeroDocCliente':# esta columna siempre debe estar vacía
                    row_dict[col] = ''
                elif col == 'RazonSocialCliente':# esta columna siempre debe estar vacía
                    row_dict[col] = ''
                else:
                    # Para el resto de columnas, tomar el valor de la primera fila del grupo (si existe)
                    row_dict[col] = primera[col] if col in primera else ''
            output_rows.append(row_dict)
            # Incrementar contadores
            id_counter += 1
            serie_counter += 1
            
            #actualizar progreso
            if progress_callback:
                progress_callback(0.20 + 0.70 * ((idx + 1) / max(total_grupos, 1)), f"Agrupando y generando registro {idx + 1} de {total_grupos}")
        
        # =========================================================================
        # 8. CREAR DATAFRAME DE SALIDA Y REESCRIBIR LA COLUMNA 'TIPO'
        # =========================================================================
        if progress_callback:
            progress_callback(0.95, 'Creando DataFrame de salida...')
        
        #df_out = pd.DataFrame(output_rows, columns=COLUMNAS_FIJAS)
        # Reemplazar la columna 'Tipo' con valores correlativos 1, 2, 3, ...
        #df_out['Tipo'] = range(1, len(df_out) + 1)
        
        # =========================================================================
        # 9. ESCRIBIR EXCEL CON UNA HOJA POR CADA CÓDIGO DE ESTABLECIMIENTO
        # =========================================================================
        if progress_callback:
            progress_callback(0.98, 'Escribiendo archivo Excel de salida...')
        
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            #Ordenar los establecimientos alfabeticamente para que las hojas del excel de salida estén ordenadas
            for establecimiento, grupo_out in sorted(df_out.groupby('CodigoEstablecimiento')):
                sheet_name = str(establecimiento)[:31]
                grupo_out.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Aplicar formato a los encabezados (primera fila)
                worksheet = writer.sheets[sheet_name]
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                font = Font(bold=True, color="FFFFFF", size=11)
                border = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                for cell in worksheet[1]:
                    if cell.value:
                        cell.fill = fill
                        cell.font = font
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = border
                
                # Ajustar ancho de columnas
                for col in worksheet.columns:
                    max_len = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value:
                                max_len = max(max_len, len(str(cell.value)))
                        except:
                            pass
                    adjusted_width = min(max_len + 2, 30)
                    worksheet.column_dimensions[col_letter].width = adjusted_width
        
        output_buffer.seek(0)
        elapsed = time.perf_counter() - start_time

        # si quiero mostrar cunatos minutos se demoro debo: 
        elapsed_minutes = elapsed / 60
        if progress_callback:
            progress_callback(1.0, f"finalizado en ({elapsed_minutes:.2f} minutos)")
        
        
        return {
            'success': True,
            'message': 'Procesamiento exitoso ✨... en ({elapsed_minutes:.2f} minutos)',
            'buffer': output_buffer,
            'sheets': list(df_out['CodigoEstablecimiento'].unique()),
            'total_rows': len(df_out),
            'elapsed_seconds': elapsed,
            'original_rows': original_rows,
            'filtered_rows': filas_finales,
            'reduction_percent': reduccion
        }
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        elapsed = time.perf_counter() - start_time
        return {
            'success': False,
            'error': f"Error en procesamiento 🔴: {str(e)}",
            'elapsed_seconds': elapsed
        }