import argparse
import csv
import glob
import io
import os
import re
import sys

import pandas as pd
import pdfplumber
from tkinter import Tk, filedialog


def seleccionar_carpeta():
    """Abre una ventana nativa de Windows para seleccionar una carpeta."""
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los PDFs a extraer")
    root.destroy()
    return carpeta


def obtener_archivos_pdf(ruta_entrada):
    ruta_entrada = os.path.expanduser(ruta_entrada)

    if os.path.isdir(ruta_entrada):
        return sorted(glob.glob(os.path.join(ruta_entrada, "*.pdf")))

    if os.path.isfile(ruta_entrada):
        if ruta_entrada.lower().endswith('.pdf'):
            return [ruta_entrada]
        raise FileNotFoundError(f"La ruta especificada existe pero no es un PDF: {ruta_entrada}")

    if '*' in ruta_entrada or '?' in ruta_entrada:
        rutas = sorted(glob.glob(ruta_entrada))
        if rutas:
            return rutas

    raise FileNotFoundError(f"No se encontró la ruta de entrada: {ruta_entrada}")


def parsear_linea_pdf(linea):
    texto = linea.replace('\t', ' ').strip()
    if not texto:
        return []

    if ',' in texto:
        return [celda.strip() for celda in next(csv.reader([texto], delimiter=',', quotechar='"'))]

    if re.search(r'\s{2,}', texto):
        return [celda.strip() for celda in re.split(r'\s{2,}', texto) if celda.strip()]

    return [celda.strip() for celda in texto.split() if celda.strip()]


def extraer_filas_de_pdf(ruta_pdf):
    filas = []
    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            texto_crudo = pagina.extract_text()
            if not texto_crudo:
                continue

            for linea in texto_crudo.splitlines():
                fila = parsear_linea_pdf(linea)
                if len(fila) < 4:
                    continue
                filas.append(fila)

    return filas


def automatizar_pdf_rodelag_v5_2(archivo_inventario, nombre_reporte_salida, carpeta_o_pdf=None):
    if carpeta_o_pdf:
        try:
            archivos_pdf = obtener_archivos_pdf(carpeta_o_pdf)
        except Exception as e:
            print(f"❌ Error de ruta: {e}")
            return
        carpeta_pdfs = os.path.dirname(os.path.abspath(archivos_pdf[0])) if archivos_pdf else os.getcwd()
    else:
        carpeta_pdfs = seleccionar_carpeta()
        if not carpeta_pdfs:
            print("❌ Proceso cancelado: No seleccionaste ninguna carpeta.")
            return
        archivos_pdf = obtener_archivos_pdf(carpeta_pdfs)

    if not archivos_pdf:
        print("❌ ¡Alerta! No encontré archivos .pdf en la ruta indicada.")
        return

    print(f"📂 Carpeta origen usada: {carpeta_pdfs}")
    ruta_completa_salida = nombre_reporte_salida
    if not os.path.isabs(ruta_completa_salida):
        ruta_completa_salida = os.path.join(carpeta_pdfs, nombre_reporte_salida)

    lista_de_filas = []

    for ruta_pdf in archivos_pdf:
        nombre_archivo = os.path.basename(ruta_pdf)
        print(f"📄 Procesando PDF: {nombre_archivo}")

        try:
            filas_pdf = extraer_filas_de_pdf(ruta_pdf)
            for fila in filas_pdf:
                fila_limpia = [celda.replace('\n', ' ').strip() for celda in fila]
                codigo = fila_limpia[0]

                if codigo.isdigit() and len(codigo) == 13:
                    num_parte = fila_limpia[1] if len(fila_limpia) > 1 else ""
                    descripcion = fila_limpia[2] if len(fila_limpia) > 2 else ""
                    cantidad = fila_limpia[3] if len(fila_limpia) > 3 else "0"
                    precio = fila_limpia[4] if len(fila_limpia) > 4 else "0"
                    total = fila_limpia[5] if len(fila_limpia) > 5 else "0"

                    try:
                        cant_val = int(pd.to_numeric(cantidad.replace(',', ''), errors='coerce') or 0)
                        precio_val = float(pd.to_numeric(precio.replace('$', '').replace(',', ''), errors='coerce') or 0)
                        total_val = float(pd.to_numeric(total.replace('$', '').replace(',', ''), errors='coerce') or 0)
                    except Exception:
                        cant_val, precio_val, total_val = 0, 0.0, 0.0

                    lista_de_filas.append({
                        'Source.Name': nombre_archivo,
                        'CODIGO': codigo,
                        'NUMERO DE PARTE': num_parte,
                        'DESCRIPCION': descripcion,
                        'CANTIDAD': cant_val,
                        'PRECIO UNITARIO': precio_val,
                        'TOTAL': total_val
                    })

        except Exception as e:
            print(f"❌ Error procesando {nombre_archivo}: {e}")

    if not lista_de_filas:
        print("❌ Error de mapeo: No se pudo extraer data de los archivos PDF. Verifica los archivos y el formato.")
        return

    df_crudo = pd.DataFrame(lista_de_filas)

    print("🔄 CONSOLIDANDO REGISTROS REPETIDOS...")
    df_consolidado = df_crudo.groupby(
        ['Source.Name', 'CODIGO', 'NUMERO DE PARTE', 'DESCRIPCION'],
        as_index=False
    ).agg({
        'CANTIDAD': 'sum',
        'PRECIO UNITARIO': 'first',
        'TOTAL': 'sum'
    })

    print("📊 Creando la matriz de distribución...")
    df_matriz = df_consolidado.pivot_table(
        index=['CODIGO', 'DESCRIPCION'],
        columns='Source.Name',
        values='CANTIDAD',
        aggfunc='sum'
    ).reset_index().fillna(0)

    columnas_sucursales = [col for col in df_matriz.columns if col not in ['CODIGO', 'DESCRIPCION']]
    df_matriz['Total_Pedido'] = df_matriz[columnas_sucursales].sum(axis=1) if columnas_sucursales else 0

    print("🔍 Cruzando datos con el inventario...")
    try:
        df_inventario = pd.read_excel(archivo_inventario)
        df_inventario['SKU'] = df_inventario['SKU'].astype(str).str.strip()
        df_matriz = pd.merge(
            df_matriz,
            df_inventario[['SKU', 'Stock_Disponible', 'Precio_Sistema']],
            left_on='CODIGO',
            right_on='SKU',
            how='left'
        )
        df_matriz['Diferencia_Stock'] = df_matriz['Stock_Disponible'] - df_matriz['Total_Pedido']

        def generar_alertas(row):
            if pd.isna(row['Stock_Disponible']):
                return '❌ SKU INEXISTENTE'
            if row['Diferencia_Stock'] < 0:
                return f'⚠️ FALTAN {abs(int(row["Diferencia_Stock"]))} UNIDADES'
            return '✅ Stock Disponible'

        df_matriz['Estatus_Inventario'] = df_matriz.apply(generar_alertas, axis=1)
        df_matriz.drop(columns=['SKU'], inplace=True)
    except Exception as e:
        print(f"⚠️ Alerta: Matriz generada sin cruce de inventario. Detalle: {e}")

    print(f"💾 Guardando reporte en: {ruta_completa_salida}")
    with pd.ExcelWriter(ruta_completa_salida, engine='openpyxl') as writer:
        df_consolidado.to_excel(writer, sheet_name='xls', index=False)
        df_matriz.to_excel(writer, sheet_name='Matriz_Validada', index=False)

    print("🎉 ¡Proceso terminado! Revisa el archivo Excel generado.")


def main():
    parser = argparse.ArgumentParser(description='Extrae y consolida datos de PDFs de Rodelag.')
    parser.add_argument('-r', '--ruta', help='Carpeta o ruta al archivo PDF a procesar.')
    parser.add_argument('-i', '--inventario', default='inventario_maestro.xlsx', help='Archivo de inventario maestro (.xlsx).')
    parser.add_argument('-s', '--salida', default='Consolidado_Final_Rodelag.xlsx', help='Nombre o ruta del archivo Excel de salida.')
    args = parser.parse_args()

    automatizar_pdf_rodelag_v5_2(args.inventario, args.salida, args.ruta)


if __name__ == '__main__':
    main()
