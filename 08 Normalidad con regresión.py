import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import matplotlib.patheffects as pe
from scipy.interpolate import make_interp_spline

valor_umbral_estadistico = None

def estilo_grafico():
    """Devuelve un diccionario con configuraciones de estilo para la gráfica."""
    return {
        "grosor_linea": 2,
        "tamaño_puntos": 3,
        "fontsize_valores": 8.5,
        "negrita_valores": "bold",
        "fontsize_promedios": 10,
        "negrita_promedios": "bold"
    }

def validacion_archivo():
    """
    Busca un único archivo .xlsx en el mismo directorio del script.

    Returns:
        tuple: Ruta completa del archivo y su nombre.
    
    Raises:
        FileNotFoundError: Si no hay archivos .xlsx en la carpeta.
        ValueError: Si hay más de un archivo .xlsx.
    """
    carpeta = os.path.dirname(os.path.abspath(__file__))
    archivos = [f for f in os.listdir(carpeta) if f.endswith(".xlsx")]
    if not archivos:
        raise FileNotFoundError("No se encontró ningún archivo .xlsx en la carpeta.")
    elif len(archivos) > 1:
        raise ValueError(f"Se encontró más de un archivo .xlsx en la carpeta: {archivos}")
    archivo = archivos[0]
    ruta_completa = os.path.join(carpeta, archivo) 
    return ruta_completa, archivo

def validacion_pestaña(archivo, hoja):
    """
    Verifica que la hoja especificada exista en el archivo Excel.

    Args:
        archivo (str): Ruta del archivo.
        hoja (str): Nombre de la hoja a validar.
    
    Raises:
        ValueError: Si la hoja no se encuentra en el archivo.
    """
    archivo_carga = pd.ExcelFile(archivo)
    if hoja not in archivo_carga.sheet_names:
        raise ValueError(f"No se encontró la pestaña '{hoja}' en el archivo.")

def carga_datos(df):
    """
    Extrae y redondea las columnas de 'Mes' y 'Normalidad' desde un DataFrame.

    Args:
        df (DataFrame): DataFrame cargado desde Excel.

    Returns:
        tuple: Lista de meses y lista de valores de normalidad.
    """
    meses = df.iloc[:, 0].astype(str).tolist()
    normalidad = df.iloc[:, 1].round(1).tolist()
    return meses, normalidad

def validacion_columnas(df, columnas_requeridas):
    """
    Verifica que existan todas las columnas necesarias en el DataFrame.
    """
    faltantes = [col for col in columnas_requeridas if col not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas requeridas: {faltantes}")
    return True

def generador_tramos(n_meses):
    """
    Divide los meses en 3 tramos (Q1, Q2, Q3) con colores predefinidos.

    Args:
        n_meses (int): Número total de meses.

    Returns:
        dict: Diccionario con nombre de tramo como clave y tupla con info (inicio, fin, etiqueta, color).
    """
    colores = ["#0033A0", "#FF6F00", "#007E33"]
    tramos = {}
    q_len = n_meses // 3
    for i in range(3):
        inicio = i * q_len
        fin = (i + 1) * q_len if i < 2 else n_meses
        nombre = f"Q{i+1}"
        tramos[nombre] = (inicio, fin, nombre, colores[i])
    return tramos

def regresion_lineal(x, y):
    """
    Calcula la regresión lineal simple sobre un conjunto de puntos.

    Args:
        x (list): Valores de X.
        y (list): Valores de Y.

    Returns:
        tuple: Valores ajustados de X e Y, y la pendiente.
    """
    x_np = np.array(x)
    y_np = np.array(y)
    coef = np.polyfit(x_np, y_np, deg=1)
    x_fit = np.linspace(min(x), max(x), 100)
    y_fit = np.polyval(coef, x_fit)
    return x_fit, y_fit, coef[0]

def umbral_estadistico(pendientes, k=0.5):
    """
    Calcula el umbral de estabilidad basado en la desviación estándar de las pendientes
    de regresión lineal obtenidas para cada tramo (trimestre).

    Este umbral sirve para clasificar si una tendencia en los datos es considerada
    estable, creciente o decreciente. La lógica es la siguiente:

    - Si la pendiente de un tramo es menor o igual al umbral calculado (en valor absoluto),
      se considera una tendencia estable.
    - Si la pendiente supera positivamente el umbral, se considera una tendencia creciente.
    - Si la pendiente supera negativamente el umbral, se considera una tendencia decreciente.

    Parámetros:

    -Lista de pendientes de regresión lineal por tramo (ej. Q1, Q2, Q3).
    -Factor de sensibilidad que ajusta qué tan estricto es el umbral.
        - Valores bajos (ej. 0.1) hacen que casi cualquier cambio sea interpretado como tendencia.
        - Valores altos (ej. 1.0) hacen que solo cambios grandes se clasifiquen como tendencia.
        - k = 0.5 representa un equilibrio entre sensibilidad y tolerancia.
    """
    global valor_umbral_estadistico
    std = np.std(pendientes)
    valor_umbral_estadistico = k * std
    print(f"\n>>> Umbral estadístico calculado: {valor_umbral_estadistico:.4f}")
    print(f"Pendientes analizadas: {', '.join(f'{p:.4f}' for p in pendientes)}\n")
    return valor_umbral_estadistico

def estructura_grafica(meses, normalidad, estilo):
    """
    Crea la estructura básica de la gráfica con puntos y línea suavizada.

    Args:
        meses (list): Lista de meses.
        normalidad (list): Valores de normalidad.
        estilo (dict): Diccionario con estilos.

    Returns:
        tuple: Figura, ejes, y valores X.
    """
    fig, ax = plt.subplots(figsize=(14, 6))
    x = np.arange(len(meses))
    ax.plot(x, normalidad, 'D', color='black',
            markersize=estilo["tamaño_puntos"], markeredgewidth=3)
    x_smooth = np.linspace(min(x), max(x), 300)
    spl = make_interp_spline(x, normalidad, k=3)
    y_smooth = spl(x_smooth)
    ax.plot(x_smooth, y_smooth, color='black', linewidth=1.8)
    return fig, ax, x

def colores_tramos(ax, x, normalidad, tramos):
    """Colorea los tramos (Q1, Q2, Q3) en el fondo de la gráfica."""
    for _, (inicio, fin, _, color) in tramos.items():
        ax.axvspan(inicio - 0.5, fin - 0.5, color=color, alpha=0.3)

def etiquetas_promedios(ax, x, normalidad, tramos, estilo):
    """Agrega etiquetas de promedio por tramo."""
    for nombre, (inicio, fin, _, _) in tramos.items():
        valores = normalidad[inicio:fin]
        promedio = np.mean(valores)
        x_pos = (inicio + fin - 1) / 2
        max_local = max(valores)
        min_local = min(valores)
        rango_visible = max(normalidad) - min(normalidad)
        margen_relativo = max(0.5, rango_visible * 0.05)
        y_pos = max_local + margen_relativo if max_local - min(valores) > min_local - min(normalidad) else min_local - margen_relativo

        ax.text(
            x_pos, y_pos, f"""Promedio\nNormalidad: {promedio:.1f}%""",
            ha='center', va='center',
            fontsize=estilo["fontsize_promedios"],
            weight=estilo["negrita_promedios"],
            color='navy',
            bbox=dict(boxstyle="round,pad=0.3", facecolor='white', edgecolor='navy', alpha=0.85),
            path_effects=[pe.withStroke(linewidth=1.5, foreground='white')]
        )

def promedio_global(ax, normalidad):
    """Dibuja la línea del promedio global y la devuelve."""
    promedio = np.mean(normalidad)
    ax.axhline(y=promedio, color='gray', linestyle='--', linewidth=1.5, label=f'Promedio global: {promedio:.1f}%')
    return promedio

def leyendas_tendencia(x, normalidad, tramos):
    """
    Clasifica la tendencia de cada tramo (creciente, estable, decreciente).

    Returns:
        list: Objetos de leyenda para cada tramo.
    """
    pendientes = []
    pendientes_por_tramo = {}
    for nombre, (inicio, fin, etiqueta, color) in tramos.items():
        x_local = x[inicio:fin]
        valores = normalidad[inicio:fin]
        _, _, pendiente = regresion_lineal(x_local, valores)
        pendientes.append(pendiente)
        pendientes_por_tramo[nombre] = (pendiente, etiqueta, color)
    umbral_estadistico(pendientes, k=0.5)
    leyendas = []
    for nombre, (pendiente, etiqueta, color) in pendientes_por_tramo.items():
        if abs(pendiente) <= valor_umbral_estadistico:
            etiqueta_tendencia = "Tendencia estable"
        elif pendiente > 0:
            etiqueta_tendencia = "Tendencia creciente"
        else:
            etiqueta_tendencia = "Tendencia decreciente"
        leyendas.append(
            plt.Rectangle((0, 0), 1, 1, color=color, alpha=0.3, label=f"{etiqueta} – {etiqueta_tendencia}")
        )
    return leyendas

def construccion_grafica(ax, leyenda_normalidad, leyendas_tramos):
    """Agrega la leyenda completa al gráfico."""
    ax.legend(
        handles=[leyenda_normalidad] + leyendas_tramos,
        loc='upper right', frameon=True, facecolor='white', edgecolor='gray',
        fancybox=True, framealpha=0.7
    )

def grafica_final(meses, normalidad, tramos, estilo, nombre_archivo):
    """
    Construye la figura final completa con todos los elementos visuales.

    Returns:
        matplotlib.figure.Figure: Objeto figura lista para guardar o mostrar.
    """
    fig, ax, x = estructura_grafica(meses, normalidad, estilo)
    colores_tramos(ax, x, normalidad, tramos)
    etiquetas_promedios(ax, x, normalidad, tramos, estilo)
    promedio = promedio_global(ax, normalidad)
    leyenda_promedio_global = Line2D([0], [0], color='gray', linestyle='--', linewidth=1.5, label=f'Promedio global: {promedio:.1f}%')
    leyendas_tramos = leyendas_tendencia(x, normalidad, tramos)

    for i, val in enumerate(normalidad):
        ax.text(i, val + 0.4, f"{val:.1f} %",
                ha='center', fontsize=10, weight='bold', color='black', fontname='Segoe UI')

    leyenda_normalidad = Line2D([0], [0], color='black', marker='D',
                                linewidth=estilo["grosor_linea"], label='Normalidad mensual')
    construccion_grafica(ax, leyenda_normalidad, [leyenda_promedio_global] + leyendas_tramos)

    ax.text(0.5, 1.04, " " + " ".join("Normalidad Nivel Nacional"), fontsize=14, weight='normal',
            ha='center', va='center', transform=ax.transAxes, fontname='Arial')
    ax.text(1.0, 1.04, " " + " ".join("Meta de normalidad: 92%"), fontsize=14, color='black',
            weight='normal', ha='right', va='center', transform=ax.transAxes, fontname='Arial')

    ax.set_ylabel("%  " + " ".join("Normalidad"), fontname='Arial', fontsize=12)
    ax.set_xticks(ticks=x)
    ax.set_xticklabels(meses)
    margen = max(1, (max(normalidad) - min(normalidad)) * 0.2)
    ax.set_ylim(min(normalidad) - margen, max(normalidad) + margen)
    ax.set_xlim(-0.5, len(meses) - 0.5)
    ax.grid(True, linestyle="--", alpha=0.5)

    fig.tight_layout()
    return fig

def guardar_grafica(fig, nombre_salida):
    """Guarda la figura como archivo PNG."""
    fig.savefig(nombre_salida, dpi=300)
    print(f"Gráfico guardado como {nombre_salida}")

def mostrar_grafica(fig):
    """Muestra la figura en pantalla."""
    plt.show()

def main():
    """Función principal del script."""
    archivo_ruta, archivo_nombre = validacion_archivo()
    validacion_pestaña(archivo_ruta, hoja="Hoja1")
    df = pd.read_excel(archivo_ruta, sheet_name="Hoja1")
    validacion_columnas(df, ["Mes", "Normalidad"])
    meses, normalidad = carga_datos(df)
    tramos = generador_tramos(len(meses))
    estilo = estilo_grafico()
    figura = grafica_final(meses, normalidad, tramos, estilo, os.path.splitext(archivo_nombre)[0])
    guardar_grafica(figura, "Normalidad_01.png")
    mostrar_grafica(figura)

if __name__ == "__main__":
    main()