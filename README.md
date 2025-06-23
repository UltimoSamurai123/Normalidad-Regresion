# Análisis de Normalidad Mensual mediante Regresión Lineal y Clasificación de Tendencia

Este repositorio contiene un script en Python que permite analizar la evolución mensual de un indicador denominado "normalidad", evaluando su comportamiento a lo largo del tiempo. El análisis se realiza mediante regresiones lineales por tramos (trimestres) y clasificación automática de tendencias, empleando un umbral estadístico dinámico calculado a partir de la desviación estándar de las pendientes.

## Objetivo

El objetivo principal del script es facilitar la detección visual y cuantitativa de cambios significativos en la tendencia de un indicador mensual. Esta herramienta es útil en contextos donde se requiere evaluar la estabilidad o evolución de procesos, como seguimiento de metas de desempeño, control de calidad o monitoreo de políticas públicas.

## Requisitos del archivo de entrada

El script está diseñado para detectar automáticamente un archivo Excel (`.xlsx`) ubicado en el mismo directorio. El archivo debe cumplir con las siguientes condiciones:

- Contener una hoja llamada `Hoja1`.
- Incluir al menos dos columnas:
  - `Mes`: nombre o número del mes (texto o entero).
  - `Normalidad`: porcentaje de normalidad (por ejemplo: 87.5).

Si no se encuentra un archivo válido, o si existen múltiples archivos `.xlsx` en la carpeta, el script detendrá la ejecución y mostrará un mensaje de error.

## Descripción general del proceso

1. **Carga automática del archivo**: el script identifica y carga un archivo Excel con el formato esperado.
2. **Segmentación en trimestres**: los datos se dividen en tres tramos iguales (Q1, Q2 y Q3).
3. **Cálculo de regresiones lineales**: se calcula una pendiente de regresión lineal para cada tramo.
4. **Cálculo del umbral estadístico**: se obtiene a partir de la desviación estándar de las pendientes.
5. **Clasificación de tendencias**: se asigna una etiqueta a cada tramo (estable, creciente o decreciente).
6. **Visualización final**: se genera una gráfica con todos los elementos anteriores y se guarda como imagen.

## Umbral estadístico

El umbral estadístico permite determinar si un cambio en la pendiente de un tramo es suficientemente significativo como para considerarse una tendencia. Su valor se calcula según la siguiente fórmula:

```python
umbral = k * std
