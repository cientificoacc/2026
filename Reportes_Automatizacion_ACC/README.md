<p align="center">
  <img src="assets/logo_acc.png" alt="ACC" height="70" style="margin-right:20px;">
  <img src="assets/logo_gobernacion_cundinamarca.png" alt="Gobernación de Cundinamarca" height="70">
</p>

# Reportes de Trámites Catastrales (ACC) – Generación automática 2024–2025

Este proyecto genera, **por municipio**, un informe en **Word (.docx)** con:

- **Tabla 1**: Clasificación de radicados por estado (2024 vs 2025) + total general  
- **Tabla 2**: Radicados por tipo y quién resuelve, con **subtotales jerárquicos** (Central / Territorio) + total general  
- **Gráfica 1**: Radicados por mes (2024 vs 2025)  
- **Gráfica 2**: Radicados por día de la semana (2024 vs 2025)  
- **Tabla 3**: Radicados resueltos/en proceso por año y resolución, con **subtotales jerárquicos por año y nivel** + total general  

Las tablas se insertan **debajo de los títulos** dentro de la plantilla institucional y aplican formato:

- **Encabezados**: fondo azul `#2F5496`, letra blanca, negrita  
- **Subtotales y Total general**: fondo blanco, letra negra, negrita

---

## 1. Requisitos

- Cuenta de Google (para usar Google Colab)
- Archivos de entrada:
  - **Excel**: `Plantilla_Reporte_TuCatastro.xlsx` (contiene hoja `CRUDOS`)
  - **Plantilla Word**: `Mpio_Informe_tramites_Catastral_ACC.docx`

---

## 2. Estructura esperada de archivos (en Colab)

En tu sesión de Colab, ubica los archivos en `/content/`:

```
/content/
  Plantilla_Reporte_TuCatastro.xlsx
  Mpio_Informe_tramites_Catastral_ACC.docx
```

> Si cambias el nombre o la ruta, ajusta las variables `BASE_EXCEL` y `PLANTILLA_WORD` en el código.

---

## 3. Cómo ejecutar en Google Colab (paso a paso)

1) **Abre un Notebook en Colab**  
   - Ve a Google Colab y crea un Notebook nuevo.

2) **Sube los archivos de entrada**  
   - En el panel izquierdo de Colab, abre **Files (Archivos)**  
   - Clic en **Upload (Subir)** y sube:
     - `Plantilla_Reporte_TuCatastro.xlsx`
     - `Mpio_Informe_tramites_Catastral_ACC.docx`

3) **Pega el código completo en una celda**  
   - Copia el script final y pégalo en una celda.

4) **Ejecuta la celda**  
   - El código instalará dependencias y generará los reportes.

---

## 4. Parámetros clave a revisar

### 4.1 Rutas y archivos
En el código, asegúrate de tener:

```python
BASE_EXCEL = "/content/Plantilla_Reporte_TuCatastro.xlsx"
PLANTILLA_WORD = "/content/Mpio_Informe_tramites_Catastral_ACC.docx"
CARPETA_RAIZ = "/content/Reportes_ACC_2024_2025"
ANIOS = [2024, 2025]
```

### 4.2 Ejecutar un municipio o todos

**Todos los municipios** (según `NOM_MUN` en `CRUDOS`):

```python
generar_sistema_reportes()
```

**Un municipio específico**:

```python
generar_sistema_reportes("GUATAVITA")
```

---

## 5. Salida del proceso (dónde quedan los reportes)

Para cada municipio se crea una carpeta:

```
/content/Reportes_ACC_2024_2025/<MUNICIPIO>/
  01_Word/
    <MUNICIPIO>_Informe_Final_ACC_2024_2025.docx
  02_Graficas/
    Radicados_por_mes.png
    Radicados_por_dia.png
```

---

## 6. Recomendaciones de calidad de datos

Para evitar errores o resultados vacíos:

- Verifica que `CRUDOS` contenga columnas:
  - `NOM_MUN`, `RAD_ANO`, `Numero radicado`, `Estado`, `OFICINA DE GESTION`, `Tipo`, `RES_ANO`
- La columna de fecha de radicación debe contener “fecha” y “rad” en el nombre (por ejemplo: `Fecha radicación`), ya que el script la detecta automáticamente.

---

## 7. Solución de problemas comunes

### “❌ No se encontró columna de fecha de radicación…”
- Revisa el nombre de la columna de fecha en `CRUDOS`.
- Debe incluir las palabras **fecha** y **rad** (sin importar mayúsculas/minúsculas).

### “⚠ No se encontró el título…”
- Revisa que la plantilla Word contenga exactamente (o al menos como substring) los títulos donde se insertan tablas/imágenes:
  - “Clasificación de los radicados…”
  - “Radicados por tipo y quien resuelve”
  - “Radicados por mes”
  - “Radicados por día de la semana”
  - “Radicados resueltos desagregados…”

---

## 8. Buenas prácticas de uso institucional

- Mantén una **plantilla Word maestra** con el texto institucional aprobado.
- Versiona el Excel y la plantilla Word por vigencia (ej. 2024–2025).
- Si cambia el criterio de **Central vs Territorio**, ajusta la regla:
  - Actualmente: se clasifica como **Central** si `OFICINA DE GESTION` contiene “CENTRAL”.

---

## 9. Créditos institucionales

Desarrollado para la generación automática de reportes municipales de trámites de conservación catastral (ACC) con lineamientos de presentación institucional.

