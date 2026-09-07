# Xpendit a Nubatic

Convierte el Reporte Consolidado de Xpendit (Excel) directamente en los
archivos de importación para Nubatic y las planillas de facturas por
persona, sin procesamiento de imágenes. Todos los gastos del Excel se
tratan como aprobados por el solo hecho de estar en la exportación de
Xpendit.

## Uso

1. Copia el Reporte Consolidado de Xpendit (`.xlsx`) en esta carpeta (debe
   ser el único Excel de entrada presente).
2. Ejecuta el script para tu sistema operativo:
   - **Windows**: doble clic en `Windows_Ejecutar.bat`
   - **Mac**: doble clic en `Exe_Segun_OS/Mac_Ejecutar.command`
   - **Linux**: `./Exe_Segun_OS/Linux_Ejecutar.sh`

La primera vez prepara un entorno virtual propio (`venv/`) e instala las
dependencias (`pandas`, `openpyxl`); las siguientes veces lo reutiliza.

Genera en la misma carpeta:

- `importador_{OBRA}_{PERSONA}_{semana}.xlsx` — boletas/comprobantes por persona/obra
- `facturas_{PERSONA}_{semana}.xlsx` — facturas por persona
- `informe_manual_{semana}.xlsx` — folio > 7 caracteres u obra sin CC configurado

## Cambios recientes

- **Sintaxis Python 2 corregida**: `except X, Y:` (inválido en Python 3)
  reemplazado por `except (X, Y):` en 4 lugares del script — antes no
  compilaba en Python 3.
- **Salida UTF-8 forzada**: los emojis de los mensajes de progreso ya no
  rompen con `UnicodeEncodeError` en consolas Windows (codepage cp1252).
- **`Windows_Ejecutar.bat` evita Python "free-threaded" (3.13t)**: esa
  build no tiene wheels precompilados de `pandas`/`numpy` en PyPI, lo que
  hacía fallar la instalación al intentar compilar desde código fuente.
  Ahora detecta y usa una build normal de Python 3 automáticamente.
- **Metadatos del `.xlsx` normalizados a "Microsoft Excel"**: `openpyxl`
  firma los archivos generados como propios (`docProps/app.xml` y
  `core.xml`), y algunas plataformas ERP los rechazan al importar por no
  "parecer" venir de Excel real. Ahora se reescriben automáticamente esos
  metadatos después de guardar cada archivo, sin necesidad de abrirlo en
  Excel y volver a guardarlo manualmente.
- **Shared strings en vez de inline strings**: `openpyxl` escribe todas las
  celdas de texto como "inline string" (sin tabla `xl/sharedStrings.xml`,
  la que usa Excel real). Algunos lectores de Excel de terceros no la
  soportan y devuelven el objeto crudo de la celda en vez del texto
  (aparecía como "[object Object]" en la UI del ERP). Ahora se reescribe
  automáticamente con una tabla de shared strings real, sin necesidad de
  abrir el archivo en Excel y volver a guardarlo.

## Pendiente

- Probar en Mac y Linux (Windows ya se probó y quedó funcionando).
