# Transferencias Starbucks

## Versión

`v14-fechas-auditadas`

Aplicación web estática para la auditoría crítica de transferencias entre tiendas.
Conserva la conciliación existente y agrega trazabilidad desde cada CSV mensual
hasta `Base_Directorio` y `Non Inventory`.

## Fuentes

- `data/source/monthly/*.csv`: un archivo independiente por mes.
- `data/source/Base_Transferencias.xlsx`:
  - `Base_Directorio`: fuente oficial de Tienda, Región y DM.
  - `Non Inventory`: artículos excluidos de indicadores y conciliación.

Los datos publicados en la PWA se generan en `data/chunks/`. El navegador no
requiere Python, Excel, backend ni servidor.

## Actualización mensual automática

El flujo `.github/workflows/update-monthly-data.yml` permite actualizar la
aplicación sin modificar HTML, CSS ni JavaScript.

1. Reemplaza el CSV del mes actual o agrega uno nuevo en
   `data/source/monthly/`.
2. Usa el nombre del mes correspondiente, por ejemplo `Agosto.csv`,
   `Septiembre.csv` o `Octubre.csv`.
3. Confirma el cambio en la rama `main`.

GitHub Actions detecta el CSV, instala las dependencias, ejecuta el procesador,
valida todos los meses disponibles y guarda automáticamente únicamente los
archivos de datos regenerados.

Cada CSV debe contener el acumulado completo del mes y conservar los 12
encabezados esperados. No es necesario crear carpetas, bloques JSON ni cambios
de código para los meses posteriores.

### Ejecución manual de respaldo

Si se necesita regenerar localmente:

1. Ejecuta `python -m pip install -r tools/requirements.txt`.
2. Ejecuta `python tools/process_monthly_data.py`.
3. Ejecuta `python tools/validate_monthly_update.py`.

## Controles del procesamiento

- Detecta mes por nombre y lo confirma contra las fechas internas.
- Determina el sistema de semana con registros no ambiguos; la fuente utiliza
  semana ISO.
- Conserva enero–junio como `DD/MM/YYYY` y normaliza julio desde `M/D/YYYY`,
  validando cada fila contra Año, Semana y mes del CSV.
- Separa cualquier fecha que no pueda resolverse sin adivinar y registra la
  trazabilidad agrupada en `data/audit/fechas.json`.
- Detecta codificación y delimitador.
- Rechaza por separado los CSV incompatibles sin bloquear los demás.
- Conserva los 12 campos originales y añade mes, archivo, fila y evidencia.
- Elimina de la consolidación únicamente duplicados exactos.
- Excluye `Non Inventory` mediante coincidencia exacta tras normalizar espacios
  y tipos de texto.
- Conserva exclusiones e incidencias en `data/audit/`.
- Mantiene Patrol y Proveedor sin CeCo en el registro técnico, pero ocultos de
  tablas, indicadores, gráficas y exportaciones operativas.
- Calcula Última actualización con la fecha máxima válida del archivo del mes
  actual; si falta, usa el periodo más reciente y lo informa.
- El detalle permanece cerrado inicialmente y “Evidencia de cruce” se conserva
  en los datos de auditoría, sin saturar la tabla principal.

## Reglas de conciliación conservadas

- `Cantidad < 0`: salida desde el CeCo de la fila hacia el CeCo inicial del
  proveedor.
- `Cantidad > 0`: entrada en el CeCo de la fila desde el CeCo inicial del
  proveedor.
- La relación origen/destino se valida en ambos sentidos con CeCo exacto.
- Una entrada mayor que su salida se concilia hasta la cantidad enviada.
- Las entradas independientes se agrupan por fecha y pareja de CeCos.
- La coincidencia exacta de ingrediente tiene prioridad.
- Si los nombres difieren, solo se usa equivalencia económica cuando unidad,
  cantidad y costo son compatibles y la opción es inequívoca.
- Cada entrada se utiliza una sola vez.
- El resumen respeta los filtros activos.

## PWA y GitHub Pages

- Rutas relativas compatibles con la subruta `Transferencia`.
- Service Worker con caché `transferencias-v13-mensual-auditada`.
- Los datos se consultan con estrategia network-first y se actualizan en caché,
  evitando servir meses anteriores cuando existe conexión.
- Después de la primera consulta de un mes, su bloque queda disponible sin
  conexión en ese dispositivo.
