# Auditoría del panel BAM-CMA — 22/10/2026

## Errores encontrados y corregidos

### Datos (ordenes_backup.json)
| Problema | Efecto | Corrección |
|---|---|---|
| 3 órdenes de febrero (02/02, 07/02, 23/02) cargadas con fecha **01/01/2026** (el parser ponía 01/01 cuando no encontraba fecha) | Esos días figuraban "sin orden" y enero tenía 3 órdenes de más | Fechas corregidas por nombre de archivo + numeración correlativa |
| Orden Nº 70 (archivo 02-04-26) guardada como **01/04** | Dos órdenes el 01/04 y ninguna el 02/04 | Fecha corregida a 02/04 |
| 963 pedidos en tipo **"Otro"** cuyo texto decía el servicio ("Limpieza de cesto y basura dispersa", "Barrido de calles", "Residuos no habituales…") | Totales por tipo subestimados | Reclasificados: 249 cesto, 213 RNH, 190 barrido, 111 basura en vía pública, 89 levante, 33 poda, 28 volcadero, 23 basural, 17 platabanda, 10 recolección |
| Tipo "Basura en vía pública" no existía en la paleta | Se mostraba con color/ícono de "Otro" | Agregado |

### Normalización de direcciones (causa principal de números erróneos)
Había **4 normalizadores distintos** (Análisis, Mapa, Cruce, Puntos verdes) y cada uno contaba distinto.
Ahora hay **uno solo (ADDR)** que usan todas las vistas:
- Esquinas en cualquier orden ("Moreno y F. García" = "Fortunata García y Moreno").
- Acentos, mayúsculas, abreviaturas (Av., Pje., Psj, Bbe., Gral., Cnel., Dr., C.), iniciales ("F. García" → Fortunata García), typos frecuentes (Quintero/Quinteros, Garcilazo/Garcilaso).
- Quita ruido: nº de A.C., "pedido WhatsApp", teléfonos, horarios ("Hs. 12.05"), expedientes, "ambas aceras", paréntesis.
- Reconoce tramos ("X entre A y B", "X desde A hasta B", "e/"), alturas ("al frente del 928" → cuadra 900) y calles con "y" en el nombre (López y Planes, Salas y Valdez, Trejo y Sanabria, Vera y Aragón).
- Resultado: 25.160 formas escritas distintas → ~10.800 lugares.

### Cruces
| Problema | Corrección |
|---|---|
| `isSubset()` usaba `split()` sin separador → **nunca funcionaba** | Reemplazado por el índice de lugares |
| Cada punto del mapa contaba solo la variante más parecida (las demás formas de escribir el mismo lugar se perdían) | Suma todas las variantes del lugar |
| Similitud Jaccard ≥ 50 % generaba falsos positivos (ej. "Nougues y Estrada" = "Mendoza y Nougues", "Corrientes 2728" = "Bs. As. 2699") | Coincidencia exacta de lugar (esquina/tramo/altura ±100) |
| Conteo por pedidos: filas duplicadas en una misma orden inflaban el número | Se cuenta en **días con pedido** (y se muestran también los pedidos) |
| Tramos vecinos se tomaban como el mismo punto | Tramo = misma calle y mismas dos transversales |

Comparación 2026 (antes → ahora): activos con órdenes 226 → 209 · activos sin órdenes 17 → 34 · nuevos candidatos 411 → 390 · inactivos que reaparecen 22 → 100.

### Incumplimientos
- La pestaña de incumplimientos rotulaba los numerales como "Presentación informes / Cobertura de barrido / …" mientras el resto del panel los usaba como Barrido / Basurales / RNH / Volcaderos / Falta de recolección. Unificado a esta última (confirmado por Daniel).
- La copia embebida de incumplimientos tenía 17 de 73 días con valores cruzados respecto de `incumplimientos.json`. Regenerada desde el JSON.

### Importación de órdenes (.docx)
- El parser exigía columnas fijas (Nº, tipo, ubicación). Ahora detecta las columnas por encabezado, acepta celdas combinadas y tablas anidadas → causa probable del "Sin pedidos" del 22/10.
- Si no encuentra fecha ya no inventa 01/01: usa el nombre del archivo y avisa si la fecha del documento no coincide con la del archivo.
- Detecta reimportación del mismo archivo y avisa si ya había otra orden del mismo día.
- Guarda el tipo original (`tipoRaw`) para poder auditar reclasificaciones.
- `restoreData()` comparaba claves distintas y podía duplicar órdenes al restaurar un backup. Corregido.

### Persistencia
- `window.storage` no existe en GitHub Pages: lo importado se perdía al recargar si no se hacía Sync. Ahora se guarda en el navegador (IndexedDB) y se suma a lo de GitHub al abrir el panel, con aviso de "órdenes sin subir".
- El JSON se pide sin caché para ver siempre la última versión.

### Calendario
- Feriados incompletos (faltaban puentes, Güemes, San Martín, Soberanía trasladado y el 24/09 provincial). Completados 2024-2026.

### Nuevo
- **Buscador de direcciones** en Análisis: escribís una esquina, tramo o calle y muestra pedidos y días por mes, por tipo y todas las formas en que fue escrita.

## Pendientes / a revisar
- 2025 tiene saltos de numeración (Nº 155 el 09/07 y el 28/07, Nº 241 el 08/10, etc.): probables fechas mal cargadas en el historial CSV. No se tocaron sin el documento original.
- La orden del 04/04/2026 trae Nº 000069 (repetido con la del 01/04): error de tipeo en el documento.
- La contraseña del panel está escrita en el código y los JSON son públicos en GitHub Pages: no protege los datos.
