# Limpiador de Excel de ventas

Herramienta de línea de comandos para revisar un Excel de ventas, aplicar limpieza conservadora y generar un informe auditable. El archivo original nunca se sobrescribe.

## Instalación

Requiere Python 3.10 o posterior.

```bash
python -m pip install -r requirements.txt
```

## Uso

```bash
python limpiar_excel.py datos_prueba.xlsx
```

El programa crea `ventas_limpias.xlsx` e `informe_errores.xlsx` en la misma carpeta que el archivo de entrada. La entrada debe contener la hoja `ventas` y las columnas `nombre`, `fecha`, `monto` y `ciudad`.

## Reglas aplicadas

- Elimina duplicados exactos comparando los cuatro valores originales, antes de recortar espacios. Conserva la primera aparición y registra el número de fila de Excel de cada copia eliminada.
- Recorta espacios iniciales y finales de los campos de texto y normaliza los nombres con formato título, respetando las tildes y dejando en minúscula partículas como «de», «la» o «del» (por ejemplo, «Juan de la Cruz»).
- Acepta fechas `dd/mm/aaaa`, `aaaa-mm-dd` y `dd-mm-aa`; las fechas reales se guardan como `aaaa-mm-dd`. Una fecha inválida no se adivina: se deja vacía y se registra con el valor original.
- Conserva las filas con nombre vacío y las reporta como incidencia. No las elimina para evitar descartar ventas sin autorización.
- Si un monto no es numérico, lo deja vacío y lo registra. Si es negativo, lo conserva como número y lo registra para que una persona decida cómo tratarlo.
- `informe_errores.xlsx` contiene `incidencias` (fila original, columna, tipo de error y valor original) y `resumen` (cantidad por tipo de incidencia).
- El resumen de consola muestra filas leídas, duplicados, fechas convertidas e inválidas, nombres vacíos, montos con error y rutas de salida.

## Pruebas

```bash
python -m pytest -v test_limpiar_excel.py
```

Las pruebas construyen sus propios Excel temporales; no importan ni ejecutan `generar_datos_prueba.py`. Incluyen duplicados, espacios, fecha imposible, monto textual, monto negativo, nombres con tilde y partículas, formatos de fecha `dd-mm-aa` y `aaaa-mm-dd`, montos con «S/» y comprobación SHA-256 para asegurar que la entrada no cambia.

## Limitaciones y decisiones

- La validación de este proyecto comprueba estructura y reglas de limpieza, no determina si una venta es comercialmente verdadera ni si los datos pertenecen a clientes reales.
- Los duplicados se determinan por igualdad exacta en las cuatro columnas originales. Dos filas que solo difieren por espacios se consideran distintas en esa etapa, aunque después puedan verse iguales al limpiar.
- Los formatos de fecha se interpretan estrictamente; no se intenta corregir fechas imposibles ni interpretar formatos no especificados.
- Para años de dos cifras se usa la convención estándar de Python (`00–68` como `2000–2068`, `69–99` como `1969–1999`).
- Un valor vacío en la fecha se deja vacío sin contarlo como fecha inválida; un valor no vacío que no se puede interpretar sí se registra.
