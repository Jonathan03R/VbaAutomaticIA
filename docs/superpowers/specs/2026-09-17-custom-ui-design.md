# Diseño: Custom UI separado de VBA

## Objetivo

Agregar comando `vba ui` para extraer, validar y guardar Ribbon Custom UI de libros `.xlsm`, sin mezclarlo con sincronización viva de VBA.

## Comandos

- `vba dev .`: comportamiento actual. Abre Excel y sincroniza solo componentes VBA.
- `vba ui .`: requiere libro y Excel cerrados. Extrae Custom UI si aún no existe en `src/custom-ui`; luego vigila XML local y lo guarda en el libro.
- `vba ui . -Once`: realiza extracción o sincronización una vez y termina.

Primera entrega no crea Ribbon nuevo ni admite `.xlsb`. Solo maneja partes Ribbon existentes: `customUI/customUI.xml` y `customUI/customUI14.xml`.

## Archivos fuente

```text
src/custom-ui/customUI.xml
src/custom-ui/customUI14.xml
```

Cada archivo solo existe si existe la parte equivalente dentro del `.xlsm`.

## Exclusión mutua

Cada proyecto tendrá `.vba/session.lock`, creado de forma exclusiva por `dev` o `ui`.

- Si `vba dev .` está activo, `vba ui .` termina sin modificar nada e informa en español que debe detenerse con `Ctrl+C`.
- Si `vba ui .` está activo, `vba dev .` termina con mensaje equivalente.
- Al terminar normalmente, con error o con `Ctrl+C`, se elimina únicamente el lock que pertenece al proceso actual.
- Un lock de proceso ya inexistente se detecta y elimina antes de continuar.

## Flujo UI

1. Validar extensión `.xlsm`, ausencia de Excel abierto y exclusión mutua.
2. Copiar libro a `.vba/backups` antes de todo cambio.
3. Leer partes Custom UI desde paquete ZIP y exportarlas como UTF-8 a `src/custom-ui`.
4. Antes de guardar XML, validar formato XML y raíz RibbonX `customUI`.
5. Reemplazar la parte ZIP de forma temporal y atómica. Si falla, conservar libro original y mostrar error español.
6. No abrir Excel, no generar imágenes, no capturar pantalla.

## Errores

Todos mensajes nuevos estarán en español: Excel abierto, libro no `.xlsm`, Ribbon ausente, XML inválido, XML sin raíz `customUI`, lock activo, ZIP dañado, archivo bloqueado y fallo al respaldar/reemplazar.

## Pruebas

Pruebas PowerShell crearán `.xlsm` ZIP temporales sin Excel. Cubrirán extracción, reemplazo, XML inválido, Ribbon ausente, rechazo `.xlsb` y bloqueo mutuo. Prueba de integración Excel existente no es necesaria para Custom UI porque operación ocurre sobre paquete cerrado.
