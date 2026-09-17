# VBA Dev

Edita tus macros en VS Code. Al guardar el código, se actualiza en Excel.

## 1. Instala

Necesitas **Windows y Excel de escritorio**.

1. Descarga este repositorio: **Code > Download ZIP**.
2. Extrae la carpeta, por ejemplo en `C:\Herramientas\VbaAutomaticIA`.
3. Agrega esa carpeta al **Path** de las variables de entorno de tu usuario.
4. Cierra y vuelve a abrir VS Code.

En Excel, activa:

**Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza > Configuración de macros > Confiar en el acceso al modelo de objetos de proyectos de VBA**.

## 2. Usa tu Excel

Pon tu archivo `.xlsm` o `.xlsb` en una carpeta. Abre esa carpeta en VS Code y ejecuta en la terminal:

```powershell
vba dev .
```

Si el libro está en la carpeta principal, ciérralo antes de ejecutar el comando: la herramienta lo mueve a `excel`. Luego lo abre y crea `src` con tu código. **Edita sus archivos y guarda con Ctrl+S**: los cambios pasan a Excel.

Tanto `new` como `dev` dejan este orden:

```text
TuProyecto/
├── excel/   Libro Excel
├── src/     Código VBA
└── .vba/    Configuración y respaldos
```

Si pide contraseña, escríbela en Excel. Si aparece la ventana de propiedades, ciérrala para continuar.

## 3. Comandos

El punto `.` significa **la carpeta donde estás**.

| Comando | Para qué sirve |
|---|---|
| `vba dev .` | Trabajar con un Excel existente y sincronizar mientras editas. |
| `vba new .` | Crear Excel y código en una carpeta vacía, y empezar a sincronizar. |
| `vba dev . -Once` | Aplicar los cambios pendientes una vez y terminar. |
| `vba new . -Once` | Crear el proyecto y terminar sin seguir sincronizando. |
| `vba ui .` | Editar Ribbon Custom UI con Excel cerrado. |
| `vba ui . -Once` | Extraer o aplicar Custom UI una vez y terminar. |

**Para detener:** pulsa `Ctrl+C` o cierra el libro de Excel. `Ctrl+C` guarda y cierra el libro manejado; si la herramienta abrió Excel, también cierra esa ventana.

**Para continuar otro día:** ejecuta otra vez `vba dev .`.

## Qué ocurre con tus cambios

- Puedes editar con Excel cerrado. Al ejecutar `vba dev .`, se aplican los cambios pendientes.
- Con el proceso encendido, guardar código en Excel actualiza `src`, y guardar en `src` actualiza Excel.
- Al reiniciar, tiene prioridad `src`. Si el mismo módulo cambia en ambos lados antes de sincronizar, también gana `src`.
- Crear o borrar un módulo en `src` lo crea o borra en Excel.
- Puedes crear subcarpetas dentro de `modules`, `classes` o `forms`. Se conservan al sincronizar con Excel.
- Puedes usar nombres descriptivos de archivo; `Attribute VB_Name` conserva identidad interna de Excel y macros.
- Borrar un módulo en Excel y guardar elimina su archivo, salvo que tenga cambios locales pendientes: esos cambios ganan.
- Conserva `.vba`: guarda el estado y los respaldos. Las hojas y `ThisWorkbook` no se borran eliminando su archivo.

Para agregar módulos, trabajar con formularios o usar opciones avanzadas: [guía completa](docs/GUIA-AVANZADA.md).

## Custom UI Ribbon

`vba ui .` trabaja botones y pestañas Ribbon desde `src\custom-ui\customUI.xml`. Solo admite `.xlsm` que ya contengan Custom UI.

Primera vez, con Excel cerrado:

```powershell
vba ui . -Once
```

Extrae XML. Edita texto, IDs o callbacks en VS Code. Para aplicar cambio, con Excel cerrado, ejecuta otra vez `vba ui . -Once`; o usa `vba ui .` para vigilar guardados. La herramienta guarda copia del libro en `.vba\backups` antes de modificarlo.

No ejecutes `vba dev .` y `vba ui .` a la vez. El segundo comando se detiene y muestra error en español. Para pasar de uno a otro, detén primero con `Ctrl+C`.
