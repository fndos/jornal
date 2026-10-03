# Jornal

Convierte el registro del reloj biométrico en el formato de la quincena.

Todo ocurre en el navegador: el archivo que sube la persona no se envía a
ningún servidor. El cálculo lo hace el mismo `motor.py` que usa la versión de
escritorio, ejecutado sobre Pyodide (Python compilado a WebAssembly).

## Qué contiene

| | |
|---|---|
| `index.html`, `estilo.css`, `app.js` | La interfaz |
| `py/motor.py` | El cálculo. Copia exacta de la versión de escritorio |
| `py/almacen.py` | Lectura de la configuración y la tabla de roles |
| `py/datos/configuracion.xlsx` | Tarifas y umbrales. No contiene datos de personas |

## Lo que NO está aquí, a propósito

La tabla de roles y la plantilla llevan cédulas y nombres. En un sitio estático
cualquiera con la URL podría descargarlas, así que no se publican: cada persona
carga la tabla una vez desde la propia aplicación y queda guardada en su
navegador.

## Desarrollo

```bash
python3 servidor.py        # http://127.0.0.1:8000 sin caché
./sincronizar.sh           # copia el motor desde el proyecto de escritorio
```

`sincronizar.sh` hay que correrlo antes de publicar: si no, el sitio se queda
con un motor viejo y empieza a dar resultados distintos a los del escritorio.
