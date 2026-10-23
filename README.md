# Descargador de imágenes Pro

Programa de escritorio para catálogos de producto: le das un Excel o CSV con
el código de cada producto y los enlaces a sus fotos, y **descarga todas las
fotos en paralelo**, las nombra con el código del producto y, si quieres,
**les quita el fondo automáticamente**.

## Qué hace

- **Encuentra solo las columnas**, aunque se llamen distinto en cada listado
  (`SKU`, `REFERENCIA`, `CODIGO`… para el nombre; `URL`, `IMAGEN`, `LINK`…
  para las fotos).
- **Entiende varias fotos por producto**, separadas por comas o en formato
  lista (`["url1", "url2"]`). La primera se guarda como `SKU` y las demás como
  `SKU_1`, `SKU_2`…
- **Descarga en paralelo** (de 1 a 50 descargas a la vez, con un control
  deslizante).
- **Quita el fondo** con [rembg](https://github.com/danielgatis/rembg), un
  modelo de IA que se ejecuta en tu ordenador. Se activa por producto con una
  columna `FONDO` = `SI`, y se puede desactivar entero con un interruptor.
- **No pierde ningún fallo**: las fotos que no se pueden descargar quedan
  apuntadas, con el motivo, en `errores_recientes.xlsx`.
- Lee **CSV con `,` o con `;`** (el formato que usa Excel en España) y `.xlsx`.

## Instalación y uso

Necesitas Python 3.

```bash
pip install -r requirements.txt
python descargador_pro.py
```

1. Pulsa **Seleccionar Excel o CSV**.
2. Ajusta la velocidad y si quieres quitar fondos.
3. Pulsa **INICIAR DESCARGA**.

Las fotos aparecen en las carpetas `originales/` y `sin_fondo/` (hay botones
para abrirlas).

### Crear el ejecutable (.exe)

El archivo `descargador_pro.spec` sirve para empaquetarlo con
[PyInstaller](https://pyinstaller.org) en un solo `.exe` que funciona sin
tener Python instalado:

```bash
pip install pyinstaller
pyinstaller descargador_pro.spec
```

## Detalles técnicos

- La descarga va en un **hilo aparte** para que la ventana no se congele. Ese
  hilo nunca toca la ventana directamente: deja sus mensajes en una cola y la
  ventana los recoge cada 100 ms, porque Tkinter no permite que otro hilo la
  modifique.
- Cada foto se guarda con la **extensión de su formato real** (`.jpg`, `.png`,
  `.webp`), y si un código de producto se repite en el listado no se pisa la
  foto anterior.
