#!/bin/bash
# Deja la versión web con el mismo motor que la de escritorio.
#
# A PROPÓSITO no se copian roles.xlsx ni plantilla.xlsx: llevan cédulas y
# nombres, y en un sitio estático cualquiera con la URL podría descargarlos.
# La tabla de roles la carga cada persona en su navegador y queda ahí.
cd "$(dirname "$0")" || exit 1
cp ../motor.py ../almacen.py py/
mkdir -p py/datos
cp ../datos/configuracion.xlsx py/datos/
rm -f py/datos/roles.xlsx py/datos/plantilla.xlsx
echo "web sincronizada (sin datos personales):"
for f in py/motor.py py/almacen.py py/datos/configuracion.xlsx; do
  printf "  %-34s %8s bytes\n" "$f" "$(stat -f%z "$f" 2>/dev/null || stat -c%s "$f")"
done
