# -*- coding: utf-8 -*-
"""
Diagnostico comun de los cruces que terminan en cero.
"""

from .utiles import _texto_seguro, _tiene_valor, normalizar


def _avisar_claves_sin_mapeo(
    valores, mapa, descripcion, origen, registrar=print, maximo=15
):
    """Avisa, una vez por valor, las claves ausentes o con dato vacio.

    Los VLOOKUP/diccionarios de la planilla suelen terminar en blanco y
    varios calculos posteriores convierten ese blanco en cero. Este helper
    hace visible el problema antes de esa conversion, sin inundar el log
    con una linea por cada cuarto de hora.
    """

    # Las columnas de entrada tienen una fila por cuarto de hora
    # (decenas de miles); los valores distintos son unas pocas
    # centrales. Se deduplica ANTES de normalizar para no repetir el
    # normalizar() -- que hace unicodedata + regex -- una vez por fila.
    distintos = {valor for valor in valores if _tiene_valor(valor)}

    faltantes = sorted({
        _texto_seguro(valor)
        for valor in distintos
        if not _tiene_valor(mapa.get(normalizar(valor)))
    })

    if not faltantes:
        return []

    muestra = ", ".join(repr(valor) for valor in faltantes[:maximo])
    resto = len(faltantes) - maximo
    sufijo = f" (y {resto:,} mas)" if resto > 0 else ""
    registrar(
        f"  [AVISO] {descripcion}: {len(faltantes):,} valor(es) sin "
        f"correspondencia en {origen}: {muestra}{sufijo}. Los resultados "
        f"dependientes pueden quedar vacios o en 0."
    )
    return faltantes
