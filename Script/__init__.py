# -*- coding: utf-8 -*-
"""
Paquete de los scripts de calculo de Balance BESS.

Balance_BESS.py (la ventana) vive un nivel mas arriba y es lo unico
que se ejecuta: todo lo que hay aca adentro es importable, no se corre
suelto.

    Balance_BESS.py
    Script/
        nucleo.py          <- todo el calculo del caso
        Cmg/
            Extrae_CMG_barras.py   <- arma cmg.xlsx desde el CSV
                                      15-minutal
        Subastas/
            Ofertas_Adjudicadas.py <- trae y lee los Access
                                      OfertasSSCCAdj*.accdb

Por ahora `nucleo.py` sigue siendo un solo archivo grande; la idea
(conversada con el usuario) es ir sacando de ahi un modulo por etapa,
como ya se hizo con Cmg/, a medida que se agreguen mas entradas.
"""
