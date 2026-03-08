# -*- coding: utf-8 -*-
def classFactory(iface):
    from .xlsx_furigana_delete import XlsxFuriganaDelete
    return XlsxFuriganaDelete(iface)
