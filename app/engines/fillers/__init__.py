"""
fillers/ — заполнители шаблонов. Читают ТОЛЬКО canonical.Shipment и пишут в
копию шаблона, не зная, как выглядел исходник.
"""
from .optiauto import fill_optiauto_whole  # noqa: F401
