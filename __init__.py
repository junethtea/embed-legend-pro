# -*- coding: utf-8 -*-
"""
/***************************************************************************
 Embed Legend Pro
                                 A QGIS plugin
 Interactive Floating Legend & Advanced Thematic Export Tool
                             -------------------
        begin                : 2025
        copyright            : (C) 2025 by Jujun Junaedi
        email                : jujun.junaedi@outlook.com
 ***************************************************************************/
"""

def classFactory(iface):
    """Load EmbedLegendPlugin class from file embed_legend.
    
    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    
    # Import class utama dari file embed_legend.py
    from .embed_legend import EmbedLegendPlugin
    return EmbedLegendPlugin(iface)