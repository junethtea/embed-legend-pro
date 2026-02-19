#!/usr/bin/env python
# -*- coding: utf-8 -*-

# ==============================================================================
#  PLUGIN      : Embed Legend
#  VERSION     : 6.9.0
#  AUTHOR      : Jujun Junaedi
#  EMAIL       : jujun.junaedi@outlook.com
#  COPYRIGHT   : (c) 2026 by Jujun Junaedi
#  LICENSE     : GPL-2.0-or-later
#  DESCRIPTION : Interactive Floating Legend & Advanced Thematic Export Tool
# ==============================================================================

"""
LICENSE AGREEMENT:
This program is free software; you can redistribute it and/or modify it under 
the terms of the GNU General Public License as published by the Free Software Foundation.
To support the developer and ensure you have the latest stable version, 
please download directly from the Official QGIS Repository.
"""

import os
import csv
import re
import zipfile
import shutil
import tempfile
import sip 

# GUI & Core Imports
from qgis.PyQt.QtCore import Qt, QUrl, QVariant
from qgis.PyQt.QtGui import (
    QColor, QIcon, QFont, QCursor, QFontMetrics, QDesktopServices, QBrush
)
from qgis.PyQt.QtWidgets import (
    QAction, QDockWidget, QListWidget, QListWidgetItem, 
    QVBoxLayout, QWidget, QLabel, QFileDialog, QMenu, 
    QColorDialog, QFontDialog, QMessageBox, QProgressDialog
)
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsCoordinateReferenceSystem, 
    QgsCoordinateTransform, QgsRenderContext, QgsWkbTypes,
    QgsGeometry, QgsSettings
)
from qgis.utils import iface

# ==============================================================================
#  CONFIGURATION & CONSTANTS
# ==============================================================================

# Internationalization (I18N) Dictionary
LANG_DICT = {
    "en": {
        "header": "Layer Info",
        "header_multi": "{} Legend",
        "menu_config": "--- CONFIGURATION ---",
        "menu_font": "🔠 Change Font",
        "menu_text_color": "🎨 Text Color",
        "menu_bg_color": "⬜ Background Color",
        "menu_border_color": "🖼️ Border Color",
        "menu_style": "🎭 Switch Style",
        "style_std": "📦 Modern (Box)",
        "style_mini": "✨ Minimalist (Clean)",
        "show_count": "🔢 Show Count",
        "show_percent": "％ Show Percentage",
        "export_mif": "📝 Export MIF (Hardcode Thematic)",
        "export_kmz": "🌏 Export KMZ (Google Earth)",
        "about": "ℹ️ About & Help",
        "lang": "🌐 Language / Bahasa",
        "success": "Success",
        "export_success": "Export Successful!",
        "file_saved": "File saved at:\n{}",
        "warning": "Warning",
        "select_layer": "Please select a vector layer first!"
    },
    "id": {
        "header": "Info Layer",
        "header_multi": "{} Legend",
        "menu_config": "--- KONFIGURASI ---",
        "menu_font": "🔠 Ganti Font",
        "menu_text_color": "🎨 Warna Teks",
        "menu_bg_color": "⬜ Warna Background",
        "menu_border_color": "🖼️ Warna Border",
        "menu_style": "🎭 Ganti Tampilan",
        "style_std": "📦 Modern (Box)",
        "style_mini": "✨ Minimalis (Clean)",
        "show_count": "🔢 Tampilkan Jumlah",
        "show_percent": "％ Tampilkan Persentase",
        "export_mif": "📝 Export MIF (Hardcode Thematic)",
        "export_kmz": "🌏 Export KMZ (Google Earth)",
        "about": "ℹ️ Tentang & Bantuan",
        "lang": "🌐 Bahasa / Language",
        "success": "Sukses",
        "export_success": "Export Berhasil!",
        "file_saved": "File disimpan di:\n{}",
        "warning": "Peringatan",
        "select_layer": "Pilih layer vektor aktif dulu, Lur!"
    }
}

# Stylesheets
STYLE_STANDARD_LBL = """
    QLabel {
        background-color: #f1f2f6; 
        color: #2f3542;
        border-bottom: 1px solid #ced6e0;
        padding-left: 8px;
    }
"""

STYLE_MINIMAL_LBL = """
    QLabel {
        background-color: transparent; 
        color: #000000;
        border: none;
        border-bottom: 1px solid #bdc3c7; 
        padding-left: 8px;
        padding-bottom: 2px;
        margin-bottom: 2px;
    }
"""

STYLE_PANEL_MINIMAL = """
    QWidget#qt_scrollarea_viewport { background-color: transparent; }
    QWidget { background-color: rgba(255, 255, 255, 220); border: 1px solid #333; border-radius: 4px; }
"""

# ==============================================================================
#  UI COMPONENT: DRAGGABLE HEADER
# ==============================================================================
class DraggableHeader(QLabel):
    """Custom QLabel that allows the parent dock widget to be dragged."""
    
    def __init__(self, parent_dock, plugin_instance):
        super().__init__()
        self.parent_dock = parent_dock 
        self.plugin = plugin_instance 
        self.drag_start_position = None
        self.setText("Layer Info")
        self.setCursor(QCursor(Qt.OpenHandCursor))
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.plugin.show_context_menu)

    def set_mode(self, mode):
        if sip.isdeleted(self): return
        
        if mode == "standard":
            self.setFont(QFont("Segoe UI", 9, QFont.Bold))
            self.setFixedHeight(22)
            self.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.setStyleSheet(STYLE_STANDARD_LBL)
        else:
            self.setFont(QFont("Segoe UI", 9, QFont.Bold))
            self.setFixedHeight(25)
            self.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.setStyleSheet(STYLE_MINIMAL_LBL)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self.parent_dock and not sip.isdeleted(self.parent_dock):
                self.drag_start_position = event.globalPos() - self.parent_dock.frameGeometry().topLeft()
                self.setCursor(QCursor(Qt.ClosedHandCursor))
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if event.buttons() & Qt.LeftButton and self.drag_start_position:
            if self.parent_dock and not sip.isdeleted(self.parent_dock):
                self.parent_dock.move(event.globalPos() - self.drag_start_position)

    def mouseReleaseEvent(self, event):
        self.setCursor(QCursor(Qt.OpenHandCursor))


# ==============================================================================
#  MAIN PLUGIN CLASS
# ==============================================================================
class EmbedLegendPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.dock_widget = None
        self.list_widget = None
        self.settings = QgsSettings()
        
        # UI Properties
        self.bg_color = QColor(255, 255, 255)
        self.border_color = QColor(200, 200, 200)
        self.text_color = QColor("#2f3542")
        self.font_item = QFont("Segoe UI", 9)
        self.show_count = True
        self.show_percent = True
        self.style_mode = "minimalist" 
        self.lang_code = self.settings.value("EmbedLegend/Lang", "en")

    # --- Utilities ---
    def tr(self, key):
        """Translates text based on current language setting."""
        return LANG_DICT.get(self.lang_code, LANG_DICT["en"]).get(key, key)

    def initGui(self):
        icon_path = os.path.join(os.path.dirname(__file__), 'icon.png')
        self.action_toggle = QAction(QIcon(icon_path), 'Embed Legend Panel', self.iface.mainWindow())
        self.action_toggle.setCheckable(True)
        self.action_toggle.triggered.connect(self.run_toggle)
        self.iface.addToolBarIcon(self.action_toggle)
        self.iface.addPluginToMenu('&Embed Legend', self.action_toggle)

    def unload(self):
        self.disconnect_signals()
        self.iface.removePluginMenu('&Embed Legend', self.action_toggle)
        self.iface.removeToolBarIcon(self.action_toggle)
        self.cleanup_widget()

    def disconnect_signals(self):
        try: self.iface.layerTreeView().selectionModel().selectionChanged.disconnect(self.update_legend)
        except: pass
        try: self.iface.mapCanvas().mapCanvasRefreshed.disconnect(self.update_legend)
        except: pass

    def cleanup_widget(self):
        if self.dock_widget:
            try:
                self.iface.removeDockWidget(self.dock_widget)
                if not sip.isdeleted(self.dock_widget):
                    self.dock_widget.deleteLater()
            except: pass
            self.dock_widget = None

    # --- UI Initialization ---
    def run_toggle(self):
        if not self.dock_widget or sip.isdeleted(self.dock_widget):
            self.create_widget()
        
        vis = self.dock_widget.isVisible()
        if vis:
            self.dock_widget.close()
            self.dock_widget.setVisible(False)
        else:
            self.dock_widget.setVisible(True)
            self.update_legend()
        self.action_toggle.setChecked(not vis)

    def create_widget(self):
        self.cleanup_widget()
        
        # Setup Native Tool Window
        self.dock_widget = QDockWidget("Embed Legend", self.iface.mainWindow())
        self.dock_widget.setFloating(True)
        self.dock_widget.setTitleBarWidget(QWidget()) 
        self.dock_widget.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint)
        
        self.panel = QWidget()
        self.panel.setContextMenuPolicy(Qt.CustomContextMenu)
        self.panel.customContextMenuRequested.connect(self.show_context_menu)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(0,0,0,0) 
        layout.setSpacing(0) 
        
        self.header_widget = DraggableHeader(self.dock_widget, self)
        layout.addWidget(self.header_widget)
        
        self.list_widget = QListWidget()
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff) 
        self.list_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_widget.customContextMenuRequested.connect(self.show_context_menu)
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        
        layout.addWidget(self.list_widget)
        self.panel.setLayout(layout)
        self.dock_widget.setWidget(self.panel)
        
        self.apply_styles()
        self.iface.addDockWidget(Qt.LeftDockWidgetArea, self.dock_widget)
        
        self.disconnect_signals()
        self.iface.layerTreeView().selectionModel().selectionChanged.connect(self.update_legend)
        self.iface.mapCanvas().mapCanvasRefreshed.connect(self.update_legend)

    # --- Core Logic: Visibility Toggle ---
    def on_item_clicked(self, item):
        """Toggles visibility of the selected legend item (Strikethrough effect)."""
        if item.flags() & Qt.NoItemFlags: return # Ignore Header Items

        try:
            layer = item.data(Qt.UserRole)
            rule_key = item.data(Qt.UserRole + 1)
            if not layer or not layer.isValid(): return
            
            model = self.iface.layerTreeView().layerTreeModel()
            tree_layer = QgsProject.instance().layerTreeRoot().findLayer(layer.id())
            if not tree_layer: return
            
            # Find matching node and toggle check state
            nodes = model.layerLegendNodes(tree_layer)
            for node in nodes:
                node_key = node.data(Qt.UserRole)
                if node_key == rule_key or str(node_key) == str(rule_key):
                    current_state = node.data(Qt.CheckStateRole)
                    new_state = Qt.Unchecked if current_state == Qt.Checked else Qt.Checked
                    node.setData(new_state, Qt.CheckStateRole)
                    break
            
            layer.triggerRepaint()
            self.iface.mapCanvas().refresh()
            self.update_legend()
        except Exception:
            pass

    # --- Styles & Render ---
    def set_style_mode(self, mode):
        self.style_mode = mode
        if self.dock_widget and not sip.isdeleted(self.dock_widget):
            old_geo = self.dock_widget.geometry()
            self.create_widget() 
            self.update_legend()
            if self.dock_widget and not sip.isdeleted(self.dock_widget):
                self.dock_widget.setGeometry(old_geo)
                self.dock_widget.show()
        else:
            self.create_widget()
            self.dock_widget.show()

    def apply_styles(self):
        """Applies CSS styles based on selected mode."""
        if not self.dock_widget or sip.isdeleted(self.dock_widget): return
        
        self.header_widget.set_mode(self.style_mode)

        if self.style_mode == "standard":
            self.dock_widget.setAttribute(Qt.WA_TranslucentBackground, False)
            self.panel.setAttribute(Qt.WA_TranslucentBackground, False)
            panel_style = f"QWidget {{ background-color: {self.bg_color.name()}; border: 1px solid {self.border_color.name()}; }}"
            list_border = "none"; list_bg = "transparent"
        else:
            self.dock_widget.setAttribute(Qt.WA_TranslucentBackground, True)
            self.panel.setAttribute(Qt.WA_TranslucentBackground, False) 
            panel_style = STYLE_PANEL_MINIMAL
            list_border = "none"; list_bg = "transparent"

        self.list_widget.setStyleSheet(f"""
            QListWidget {{ background-color: {list_bg}; border: {list_border}; outline: none; spacing: 1px; }}
            QListWidget::item {{ height: 22px; padding-left: 8px; color: {self.text_color.name()}; }}
            QListWidget::item:selected {{ background-color: transparent; color: #3498db; font-weight: bold; }}
            QListWidget::item:hover {{ background-color: rgba(0,0,0,10); border-radius: 4px; }}
        """)
        self.panel.setStyleSheet(panel_style)

    def update_legend(self):
        """Refreshes the legend list based on ALL selected layers."""
        if not self.dock_widget or sip.isdeleted(self.dock_widget): return
        try:
            if not self.dock_widget.isVisible(): return
            
            self.list_widget.clear()
            
            # [FIX v6.9] Multi-Select Support
            selected_layers = self.iface.layerTreeView().selectedLayers()
            valid_layers = [l for l in selected_layers if isinstance(l, QgsVectorLayer)]
            
            if not valid_layers: return
            
            # Update Header Title
            if len(valid_layers) > 1:
                self.header_widget.setText(self.tr("header_multi").format(len(valid_layers)))
            else:
                if not sip.isdeleted(self.header_widget):
                    self.header_widget.setText(valid_layers[0].name())
            
            fm = QFontMetrics(self.font_item)
            max_width = 0
            total_items_count = 0
            
            # Loop through ALL valid selected layers
            for layer in valid_layers:
                renderer = layer.renderer()
                if not renderer: continue
                
                model = self.iface.layerTreeView().layerTreeModel()
                tree_layer = QgsProject.instance().layerTreeRoot().findLayer(layer.id())
                if not tree_layer: continue
                
                nodes = model.layerLegendNodes(tree_layer) if tree_layer else []
                r_items = renderer.legendSymbolItems()
                
                # Add Layer Separator (If multiple layers)
                if len(valid_layers) > 1:
                    sep_item = QListWidgetItem(f"◆ {layer.name()}")
                    sep_font = QFont(self.font_item)
                    sep_font.setBold(True)
                    sep_item.setFont(sep_font)
                    sep_item.setFlags(Qt.NoItemFlags) # Non-selectable
                    
                    # Styling Separator
                    if self.style_mode == "standard":
                        sep_item.setBackground(QColor("#dfe6e9"))
                        sep_item.setForeground(QColor("#2d3436"))
                    else:
                        sep_item.setForeground(QColor("#000000"))
                        
                    self.list_widget.addItem(sep_item)
                    total_items_count += 1
                    max_width = max(max_width, fm.horizontalAdvance(layer.name()) + 30)

                # Calculate totals for percentage (Per Layer)
                total_f = 0
                for n in nodes:
                    m = re.search(r"\[([\d\.,]+)\]", n.data(Qt.DisplayRole))
                    if m: total_f += int(m.group(1).replace('.', '').replace(',', ''))
                
                count = min(len(nodes), len(r_items))
                
                for i in range(count):
                    node = nodes[i]
                    raw = node.data(Qt.DisplayRole)
                    is_checked = node.data(Qt.CheckStateRole) == Qt.Checked
                    
                    # Format Text
                    match = re.search(r"\[([\d\.,]+)\]", raw)
                    cnt = int(match.group(1).replace('.', '').replace(',', '')) if match else 0
                    txt = raw if self.show_count else re.sub(r"\s*\[[\d\.,]+\]", "", raw)
                    if self.show_percent and total_f > 0: txt += f" ({(cnt/total_f)*100:.1f}%)"
                    
                    # Calc width for resizing
                    w = fm.horizontalAdvance(txt)
                    max_width = max(max_width, w)
                    
                    # Create Item
                    item = QListWidgetItem(QIcon(nodes[i].data(Qt.DecorationRole)), txt)
                    item.setData(Qt.UserRole, layer)
                    item.setData(Qt.UserRole + 1, r_items[i].ruleKey()) 
                    
                    # Apply Visual State (Strikethrough if unchecked)
                    current_font = QFont(self.font_item)
                    if not is_checked:
                        current_font.setStrikeOut(True)
                        item.setForeground(QColor("gray"))
                    else:
                        item.setForeground(self.text_color)
                    
                    item.setFont(current_font)
                    self.list_widget.addItem(item)
                    total_items_count += 1
            
            # Resize Panel Logic
            final_width = max(125, max_width + 55)
            # Dynamic Height based on items count
            self.dock_widget.setFixedWidth(final_width)
            self.dock_widget.resize(final_width, min(60 + (total_items_count * 22), 700))
            
        except RuntimeError: pass
        except Exception: pass

    # --- Menu & Interactions ---
    def show_context_menu(self, pos):
        if not self.dock_widget or sip.isdeleted(self.dock_widget): return
        menu = QMenu()
        menu.addAction(self.tr("menu_config")).setEnabled(False)
        menu.addAction(self.tr("menu_font")).triggered.connect(self.change_font)
        menu.addAction(self.tr("menu_text_color")).triggered.connect(self.change_text_color)
        
        act_bg = menu.addAction(self.tr("menu_bg_color"))
        act_bg.setEnabled(self.style_mode == "standard")
        act_bg.triggered.connect(self.change_bg)
        
        act_border = menu.addAction(self.tr("menu_border_color"))
        act_border.setEnabled(self.style_mode == "standard") 
        act_border.triggered.connect(self.change_border)
        
        menu.addSeparator()
        
        # Style Submenu
        submenu_style = menu.addMenu(self.tr("menu_style"))
        act_std = submenu_style.addAction(self.tr("style_std"))
        act_std.setCheckable(True)
        act_std.setChecked(self.style_mode == "standard")
        act_std.triggered.connect(lambda: self.set_style_mode("standard"))
        
        act_mini = submenu_style.addAction(self.tr("style_mini"))
        act_mini.setCheckable(True)
        act_mini.setChecked(self.style_mode == "minimalist")
        act_mini.triggered.connect(lambda: self.set_style_mode("minimalist"))

        menu.addSeparator()
        
        # Language Submenu
        submenu_lang = menu.addMenu(self.tr("lang"))
        act_en = submenu_lang.addAction("English")
        act_en.setCheckable(True)
        act_en.setChecked(self.lang_code == "en")
        act_en.triggered.connect(lambda: self.set_language("en"))
        
        act_id = submenu_lang.addAction("Bahasa Indonesia")
        act_id.setCheckable(True)
        act_id.setChecked(self.lang_code == "id")
        act_id.triggered.connect(lambda: self.set_language("id"))

        menu.addSeparator()
        
        # Display Options
        act_cnt = menu.addAction(self.tr("show_count"))
        act_cnt.setCheckable(True)
        act_cnt.setChecked(self.show_count)
        act_cnt.triggered.connect(lambda: self.update_data_state("count"))
        
        act_pct = menu.addAction(self.tr("show_percent"))
        act_pct.setCheckable(True)
        act_pct.setChecked(self.show_percent)
        act_pct.triggered.connect(lambda: self.update_data_state("percent"))
        
        menu.addSeparator()
        menu.addAction(self.tr("export_mif")).triggered.connect(self.export_manual_mif)
        menu.addAction(self.tr("export_kmz")).triggered.connect(self.export_kmz)
        menu.addSeparator()
        menu.addAction(self.tr("about")).triggered.connect(self.show_about)
        menu.exec_(QCursor.pos())

    # --- Actions ---
    def set_language(self, code):
        self.lang_code = code
        self.settings.setValue("EmbedLegend/Lang", code)
        self.update_legend()

    def update_data_state(self, key):
        if key == "count": self.show_count = not self.show_count
        else: self.show_percent = not self.show_percent
        self.update_legend()

    def change_font(self):
        f, ok = QFontDialog.getFont(self.font_item)
        if ok: self.font_item = f; self.update_legend()

    def change_text_color(self):
        c = QColorDialog.getColor(self.text_color)
        if c.isValid(): self.text_color = c; self.apply_styles(); self.update_legend()

    def change_bg(self):
        c = QColorDialog.getColor(self.bg_color)
        if c.isValid(): self.bg_color = c; self.apply_styles()

    def change_border(self):
        c = QColorDialog.getColor(self.border_color)
        if c.isValid(): self.border_color = c; self.apply_styles()

    def show_about(self):
        msg = QMessageBox(self.iface.mainWindow())
        msg.setWindowTitle(self.tr("about"))
        msg.setIcon(QMessageBox.Information)
        msg.setTextFormat(Qt.RichText)
        
        # HTML Styling for Global Audience + Bilingual Support
        text = (
            "<h3>Embed Legend Pro</h3>"
            "<b>Version:</b> 6.9.0<br>"
            "<b>Author:</b> Jujun Junaedi<br><br>"
            
            "<b>🚀 How to Use:</b><br>"
            "1. Select one or multiple vector layers in your Layers Panel.<br>"
            "2. Click the <i>Embed Legend</i> icon on the toolbar.<br>"
            "3. <b>Right-click</b> the panel to customize UI or export to KMZ/MIF.<br>"
            "4. <b>Left-click</b> items in the legend to toggle map visibility.<br><br>"
            
            "<b>☕ Support & Donate:</b><br>"
            "If this tool saves you hours of work, consider buying me a coffee!<br>"
            "• <b>Global:</b> Buy Me a Coffee / PayPal (junethtea)<br>"
            "• <b>Indonesia:</b> OVO / GoPay (081510027058)<br><br>"
            
            "<div style='background-color: #e8f4f8; padding: 10px; border-radius: 5px; text-align: center; color: #2d98da; border: 1px solid #bdc3c7;'>"
            "<b>💡 PRO TIP FOR SHARING 💡</b><br>"
            "<span style='font-size: 11px;'>"
            "To ensure your colleagues get the latest version without bugs, please share the <b>Official QGIS Plugin Link</b> or <b>GitHub Link</b> instead of raw ZIP files.<br><br>"
            "<i>Biar rekan kerjamu selalu dapat versi terbaru yang bebas error, yuk biasakan share link resmi QGIS/GitHub, bukan bagi-bagi file ZIP mentahan 😉</i>"
            "</span>"
            "</div><br>"
            
            "<hr>"
            "<p align='center' style='color: #636e72; font-size: 11px;'>"
            "<i>\"May this tool be a continuous charity (amal jariah),<br>especially for my beloved late parents. 🤲\"</i>"
            "</p>"
        )
        
        msg.setText(text)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
        
    # --- Export Engines ---
    def export_manual_mif(self):
        layer = self.iface.activeLayer()
        if not layer or not isinstance(layer, QgsVectorLayer):
            QMessageBox.warning(None, self.tr("warning"), self.tr("select_layer"))
            return
            
        mif_path, _ = QFileDialog.getSaveFileName(None, self.tr("export_mif"), "", "MapInfo Interchange (*.mif)")
        if not mif_path: return
        mid_path = os.path.splitext(mif_path)[0] + ".mid"
        
        total_feat = layer.featureCount()
        progress = QProgressDialog("Exporting MIF...", "Abort", 0, total_feat, self.iface.mainWindow())
        progress.setWindowModality(Qt.WindowModal); progress.setMinimumDuration(0)
        
        try:
            source_crs = layer.crs()
            dest_crs = QgsCoordinateReferenceSystem("EPSG:4326")
            tr = QgsCoordinateTransform(source_crs, dest_crs, QgsProject.instance())
            context = QgsRenderContext()
            renderer = layer.renderer()
            renderer.startRender(context, layer.fields())
            
            with open(mif_path, 'w', encoding='latin-1', errors='replace') as f_mif, \
                 open(mid_path, 'w', encoding='latin-1', errors='replace', newline='') as f_mid:
                
                # Header MIF
                f_mif.write("Version 300\nCharset \"WindowsLatin1\"\nDelimiter \",\"\nCoordSys Earth Projection 1, 104\n")
                fields = layer.fields()
                f_mif.write(f"Columns {len(fields)}\n")
                
                for field in fields:
                    col_name = "".join(x for x in field.name() if x.isalnum() or x == "_")[:10]
                    if not col_name: col_name = f"Col_{fields.indexOf(field)}"
                    f_type = "Char(254)"
                    if field.isNumeric():
                        if field.type() == QVariant.Int: f_type = "Integer"
                        elif field.type() == QVariant.Double: f_type = "Float"
                    f_mif.write(f"  {col_name} {f_type}\n")
                
                f_mif.write("Data\n\n")
                writer = csv.writer(f_mid, quotechar='"', quoting=csv.QUOTE_MINIMAL)
                
                for i, feat in enumerate(layer.getFeatures()):
                    if progress.wasCanceled(): break
                    progress.setValue(i)
                    try:
                        attrs = [str(a) if a != None else "" for a in feat.attributes()]
                        writer.writerow(attrs)
                        
                        geom = QgsGeometry(feat.geometry())
                        if not geom or geom.isEmpty(): f_mif.write("None\n"); continue
                        try: geom.transform(tr) 
                        except: pass 
                        
                        wkb_type = geom.wkbType()
                        sym = renderer.symbolForFeature(feat, context)
                        color_int = 0
                        if sym: 
                            c = sym.color()
                            color_int = (c.red() * 65536) + (c.green() * 256) + c.blue()
                            
                        # Geometry Handling
                        if QgsWkbTypes.geometryType(wkb_type) == QgsWkbTypes.PointGeometry:
                            pt = geom.asMultiPoint()[0] if geom.isMultipart() else geom.asPoint()
                            f_mif.write(f"Point {pt.x()} {pt.y()}\n")
                            f_mif.write(f'    Symbol (108, {color_int}, 8, "Wingdings", 0, 0)\n')
                        elif QgsWkbTypes.geometryType(wkb_type) == QgsWkbTypes.LineGeometry:
                            lines = geom.asMultiPolyline() if geom.isMultipart() else [geom.asPolyline()]
                            for line in lines:
                                f_mif.write(f"Pline {len(line)}\n")
                                for p in line: f_mif.write(f"{p.x()} {p.y()}\n")
                                f_mif.write(f"    Pen (2, 2, {color_int})\n")
                        elif QgsWkbTypes.geometryType(wkb_type) == QgsWkbTypes.PolygonGeometry:
                            all_rings = []
                            if geom.isMultipart():
                                for poly in geom.asMultiPolygon():
                                    for ring in poly: all_rings.append(ring)
                            else:
                                for ring in geom.asPolygon(): all_rings.append(ring)
                            f_mif.write(f"Region {len(all_rings)}\n")
                            for ring in all_rings:
                                f_mif.write(f"  {len(ring)}\n"); 
                                for p in ring: f_mif.write(f"    {p.x()} {p.y()}\n")
                            f_mif.write(f"    Pen (1, 2, {color_int})\n"); f_mif.write(f"    Brush (2, {color_int})\n")
                    except Exception: 
                        f_mif.write("None\n")
            
            renderer.stopRender(context)
            progress.setValue(total_feat)
            
            folder_path = os.path.dirname(mif_path)
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle(self.tr("success"))
            msg.setText(self.tr("export_success"))
            msg.setInformativeText(self.tr("file_saved").format(mif_path))
            btn_open = msg.addButton("Open Folder", QMessageBox.ActionRole)
            msg.addButton("Close", QMessageBox.RejectRole)
            msg.exec_()
            if msg.clickedButton() == btn_open: 
                QDesktopServices.openUrl(QUrl.fromLocalFile(folder_path))
        except Exception as e: 
            QMessageBox.critical(None, "Critical Error", str(e))

    def export_kmz(self, mode="auto"):
        layer = self.iface.activeLayer()
        if not layer: return
        
        path, _ = QFileDialog.getSaveFileName(None, self.tr("export_kmz"), "", "Google Earth (*.kmz)")
        if not path: return
        
        # STRICT IDENTIFIER DETECTION
        target_cols = ["SiteID", "Site_ID", "SITEID", "SITE_ID", "EnodeB", "eNB", "Site", "SiteName"]
        fields = layer.fields()
        field_names = [f.name() for f in fields]
        label_col = None
        
        for p in target_cols:
            for f in field_names:
                if f.lower() == p.lower():
                    label_col = f
                    break
            if label_col: break

        total_feat = layer.featureCount()
        progress = QProgressDialog("Exporting KMZ...", "Abort", 0, total_feat, self.iface.mainWindow())
        progress.setWindowModality(Qt.WindowModal); progress.setMinimumDuration(0)
        
        try:
            temp_dir = tempfile.mkdtemp()
            kml_path = os.path.join(temp_dir, "doc.kml")
            img_path = os.path.join(temp_dir, "legend.png")
            
            if self.dock_widget and self.dock_widget.isVisible():
                self.list_widget.clearSelection()
                self.dock_widget.grab().save(img_path, "PNG")
            
            context = QgsRenderContext()
            renderer = layer.renderer()
            renderer.startRender(context, layer.fields())
            tr = QgsCoordinateTransform(layer.crs(), QgsCoordinateReferenceSystem("EPSG:4326"), QgsProject.instance())
            
            kml_parts = ['<?xml version="1.0" encoding="UTF-8"?>', '<kml xmlns="http://www.opengis.net/kml/2.2">', '<Document>']
            labeled_sites = set() # Anti-overlap tracker

            for i, feat in enumerate(layer.getFeatures()):
                if progress.wasCanceled(): break
                progress.setValue(i)
                try: 
                    if not feat.hasGeometry(): continue
                    geom = feat.geometry()
                    sym = renderer.symbolForFeature(feat, context)
                    if not sym: continue
                    
                    color = sym.color()
                    kml_color = f"ff{color.blue():02x}{color.green():02x}{color.red():02x}"
                    poly_color = f"bf{color.blue():02x}{color.green():02x}{color.red():02x}"
                    
                    try: geom.transform(tr)
                    except: pass
                    
                    wkb_type = geom.wkbType()
                    
                    # Common Description Table
                    desc_table = "<table border='1' width='300'>"
                    for idx, val in enumerate(feat.attributes()): 
                        val_str = str(val) if val is not None else "-"
                        desc_table += f"<tr><td>{field_names[idx]}</td><td>{val_str}</td></tr>"
                    desc_table += "</table>"
                    
                    # Point Processing (Clean: No Name Label)
                    if QgsWkbTypes.geometryType(wkb_type) == QgsWkbTypes.PointGeometry:
                        p = geom.centroid(); pt = p.asPoint()
                        kml_parts.append('<Placemark>')
                        kml_parts.append('<name></name>') # Force Empty Name
                        kml_parts.append(f'<description><![CDATA[{desc_table}]]></description>')
                        kml_parts.append(f'<Style><IconStyle><color>{kml_color}</color><scale>0.7</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href></Icon></IconStyle></Style>')
                        kml_parts.append(f'<Point><coordinates>{pt.x()},{pt.y()},0</coordinates></Point></Placemark>')
                    
                    # Polygon Processing (Clean Grid: No Name Label)
                    elif QgsWkbTypes.geometryType(wkb_type) == QgsWkbTypes.PolygonGeometry:
                        kml_parts.append(f'<Placemark><name></name><description><![CDATA[{desc_table}]]></description>')
                        kml_parts.append(f'<Style><LineStyle><color>{kml_color}</color><width>1</width></LineStyle><PolyStyle><color>{poly_color}</color><fill>1</fill><outline>1</outline></PolyStyle></Style>')
                        
                        polys = geom.asMultiPolygon() if geom.isMultipart() else [geom.asPolygon()]
                        kml_parts.append('<MultiGeometry>')
                        for poly in polys:
                            outer_coords = " ".join([f"{p.x()},{p.y()},0" for p in poly[0]])
                            kml_parts.append(f'<Polygon><outerBoundaryIs><LinearRing><coordinates>{outer_coords}</coordinates></LinearRing></outerBoundaryIs>')
                            for i in range(1, len(poly)): 
                                inner_coords = " ".join([f"{p.x()},{p.y()},0" for p in poly[i]])
                                kml_parts.append(f'<innerBoundaryIs><LinearRing><coordinates>{inner_coords}</coordinates></LinearRing></innerBoundaryIs>')
                            kml_parts.append('</Polygon>')
                        kml_parts.append('</MultiGeometry></Placemark>')
                        
                        # --- Smart Labeling Logic (Strictly for Sectoral/Polygons with SiteID) ---
                        if label_col:
                            site_id = str(feat[label_col])
                            if site_id not in labeled_sites:
                                center = geom.centroid().asPoint()
                                kml_parts.append(f'''
                                <Placemark>
                                    <name>{site_id}</name>
                                    <Style>
                                        <IconStyle><scale>0</scale></IconStyle>
                                        <LabelStyle><scale>0.9</scale><color>ff00ffff</color></LabelStyle>
                                    </Style>
                                    <Point><coordinates>{center.x()},{center.y()},0</coordinates></Point>
                                </Placemark>
                                ''')
                                labeled_sites.add(site_id)

                    # Line Processing (Clean: No Name Label)
                    elif QgsWkbTypes.geometryType(wkb_type) == QgsWkbTypes.LineGeometry:
                        kml_parts.append(f'<Placemark><name></name><description><![CDATA[{desc_table}]]></description>')
                        kml_parts.append(f'<Style><LineStyle><color>{kml_color}</color><width>2</width></LineStyle></Style>')
                        
                        lines = geom.asMultiPolyline() if geom.isMultipart() else [geom.asPolyline()]
                        kml_parts.append('<MultiGeometry>')
                        for line in lines:
                            coords = " ".join([f"{p.x()},{p.y()},0" for p in line])
                            kml_parts.append(f'<LineString><coordinates>{coords}</coordinates></LineString>')
                        kml_parts.append('</MultiGeometry></Placemark>')

                except: continue 
            
            renderer.stopRender(context)
            progress.setValue(total_feat)
            
            # Pack KML + Image
            if os.path.exists(img_path): 
                kml_parts.append('<ScreenOverlay><name>Legend</name><Icon><href>legend.png</href></Icon><overlayXY x="0" y="1" xunits="fraction" yunits="fraction"/><screenXY x="0.01" y="0.99" xunits="fraction" yunits="fraction"/></ScreenOverlay>')
            kml_parts.append('</Document></kml>')
            
            with open(kml_path, "w", encoding="utf-8") as f: 
                f.write("\n".join(kml_parts))
            
            with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
                z.write(kml_path, "doc.kml")
                if os.path.exists(img_path): z.write(img_path, "legend.png")
            
            shutil.rmtree(temp_dir)
            folder_path = os.path.dirname(path)
            
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle(self.tr("success"))
            msg.setText(self.tr("export_success"))
            msg.setInformativeText(self.tr("file_saved").format(path))
            btn_open = msg.addButton("Open Folder", QMessageBox.ActionRole)
            msg.addButton("Close", QMessageBox.RejectRole)
            msg.exec_()
            if msg.clickedButton() == btn_open: 
                QDesktopServices.openUrl(QUrl.fromLocalFile(folder_path))
                
        except Exception as e: 
            QMessageBox.critical(None, "Error", str(e))