import os
import getpass
from datetime import datetime
from PyQt5.QtWidgets import (QAction, QFileDialog, QWidget, QVBoxLayout, QLabel, QLineEdit,
                             QPushButton, QComboBox, QHBoxLayout, QFormLayout, QLineEdit, QGroupBox, QDialog)
from PyQt5.QtGui import QIcon, QColor
from PyQt5.QtCore import Qt
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsRasterLayer,
    QgsPrintLayout, QgsLayoutItemMap, QgsReadWriteContext, QgsRectangle,
    QgsLayoutExporter, QgsLayoutItemRegistry, QgsLineSymbol, QgsSingleSymbolRenderer,
    QgsLayoutItemScaleBar, QgsUnitTypes, QgsLayerTreeLayer, QgsLayoutSize, QgsFillSymbol,
    QgsSimpleFillSymbolLayer, QgsSimpleLineSymbolLayer
)
from qgis.PyQt.QtXml import QDomDocument


class MapCraftPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        self.dialog = None

    def initGui(self):
        icon_path = os.path.join(self.plugin_dir, 'logo.png')
        self.action = QAction(QIcon(icon_path), 'MapCraft', self.iface.mainWindow())
        self.action.triggered.connect(self.open_dialog)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu('MapCraft', self.action)

    def unload(self):
        self.iface.removePluginMenu('MapCraft', self.action)
        self.iface.removeToolBarIcon(self.action)

    def open_dialog(self):
        if self.dialog is None:
            self.dialog = QWidget()
            self.dialog.setWindowTitle("MapCraft - Vattenfall Map Generator")

            # Top-level layout: horizontal split (left = form, right = help)
            main_layout = QHBoxLayout()
            self.dialog.setLayout(main_layout)

            # --- Left Side: Input Form ---
            form_widget = QWidget()
            form_layout = QVBoxLayout(form_widget)

            # Mode Selector
            self.mode_combo = QComboBox()
            self.mode_combo.addItems(["Automated", "Manual (Use QGIS layers)"])
            self.mode_combo.currentIndexChanged.connect(self.toggle_shp_inputs)
            form_layout.addWidget(QLabel("Map generation mode:"))
            form_layout.addWidget(self.mode_combo)

            # WTG SHP
            WTG_layout = QHBoxLayout()
            self.wtg_path = QLineEdit()
            browse_wtg_shp = QPushButton("Browse SHP")
            browse_wtg_shp.clicked.connect(self.browse_wtg_shp)
            WTG_layout.addWidget(QLabel("WTG layout:"))
            WTG_layout.addWidget(self.wtg_path)
            WTG_layout.addWidget(browse_wtg_shp)
            form_layout.addLayout(WTG_layout)

            # WTG Buffer File
            WTG_layout_Buff = QHBoxLayout()
            self.wtg_buff_path = QLineEdit()
            browse_wtgBuff_shp = QPushButton("Browse SHP")
            browse_wtgBuff_shp.clicked.connect(self.browse_wtgBuff_shp)
            WTG_layout_Buff.addWidget(QLabel("WTG buffer Layout:"))
            WTG_layout_Buff.addWidget(self.wtg_buff_path)
            WTG_layout_Buff.addWidget(browse_wtgBuff_shp)
            form_layout.addLayout(WTG_layout_Buff)

            # Site boundary
            site_boundary = QHBoxLayout()
            self.sibdry_path = QLineEdit()
            browse_sitebdry_shp = QPushButton("Browse SHP")
            browse_sitebdry_shp.clicked.connect(self.browse_sitebdry_shp)
            site_boundary.addWidget(QLabel("Site boundary:"))
            site_boundary.addWidget(self.sibdry_path)
            site_boundary.addWidget(browse_sitebdry_shp)
            form_layout.addLayout(site_boundary)

            # Site boundary buffer
            site_boundary_Buff = QHBoxLayout()
            self.sibdry_buff_path = QLineEdit()
            browse_sitebdryBuff_shp = QPushButton("Browse SHP")
            browse_sitebdryBuff_shp.clicked.connect(self.browse_sitebdryBuff_shp)
            site_boundary_Buff.addWidget(QLabel("Site boundary buffer:"))
            site_boundary_Buff.addWidget(self.sibdry_buff_path)
            site_boundary_Buff.addWidget(browse_sitebdryBuff_shp)
            form_layout.addLayout(site_boundary_Buff)

            # Wind priority area
            priority_area = QHBoxLayout()
            self.prio_area = QLineEdit()
            browse_priority_area_shp = QPushButton("Browse SHP")
            browse_priority_area_shp.clicked.connect(self.browse_priority_area_shp)
            priority_area.addWidget(QLabel("Wind priority area:"))
            priority_area.addWidget(self.prio_area)
            priority_area.addWidget(browse_priority_area_shp)
            form_layout.addLayout(priority_area)

            # Project Name
            self.project_name_input = QLineEdit()
            form_layout.addWidget(QLabel("Project name:"))
            form_layout.addWidget(self.project_name_input)

            # Map title
            self.Map_title_input = QLineEdit()
            self.Map_title_input.setText("Übersichtskarte")
            form_layout.addWidget(QLabel("Map title:"))
            form_layout.addWidget(self.Map_title_input)

            # Base Map
            self.basemap_combo = QComboBox()
            self.basemap_combo.addItems(["Topographic", "Satellite"])
            form_layout.addWidget(QLabel("Select base map type:"))
            form_layout.addWidget(self.basemap_combo)

            # State
            self.state_combo = QComboBox()
            self.state_combo.addItems(["Baden-Württemberg", "Niedersachsen", "Rheinland-Pfalz", "Schleswig-Holstein"])
            form_layout.addWidget(QLabel("Select german state:"))
            form_layout.addWidget(self.state_combo)

            # Scale
            self.scale_combo = QComboBox()
            self.scale_combo.addItems(["10000","15000", "25000", "50000"])
            form_layout.addWidget(QLabel("Map scale:"))
            form_layout.addWidget(self.scale_combo)

            # Output Folder
            pdf_layout = QHBoxLayout()
            self.pdf_path = QLineEdit()
            browse_pdf = QPushButton("Browse Folder")
            browse_pdf.clicked.connect(self.browse_pdf)
            pdf_layout.addWidget(QLabel("PDF Output Folder:"))
            pdf_layout.addWidget(self.pdf_path)
            pdf_layout.addWidget(browse_pdf)
            form_layout.addLayout(pdf_layout)

            # Format Selector
            self.format_combo = QComboBox()
            self.format_combo.addItems(["PDF", "PNG"])
            form_layout.addWidget(QLabel("Export Format:"))
            form_layout.addWidget(self.format_combo)

            # Reset Button
            reset_button = QPushButton("Reset")
            reset_button.clicked.connect(self.reset_fields)
            form_layout.addWidget(reset_button)

            # Run Button
            run_button = QPushButton("Run")
            run_button.clicked.connect(self.run_map_generation)
            form_layout.addWidget(run_button)

            # Add form widget to left side
            main_layout.addWidget(form_widget)

            # --- Right Side: Help Box ---
            help_group = QGroupBox("Help / Info")
            help_group.setStyleSheet(
                "QGroupBox { background-color: white; border: 1px solid lightgray; border-radius: 5px; }")

            help_label = QLabel("""
            📌 Description of Parameters:
            • SHP File: Upload the main shapefile for the wind park.
            • Optional SHP: Upload an optional area outline shapefile.
            • State: Select the federal state where the wind park is located.
            • Project Name: This name will be included in the map title and file name.
            • Scale: Choose the map scale (e.g., 1:25000).
            • Format: Select PDF or PNG as output format.
            • Output Folder: Choose where the file will be saved.

            🛈 Click 'Run' to generate the map automatically.
            """)
            help_label.setWordWrap(True)
            help_layout = QVBoxLayout()
            help_layout.addWidget(help_label)
            help_group.setLayout(help_layout)

            # Add help box to right side
            main_layout.addWidget(help_group)

            # Hide optional SHP if needed
            self.toggle_shp_inputs()

        self.dialog.show()

    def toggle_shp_inputs(self):
        is_automated = self.mode_combo.currentText() == "Automated"
        self.wtg_path.setEnabled(is_automated)
        self.wtg_buff_path.setEnabled(is_automated)
        self.sibdry_path.setEnabled(is_automated)
        self.sibdry_buff_path.setEnabled(is_automated)
        self.prio_area.setEnabled(is_automated)
        # Also disable the browse buttons
        for button in self.dialog.findChildren(QPushButton):
            if button.text() in ["Browse SHP"]:
                button.setEnabled(is_automated)

    def browse_wtg_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.wtg_path.setText(filename)

    def browse_wtgBuff_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.wtg_buff_path.setText(filename)

    def browse_sitebdry_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.sibdry_path.setText(filename)

    def browse_sitebdryBuff_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.sibdry_buff_path.setText(filename)

    def browse_priority_area_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.prio_area.setText(filename)

    def browse_pdf(self):
        filename = QFileDialog.getExistingDirectory(None, "Select Output Folder", "")
        if filename:
            self.pdf_path.setText(filename)

    def reset_fields(self):
        self.wtg_path.clear()
        self.wtg_buff_path.clear()
        self.sibdry_path.clear()
        self.sibdry_buff_path.clear()
        self.prio_area.clear()
        self.project_name_input.clear()
        # self.Map_title_input.clear()
        self.state_combo.setCurrentIndex(0)
        self.scale_combo.setCurrentIndex(0)
        self.pdf_path.clear()
        self.format_combo.setCurrentIndex(0)
        self.mode_combo.setCurrentIndex(0)
        self.toggle_shp_inputs()

    def load_wms_layer(self, state_selected, scale, basemap_type):
        """
        Loads a WMS or XYZ raster layer based on the selected state, scale, and basemap type.

        Args:
            state_selected (str): Name of the German federal state.
            scale (int or str): Map scale, e.g., 10000, 25000, 50000.
            basemap_type (str): Either "Topographic" or "Satellite".

        Returns:
            Tuple(QgsRasterLayer, dict or None, dict or None): The loaded raster layer, config (state or satellite), and scale config (or None for Satellite).
        """

        # --- State-specific topographic WMS configurations ---
        state_settings = {
            "Baden-Württemberg": {
                "scales": {
                    "10000": {
                        "wms_url": "https://owsproxy.lgl-bw.de/owsproxy/ows/WMS_LGL-BW_ATKIS_DTK_10_K?",
                        "layer_name": "RDS.LY_DTK10K_COL",
                        "title": "DTK10 Color"
                    },
                    "15000": {
                        "wms_url": "https://owsproxy.lgl-bw.de/owsproxy/ows/WMS_LGL-BW_ATKIS_DTK_10_K?",
                        "layer_name": "RDS.LY_DTK10K_COL",
                        "title": "DTK10 Color"
                    },
                    "25000": {
                        "wms_url": "https://owsproxy.lgl-bw.de/owsproxy/ows/WMS_LGL-BW_ATKIS_DTK_25_K_A?",
                        "layer_name": "RDS_LY_DTK25K_COL",
                        "title": "DTK25 Color"
                    },
                    "50000": {
                        "wms_url": "https://owsproxy.lgl-bw.de/owsproxy/ows/WMS_LGL-BW_ATKIS_DTK_50_K_A?",
                        "layer_name": "RDS_LY_DTK50K_COL",
                        "title": "DTK50 Color"
                    }
                },
                "copyright": "LGL-BW(2025) Datenlizenz Deutschland-Namensnennung-Version 2.0, www.lgl-bw.de"
            },
            "Rheinland-Pfalz": {
                "scales": {
                    "10000": {
                        "wms_url": "https://geo4.service24.rlp.de/wms/rp_dtk10.fcgi?",
                        "layer_name": "rp_dtk10",
                        "title": "DTK10 RLP"
                    },
                    "15000": {
                        "wms_url": "https://geo4.service24.rlp.de/wms/rp_dtk10.fcgi?",
                        "layer_name": "rp_dtk10",
                        "title": "DTK10 RLP"
                    },
                    "25000": {
                        "wms_url": "https://geo4.service24.rlp.de/wms/rp_dtk25.fcgi?",
                        "layer_name": "rp_dtk25",
                        "title": "DTK25 RLP"
                    },
                    "50000": {
                        "wms_url": "https://geo4.service24.rlp.de/wms/rp_dtk50.fcgi?",
                        "layer_name": "rp_dtk50",
                        "title": "DTK50 RLP"
                    },
                    "satellite": {
                        "wms_url": "https://www.geoportal.rlp.de/mapbender/php/wms.php?layer_id=61675",
                        "layer_name": "dop20",
                        "title": "RLP DOP20"
                    }
                },
                "copyright": "GeoBasis-DE/LVermGeoRP (2005) dl-de/by-2-0"
            },
            "Niedersachsen": {
                "scales": {
                    "10000": {},  # Not available
                    "25000": {
                        "wms_url": "https://www.geobasisdaten.niedersachsen.de/wms/dtk25?",
                        "layer_name": "DTK25",
                        "title": "DTK25 NI"
                    },
                    "50000": {
                        "wms_url": "https://www.geobasisdaten.niedersachsen.de/wms/dtk50?",
                        "layer_name": "DTK50",
                        "title": "DTK50 NI"
                    }
                },
                "copyright": "LGLN 2025"
            },
            "Schleswig-Holstein": {
                "scales": {
                    "10000": {
                        "wms_url": "https://service.gdi-sh.de/WMS_SH_DTK5_OpenGBD?",
                        "layer_name": "sh_dtk5_col",
                        "title": "DTK5 SH"
                    },
                    "15000": {
                        "wms_url": "https://service.gdi-sh.de/WMS_SH_DTK25_OpenGBD?",
                        "layer_name": "sh_dtk25_col",
                        "title": "DTK25 SH"
                    },
                    "25000": {
                        "wms_url": "https://service.gdi-sh.de/WMS_SH_DTK25_OpenGBD?",
                        "layer_name": "sh_dtk25_col",
                        "title": "DTK25 SH"
                    },
                    "50000": {
                        "wms_url": "https://service.gdi-sh.de/WMS_SH_DTK50_OpenGBD?",
                        "layer_name": "sh_dtk50_col",
                        "title": "DTK50 SH"
                    },
                    "satellite": {
                        "wms_url": "https://dienste.gdi-sh.de/WMS_SH_DOP20col_OpenGBD?",
                        "layer_name": "sh_dop20_col",
                        "title": "DOP20 SH"
                    }
                },
                "copyright": "GeoBasis-DE/LVermGeo 2025 SH/CC BY 4.0"
            }
        }

        # --- Global Satellite (XYZ) Basemap ---
        satellite_settings = {
            "basemap": {
                "zmin": 0,
                "zmax": 19,
                "crs": "EPSG:3857",
                "url": "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
                "title": "World Imagery"
            },
            "copyright": "Esri, Maxar, Earthstar Geographics, and the GIS User Community"
        }

        scale_str = str(scale)

        # --- Satellite Mode ---
        if basemap_type == "Satellite":
            satellite_conf = satellite_settings["basemap"]
            url = satellite_conf["url"]
            zmin = satellite_conf["zmin"]
            zmax = satellite_conf["zmax"]
            crs = satellite_conf["crs"]
            title = satellite_conf["title"]

            encoded_url = url.replace("=", "%3D").replace("&", "%26")
            uri = f"type=xyz&url={encoded_url}&zmin={zmin}&zmax={zmax}&crs={crs}"

            layer = QgsRasterLayer(uri, title, "wms")
            if not layer.isValid():
                self.iface.messageBar().pushCritical("MapCraft Plugin", "Could not load satellite basemap.")
                return None, None, None

            # layer.setOpacity(0.2)
            QgsProject.instance().addMapLayer(layer)

            # Return satellite settings so they can be used in layout updates
            return layer, satellite_settings, None

        # --- Topographic Mode ---
        elif basemap_type == "Topographic":
            state_conf = state_settings.get(state_selected)
            if not state_conf:
                self.iface.messageBar().pushCritical("MapCraft Plugin",
                                                     f"No configuration found for state '{state_selected}'.")
                return None, None, None

            scale_conf = state_conf["scales"].get(scale_str)
            if not scale_conf:
                self.iface.messageBar().pushCritical("MapCraft Plugin",
                                                     f"No topographic WMS for {state_selected} at scale {scale_str}.")
                return None, None, None

            wms_url = (
                f"contextualWMSLegend=0&crs=EPSG:25832&dpiMode=7&featureCount=10"
                f"&format=image/png&layers={scale_conf['layer_name']}&styles=&url={scale_conf['wms_url']}"
            )

            wms_layer = QgsRasterLayer(wms_url, f"{state_selected} Basemap", "wms")
            if not wms_layer.isValid():
                self.iface.messageBar().pushCritical("MapCraft Plugin",
                                                     f"Could not load WMS for {state_selected} at scale {scale_str}.")
                return None, None, None

            wms_layer.setOpacity(0.5)
            QgsProject.instance().addMapLayer(wms_layer)
            return wms_layer, state_conf, scale_conf

        else:
            self.iface.messageBar().pushCritical("MapCraft Plugin", f"Unknown basemap type: {basemap_type}")
            return None, None, None

    def adjust_legend_font_size(self, legend_item, max_items=5, base_size=6, min_size=6):
        # max_items:	Max number of legend entries before font starts shrinking
        # base_size:	Default font size when entries are ≤ max_items
        # min_size:	    Minimum font size allowed even if many entries

        from qgis.core import QgsLegendStyle

        if not legend_item:
            return

        # Count the total number of symbol entries
        model = legend_item.model()
        root_group = model.rootGroup()
        count = sum(len(group.children()) for group in root_group.children())

        # Reduce font size if too many entries
        if count > max_items:
            scale_factor = max(min_size, base_size - (count - max_items) // 2)
        else:
            scale_factor = base_size

        # Apply font size to legend styles
        for style in [QgsLegendStyle.Group, QgsLegendStyle.Subgroup, QgsLegendStyle.SymbolLabel]:
            font = legend_item.styleFont(style)
            font.setPointSize(scale_factor)
            legend_item.setStyleFont(style, font)

        legend_item.update()

    def run_map_generation(self):
        mode = self.mode_combo.currentText()
        if mode == "Automated":
            self.run_automated_map()
        else:
            self.run_manual_map()

    def run_automated_map(self):
        Layout = self.wtg_path.text()
        Layout_buff = self.wtg_buff_path.text()
        Site_Bdry = self.sibdry_path.text()
        Site_Bdry_buff = self.sibdry_buff_path.text()
        wind_prio_area = self.prio_area.text()
        project_name = self.project_name_input.text()
        Map_title = self.Map_title_input.text()
        basemap_type = self.basemap_combo.currentText()
        state_selected = self.state_combo.currentText()
        scale = int(self.scale_combo.currentText())
        export_format = self.format_combo.currentText()
        output_folder = self.pdf_path.text()

        today_name = datetime.today().strftime("%Y%m%d")
        pdf_filename = f"{today_name}_Windpark_{project_name}_{Map_title}"
        filename_base = os.path.join(output_folder, pdf_filename)

        if export_format == "PDF":
            output_path = os.path.join(self.pdf_path.text(), f"{filename_base}.pdf")
        else:
            output_path = os.path.join(self.pdf_path.text(), f"{filename_base}.png")


        if not self.wtg_path.text() or not project_name or not output_folder:
            print("Please complete all fields before running.")
            return


        layout_path = os.path.join(self.plugin_dir, "Übersichskarte_template_v5.qpt")
        style_path = os.path.join(self.plugin_dir, "WEA.qml")

        wms_layer, conf_dict, scale_conf = self.load_wms_layer(state_selected, scale, basemap_type)

        shp_layers_ref = []  # Create a list to be used a REF

        # Load WTG SHP
        layer_name = os.path.basename(Layout)  # This is to get the SHP name in the ref
        WTG_layer = QgsVectorLayer(Layout, layer_name, "ogr")
        if WTG_layer.isValid():

            # Load style and add to project
            WTG_layer.loadNamedStyle(style_path)
            WTG_layer.triggerRepaint()
            QgsProject.instance().addMapLayer(WTG_layer)
            shp_layers_ref.append(layer_name)


            # Load WTG Buffer SHP
            WTG_buff_layer = None
            if Layout_buff:
                layer_name_1 = os.path.basename(Layout_buff)  # Get the SHP name
                WTG_buff_layer = QgsVectorLayer(Layout_buff, layer_name_1, "ogr")

                if WTG_buff_layer.isValid():
                    # Create a transparent fill with red outline
                    symbol = QgsFillSymbol.createSimple({
                        'outline_color': 'blue',
                        'outline_width': '0.4',
                        'outline_style': 'dash dot',
                        'color': 'transparent'
                    })

                    WTG_buff_layer.setRenderer(QgsSingleSymbolRenderer(symbol))
                    WTG_buff_layer.triggerRepaint()
                    QgsProject.instance().addMapLayer(WTG_buff_layer)
                    shp_layers_ref.append(layer_name_1)

        # Load Site Boundary
        Site_Bdry_layer = None
        if Site_Bdry:
            layer_name_2 = os.path.basename(Site_Bdry)  # This is to get the SHP name in the ref
            Site_Bdry_layer = QgsVectorLayer(Site_Bdry, layer_name_2, "ogr")
            if Site_Bdry_layer.isValid():
                # Create a transparent fill with red outline
                symbol = QgsFillSymbol.createSimple({
                    'outline_color': 'red',
                    'outline_width': '0.5',
                    'color': 'transparent'
                })
                Site_Bdry_layer.setRenderer(QgsSingleSymbolRenderer(symbol))
                Site_Bdry_layer.triggerRepaint()
                QgsProject.instance().addMapLayer(Site_Bdry_layer)
                shp_layers_ref.append(layer_name_2)

        # Load Site Boundary Buffer
        Site_Bdry_buff_layer = None
        if Site_Bdry_buff:
            layer_name_3 = os.path.basename(Site_Bdry_buff)  # Get the SHP name
            Site_Bdry_buff_layer = QgsVectorLayer(Site_Bdry_buff, layer_name_3, "ogr")

            if Site_Bdry_buff_layer.isValid():
                # Create the bottom stroke: thick, light red, semi-transparent
                bottom_line = QgsSimpleLineSymbolLayer()
                bottom_line.setColor(QColor(153, 0, 0, 128))  # Light red with 50% transparency
                bottom_line.setWidth(1.5)

                # Create the top stroke: thinner, dark red
                top_line = QgsSimpleLineSymbolLayer()
                top_line.setColor(QColor(153, 0, 0))  # Dark red
                top_line.setWidth(0.3)

                # Create a transparent fill symbol and disable the default fill
                symbol = QgsFillSymbol.createSimple({'color': 'transparent'})
                symbol.symbolLayer(0).setEnabled(False)

                # Add the two outline layers
                symbol.appendSymbolLayer(bottom_line)
                symbol.appendSymbolLayer(top_line)

                # Apply the symbol to the layer
                Site_Bdry_buff_layer.setRenderer(QgsSingleSymbolRenderer(symbol))
                Site_Bdry_buff_layer.triggerRepaint()

                # Add the layer to the project
                QgsProject.instance().addMapLayer(Site_Bdry_buff_layer)
                shp_layers_ref.append(layer_name_3)

        # Load wind priority area
        priority_area_layer = None
        if wind_prio_area:
            layer_name_4 = os.path.basename(wind_prio_area)  # Get the SHP name
            priority_area_layer = QgsVectorLayer(wind_prio_area, layer_name_4, "ogr")

            if priority_area_layer.isValid():
                # Create the bottom stroke: thick, light red, semi-transparent
                bottom_line = QgsSimpleLineSymbolLayer()
                bottom_line.setColor(QColor(200, 160, 255, 128))  # Light purple with 50% transparency
                bottom_line.setWidth(1.5)

                # Create the top stroke: thinner, dark red
                top_line = QgsSimpleLineSymbolLayer()
                top_line.setColor(QColor(76, 0, 153))  # Dark purple
                top_line.setWidth(0.3)

                # Create a transparent fill symbol and disable the default fill
                symbol = QgsFillSymbol.createSimple({'color': 'transparent'})
                symbol.symbolLayer(0).setEnabled(False)

                # Add the two outline layers
                symbol.appendSymbolLayer(bottom_line)
                symbol.appendSymbolLayer(top_line)

                # Apply the symbol to the layer
                priority_area_layer.setRenderer(QgsSingleSymbolRenderer(symbol))
                priority_area_layer.triggerRepaint()

                # Add the layer to the project
                QgsProject.instance().addMapLayer(priority_area_layer)
                shp_layers_ref.append(layer_name_4)

        # Load Layout
        with open(layout_path, 'r') as f:
            template_content = f.read()
        document = QDomDocument()
        document.setContent(template_content)

        layout = QgsPrintLayout(QgsProject.instance())
        layout.initializeDefaults()
        layout.loadFromTemplate(document, QgsReadWriteContext())

        # Map Item
        map_item = next((item for item in layout.items() if isinstance(item, QgsLayoutItemMap) and item.id() == "Map"),
                        None)

        if map_item:
            map_item.setScale(scale)
            map_width_m = (map_item.rect().width() * scale) / 1000
            map_height_m = (map_item.rect().height() * scale) / 1000
            center = WTG_layer.extent().center()
            extent = QgsRectangle(center.x() - map_width_m / 2, center.y() - map_height_m / 2,
                                  center.x() + map_width_m / 2, center.y() + map_height_m / 2)
            map_item.setExtent(extent)
            map_item.refresh()

            # === SCALE BAR SETUP ===
            scale_bar_item = layout.itemById('scale')
            if isinstance(scale_bar_item, QgsLayoutItemScaleBar):
                scale_bar_item.setStyle('Line Ticks Up')
                scale_bar_item.setUnits(QgsUnitTypes.DistanceKilometers)
                scale_bar_item.setNumberOfSegments(2)
                scale_bar_item.setNumberOfSegmentsLeft(0)
                scale_bar_item.setLinkedMap(map_item)

                if scale == 10000:
                    scale_bar_item.setUnitsPerSegment(0.25)
                elif scale == 15000:
                    scale_bar_item.setUnitsPerSegment(0.3)  # 0.3 km per segment
                elif scale == 25000:
                    scale_bar_item.setUnitsPerSegment(0.5)  # 0.5 km per segment
                elif scale == 50000:
                    scale_bar_item.setUnitsPerSegment(1.0)

            # === LEGEND SETUP ===
            legend_item = layout.itemById("simbology")  # Make sure your layout legend ID is 'simbology'

            if legend_item and map_item:
                legend_item.setLinkedMap(map_item)

                # Disable auto-update to manually control legend entries
                legend_item.setAutoUpdateModel(False)

                # Access the legend model and clear all current entries
                legend_model = legend_item.model()
                root_group = legend_model.rootGroup()
                root_group.removeAllChildren()

                # Add only the SHP layer manually
                root_group.addLayer(WTG_layer)

                # Rename the legend label for the SHP WTG_layer
                for child in root_group.findLayers():
                    if child.layer() == WTG_layer:
                        child.setName("WEA - Neuplanung")  # Custom name shown in legend

                legend_item.refresh()

                # Add optional layer if available
                if WTG_buff_layer:
                    root_group.addLayer(WTG_buff_layer)
                    for child in root_group.findLayers():
                        if child.layer() == WTG_buff_layer:
                            child.setName("Rotor")
                if Site_Bdry_layer:
                    root_group.addLayer(Site_Bdry_layer)
                    for child in root_group.findLayers():
                        if child.layer() == Site_Bdry_layer:
                            child.setName("Projektfläche")
                if Site_Bdry_buff_layer:
                    root_group.addLayer(Site_Bdry_buff_layer)
                    for child in root_group.findLayers():
                        if child.layer() == Site_Bdry_buff_layer:
                            child.setName("Projektfläche Puffer")
                if priority_area_layer:
                    root_group.addLayer(priority_area_layer)
                    for child in root_group.findLayers():
                        if child.layer() == priority_area_layer:
                            child.setName("Potenzialfläche")

        # Dynamic Labels

        projection = WTG_layer.crs().description()
        today = datetime.today().strftime("%d/%m/%y")
        username = getpass.getuser()
        ref_text = " | ".join(shp_layers_ref)
        if basemap_type == "Topographic" and conf_dict:
            copyright_text = conf_dict.get("copyright", "")
        elif basemap_type == "Satellite" and conf_dict:
            copyright_text = conf_dict.get("copyright", "")

        for item in layout.items():
            if item.type() == QgsLayoutItemRegistry.LayoutLabel:
                if item.id() == 'label_proj':
                    item.setText(f"CRS: {projection}")
                elif item.id() == 'label_creator':
                    item.setText(f"Map produced on {today} by {username}")
                elif item.id() == 'label_title':
                    item.setText(f"{Map_title}")
                elif item.id() == 'label_Windpark':
                    item.setText(f"Windpark {project_name}")
                elif item.id() == 'label_ref':
                    item.setText(f"Ref: {ref_text}")
                elif item.id() == 'label_CR':
                    item.setText(f"Hintergrund: ©{copyright_text}")
                    font = item.font()
                    font.setPointSize(4)
                    item.setFont(font)
                    item.refresh()


        # Export based on selected format
        exporter = QgsLayoutExporter(layout)

        if export_format == "PDF":
            pdf_settings = QgsLayoutExporter.PdfExportSettings()
            result = exporter.exportToPdf(output_path, pdf_settings)
            if result == QgsLayoutExporter.Success:
                self.iface.messageBar().pushSuccess('Success', 'PDF exported successfully!')
            else:
                self.iface.messageBar().pushCritical('Error', 'PDF export failed.')

        elif export_format == "PNG":
            image_settings = QgsLayoutExporter.ImageExportSettings()
            image_settings.dpi = 300  # ✅ Set high resolution
            result = exporter.exportToImage(output_path, image_settings)
            if result == QgsLayoutExporter.Success:
                self.iface.messageBar().pushSuccess('Success', 'PNG exported successfully!')
            else:
                self.iface.messageBar().pushCritical('Error', 'PNG export failed.')

        # ✅ Remove added layers from canvas
        QgsProject.instance().removeMapLayer(WTG_layer)
        QgsProject.instance().removeMapLayer(wms_layer)
        if WTG_buff_layer:
            QgsProject.instance().removeMapLayer(WTG_buff_layer)
        elif Site_Bdry_layer:
            QgsProject.instance().removeMapLayer(Site_Bdry_layer)
        elif Site_Bdry_buff_layer:
            QgsProject.instance().removeMapLayer(Site_Bdry_buff_layer)
        elif priority_area_layer:
            QgsProject.instance().removeMapLayer(priority_area_layer)


        self.iface.mapCanvas().refresh()

    def run_manual_map(self):
        # Basic validation
        project_name = self.project_name_input.text()
        Map_title = self.Map_title_input.text()
        state_selected = self.state_combo.currentText()
        basemap_type = self.basemap_combo.currentText()
        scale = int(self.scale_combo.currentText())
        export_format = self.format_combo.currentText()
        output_folder = self.pdf_path.text()
        today_name = datetime.today().strftime("%Y%m%d")
        pdf_filename = f"{today_name}_Windpark_{project_name}_{Map_title}"
        output_path = os.path.join(output_folder, f"{pdf_filename}.{export_format.lower()}")

        layout_path = os.path.join(self.plugin_dir, "Übersichskarte_template_v5.qpt")

        # Load template
        with open(layout_path, 'r') as f:
            template_content = f.read()
        document = QDomDocument()
        document.setContent(template_content)
        layout = QgsPrintLayout(QgsProject.instance())
        layout.initializeDefaults()
        layout.loadFromTemplate(document, QgsReadWriteContext())

        # Get the first layer of the project
        layers = list(QgsProject.instance().mapLayers().values())
        layer = layers[0] # Select the first layer

        wms_layer, conf_dict, scale_conf = self.load_wms_layer(state_selected, scale, basemap_type)

        # Move WMS layer to the bottom in Layer Tree
        layer_tree = QgsProject.instance().layerTreeRoot()
        wms_node = layer_tree.findLayer(wms_layer.id())
        if wms_node:
            wms_layer_ref = wms_node.layer()  # Keep reference to the QgsMapLayer
            layer_tree.removeChildNode(wms_node)
            new_wms_node = QgsLayerTreeLayer(wms_layer_ref)
            layer_tree.insertChildNode(0, new_wms_node)

        # Map item
        map_item = next((item for item in layout.items() if isinstance(item, QgsLayoutItemMap) and item.id() == "Map"),
                        None)

        if not map_item:
            self.iface.messageBar().pushCritical("Error", "Map item with ID 'Map' not found.")
            return

        # Set scale and extent
        map_item.setScale(scale)
        map_width_m = (map_item.rect().width() * scale) / 1000
        map_height_m = (map_item.rect().height() * scale) / 1000
        center = layer.extent().center()
        extent = QgsRectangle(center.x() - map_width_m / 2, center.y() - map_height_m / 2,
                              center.x() + map_width_m / 2, center.y() + map_height_m / 2)
        map_item.setExtent(extent)

        # Organize layers: WMS at the bottom
        all_layers = list(QgsProject.instance().mapLayers().values())
        vector_layers = [l for l in all_layers if isinstance(l, QgsVectorLayer)]
        wms_layers = [l for l in all_layers if isinstance(l, QgsRasterLayer) and 'wms' in l.source().lower()]
        final_order = vector_layers + wms_layers
        map_item.setLayers(final_order)

        map_item.refresh()

        # Get only visible layers (including the WMS layer if visible)
        layer_tree = QgsProject.instance().layerTreeRoot()
        visible_layers = []

        for child in layer_tree.children():
            if isinstance(child, QgsLayerTreeLayer) and child.isVisible():
                visible_layers.append(child.layer())

        # Layer ordering: WMS/raster layers at the bottom
        vector_layers = [l for l in visible_layers if isinstance(l, QgsVectorLayer)]
        final_order = vector_layers + [wms_layer]

        map_item.setLayers(final_order)
        map_item.refresh()

        # === SCALE BAR SETUP ===
        scale_bar_item = layout.itemById('scale')
        if isinstance(scale_bar_item, QgsLayoutItemScaleBar):
            scale_bar_item.setStyle('Line Ticks Up')
            scale_bar_item.setUnits(QgsUnitTypes.DistanceKilometers)
            scale_bar_item.setNumberOfSegments(2)
            scale_bar_item.setNumberOfSegmentsLeft(0)
            scale_bar_item.setLinkedMap(map_item)

            if scale == 10000:
                scale_bar_item.setUnitsPerSegment(0.25)
            elif scale == 15000:
                scale_bar_item.setUnitsPerSegment(0.3)  # 0.25 km per segment
            elif scale == 25000:
                scale_bar_item.setUnitsPerSegment(0.5)  # 0.5 km per segment
            elif scale == 50000:
                scale_bar_item.setUnitsPerSegment(1.0)

        # Legend setup

        legend_item = layout.itemById("simbology")
        if legend_item and map_item:
            legend_item.setLinkedMap(map_item)
            legend_item.setAutoUpdateModel(False)

            legend_model = legend_item.model()
            root_group = legend_model.rootGroup()
            root_group.removeAllChildren()

            # Access the current layer tree
            layer_tree = QgsProject.instance().layerTreeRoot()

            # Add only visible vector layers (skip WMS/raster/invisible)
            for child in layer_tree.children():
                if isinstance(child, QgsLayerTreeLayer) and child.isVisible():
                    layer = child.layer()
                    if isinstance(layer, QgsVectorLayer):
                        root_group.addLayer(layer)


            legend_item.refresh()

            # Adjust font size based on number of legend entries
            legend_item.attemptResize(QgsLayoutSize(55, 100))
            self.adjust_legend_font_size(legend_item)

        # Create a list to be used a REF
        shp_layers_ref = []
        layer_tree = QgsProject.instance().layerTreeRoot()

        for child in layer_tree.children():
            if isinstance(child, QgsLayerTreeLayer) and child.isVisible():
                layer = child.layer()
                if isinstance(layer, QgsVectorLayer):
                    layer_path = layer.source()
                    layer_name = os.path.basename(layer_path)
                    shp_layers_ref.append(layer_name)

        # Dynamic labels
        username = getpass.getuser()
        today = datetime.today().strftime("%d/%m/%y")
        projection = layer.crs().description()
        ref_text = " | ".join(shp_layers_ref)
        if basemap_type == "Topographic" and conf_dict:
            copyright_text = conf_dict.get("copyright", "")
        elif basemap_type == "Satellite" and conf_dict:
            copyright_text = conf_dict.get("copyright", "")


        for item in layout.items():
            if item.type() == QgsLayoutItemRegistry.LayoutLabel:
                if item.id() == 'label_proj':
                    item.setText(f"CRS: {projection}")
                elif item.id() == 'label_creator':
                    item.setText(f"Map produced on {today} by {username}")
                elif item.id() == 'label_title':
                    item.setText(f"{Map_title}")
                elif item.id() == 'label_Windpark':
                    item.setText(f"Windpark {project_name}")
                elif item.id() == 'label_ref':
                    item.setText(f"Ref: {ref_text}")
                elif item.id() == 'label_CR':
                    item.setText(f"Hintergrund: ©{copyright_text}")

        # Export
        exporter = QgsLayoutExporter(layout)
        if export_format == "PDF":
            result = exporter.exportToPdf(output_path, QgsLayoutExporter.PdfExportSettings())
        else:
            settings = QgsLayoutExporter.ImageExportSettings()
            settings.dpi = 300
            result = exporter.exportToImage(output_path, settings)

        if result == QgsLayoutExporter.Success:
            self.iface.messageBar().pushSuccess('Success', f'{export_format} exported successfully!')
        else:
            self.iface.messageBar().pushCritical('Error', f'{export_format} export failed.')
