import os
import getpass
from datetime import datetime
from PyQt5.QtWidgets import (QAction, QFileDialog, QWidget, QVBoxLayout, QLabel, QLineEdit,
                             QPushButton, QComboBox, QHBoxLayout, QFormLayout, QLineEdit,
                             QGroupBox, QDialog, QScrollArea, QWidget, QCheckBox)
from PyQt5.QtGui import QIcon, QColor, QFontMetricsF, QFont
from PyQt5.QtCore import Qt, QSizeF, QRectF
from qgis.utils import iface
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsRasterLayer,
    QgsPrintLayout, QgsLayoutItemMap, QgsReadWriteContext, QgsRectangle,
    QgsLayoutExporter, QgsLayoutItemRegistry, QgsLineSymbol, QgsSingleSymbolRenderer,
    QgsLayoutItemScaleBar, QgsUnitTypes, QgsLayerTreeLayer, QgsLayoutSize, QgsFillSymbol,
    QgsSimpleFillSymbolLayer, QgsSimpleLineSymbolLayer, QgsLayoutPoint, QgsLayerTreeGroup
)
from qgis.PyQt.QtXml import QDomDocument

# START OF PLUG-IN CONFIGURATION
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
            # self.wtg_path.setText("C:/Users/cfp29/Documents/Testing_SHP/DEWTL01WN_DWTG_LWTL01007_v01_250509jmmr25832.shp")  # Deactivate
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

            # WTG Buffer Size Input (initially hidden)
            self.wtg_buff_size_input = QLineEdit()
            self.wtg_buff_size_input.setPlaceholderText("Enter WTG buffer size (e.g. 87.5)")
            self.wtg_buff_size_input.hide()
            form_layout.addWidget(self.wtg_buff_size_input)

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

            # Site Boundary Buffer Size Input
            self.sibdry_buff_size_input = QLineEdit()
            self.sibdry_buff_size_input.setPlaceholderText("Enter Site boundary buffer size (e.g. 100)")
            self.sibdry_buff_size_input.hide()
            form_layout.addWidget(self.sibdry_buff_size_input)

            # Wind priority area
            priority_area = QHBoxLayout()
            self.priory_area = QLineEdit()
            browse_priority_area_shp = QPushButton("Browse SHP")
            browse_priority_area_shp.clicked.connect(self.browse_priority_area_shp)
            priority_area.addWidget(QLabel("Wind priority area:"))
            priority_area.addWidget(self.priory_area)
            priority_area.addWidget(browse_priority_area_shp)
            form_layout.addLayout(priority_area)

            # Potential wind area
            potential_area = QHBoxLayout()
            self.potential_area = QLineEdit()
            browse_potential_area_shp = QPushButton("Browse SHP")
            browse_potential_area_shp.clicked.connect(self.browse_potential_area_shp)
            potential_area.addWidget(QLabel("Wind potential area:"))
            potential_area.addWidget(self.potential_area)
            potential_area.addWidget(browse_potential_area_shp)
            form_layout.addLayout(potential_area)

            # Project Name
            self.project_name_input = QLineEdit()
            # self.project_name_input.setText("Winterlingen") # Deactivate
            form_layout.addWidget(QLabel("Project name:"))
            form_layout.addWidget(self.project_name_input)

            # Map title
            self.Map_title_input = QLineEdit()
            self.Map_title_input.setMaxLength(95) # Limit the number of characters
            self.Map_title_input.setText("Übersichtskarte")
            form_layout.addWidget(QLabel("Map title:"))
            form_layout.addWidget(self.Map_title_input)

            # Layout Size Selector
            self.layout_size_combo = QComboBox()
            self.layout_size_combo.addItems(["A3", "A4", ])  # Add more if needed
            form_layout.addWidget(QLabel("Map layout size:"))
            form_layout.addWidget(self.layout_size_combo)

            # Base Map
            self.basemap_combo = QComboBox()
            self.basemap_combo.addItems(["Topographic", "Satellite", "OpenStreetMap"])
            form_layout.addWidget(QLabel("Select base map type:"))
            form_layout.addWidget(self.basemap_combo)

            # State
            self.state_combo = QComboBox()
            self.state_combo.addItems(["Baden-Württemberg", "Niedersachsen", "Rheinland-Pfalz", "Schleswig-Holstein"])
            form_layout.addWidget(QLabel("Select german state:"))
            form_layout.addWidget(self.state_combo)

            # Scale
            self.scale_combo = QComboBox()
            self.scale_combo.addItems(["25000", "10000","15000", "50000"])
            form_layout.addWidget(QLabel("Map scale:"))
            form_layout.addWidget(self.scale_combo)

            # Output Folder
            pdf_layout = QHBoxLayout()
            self.pdf_path = QLineEdit()
            # self.pdf_path.setText("C:/Users/cfp29/Downloads/map")  # Deactivate
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

            self.keepLayersCheckBox = QCheckBox("Keep layers in QGIS after Map exporting")
            form_layout.addWidget(self.keepLayersCheckBox)

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
            help_group = QGroupBox("Do you need some help?:")
            help_group.setStyleSheet(
                "QGroupBox { background-color: white; border: 1px solid lightgray; border-radius: 5px; }")

            help_label = QLabel("""
                <b>Description of Parameters</b><br><br><br>

                <b>Map Generation Mode</b> <i>(Required)</i> Select how the map will be created:<br>
                
                <ul>
                    <li><b>Automatic:</b> Uses predefined symbology with minimal input.</li>
                    <li><b>Manual:</b> Allows full control over symbology and labels using user-supplied shapefiles.<br>
                    In Manual mode, users must upload their own shapefiles to define map content.</li>
                </ul><br>

                <b>Input Layers</b><br>
                <span style="color:red;">IMPORTANT:</span> All Shapefiles must be projected.<br>
                <ul>
                    <li><b>WTG Layout:</b> <i>(Required)</i> Wind Turbine Generator layout.<br>
                    <span style="color:red;">IMPORTANT:</span> This shapefile should be the one produced by the GIS team.<br>
                    The tool requires the fields [TRB_ID] and [LAYOUT] to generate map legends and labels. If these fields are missing, the tool will fail.</li><br>

                    <li><b>WTG Buffer Layout:</b> <i>(Optional)</i> Shapefile representing a buffer area around the WTGs. If a Shapefile is provided, the buffer area distance must be also registered (e.g., 87.5, 90).</li><br>

                    <li><b>Site Boundary:</b> <i>(Optional)</i> Shapefile defining the project site boundary.</li><br>

                    <li><b>Site Boundary Buffer:</b> <i>(Optional)</i> Shapefile defining a buffer around the site boundary. If a Shapefile is provided, the buffer area distance must be also registered (e.g., 87.5, 90).</li><br>
                    
                    <li><b>Wind Priority Area (Windvorranggebiet):</b> <i>(Optional)</i> A legally designated area in regional or land-use plans where wind energy has priority.</li><br>
                    
                    <li><b>Wind Potential Area (Potenzialfläche):</b> <i>(Optional)</i> Area that has been identified as suitable for wind energy based on different evaluations.</li>
                </ul><br>

                <b>Project Metadata</b><br>
                <ul>
                    <li><b>Project Name:</b> <i>(Required)</i> Name of the project (e.g., Winterlingen).</li><br>
                    <li><b>Map Title:</b> <i>(Required)</i> Custom title to be displayed on the exported map.</li>
                </ul><br>

                <b>Map Settings</b><br>
                <ul>
                    <li><b>Map Layout Size:</b> <i>(Required)</i> Select the paper size for the map layout (e.g., A3, A4).</li><br>
                    <li><b>Select Base Map Type:</b> <i>(Required)</i> Select the background map to use (e.g., topographic, satellite).</li><br>
                    <li><b>Select German State:</b> <i>(Required)</i> Select the federal state where the project is located. This determines which WMS basemap will be used.</li><br>
                    <li><b>Map Scale:</b> <i>(Required)</i> Define the desired map scale (e.g., 1:25,000 or 1:50,000).</li>
                </ul><br>

                <b>Output Options</b><br>
                <ul>
                    <li><b>PDF Output Folder:</b> <i>(Required)</i> Select the folder where the exported map (PDF/PNG) will be saved.</li><br>
                    <li><b>Export Format:</b> <i>(Required)</i> Select the format of the output file (PDF or PNG).</li>
                </ul><br>

                <b>Actions</b><br>
                <ul>
                    <li><b>Keep layers in QGIS after Map exporting:</b> Use this option if you want to retain the layers used to create the map in your QGIS project.:</b> Use this option if you want to retain the layers used in the QGIS project after exporting the map.</li><br>
                    <li><b>Reset:</b> Clears all fields and selections in the form.</li><br>
                    <li><b>Run:</b> Starts the map generation process.</li>
                </ul><br>
                
                <b>Need more help? Keine Sorgen</b><br>
                <ul>
                    <li><a href="https://vattenfall.sharepoint.com/sites/Wind_OnDpt_WNMX/GISTeam/SitePages/Home.aspx" style="color:blue;" target="_blank">Open SharePoint Documentation</a></li><br>
                    <li><a href="https://emea01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fapps.powerapps.com%2Fplay%2Fe%2Fdefault-f8be18a6-f648-4a47-be73-86d6c5c6604d%2Fa%2Fffaf49bb-9017-4ae2-843b-a1042eb8cf5f%3FtenantId%3Df8be18a6-f648-4a47-be73-86d6c5c6604d%26source%3Demail&data=05%7C02%7Cjosemanuel.mendozareyes%40vattenfall.de%7C22154ba212974807228108dd357c43a9%7Cf8be18a6f6484a47be7386d6c5c6604d%7C0%7C0%7C638725529994715332%7CUnknown%7CTWFpbGZsb3d8eyJFbXB0eU1hcGkiOnRydWUsIlYiOiIwLjAuMDAwMCIsIlAiOiJXaW4zMiIsIkFOIjoiTWFpbCIsIldUIjoyfQ%3D%3D%7C0%7C%7C%7C&sdata=rFa7XdY4gkeD7TfZfVM%2Fe20Xt5phu4FMNC4NV0j%2FdUY%3D&reserved=0" style="color:blue;" target="_blank">Report a Problem (GIS Ticket System)</a></li>
                </ul>
            """)

            help_label.setWordWrap(True)
            help_label.setTextFormat(Qt.RichText)
            help_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
            help_label.setOpenExternalLinks(True)

            # Add scroll area
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            scroll_area.setWidget(help_label)
            scroll_area.setFixedHeight(550) # Adjust height as needed

            help_layout = QVBoxLayout()
            help_layout.addWidget(scroll_area)  # Add scroll area instead of the label
            help_group.setLayout(help_layout)

            # Add help box to right side
            main_layout.addWidget(help_group)

            # Hide optional SHP if needed
            self.toggle_shp_inputs()

        self.dialog.show()

    def toggle_shp_inputs(self):
        is_automated = self.mode_combo.currentText() == "Automated"

        # Enable/disable SHP inputs and checkbox
        self.wtg_path.setEnabled(is_automated)
        self.wtg_buff_path.setEnabled(is_automated)
        self.sibdry_path.setEnabled(is_automated)
        self.sibdry_buff_path.setEnabled(is_automated)
        self.priory_area.setEnabled(is_automated)
        self.potential_area.setEnabled(is_automated)
        self.keepLayersCheckBox.setEnabled(is_automated)

        # Disable/enable browse buttons associated with SHP files
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
            self.wtg_buff_size_input.show() # if the user add a SHP, the need to add the buffer distance
        else:
            self.wtg_buff_size_input.hide() # if the user does not add a SHP, the option does not show up

    def browse_sitebdry_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.sibdry_path.setText(filename)

    def browse_sitebdryBuff_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.sibdry_buff_path.setText(filename)
            self.sibdry_buff_size_input.show()  # if the user add a SHP, the need to add the buffer distance
        else:
            self.sibdry_buff_size_input.hide()  # if the user does not add a SHP, the option does not show up

    def browse_priority_area_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.priory_area.setText(filename)

    def browse_potential_area_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.potential_area.setText(filename)

    def browse_pdf(self):
        filename = QFileDialog.getExistingDirectory(None, "Select Output Folder", "")
        if filename:
            self.pdf_path.setText(filename)

    def reset_fields(self):
        self.wtg_path.clear()
        self.wtg_buff_path.clear()
        self.wtg_buff_size_input.clear()
        self.wtg_buff_size_input.hide()
        self.sibdry_path.clear()
        self.sibdry_buff_path.clear()
        self.sibdry_buff_size_input.clear()
        self.sibdry_buff_size_input.hide()
        self.priory_area.clear()
        self.potential_area.clear()
        self.project_name_input.clear()
        self.layout_size_combo.setCurrentIndex(0)
        # self.Map_title_input.clear()
        self.state_combo.setCurrentIndex(0)
        self.scale_combo.setCurrentIndex(0)
        self.pdf_path.clear()
        self.keepLayersCheckBox.setChecked(False)
        self.format_combo.setCurrentIndex(0)
        self.mode_combo.setCurrentIndex(0)
        self.toggle_shp_inputs()

# END OF PLUG-IN CONFIGURATION
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
                    "10000": {
                        "wms_url": "https://www.geobasisdaten.niedersachsen.de/wms/dtk25?",
                        "layer_name": "DTK25",
                        "title": "DTK25 NI"
                    },
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

        # --- Global Satellite (XYZ) ESRI Basemap ---
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

        # --- Google Satellite (XYZ) Basemap ---
        satellite_settings_2 = {
            "basemap": {
                "zmin": 0,
                "zmax": 19,
                "crs": "EPSG:3857",
                "url": "https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}",
                "title": "Google Satellite"
            },
            "copyright": "Imagery ©2025 Google, Maxar Technologies"
        }
        # --- OpenStreetMap (XYZ) Basemap ---
        osm_settings = {
            "basemap": {
                "zmin": 0,
                "zmax": 19,
                "crs": "EPSG:3857",
                "url": "https://tile.openstreetmap.org/{z}/{x}/{y}.png",
                "title": "OpenStreetMap"
            },
            "copyright": "© OpenStreetMap (and) contributors, CC-BY-SA"
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

            layer.setOpacity(0.80)
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
                print("MapCraft Plugin",f"Could not load WMS for {state_selected} at scale {scale_str}. TRY LATER!")
                return None, None, None

            wms_layer.setOpacity(0.6)
            QgsProject.instance().addMapLayer(wms_layer)
            return wms_layer, state_conf, scale_conf

        elif basemap_type == "OpenStreetMap":
            osm_conf = osm_settings["basemap"]
            url = osm_conf["url"]
            zmin = osm_conf["zmin"]
            zmax = osm_conf["zmax"]
            crs = osm_conf["crs"]
            title = osm_conf["title"]

            encoded_url = url.replace("=", "%3D").replace("&", "%26")
            uri = f"type=xyz&url={encoded_url}&zmin={zmin}&zmax={zmax}&crs={crs}"

            layer = QgsRasterLayer(uri, title, "wms")
            if not layer.isValid():
                print("MapCraft Plugin", "Could not load OpenStreetMap basemap.")
                return None, None, None

            layer.setOpacity(0.75)
            QgsProject.instance().addMapLayer(layer)

            return layer, osm_settings, None

        else:
            print("MapCraft Plugin", f"Unknown basemap type: {basemap_type}")
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

    def get_visible_layers_in_tree(self):
        """
                Creates the layer reference for the map.
                """
        visible_layers = []
        shp_layers_ref = []

        def collect_visible_layers(node):
            if isinstance(node, QgsLayerTreeLayer):
                if node.isVisible():
                    layer = node.layer()
                    visible_layers.append(layer)

                    # Exclude WMS layers from the Reference
                    if isinstance(layer, QgsRasterLayer) and 'wms' in layer.source().lower():
                        return
                    layer_path = layer.source()
                    # Try to extract the layer name from the source string
                    if "|layername=" in layer_path:
                        layer_name = layer_path.split("|layername=")[-1]
                    elif "layers=" in layer_path:
                        return  # Skip WMS layers
                    else:
                        layer_name = layer.name()  # fallback
                    shp_layers_ref.append(layer_name)

            elif isinstance(node, QgsLayerTreeGroup):
                for child in node.children():
                    collect_visible_layers(child)

        root = QgsProject.instance().layerTreeRoot()
        collect_visible_layers(root)
        return visible_layers, shp_layers_ref

    def adjust_font_size_to_fit(self, label_item, text, max_width, min_font_size, default_font_size):
        """
        Reduces font size of a label until the text fits within max_width or reaches min_font_size.
        """
        font = label_item.font()
        font.setPointSize(default_font_size)
        fm = QFontMetricsF(font)
        text_width = fm.width(text)

        while text_width > max_width and font.pointSize() > min_font_size:
            font.setPointSize(font.pointSize() - 1)
            fm = QFontMetricsF(font)
            text_width = fm.width(text)

        # print("==== FONT ADJUST DEBUG ====")
        # print("Label ID:", label_item.id())
        # print("Final font size:", font.pointSize())
        # print("Text:", text)
        # print("Text width (px):", text_width)
        # print("Max allowed width (px):", max_width)
        # print("===========================")

        label_item.setFont(font)
        label_item.setText(text)
        label_item.refresh()

    def run_map_generation(self):
        mode = self.mode_combo.currentText()
        if mode == "Automated":
            self.run_automated_map()
        else:
            self.run_manual_map()

    def run_automated_map(self):
        Layout = self.wtg_path.text()
        Layout_buff = self.wtg_buff_path.text()
        layout_buff_size = self.wtg_buff_size_input.text()
        Site_Bdry = self.sibdry_path.text()
        Site_Bdry_buff = self.sibdry_buff_path.text()
        Site_Bdry_buff_size = self.sibdry_buff_size_input.text()
        wind_priory_area = self.priory_area.text()
        wind_potential_area = self.potential_area.text()
        project_name = self.project_name_input.text()
        Map_title = self.Map_title_input.text()
        layout_size = self.layout_size_combo.currentText()
        basemap_type = self.basemap_combo.currentText()
        state_selected = self.state_combo.currentText()
        scale = int(self.scale_combo.currentText())
        export_format = self.format_combo.currentText()
        output_folder = self.pdf_path.text()

        today_name = datetime.today().strftime("%Y%m%d")
        pdf_filename = f"{today_name}_Windpark_{project_name}_{layout_size}"
        filename_base = os.path.join(output_folder, pdf_filename)

        if export_format == "PDF":
            output_path = os.path.join(self.pdf_path.text(), f"{filename_base}.pdf")
        else:
            output_path = os.path.join(self.pdf_path.text(), f"{filename_base}.png")

        # Check if one requested parameter is missing
        if not Layout or not project_name or not output_folder or not Map_title:
            print("Please complete all fields before running.")
            return

        # Check if WTG buffer SHP is selected but no buffer size entered
        if Layout_buff and not layout_buff_size.strip():
            print("Missing Input", "Please enter the WTG buffer size.")
            return

        # Check if Site boundary buffer SHP is selected but no buffer size entered
        if Site_Bdry_buff and not Site_Bdry_buff_size.strip():
            print("Missing Input", "Please enter the site boundary buffer size.")
            return


        layout_path = os.path.join(self.plugin_dir, f"Übersichskarte_{layout_size}.qpt")
        style_path = os.path.join(self.plugin_dir, "WEA.qml")

        map_layers = []
        shp_layers_ref = []  # Create a list to be used a REF

        # Load WTG SHP
        layer_name = os.path.basename(Layout) # This is to get the SHP name in the ref
        WTG_layer = QgsVectorLayer(Layout, layer_name, "ogr")
        if WTG_layer.isValid():
            # Remove any existing layer with the same data source
            for layer in QgsProject.instance().mapLayers().values():
                if isinstance(layer, QgsVectorLayer) and layer.source() == WTG_layer.source():
                    QgsProject.instance().removeMapLayer(layer.id())


            # Load style and add to project
            WTG_layer.loadNamedStyle(style_path)
            WTG_layer.triggerRepaint()
            QgsProject.instance().addMapLayer(WTG_layer)
            shp_layers_ref.append(layer_name)
            map_layers.append(WTG_layer)


            # Get the first value from the 'LAYOUT' field
            layout_value = None
            layout_field_index = WTG_layer.fields().indexOf('LAYOUT')
            if layout_field_index != -1:
                for feature in WTG_layer.getFeatures():
                    layout_value = feature['LAYOUT']
                    if layout_value:
                        break

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
                map_layers.append(WTG_buff_layer)

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
                map_layers.append(Site_Bdry_layer)

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
                map_layers.append(Site_Bdry_buff_layer)

        # Load wind potential area
        potential_area_layer = None
        if wind_potential_area:
            layer_name_5 = os.path.basename(wind_potential_area)
            potential_area_layer = QgsVectorLayer(wind_potential_area, layer_name_5, "ogr")

            if potential_area_layer.isValid():
                # --- Fill style with diagonal lines ---
                fill_layer = QgsSimpleFillSymbolLayer()
                fill_layer.setColor(QColor(0, 128, 0, 80))  # Light semi-transparent green fill
                fill_layer.setBrushStyle(Qt.BDiagPattern)  # Backward diagonal pattern

                # --- Bottom border stroke (thicker, light green) ---
                bottom_line = QgsSimpleLineSymbolLayer()
                bottom_line.setColor(QColor(144, 238, 144, 128))  # Light green with transparency
                bottom_line.setWidth(1.5)

                # --- Top border stroke (thinner, dark green) ---
                top_line = QgsSimpleLineSymbolLayer()
                top_line.setColor(QColor(0, 100, 0))  # Dark green
                top_line.setWidth(0.8)

                # --- Assemble final symbol ---
                symbol = QgsFillSymbol()
                symbol.changeSymbolLayer(0, fill_layer)
                symbol.appendSymbolLayer(bottom_line)
                symbol.appendSymbolLayer(top_line)

                # Apply the symbol to the layer
                potential_area_layer.setRenderer(QgsSingleSymbolRenderer(symbol))
                potential_area_layer.triggerRepaint()

                # Add the layer to the project
                QgsProject.instance().addMapLayer(potential_area_layer)
                shp_layers_ref.append(layer_name_5)
                map_layers.append(potential_area_layer)

        # Load wind priority area
        priority_area_layer = None
        if wind_priory_area:
            layer_name_4 = os.path.basename(wind_priory_area)  # Get the SHP name
            priority_area_layer = QgsVectorLayer(wind_priory_area, layer_name_4, "ogr")

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
                map_layers.append(priority_area_layer)

        # Load the WMS sever
        wms_layer, conf_dict, scale_conf = self.load_wms_layer(state_selected, scale, basemap_type)
        map_layers.append(wms_layer)

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
            map_item.setLayers(map_layers) # Make sure that only the loaded layers are visible on the PDF map.
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
            if isinstance(scale_bar_item, QgsLayoutItemScaleBar) and map_item:
                scale_bar_item.setStyle('Line Ticks Up')
                scale_bar_item.setUnits(QgsUnitTypes.DistanceKilometers)
                scale_bar_item.setNumberOfSegmentsLeft(0)
                scale_bar_item.setLinkedMap(map_item)

                if layout_size == "A4":

                    if scale == 10000:
                        real_world_km = 0.25
                        segments = 2

                    elif scale == 15000:
                        real_world_km = 0.5
                        segments = 2

                    elif scale == 25000:
                        real_world_km = 1
                        segments = 2

                    elif scale == 50000:
                        real_world_km = 2
                        segments = 2

                    units_per_segment = real_world_km / segments

                    # Apply to scale bar
                    scale_bar_item.setNumberOfSegments(segments)
                    scale_bar_item.setUnitsPerSegment(units_per_segment)

                    if scale == 10000:
                        # Adjust scale bar position
                        current_pos = scale_bar_item.pos()
                        adjusted_x = current_pos.x() + 2  # Move x mm to the left
                        adjusted_y = current_pos.y()  # Keep Y position unchanged
                        scale_bar_item.attemptMove(
                            QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                    elif scale == 15000:
                        # Adjust scale bar position
                        current_pos = scale_bar_item.pos()
                        adjusted_x = current_pos.x() - 3  # Move x mm to the left
                        adjusted_y = current_pos.y()  # Keep Y position unchanged
                        scale_bar_item.attemptMove(
                            QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                    elif scale == 25000 or scale == 50000:
                        # Adjust scale bar position
                        current_pos = scale_bar_item.pos()
                        adjusted_x = current_pos.x() - 5  # Move x mm to the left
                        adjusted_y = current_pos.y()  # Keep Y position unchanged
                        scale_bar_item.attemptMove(
                            QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                else:  # A3 or larger

                    if scale == 10000:
                        real_world_km = 0.5
                        segments = 2

                    elif scale == 15000:
                        real_world_km = 0.5
                        segments = 2

                    elif scale == 25000:
                        real_world_km = 1.0
                        segments = 2

                    elif scale == 50000:
                        real_world_km = 2.0
                        segments = 2

                    units_per_segment = real_world_km / segments

                    # Apply to scale bar
                    scale_bar_item.setNumberOfSegments(segments)
                    scale_bar_item.setUnitsPerSegment(units_per_segment)

                    if scale == 10000:
                        # Adjust scale bar position
                        current_pos = scale_bar_item.pos()
                        adjusted_x = current_pos.x() - 5  # Move x mm to the left
                        adjusted_y = current_pos.y()  # Keep Y position unchanged
                        scale_bar_item.attemptMove(
                            QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                    elif scale == 15000:
                        # Adjust scale bar position
                        current_pos = scale_bar_item.pos()
                        adjusted_x = current_pos.x() + 5  # Move x mm to the right
                        adjusted_y = current_pos.y()  # Keep Y position unchanged
                        scale_bar_item.attemptMove(
                            QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                    elif scale == 25000 or scale == 50000:
                        # No position adjustment defined
                        pass

                # Get the width of the scale bar in layout units (mm)
                scale_bar_width_mm = scale_bar_item.rect().width()


            # === LEGEND SETUP ===
            legend_item = layout.itemById("symbology")  # Make sure your layout legend ID is 'symbology'

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
                        if layout_value is not None:
                            child.setName(f"WEA - Neuplanung ({layout_value})")  # Custom name shown in legend
                        else:
                            child.setName(f"WEA - Neuplanung")  # Custom name shown in legend

                legend_item.refresh()

                # Add optional layer if available
                if WTG_buff_layer:
                    root_group.addLayer(WTG_buff_layer)
                    for child in root_group.findLayers():
                        if child.layer() == WTG_buff_layer:
                            name = "Rotorradius"
                            if layout_buff_size:
                                name = f"{name} ({layout_buff_size} m)"
                            child.setName(name)
                if Site_Bdry_layer:
                    root_group.addLayer(Site_Bdry_layer)
                    for child in root_group.findLayers():
                        if child.layer() == Site_Bdry_layer:
                            child.setName("Projektfläche")
                if Site_Bdry_buff_layer:
                    root_group.addLayer(Site_Bdry_buff_layer)
                    for child in root_group.findLayers():
                        if child.layer() == Site_Bdry_buff_layer:
                            name = "Abstandsfläche"
                            if Site_Bdry_buff_size:
                                name = f"{name} ({Site_Bdry_buff_size} m)"
                            child.setName(name)
                if priority_area_layer:
                    root_group.addLayer(priority_area_layer)
                    for child in root_group.findLayers():
                        if child.layer() == priority_area_layer:
                            child.setName("Windvorranggebiet")
                if potential_area_layer:
                    root_group.addLayer(potential_area_layer)
                    for child in root_group.findLayers():
                        if child.layer() == potential_area_layer:
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
        elif basemap_type == "OpenStreetMap" and conf_dict:
            copyright_text = conf_dict.get("copyright", "")

        dpi_ = layout.renderContext().dpi()
        for item in layout.items():
            if item.type() == QgsLayoutItemRegistry.LayoutLabel:

                if item.id() == 'label_proj':
                    item.setText(f"CRS: {projection}")

                elif item.id() == 'label_creator':
                    item.setText(f"Karte erzeugt am {today} von {username}")

                elif item.id() == 'label_title':
                    item.setText(f"{Map_title}")

                elif item.id() == 'label_Windpark':
                    full_name = f"Windpark {project_name}"
                    item.setText(full_name)
                    # if the name is too long, move the box up north
                    print(len(project_name))
                    if layout_size == "A3":
                        print("Layout A3")
                        if len(project_name) > 21:
                            current_pos = item.pos()
                            adjusted_x = current_pos.x()  # Move x mm to the right
                            adjusted_y = current_pos.y() -5 # Keep Y position unchanged
                            item.attemptMove(
                                QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))
                    elif layout_size == "A4":
                        print("Layout A4")
                        if len(project_name) > 15:
                            current_pos = item.pos()
                            adjusted_x = current_pos.x()  # Move x mm to the right
                            adjusted_y = current_pos.y() -3  # Keep Y position unchanged
                            item.attemptMove(
                                QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                elif item.id() == 'label_ref':
                    ref_label_text = f"Ref: {ref_text}"
                    max_width = item.rect().width()
                    max_width_px = max_width * dpi_ / 25.4
                    if layout_size == "A3":
                        self.adjust_font_size_to_fit(item, ref_label_text, max_width_px, min_font_size=2.5,
                                                     default_font_size=5)
                    else:
                        self.adjust_font_size_to_fit(item, ref_label_text, max_width_px, min_font_size=2,
                                                     default_font_size=4)

                elif item.id() == 'label_CR':
                    label_CR_text = f"Hintergrund: ©{copyright_text}"
                    max_width = item.rect().width()
                    max_width_px = max_width * dpi_ / 25.4
                    if layout_size == "A3":
                        self.adjust_font_size_to_fit(item, label_CR_text, max_width_px, min_font_size=2.5,
                                                     default_font_size=4)
                    else:
                        self.adjust_font_size_to_fit(item, label_CR_text, max_width_px, min_font_size=2.5,
                                                     default_font_size=3)

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

        # ✅ Remove WMS layers from canvas
        QgsProject.instance().removeMapLayer(wms_layer)

        # Conditionally remove other layers
        if not self.keepLayersCheckBox.isChecked():
            QgsProject.instance().removeMapLayer(WTG_layer)

            if WTG_buff_layer:
                QgsProject.instance().removeMapLayer(WTG_buff_layer)
            if Site_Bdry_layer:
                QgsProject.instance().removeMapLayer(Site_Bdry_layer)
            if Site_Bdry_buff_layer:
                QgsProject.instance().removeMapLayer(Site_Bdry_buff_layer)
            if priority_area_layer:
                QgsProject.instance().removeMapLayer(priority_area_layer)
            if potential_area_layer:
                QgsProject.instance().removeMapLayer(potential_area_layer)


        self.iface.mapCanvas().refresh()

    def run_manual_map(self):
        # Basic validation
        project_name = self.project_name_input.text()
        Map_title = self.Map_title_input.text()
        layout_size = self.layout_size_combo.currentText()
        state_selected = self.state_combo.currentText()
        basemap_type = self.basemap_combo.currentText()
        scale = int(self.scale_combo.currentText())
        export_format = self.format_combo.currentText()
        output_folder = self.pdf_path.text()

        today_name = datetime.today().strftime("%Y%m%d")
        pdf_filename = f"{today_name}_Windpark_{project_name}_{layout_size}"
        filename_base = os.path.join(output_folder, pdf_filename)

        if export_format == "PDF":
            output_path = os.path.join(self.pdf_path.text(), f"{filename_base}.pdf")
        else:
            output_path = os.path.join(self.pdf_path.text(), f"{filename_base}.png")

        layout_path = os.path.join(self.plugin_dir, f"Übersichskarte_{layout_size}.qpt")

        # Load template
        with open(layout_path, 'r') as f:
            template_content = f.read()
        document = QDomDocument()
        document.setContent(template_content)
        layout = QgsPrintLayout(QgsProject.instance())
        layout.initializeDefaults()
        layout.loadFromTemplate(document, QgsReadWriteContext())

        root = QgsProject.instance().layerTreeRoot()

        # Select the first active layer to set it as a center
        def find_first_visible_layer(children):
            for child in children:
                if isinstance(child, QgsLayerTreeLayer) and child.isVisible():
                    layer = child.layer()
                    print(f"First visible layer: {layer.name()}")
                    return layer
                elif isinstance(child, QgsLayerTreeGroup):
                    result = find_first_visible_layer(child.children())
                    if result:
                        return result
            return None

        layer = find_first_visible_layer(root.children())


        wms_layer, conf_dict, scale_conf = self.load_wms_layer(state_selected, scale, basemap_type)

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

        # Get only visible layers (including the WMS layer if visible)
        visible_layers, shp_layers_ref = self.get_visible_layers_in_tree()


        print("Number of layers in REF: ", len(shp_layers_ref))

        # Layer ordering: WMS/raster layers at the bottom
        vector_layers = [l for l in visible_layers if isinstance(l, QgsVectorLayer)]
        final_order = vector_layers + [wms_layer] # This part does the trick to send wms to the bottom


        map_item.setLayers(final_order)
        map_item.refresh()

        # === SCALE BAR SETUP ===
        scale_bar_item = layout.itemById('scale')
        if isinstance(scale_bar_item, QgsLayoutItemScaleBar) and map_item:
            scale_bar_item.setStyle('Line Ticks Up')
            scale_bar_item.setUnits(QgsUnitTypes.DistanceKilometers)
            scale_bar_item.setNumberOfSegmentsLeft(0)
            scale_bar_item.setLinkedMap(map_item)

            if layout_size == "A4":

                if scale == 10000:
                    real_world_km = 0.25
                    segments = 2

                elif scale == 15000:
                    real_world_km = 0.5
                    segments = 2

                elif scale == 25000:
                    real_world_km = 1
                    segments = 2

                elif scale == 50000:
                    real_world_km = 2
                    segments = 2

                units_per_segment = real_world_km / segments

                # Apply to scale bar
                scale_bar_item.setNumberOfSegments(segments)
                scale_bar_item.setUnitsPerSegment(units_per_segment)

                if scale == 10000:
                    # Adjust scale bar position
                    current_pos = scale_bar_item.pos()
                    adjusted_x = current_pos.x() + 2  # Move x mm to the left
                    adjusted_y = current_pos.y()  # Keep Y position unchanged
                    scale_bar_item.attemptMove(
                        QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                elif scale == 15000:
                    # Adjust scale bar position
                    current_pos = scale_bar_item.pos()
                    adjusted_x = current_pos.x() - 3  # Move x mm to the left
                    adjusted_y = current_pos.y()  # Keep Y position unchanged
                    scale_bar_item.attemptMove(
                        QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                elif scale == 25000 or scale == 50000:
                    # Adjust scale bar position
                    current_pos = scale_bar_item.pos()
                    adjusted_x = current_pos.x() - 5  # Move x mm to the left
                    adjusted_y = current_pos.y()  # Keep Y position unchanged
                    scale_bar_item.attemptMove(
                        QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

            else:  # A3 or larger

                if scale == 10000:
                    real_world_km = 0.5
                    segments = 2

                elif scale == 15000:
                    real_world_km = 0.5
                    segments = 2

                elif scale == 25000:
                    real_world_km = 1.0
                    segments = 2

                elif scale == 50000:
                    real_world_km = 2.0
                    segments = 2

                units_per_segment = real_world_km / segments

                # Apply to scale bar
                scale_bar_item.setNumberOfSegments(segments)
                scale_bar_item.setUnitsPerSegment(units_per_segment)

                if scale == 10000:
                    # Adjust scale bar position
                    current_pos = scale_bar_item.pos()
                    adjusted_x = current_pos.x() - 5  # Move x mm to the left
                    adjusted_y = current_pos.y()  # Keep Y position unchanged
                    scale_bar_item.attemptMove(
                        QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                elif scale == 15000:
                    # Adjust scale bar position
                    current_pos = scale_bar_item.pos()
                    adjusted_x = current_pos.x() + 5  # Move x mm to the right
                    adjusted_y = current_pos.y()  # Keep Y position unchanged
                    scale_bar_item.attemptMove(
                        QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))

                elif scale == 25000 or scale == 50000:
                    # No position adjustment defined
                    pass

            # Get the width of the scale bar in layout units (mm)
            #scale_bar_width_mm = scale_bar_item.rect().width()

        # Legend setup

        legend_item = layout.itemById("symbology")
        if legend_item and map_item:
            legend_item.setLinkedMap(map_item)
            legend_item.setAutoUpdateModel(False)

            legend_model = legend_item.model()
            root_group = legend_model.rootGroup()
            root_group.removeAllChildren()

            # Access the current layer tree
            layer_tree = QgsProject.instance().layerTreeRoot()

            # Add only visible vector layers (skip WMS/raster/invisible)
            # Recursively add all visible vector layers (even inside groups)
            def add_visible_vector_layers(node):
                if isinstance(node, QgsLayerTreeLayer):
                    if node.isVisible() and isinstance(node.layer(), QgsVectorLayer):
                        root_group.addLayer(node.layer())
                elif isinstance(node, QgsLayerTreeGroup):
                    for child in node.children():
                        add_visible_vector_layers(child)


            add_visible_vector_layers(layer_tree)

            legend_item.refresh()

            ## Adjust font size based on number of legend entries
            #legend_item.attemptResize(QgsLayoutSize(55, 100))
            #self.adjust_legend_font_size(legend_item)


        # Dynamic labels
        username = getpass.getuser()
        today = datetime.today().strftime("%d/%m/%y")
        projection = layer.crs().description()
        ref_text = " | ".join(shp_layers_ref)
        if basemap_type == "Topographic" and conf_dict:
            copyright_text = conf_dict.get("copyright", "")
        elif basemap_type == "Satellite" and conf_dict:
            copyright_text = conf_dict.get("copyright", "")
        elif basemap_type == "OpenStreetMap" and conf_dict:
            copyright_text = conf_dict.get("copyright", "")

        dpi_ = layout.renderContext().dpi()
        for item in layout.items():
            if item.type() == QgsLayoutItemRegistry.LayoutLabel:

                if item.id() == 'label_proj':
                    item.setText(f"CRS: {projection}")

                elif item.id() == 'label_creator':
                    item.setText(f"Karte erzeugt am {today} von {username}")

                elif item.id() == 'label_title':
                    item.setText(f"{Map_title}")

                elif item.id() == 'label_Windpark':
                    full_name = f"Windpark {project_name}"
                    item.setText(full_name)
                    # if the name is too long, move the box up north
                    print(len(project_name))
                    if layout_size == "A3":
                        print("Layout A3")
                        if len(project_name) > 21:
                            current_pos = item.pos()
                            adjusted_x = current_pos.x()  # Move x mm to the right
                            adjusted_y = current_pos.y() -5 # Keep Y position unchanged
                            item.attemptMove(
                                QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))
                    elif layout_size == "A4":
                        print("Layout A4")
                        if len(project_name) > 15:
                            current_pos = item.pos()
                            adjusted_x = current_pos.x()  # Move x mm to the right
                            adjusted_y = current_pos.y() -3  # Keep Y position unchanged
                            item.attemptMove(
                                QgsLayoutPoint(adjusted_x, adjusted_y, QgsUnitTypes.LayoutMillimeters))


                elif item.id() == 'label_ref':
                    ref_label_text = f"Ref: {ref_text}"
                    max_width = item.rect().width()
                    max_width_px = max_width * dpi_ / 25.4
                    if layout_size == "A3":
                        self.adjust_font_size_to_fit(item, ref_label_text, max_width_px, min_font_size=2.5,
                                                     default_font_size=5)
                    else:
                        self.adjust_font_size_to_fit(item, ref_label_text, max_width_px, min_font_size=2,
                                                     default_font_size=4)

                elif item.id() == 'label_CR':
                    label_CR_text = f"Hintergrund: ©{copyright_text}"
                    max_width = item.rect().width()
                    max_width_px = max_width * dpi_ / 25.4
                    if layout_size == "A3":
                        self.adjust_font_size_to_fit(item, label_CR_text, max_width_px, min_font_size=2.5,
                                                     default_font_size=4)
                    else:
                        self.adjust_font_size_to_fit(item, label_CR_text, max_width_px, min_font_size=2.5,
                                                     default_font_size=3)

        # Export based on selected format
        exporter = QgsLayoutExporter(layout)

        if export_format == "PDF":
            pdf_settings = QgsLayoutExporter.PdfExportSettings()
            result = exporter.exportToPdf(output_path, pdf_settings)
            if result == QgsLayoutExporter.Success:
                print('Success', 'PDF exported successfully!')
            else:
                print('Error', 'PDF export failed.')

        elif export_format == "PNG":
            image_settings = QgsLayoutExporter.ImageExportSettings()
            image_settings.dpi = 300  # ✅ Set high resolution
            result = exporter.exportToImage(output_path, image_settings)
            if result == QgsLayoutExporter.Success:
                print('Success', 'PNG exported successfully!')
            else:
                print('Error', 'PNG export failed.')

        #  Remove WMS layers from canvas
        QgsProject.instance().removeMapLayer(wms_layer)
        iface.mapCanvas().refresh()




