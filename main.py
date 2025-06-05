import os
import getpass
from datetime import datetime
from PyQt5.QtWidgets import (QAction, QFileDialog, QWidget, QVBoxLayout, QLabel, QLineEdit,
                             QPushButton, QComboBox, QHBoxLayout)
from PyQt5.QtGui import QIcon, QColor
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsRasterLayer,
    QgsPrintLayout, QgsLayoutItemMap, QgsReadWriteContext, QgsRectangle,
    QgsLayoutExporter, QgsLayoutItemRegistry, QgsLineSymbol, QgsSingleSymbolRenderer
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
            layout = QVBoxLayout()

            # SHP File
            shp_layout = QHBoxLayout()
            self.shp_path = QLineEdit()
            browse_shp = QPushButton("Browse SHP")
            browse_shp.clicked.connect(self.browse_shp)
            shp_layout.addWidget(QLabel("WTG Layout:"))
            shp_layout.addWidget(self.shp_path)
            shp_layout.addWidget(browse_shp)
            layout.addLayout(shp_layout)

            # Optional SHP File (Projekt Fläche)
            opt_shp_layout = QHBoxLayout()
            self.opt_shp_path = QLineEdit()
            browse_opt_shp = QPushButton("Browse SHP")
            browse_opt_shp.clicked.connect(self.browse_opt_shp)
            opt_shp_layout.addWidget(QLabel("Site boundary:"))
            opt_shp_layout.addWidget(self.opt_shp_path)
            opt_shp_layout.addWidget(browse_opt_shp)
            layout.addLayout(opt_shp_layout)

            # Project Name
            self.project_name_input = QLineEdit()
            layout.addWidget(QLabel("Project Name:"))
            layout.addWidget(self.project_name_input)

            # State Selector
            self.state_combo = QComboBox()
            self.state_combo.addItems(["Rheinland-Pfalz", "Baden-Württemberg", "Niedersachsen", "Schleswig-Holstein"])
            layout.addWidget(QLabel("Select German State:"))
            layout.addWidget(self.state_combo)

            # Scale Selector
            self.scale_combo = QComboBox()
            self.scale_combo.addItems(["25000", "50000"])
            layout.addWidget(QLabel("Map Scale:"))
            layout.addWidget(self.scale_combo)

            # Output PDF Path
            pdf_layout = QHBoxLayout()
            self.pdf_path = QLineEdit()
            browse_pdf = QPushButton("Browse Folder")
            browse_pdf.clicked.connect(self.browse_pdf)
            pdf_layout.addWidget(QLabel("PDF Output Folder:"))
            pdf_layout.addWidget(self.pdf_path)
            pdf_layout.addWidget(browse_pdf)
            layout.addLayout(pdf_layout)

            # Export Format Selector
            self.format_combo = QComboBox()
            self.format_combo.addItems(["PDF", "PNG"])
            layout.addWidget(QLabel("Export Format:"))
            layout.addWidget(self.format_combo)

            # Run Button
            run_button = QPushButton("Run")
            run_button.clicked.connect(self.run_map_generation)
            layout.addWidget(run_button)

            self.dialog.setLayout(layout)

        self.dialog.show()

    def browse_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.shp_path.setText(filename)

    def browse_opt_shp(self):
        filename, _ = QFileDialog.getOpenFileName(None, "Select SHP File", "", "Shapefiles (*.shp)")
        if filename:
            self.opt_shp_path.setText(filename)

    def browse_pdf(self):
        filename = QFileDialog.getExistingDirectory(None, "Select Output Folder", "")
        if filename:
            self.pdf_path.setText(filename)

    def run_map_generation(self):
        Layout = self.shp_path.text()
        Site_Bdry = self.opt_shp_path.text()
        project_name = self.project_name_input.text()
        state_selected = self.state_combo.currentText()
        scale = int(self.scale_combo.currentText())
        export_format = self.format_combo.currentText()
        output_folder = self.pdf_path.text()

        today_name = datetime.today().strftime("%d%m%y")
        pdf_filename = f"Übersichtskarte_Windpark{project_name}_{today_name}"
        filename_base = os.path.join(output_folder, pdf_filename)

        if export_format == "PDF":
            output_path = os.path.join(self.pdf_path.text(), f"{filename_base}.pdf")
        else:
            output_path = os.path.join(self.pdf_path.text(), f"{filename_base}.png")


        if not all([Layout, project_name, output_path]):
            self.iface.messageBar().pushWarning("Warning", "Please complete all fields before running.")
            return

        layout_path = os.path.join(self.plugin_dir, "Übersichskarte_template_v4.qpt")
        style_path = os.path.join(self.plugin_dir, "WEA.qml")

        state_settings = {
            "Baden-Württemberg": {
                "scales": {
                    "25000": {
                        "wms_url": "https://owsproxy.lgl-bw.de/owsproxy/ows/WMS_LGL-BW_ATKIS_DTK_25_K_A?",
                        "layer_name": "RDS_LY_DTK25K_COL",
                        "title": "DTK25 Farbkombination"
                    },
                    "50000": {
                        "wms_url": "https://owsproxy.lgl-bw.de/owsproxy/ows/WMS_LGL-BW_ATKIS_DTK_50_K_A?",
                        "layer_name": "RDS_LY_DTK50K_COL",
                        "title": "DTK50 Farbkombination"
                    }
                },
                "copyright": "LGL-BW (2025)"
            },
            "Rheinland-Pfalz": {
                "scales": {
                    "25000": {
                        "wms_url": "https://geo4.service24.rlp.de/wms/rp_dtk25.fcgi?",
                        "layer_name": "rp_dtk25",
                        "title": "WMS RP DTK25"
                    },
                    "50000": {
                        "wms_url": "https://geo4.service24.rlp.de/wms/rp_dtk50.fcgi?",
                        "layer_name": "rp_dtk50",
                        "title": "WMS RP DTK50"
                    }
                },
                "copyright": "GeoBasis-DE/LVermGeoRP (2005) dl-de/by-2-0"
            },
            "Niedersachsen": {
                "scales": {
                    "25000": {
                        "wms_url": "https://www.geobasisdaten.niedersachsen.de/wms/dtk25?",
                        "layer_name": "DTK25",
                        "title": "WMS NI DTK25"
                    },
                    "50000": {
                        "wms_url": "https://www.geobasisdaten.niedersachsen.de/wms/dtk50?",
                        "layer_name": "DTK50",
                        "title": "WMS NI DTK50"
                    }
                },
                "copyright": "LGLN 2025"
            },
            "Schleswig-Holstein": {
                "scales": {
                    "25000": {
                        "wms_url": "https://www.geodaten-sh.de/ows/dtk25?",
                        "layer_name": "sh_dtk25",
                        "title": "WMS SH DTK25"
                    },
                    "50000": {
                        "wms_url": "https://www.geodaten-sh.de/ows/dtk50?",
                        "layer_name": "sh_dtk50",
                        "title": "WMS SH DTK50"
                    }
                },
                "copyright": "GeoBasis-DE/LVermGeo 2025 SH/CC BY 4.0"
            }
        }

        # Load WMS
        state_conf = state_settings[state_selected]
        scale_conf = state_conf["scales"][str(scale)]

        wms_url = (
            f"contextualWMSLegend=0&crs=EPSG:25832&dpiMode=7&featureCount=10"
            f"&format=image/png&layers={scale_conf['layer_name']}&styles=&url={scale_conf['wms_url']}"
        )

        wms_layer = QgsRasterLayer(wms_url, f"{state_selected} Basemap", "wms")
        if wms_layer.isValid():
            wms_layer.setOpacity(0.5)
            QgsProject.instance().addMapLayer(wms_layer)
        else:
            raise Exception(f"Failed to load WMS layer for {state_selected} at scale {scale}")

        shp_layers_ref = [] # Create a list to be used a REF

        # Load SHP
        layer_name = os.path.basename(Layout) # This is to get the SHP name in the ref
        layer = QgsVectorLayer(Layout, layer_name, "ogr")
        if layer.isValid():
            layer.loadNamedStyle(style_path)
            layer.triggerRepaint()
            QgsProject.instance().addMapLayer(layer)
            shp_layers_ref.append(layer_name)

        # Load Optional SHP 1 (Projekt Fläche)
        opt_layer = None
        if Site_Bdry:
            layer_name_2 = os.path.basename(Site_Bdry)  # This is to get the SHP name in the ref
            opt_layer = QgsVectorLayer(Site_Bdry, "Projekt Fläche", "ogr")
            if opt_layer.isValid():
                # Create a transparent fill with red outline
                from qgis.core import QgsFillSymbol, QgsSimpleFillSymbolLayer
                symbol = QgsFillSymbol.createSimple({
                    'outline_color': 'red',
                    'outline_width': '0.3',
                    'color': 'transparent'
                })
                opt_layer.setRenderer(QgsSingleSymbolRenderer(symbol))
                opt_layer.triggerRepaint()
                QgsProject.instance().addMapLayer(opt_layer)
                shp_layers_ref.append(layer_name_2)

        # Load Layout
        with open(layout_path, 'r') as f:
            template_content = f.read()
        document = QDomDocument()
        document.setContent(template_content)

        layout = QgsPrintLayout(QgsProject.instance())
        layout.initializeDefaults()
        layout.loadFromTemplate(document, QgsReadWriteContext())

        # Map Item
        map_item = next((item for item in layout.items() if isinstance(item, QgsLayoutItemMap) and item.id() == "Map"), None)

        if map_item:
            map_item.setScale(scale)
            map_width_m = (map_item.rect().width() * scale) / 1000
            map_height_m = (map_item.rect().height() * scale) / 1000
            center = layer.extent().center()
            extent = QgsRectangle(center.x() - map_width_m / 2, center.y() - map_height_m / 2,
                                  center.x() + map_width_m / 2, center.y() + map_height_m / 2)
            map_item.setExtent(extent)
            map_item.refresh()

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
                root_group.addLayer(layer)

                # Rename the legend label for the SHP layer
                for child in root_group.findLayers():
                    if child.layer() == layer:
                        child.setName("WEA - Neuplanung")  # Custom name shown in legend

                legend_item.refresh()

                # Add optional layer if available
                if opt_layer:
                    root_group.addLayer(opt_layer)
                    for child in root_group.findLayers():
                        if child.layer() == opt_layer:
                            child.setName("Projektfläche")

        # Dynamic Labels

        projection = layer.crs().description()
        today = datetime.today().strftime("%d/%m/%y")
        username = getpass.getuser()
        ref_text = " | ".join(shp_layers_ref)
        copyright_text = state_conf["copyright"]

        for item in layout.items():
            if item.type() == QgsLayoutItemRegistry.LayoutLabel:
                if item.id() == 'label_proj':
                    item.setText(f"CRS: {projection}")
                elif item.id() == 'label_creator':
                    item.setText(f"Map produced on {today} by {username}")
                elif item.id() == 'label_title':
                    item.setText("\u00dcbersichtskarte")
                elif item.id() == 'label_Windpark':
                    item.setText(f"Windpark {project_name}")
                elif item.id() == 'label_ref':
                    item.setText(f"Ref: {ref_text}")
                elif item.id() == 'label_CR':
                    item.setText(f"Hintergrund: ©{copyright_text}")


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
        QgsProject.instance().removeMapLayer(layer)
        QgsProject.instance().removeMapLayer(wms_layer)
        if opt_layer:
            QgsProject.instance().removeMapLayer(opt_layer)

        self.iface.mapCanvas().refresh()
