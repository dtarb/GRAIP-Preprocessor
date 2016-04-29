import os
import time

import pyodbc
from osgeo import ogr, osr
from gdalconst import *

from datetime import datetime
from PySide.QtGui import *
from PySide.QtCore import *

MS_ACCESS_CONNECTION = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;"
DP_ERROR_LOG_TABLE_NAME = 'DPErrorLog'
RD_ERROR_LOG_TABLE_NAME = 'RDErrorLog'
GRAIP_ICON_FILE = "GRAIPIcon.ico"


class GDALFileDriver(object):
    @classmethod
    def ShapeFile(cls):
        return "ESRI Shapefile"

    @classmethod
    def TifFile(cls):
        return "GTiff"


class GraipMessageBox(QMessageBox):
    def __init__(self, *args):
        QMessageBox.__init__(self, *args)
        self.setWindowIcon(QIcon(GRAIP_ICON_FILE))


class FileDialog(QFileDialog):
    """
    Custom file dialog for selecting multiple files
    """
    def __init__(self, *args):
        QFileDialog.__init__(self, *args)
        self.setOption(self.DontUseNativeDialog, True)
        self.setFileMode(self.ExistingFiles)
        btns = self.findChildren(QPushButton)
        self.openBtn = [x for x in btns if 'open' in str(x.text()).lower()][0]
        self.openBtn.clicked.disconnect()
        self.openBtn.clicked.connect(self.openClicked)
        self.tree = self.findChild(QTreeView)
        self.selectedFiles = []

    def openClicked(self):
        indexes = self.tree.selectionModel().selectedIndexes()
        files = []
        for i in indexes:
            if i.column() == 0:
                #print("i.data():" + str(i.data()))
                files.append(os.path.join(str(self.directory().absolutePath()),str(i.data())))
        self.selectedFiles = files
        self.hide()

    def filesSelected(self):
        return self.selectedFiles


class ConsolidateShapeFiles(QDialog):
    def __init__(self, graip_db_file, dp_shp_files, rd_shp_files, dp_log_file, rd_log_file, working_directory, parent=None):
        super(ConsolidateShapeFiles, self).__init__(parent)
        self.graip_db_file = graip_db_file
        self.dp_shp_files = dp_shp_files
        self.rd_shp_files = rd_shp_files
        self.dp_log_file = dp_log_file
        self.rd_log_file = rd_log_file
        self.working_directory = working_directory

        v_layout = QVBoxLayout()
        msg = "Checking for orphan drain points, road segments, and duplicate ids ..."
        self.message = QLabel(msg)
        self.message.wordWrap()
        v_layout.addWidget(self.message)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        v_layout.addWidget(self.progress_bar)
        self.btn_close = QDialogButtonBox(QDialogButtonBox.Close,
                                        Qt.Horizontal, self)
        self.btn_close.setEnabled(False)
        self.btn_close.rejected.connect(self.reject)
        v_layout.addWidget(self.btn_close)
        self.setLayout(v_layout)
        self.setWindowTitle("Consolidating")
        self.setWindowIcon(QIcon(GRAIP_ICON_FILE))
        self.setFixedSize(600, 100)
        self.setModal(True)
        self.setFocus()

    def do_process(self):
        self.check_for_orphan_drain_points()
        self.consolidate_dp_shp_files()
        self.consolidate_rd_shp_files()
        self.message.setText("Preprocessing successful")
        self.btn_close.setEnabled(True)

    def check_for_orphan_drain_points(self):
        conn = pyodbc.connect(MS_ACCESS_CONNECTION % self.graip_db_file)
        cursor = conn.cursor()
        dp_rows = cursor.execute("SELECT * FROM DrainPoints").fetchall()
        dp_record_count = len(dp_rows)
        self.progress_bar.setMaximum(dp_record_count)
        for i, dp_row in enumerate(dp_rows):
            graipdid = dp_row.GRAIPDID
            dp_def_row = cursor.execute("SELECT * FROM DrainTypeDefinitions WHERE DrainTypeID=?", dp_row.DrainTypeID).fetchone()
            sql_select = "SELECT GRAIPDID1, GRAIPDID2 FROM RoadLines WHERE GRAIPDID1=? OR GRAIPDID2=?"
            rd_rows = cursor.execute(sql_select, (graipdid, graipdid)).fetchall()
            if rd_rows is None:
                add_entry_to_log_file(self.dp_log_file, graipdid, dp_def_row.DrainTypeName, "Orphan Drain Point",
                                      "Nothing")
                add_entry_to_error_table(self.graip_db_file, DP_ERROR_LOG_TABLE_NAME, graipdid,
                                         dp_def_row.DrainTypeName, "Orphan Drain Point", "Nothing")

            # check if there are duplicate drainids in DrainPoints table
            dp_rows_for_drainid = cursor.execute("SELECT * FROM DrainPoints WHERE DrainID=?", dp_row.DrainID).fetchall()
            for dp_row_di in dp_rows_for_drainid:
                msg = "Duplicate DrainID:{}".format(dp_row.DrainID)
                add_entry_to_log_file(self.dp_log_file, dp_row_di.GRAIPDID, dp_def_row.DrainTypeName, msg,
                                      "Nothing")
                add_entry_to_error_table(self.graip_db_file, DP_ERROR_LOG_TABLE_NAME, dp_row_di.GRAIPDID,
                                         dp_def_row.DrainTypeName, msg, "Nothing")

            # update progressbar
            progress_value = i + 1
            self.progress_bar.setValue(progress_value)
            QApplication.processEvents()

        rd_rows = cursor.execute("SELECT * FROM RoadLines").fetchall()
        rd_record_count = len(rd_rows)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(rd_record_count)
        for i, rd_row in enumerate(rd_rows):
            graiprid = rd_row.GRAIPRID
            sql_select = "SELECT * FROM DrainPoints WHERE GRAIPDID=? OR GRAIPDID=?"
            dp_rows = cursor.execute(sql_select, (rd_row.GRAIPDID1, rd_row.GRAIPDID2)).fetchall()
            if dp_rows is None:
                add_entry_to_log_file(self.dp_log_file, graiprid, "Road Line", "Orphan Road Segment",
                                      "Nothing")
                add_entry_to_error_table(self.graip_db_file, RD_ERROR_LOG_TABLE_NAME, graiprid,
                                         "Road Line", "Orphan Road Segment", "Nothing")
            # update progressbar
            progress_value = i + 1

            self.progress_bar.setValue(progress_value)

        conn.close()

    def consolidate_dp_shp_files(self):
        self.message.setText("Consolidating multiple drain points shapefiles...")
        conn = pyodbc.connect(MS_ACCESS_CONNECTION % self.graip_db_file)
        cursor = conn.cursor()
        dp_rows = cursor.execute("SELECT GRAIPDID FROM DrainPoints ORDER BY GRAIPDID").fetchall()
        dp_record_count = len(dp_rows)
        conn.close()
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(dp_record_count)
        # set up the shapefile driver
        driver = ogr.GetDriverByName(GDALFileDriver.ShapeFile())

        # create the data source
        dp_consolidated_shp_file = os.path.join(self.working_directory, "DrainPoints.shp")
        dp_data_source = driver.CreateDataSource(dp_consolidated_shp_file)
        # create the layer of the shapefile
        srs = osr.SpatialReference()
        dp_layer = dp_data_source.CreateLayer("DrainPoints", srs, ogr.wkbPoint)
        # Add the field GRAIPDID
        field_name = ogr.FieldDefn("GRAIPDID", ogr.OFTInteger)
        dp_layer.CreateField(field_name)

        # for dp_row in dp_rows:
        #     feature = ogr.Feature(dp_layer.GetLayerDefn())
        #     # Set the attributes using the values from the delimited text file
        #     feature.SetField("GRAIPDID", dp_row.GRAIPDID)
        #     # Create the feature in the layer (shapefile)
        #     dp_layer.CreateFeature(feature)
        #     # Destroy the feature to free resources
        #     feature.Destroy()
        #
        #     self.progress_bar.setValue(dp_row.GRAIPDID + 1)
        #     QApplication.processEvents()
        #     time.sleep(0.01)

        # NOTE: This how the old graip used to do
        graipdid = 0
        gdal_driver = ogr.GetDriverByName(GDALFileDriver.ShapeFile())
        for shp_file in self.dp_shp_files:
            data_source = gdal_driver.Open(shp_file, GA_ReadOnly)
            layer = data_source.GetLayer(0)
            for feature in layer:
                dp_feature = ogr.Feature(dp_layer.GetLayerDefn())
                geom = feature.GetGeometryRef()
                # Set the feature geometry
                dp_feature.SetGeometry(geom)
                # Set the attributes using the values from the delimited text file
                dp_feature.SetField("GRAIPDID", graipdid)
                # Create the feature in the layer (shapefile)
                dp_layer.CreateFeature(dp_feature)
                # Destroy the feature to free resources
                dp_feature.Destroy()
                graipdid += 1
                self.progress_bar.setValue(graipdid)
                QApplication.processEvents()
                time.sleep(0.01)
            data_source.Destroy()

        dp_data_source.Destroy()

    def consolidate_rd_shp_files(self):
        self.message.setText("Consolidating multiple road lines shapefiles...")
        conn = pyodbc.connect(MS_ACCESS_CONNECTION % self.graip_db_file)
        cursor = conn.cursor()
        rd_rows = cursor.execute("SELECT GRAIPRID FROM RoadLines ORDER BY GRAIPRID").fetchall()
        rd_record_count = len(rd_rows)
        conn.close()
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(rd_record_count)
        # set up the shapefile driver
        driver = ogr.GetDriverByName(GDALFileDriver.ShapeFile())

        # create the data source
        rd_consolidated_shp_file = os.path.join(self.working_directory, "RoadLines.shp")
        rd_data_source = driver.CreateDataSource(rd_consolidated_shp_file)
        # create the layer
        srs = osr.SpatialReference()
        rd_layer = rd_data_source.CreateLayer("RoadLines", srs, ogr.wkbLineString)
        # Add the field GRAIPRID
        field_name = ogr.FieldDefn("GRAIPRID", ogr.OFTInteger)
        rd_layer.CreateField(field_name)

        # for rd_row in rd_rows:
        #     feature = ogr.Feature(layer.GetLayerDefn())
        #     # Set the attributes using the values from the delimited text file
        #     feature.SetField("GRAIPRID", rd_row.GRAIPRID)
        #     # Create the feature in the layer (shapefile)
        #     layer.CreateFeature(feature)
        #     # Destroy the feature to free resources
        #     feature.Destroy()
        #     self.progress_bar.setValue(rd_row.GRAIPRID + 1)
        #     QApplication.processEvents()
        #     time.sleep(0.01)
        # data_source.Destroy()

        graiprid = 0
        gdal_driver = ogr.GetDriverByName(GDALFileDriver.ShapeFile())
        for shp_file in self.rd_shp_files:
            data_source = gdal_driver.Open(shp_file, GA_ReadOnly)
            layer = data_source.GetLayer(0)
            for feature in layer:
                rd_feature = ogr.Feature(rd_layer.GetLayerDefn())
                geom = feature.GetGeometryRef()
                # Set the feature geometry
                rd_feature.SetGeometry(geom)
                # Set the attributes using the values from the delimited text file
                rd_feature.SetField("GRAIPRID", graiprid)
                # Create the feature in the layer (shapefile)
                rd_layer.CreateFeature(rd_feature)
                # Destroy the feature to free resources
                rd_feature.Destroy()
                graiprid += 1
                self.progress_bar.setValue(graiprid)
                QApplication.processEvents()
                time.sleep(0.01)
            data_source.Destroy()

        rd_data_source.Destroy()


class AddRoadNetworkDefinitionsDialog(QDialog):
    rd_network_name = None
    rd_base_rate = None
    rd_description = None

    def __init__(self, graip_db_file, parent=None):
        super(AddRoadNetworkDefinitionsDialog, self).__init__(parent)
        self.graip_db_file = graip_db_file
        grid_layout = QGridLayout()
        self.label_rd_network_name = QLabel("Road Network Name")
        self.line_edit_rd_network_name = QLineEdit()
        grid_layout.addWidget(self.label_rd_network_name, 0, 0)
        grid_layout.addWidget(self.line_edit_rd_network_name, 0, 1)

        self.label_rd_base_rate = QLabel("Base Rate(kg/m/yr)")
        self.line_edit_rd_base_rate = QLineEdit()
        grid_layout.addWidget(self.label_rd_base_rate, 1, 0)
        grid_layout.addWidget(self.line_edit_rd_base_rate, 1, 1)

        self.label_rd_description = QLabel("Description")
        self.line_edit_rd_description = QLineEdit()
        grid_layout.addWidget(self.label_rd_description, 2, 0)
        grid_layout.addWidget(self.line_edit_rd_description, 2, 1)

        # OK and Cancel buttons
        btn_layout = QHBoxLayout()
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self)

        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        btn_layout.addWidget(self.buttons)
        grid_layout.addLayout(btn_layout, 3, 1)

        self.setWindowTitle("Add Road Network Definition")
        self.resize(400, 200)
        self.setLayout(grid_layout)
        self.setWindowIcon(QIcon(GRAIP_ICON_FILE))
        self.setModal(True)

    def accept(self, *args, **kwargs):
        msg_box = GraipMessageBox()
        msg_box.setWindowTitle("Invalid Data")

        if len(self.line_edit_rd_network_name.text().strip()) == 0:
            msg_box.setText("Network name is missing.")
            msg_box.exec_()
            return
        else:
            # check that the network name not already exists
            conn = pyodbc.connect(MS_ACCESS_CONNECTION % self.graip_db_file)
            cursor = conn.cursor()
            network_name = self.line_edit_rd_network_name.text().strip()
            sql_select = "SELECT * FROM RoadNetworkDefinitions WHERE RoadNetwork=?"
            network_def_row = cursor.execute(sql_select, network_name).fetchone()
            conn.close()
            if network_def_row is not None:
                msg_box.setText("Network name already exists.")
                msg_box.exec_()
                return
            self.rd_network_name = network_name
        if len(self.line_edit_rd_base_rate.text().strip()) == 0:
            msg_box.setText("Base rate is missing.")
            msg_box.exec_()
            return
        else:
            try:
                int(self.line_edit_rd_base_rate.text())
                self.rd_base_rate = self.line_edit_rd_base_rate.text()
            except ValueError:
                msg_box.setText("Base rate should be a numerical value.")
                msg_box.exec_()
                return

        if len(self.line_edit_rd_description.text().strip()) == 0:
            msg_box.setText("Network description is missing.")
            msg_box.exec_()
            return
        else:
            self.rd_description = self.line_edit_rd_description.text().strip()

        super(AddRoadNetworkDefinitionsDialog, self).accept(*args, **kwargs)


class OptionsDialog(QDialog):
    def __init__(self, parent=None):
        super(OptionsDialog, self).__init__(parent)
        self.form_layout = QFormLayout()
        self.grp_box_processing = QGroupBox("Processing")
        self.radio_btn_uninterrupted = QRadioButton("Uninterrupted")
        self.radio_btn_step_by_step = QRadioButton("Step by step")
        # set this radio button as checked
        self.radio_btn_step_by_step.toggle()
        v_layout_processing = QVBoxLayout()
        v_layout_processing.addWidget(self.radio_btn_uninterrupted)
        v_layout_processing.addWidget(self.radio_btn_step_by_step)
        self.grp_box_processing.setLayout(v_layout_processing)
        self.form_layout.addRow(self.grp_box_processing)
        # layout = QGridLayout()
        # layout.setColumnStretch(1, 1)
        self.grp_box_log_files = QGroupBox("Log Files")
        self.dp_label = QLabel("Drain Point")
        self.line_edit_dp_log_file = QLineEdit()
        self.btn_browse_dp_log_file = QPushButton('....')
        # connect the browse button to the function
        self.btn_browse_dp_log_file.clicked.connect(self.browse_dp_log_file)

        v_layout_log_files = QVBoxLayout()
        v_layout_dp_log = QVBoxLayout()
        h_layout_dp_log = QHBoxLayout()
        h_layout_dp_log.addWidget(self.line_edit_dp_log_file)
        h_layout_dp_log.addWidget(self.btn_browse_dp_log_file)
        v_layout_dp_log.addWidget(self.dp_label)
        v_layout_dp_log.addLayout(h_layout_dp_log)
        v_layout_log_files.addLayout(v_layout_dp_log)
        self.grp_box_log_files.setLayout(v_layout_log_files)

        self.rd_label = QLabel("Road")
        self.line_edit_rd_log_file = QLineEdit()
        self.btn_browse_rd_log_file = QPushButton('....')
        # connect the browse button to the function
        self.btn_browse_rd_log_file.clicked.connect(self.browse_rd_log_file)

        v_layout_rd_log = QVBoxLayout()
        h_layout_rd_log = QHBoxLayout()
        h_layout_rd_log.addWidget(self.line_edit_rd_log_file)
        h_layout_rd_log.addWidget(self.btn_browse_rd_log_file)
        v_layout_rd_log.addWidget(self.rd_label)
        v_layout_rd_log.addLayout(h_layout_rd_log)
        v_layout_log_files.addLayout(v_layout_rd_log)
        self.grp_box_log_files.setLayout(v_layout_log_files)

        self.form_layout.addRow(self.grp_box_log_files)

        # OK and Cancel buttons
        btn_layout = QHBoxLayout()
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self)

        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        btn_layout.addWidget(self.buttons)
        self.form_layout.addRow(btn_layout)

        self.setWindowTitle("Options")
        self.setWindowIcon(QIcon(GRAIP_ICON_FILE))
        self.resize(600, 250)
        self.setLayout(self.form_layout)
        self.setModal(True)

    def browse_dp_log_file(self):
        #working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        dp_log_file, _ = QFileDialog.getSaveFileName(None, 'Enter Drain Point Log Filename', os.getcwd(),
                                                     filter="Drain Point Logfile (*.log)")
        # check if the cancel was clicked on the file dialog
        if len(dp_log_file) == 0:
            return

        self.line_edit_dp_log_file.setText(dp_log_file)

    def browse_rd_log_file(self):
        # working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        rd_log_file, _ = QFileDialog.getSaveFileName(None, 'Enter Road Log Filename', os.getcwd(),
                                                     filter="Road Logfile (*.log)")
        # check if the cancel was clicked on the file dialog
        if len(rd_log_file) == 0:
            return

        self.line_edit_rd_log_file.setText(rd_log_file)

    def get_selected_options(self):
        is_uniterrupted = self.radio_btn_uninterrupted.isChecked()

        return self.line_edit_dp_log_file.text(), self.line_edit_rd_log_file.text(), is_uniterrupted

    @staticmethod
    def get_data_from_dialog(dp_log_file, rd_log_file, is_uninterrupted=False):
        dialog = OptionsDialog()
        if is_uninterrupted:
            dialog.radio_btn_uninterrupted.toggle()
        else:
            dialog.radio_btn_step_by_step.toggle()

        dialog.line_edit_dp_log_file.setText(dp_log_file)
        dialog.line_edit_rd_log_file.setText(rd_log_file)
        dialog.exec_()
        return dialog.get_selected_options()


class DefineValueDialog(QDialog):
    use_default = False
    reassign_value = False
    add_new = False
    is_cancel = False
    definition_id = 0
    action_taken_msg = ""

    def __init__(self, missing_field_value, missing_field_name, graip_db_file, def_table_name,
                 is_multiplier=False, parent=None):
        super(DefineValueDialog, self).__init__(parent)
        self.missing_field_value = missing_field_value
        self.missing_field_name = missing_field_name
        self.graip_db_file = graip_db_file
        self.def_table_name = def_table_name
        self.is_multiplier = is_multiplier
        self.def_default_id = None
        self.form_layout = QFormLayout()
        v_main_layout = QVBoxLayout()
        self.radio_btn_use_default = QRadioButton("Use default value")
        # set this radio button as checked
        #self.radio_btn_use_default.toggle()
        self.radio_btn_reassign = QRadioButton("Reassign this value to an existing value in definitions table")
        self.radio_btn_add_new = QRadioButton("Add new entry to definitions table")
        missing_value_msg = "Value '{}' in field '{}' is not in the definitions " \
                            "table".format(self.missing_field_value, self.missing_field_name)
        self.lbl_missing_value = QLabel(missing_value_msg)
        v_main_layout.addWidget(self.lbl_missing_value)
        h_layout_default = QHBoxLayout()
        self.line_edit_default = QLineEdit()
        self.line_edit_default.setEnabled(False)
        h_layout_default.addWidget(self.radio_btn_use_default)
        h_layout_default.addWidget(self.line_edit_default)
        v_main_layout.addLayout(h_layout_default)
        v_main_layout.addWidget(self.radio_btn_reassign)
        h_layout_definitions = QHBoxLayout()
        self.label_definitions = QLabel("Definitions")
        self.cmb_definitions = QComboBox()
        h_layout_definitions.addWidget(self.label_definitions)
        h_layout_definitions.addWidget(self.cmb_definitions)
        v_main_layout.addLayout(h_layout_definitions)
        v_main_layout.addWidget(self.radio_btn_add_new)

        self.grp_box_new_entry = QGroupBox()
        grid_layout_new_entry = QGridLayout()
        self.label_table_name = QLabel("Table Name")
        self.line_edit_table_name = QLineEdit()
        self.line_edit_table_name.setText(def_table_name)
        self.line_edit_table_name.setEnabled(False)
        #self.label_table_name.setText(self.def_table_name)
        grid_layout_new_entry.addWidget(self.label_table_name, 0, 0)
        grid_layout_new_entry.addWidget(self.line_edit_table_name, 0, 1)

        self.label_id = QLabel("ID")
        self.line_edit_id = QLineEdit()
        self.line_edit_id.setEnabled(False)
        grid_layout_new_entry.addWidget(self.label_id, 1, 0)
        grid_layout_new_entry.addWidget(self.line_edit_id, 1, 1)

        self.label_def = QLabel("Definition")
        self.line_edit_def = QLineEdit()
        grid_layout_new_entry.addWidget(self.label_def, 2, 0)
        grid_layout_new_entry.addWidget(self.line_edit_def, 2, 1)

        self.label_description = QLabel("Description")
        self.line_edit_description = QLineEdit()
        grid_layout_new_entry.addWidget(self.label_description, 3, 0)
        grid_layout_new_entry.addWidget(self.line_edit_description, 3, 1)
        if self.is_multiplier:
            self.label_multiplier = QLabel("Multiplier")
            self.line_edit_multiplier = QLineEdit()
            grid_layout_new_entry.addWidget(self.label_multiplier, 4, 0)
            grid_layout_new_entry.addWidget(self.line_edit_multiplier, 4, 1)

        self.grp_box_new_entry.setLayout(grid_layout_new_entry)
        v_main_layout.addWidget(self.grp_box_new_entry)
        self.form_layout.addRow(v_main_layout)

        # OK and Cancel buttons
        btn_layout = QHBoxLayout()
        self.buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self)

        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        btn_layout.addWidget(self.buttons)
        self.form_layout.addRow(btn_layout)

        self.definition_ids = []
        self._initial_setup()

        self.setWindowTitle("Define Value")
        self.setWindowIcon(QIcon(GRAIP_ICON_FILE))
        self.resize(600, 300)
        self.setLayout(self.form_layout)
        self.setModal(True)

    def _initial_setup(self):
        try:
            conn = pyodbc.connect(MS_ACCESS_CONNECTION % self.graip_db_file)
            cursor = conn.cursor()
            sql_select = "SELECT * FROM {}".format(self.def_table_name)
            def_rows = cursor.execute(sql_select).fetchall()
            # definitions are in the 2nd column of the definitions table
            definitions = [row[1] for row in def_rows]
            self.cmb_definitions.addItems(definitions)
            # definitions ID values are in the 1st column of the definitions table
            self.definition_ids = [row[0] for row in def_rows]
            # set the current index of the combobox
            initial_combobox_index = 0
            field_match_found = False
            # if isinstance(self.missing_field_value, int):
            #     print str(self.missing_field_value)
            missing_field_value = self.missing_field_value.replace(" ", "")
            if len(missing_field_value) > 2:
                for index in range(self.cmb_definitions.count()):
                    definition = self.cmb_definitions.itemText(index)
                    definition = definition.replace(" ", "")
                    if len(definition) > 2:
                        # match at least any first 3 characters
                        match_count = 0
                        for i in range(len(missing_field_value)):
                            if i < len(definition):
                                if missing_field_value[0:i+1].lower() == definition[0:i+1].lower():
                                    match_count += 1

                        if match_count > 2:
                            initial_combobox_index = index
                            field_match_found = True
                            break
                        elif match_count == 1:
                            initial_combobox_index = index

            self.cmb_definitions.setCurrentIndex(initial_combobox_index)
            # matching found
            if field_match_found:
                # set the reassign radio button as checked
                self.radio_btn_reassign.toggle()

            # see default value exists in definitions table
            sql_select = "SELECT * FROM {} WHERE Description LIKE '*Default*'".format(self.def_table_name)
            def_row = cursor.execute(sql_select).fetchone()
            if def_row:
                self.line_edit_default.setText(def_row[1])
                self.def_default_id = def_row[0]
                if not field_match_found:
                    self.radio_btn_use_default.toggle()
            else:
                self.line_edit_default.setText("No Default Specified")
                self.def_default_id = 0
                # disable the radio button for default
                self.radio_btn_use_default.setEnabled(False)
                if not field_match_found:
                    self.radio_btn_add_new.toggle()

            # TODO: Continue here (ref getIDFromDefinitionTable function in mod modGeneralFunctions)
            # frmAddMultiplier.txtID = defID
            # in case of adding a new entry to the definition table, find the next id that will be used for this
            # new definition record
            # first determine the field name of the ID field in the definition table (this is the first column of the
            # table)
            def_column = cursor.columns(table=self.def_table_name).fetchone()
            # now use the def_column.column_name to find the maximum value for this field
            sql_select = "SELECT MAX({}) AS max_id FROM {}".format(def_column.column_name, self.def_table_name)
            row = cursor.execute(sql_select).fetchone()
            def_id = 0
            if row:
                def_id = row.max_id + 1

            self.line_edit_id.setText(str(def_id))
            self.line_edit_def.setText(self.missing_field_value)
            self.line_edit_description.setText(self.missing_field_value)
            if self.is_multiplier:
                self.line_edit_multiplier.setText("1")
        except Exception:
            raise
        finally:
            if conn:
                conn.close()

    def accept(self, *args, **kwargs):
        # This function is called when the OK button of this dialog is clicked
        # ref to the code for the OK button of the frmAddNew
        try:
            conn = pyodbc.connect(MS_ACCESS_CONNECTION % self.graip_db_file)
            cursor = conn.cursor()
            if self.radio_btn_reassign.isChecked():
                self.reassign_value = True
                sql_insert = "INSERT INTO ValueReassigns (FromField, ToField, DefinitionID, DefinitionTable) " \
                             "VALUES (?, ?, ?, ?)"
                self.definition_id = self.definition_ids[self.cmb_definitions.currentIndex()]
                data = (self.line_edit_def.text(), self.cmb_definitions.currentText(), self.definition_id,
                        self.def_table_name)
                self.action_taken_msg = "Reassigned value as {}".format(self.cmb_definitions.currentText())
            elif self.radio_btn_use_default.isChecked():
                self.use_default = True
                self.definition_id = int(self.line_edit_id.text())
                sql_insert = "INSERT INTO ValueReassigns (FromField, ToField, DefinitionID, DefinitionTable) " \
                             "VALUES (?, ?, ?, ?)"
                data = (self.line_edit_def.text(), self.line_edit_default.text(), self.definition_id,
                        self.def_table_name)
                self.action_taken_msg = "Used Default - {}".format(self.line_edit_default.text())
            elif self.radio_btn_add_new.isChecked():
                self.add_new = True
                self.definition_id = int(self.line_edit_id.text())
                if not self.is_multiplier:
                    sql_insert = "INSERT INTO {} VALUES (?, ?, ?)".format(self.def_table_name)
                    data = (self.definition_id, self.line_edit_def.text(), self.line_edit_description.text())
                else:
                    sql_insert = "INSERT INTO {} VALUES (?, ?, ?, ?)".format(self.def_table_name)
                    data = (self.definition_id, self.line_edit_def.text(), self.line_edit_description.text(),
                            self.line_edit_multiplier.text())
                self.action_taken_msg = "Added to definition table as ID {}".format(self.definition_id)
            cursor.execute(sql_insert, data)
            conn.commit()

            # print ("You clicked OK")
            super(DefineValueDialog, self).accept()
        except Exception:
            raise
        finally:
            if conn:
                conn.close()

    def reject(self, *args, **kwargs):
        self.is_cancel = True
        super(DefineValueDialog, self).reject()


class FileDeleteMessageBox(GraipMessageBox):
    def __init__(self, file_to_delete, parent=None):
        super(FileDeleteMessageBox, self).__init__(parent)
        self.setWindowTitle("File action")
        self.setText("{} file exists.".format(file_to_delete))
        self.setInformativeText("Do you want to delete this file?")
        self.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        self.setDefaultButton(QMessageBox.No)


class TableModel(QAbstractTableModel):
    """
    A simple 5x4 table model to demonstrate the delegates
    """
    def __init__(self, parent, data_list, header, *args):
        QAbstractTableModel.__init__(self, parent, *args)
        self.table_data = data_list
        self.header = header

    def rowCount(self, parent=QModelIndex()):
        return len(self.table_data)

    def columnCount(self, parent=QModelIndex()):
        return len(self.table_data[0])

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        if not role == Qt.DisplayRole:
            return None
        return self.table_data[index.row()][index.column()]

    def setData(self, index, value, role=Qt.DisplayRole):
        # update the data bound to the TableView object
        self.table_data[index.row()][index.column()] = value
        print "setData", index.row(), index.column(), value

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header[col]
        return None

    def flags(self, index):
        if index.column() == 1:
            return Qt.ItemIsEditable | Qt.ItemIsEnabled
        else:
            return Qt.ItemIsEnabled


class ComboDelegate(QItemDelegate):
    """
    A delegate that places a fully functioning QComboBox in every
    cell of the column to which it's applied
    """
    def __init__(self, parent, cmb_data_list):
        self.data = cmb_data_list
        QItemDelegate.__init__(self, parent)

    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.addItems(self.data)
        self.connect(combo, SIGNAL("currentIndexChanged(int)"), self, SLOT("currentIndexChanged()"))
        return combo

    def setEditorData(self, editor, index):
        editor.blockSignals(True)
        # get data for the current row in the table
        data = index.model().data(index)
        # get the index for the combobox matching the data for the table
        index = editor.findText(data)
        # if no match was found then set index to 0 (combobox item <No Match Use Default>')
        if index == -1:
            index = 0
        #editor.setCurrentIndex(index.row())
        editor.setCurrentIndex(index)
        #editor.setCurrentIndex(str(index.model().data(index)))
        #print ('editor data:' + str(index.model().data(index)))
        editor.blockSignals(False)

    def setModelData(self, editor, model, index):
        #model.setData(index, editor.currentIndex())
        model.setData(index, editor.currentText())

    @Slot()
    def currentIndexChanged(self):
        self.commitData.emit(self.sender())


class TableView(QTableView):
    """
    A simple table to demonstrate the QComboBox delegate.
    """
    def __init__(self, *args, **kwargs):
        self.combobox_data = kwargs.pop('combobox_data', None)
        QTableView.__init__(self, *args, **kwargs)
        self.setColumnWidth(0, 300)
        self.setColumnWidth(1, 300)
        self.resizeColumnsToContents()

        # Set the delegate for column 1 (2nd column) of our table (col number starts with 0)
        # self.setItemDelegateForColumn(0, ButtonDelegate(self))
        self.setItemDelegateForColumn(1, ComboDelegate(self, cmb_data_list=self.combobox_data))


class TableWidget(QWidget):
        """
        A simple test widget to contain and own the model and table.
        """
        def __init__(self, table_data, table_header, cmb_data, parent=None):
            QWidget.__init__(self, parent)
            self.data = table_data
            l = QVBoxLayout(self)
            self._tm = TableModel(self, data_list=table_data, header=table_header)
            self._tv = TableView(self, combobox_data=cmb_data)
            self._tv.setModel(self._tm)
            self._tv.setColumnWidth(0, 300)
            self._tv.setColumnWidth(1, 300)
            self._tv.resizeColumnsToContents()

            for row in range(0, self._tm.rowCount()):
                self._tv.openPersistentEditor(self._tm.index(row, 1))

            l.addWidget(self._tv)

        @property
        def table_model(self):
            return self._tm

        def closeEvent(self, *args, **kwargs):
            # This is test that we can retrieve the updated data bound to the tableview object
            for row in range(self._tm.rowCount()):
                for col in range(self._tm.columnCount()):
                    #index = QtCore.QModelIndex(row, col)
                    value = str(self._tm.index(row, col).data())
                    print("row:" + str(row) + " col:" + str(col) + " value:" + value)


class ImportWizardPage(QWizardPage):
    def __init__(self, shp_type='DP', shp_file_index=0, shp_file="", shp_file_count=0, parent=None):
        super(ImportWizardPage, self).__init__(parent=parent)
        self.wizard = parent
        self.file_index = shp_file_index
        self.working_directory = None
        self.lst_widget_dp_shp_files = None
        self.no_match_use_default = None
        self.form_layout = QFormLayout()
        self.msg_label = QLabel()
        self.msg_label.setText("Match a source field from the input file to the appropriate target field "
                               "in the database")
        self.msg_label.setWordWrap(True)
        self.form_layout.addRow("", self.msg_label)

        self.group_box_imported_file = QGroupBox("File being imported")
        self.line_edit_imported_file = QLineEdit()
        self.line_edit_imported_file.setText(shp_file)
        v_file_layout = QVBoxLayout()
        v_file_layout.addWidget(self.line_edit_imported_file)
        self.group_box_imported_file.setLayout(v_file_layout)
        self.form_layout.addRow(self.group_box_imported_file)

        self.v_set_fields_layout = QVBoxLayout()
        self.group_box_set_field_names = QGroupBox("Set Field Names")
        if shp_type == "DP":
            self.dp_label = QLabel("Drain Point Type")
            self.dp_type_combo_box = QComboBox()
            self.v_set_fields_layout.addWidget(self.dp_label)
            self.v_set_fields_layout.addWidget(self.dp_type_combo_box)
        else:
            h_rn_layout = QHBoxLayout()
            v_rn_layout = QVBoxLayout()
            self.rd_rn_label = QLabel("Road Network")
            self.rd_network_combo_box = QComboBox()
            self.rd_network_combo_box.setMinimumWidth(150)
            self.rd_network_combo_box.currentIndexChanged.connect(self.update_rd_network_gui_elements)
            v_rn_layout.addWidget(self.rd_rn_label)
            v_rn_layout.addWidget(self.rd_network_combo_box)
            h_rn_layout.addLayout(v_rn_layout)

            v_ber_layout = QVBoxLayout()
            self.rd_ber_label = QLabel("Base Erosion Rate (kg/m/yr)")
            self.rd_ber_line_edit = QLineEdit()
            self.rd_ber_line_edit.setEnabled(False)
            self.rd_ber_line_edit.setMaximumWidth(75)
            v_ber_layout.addWidget(self.rd_ber_label)
            v_ber_layout.addWidget(self.rd_ber_line_edit)
            h_rn_layout.addLayout(v_ber_layout)

            v_des_layout = QVBoxLayout()
            self.rd_des_label = QLabel("Description")
            self.rd_des_line_edit = QLineEdit()
            self.rd_des_line_edit.setEnabled(False)
            v_des_layout.addWidget(self.rd_des_label)
            v_des_layout.addWidget(self.rd_des_line_edit)
            h_rn_layout.addLayout(v_des_layout)

            v_add_btn_layout = QVBoxLayout()
            v_add_btn_layout.addWidget(QLabel("Add"))
            self.add_rn_button = QPushButton()
            self.add_rn_button.setText('+')
            self.add_rn_button.clicked.connect(self.show_add_network_dlg)
            v_add_btn_layout.addWidget(self.add_rn_button)
            h_rn_layout.addLayout(v_add_btn_layout)

            v_del_btn_layout = QVBoxLayout()
            v_del_btn_layout.addWidget(QLabel("Delete"))
            self.del_rn_button = QPushButton()
            self.del_rn_button.setText('-')
            self.del_rn_button.clicked.connect(self.delete_rn_def_record)
            v_del_btn_layout.addWidget(self.del_rn_button)
            h_rn_layout.addLayout(v_del_btn_layout)

            self.v_set_fields_layout.addLayout(h_rn_layout)

        self.table_label = QLabel("For each target field, select the source field that should be loaded into it")
        # TODO: need to pass data for the table to the TableWidget()
        self.field_match_table_wizard = None
        self.v_set_fields_layout.addWidget(self.table_label)

        self.group_box_set_field_names.setLayout(self.v_set_fields_layout)
        self.form_layout.addRow(self.group_box_set_field_names)

        # progress bar
        self.group_box_import_progress = QGroupBox("Import Progress")
        v_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setValue(0)
        v_layout.addWidget(self.progress_bar)
        self.group_box_import_progress.setLayout(v_layout)

        self.form_layout.addRow(self.group_box_import_progress)
        self.setLayout(self.form_layout)
        if shp_type == "DP":
            self.setTitle("Import Drain Point Shapefile: {} of {}".format(self.file_index + 1, shp_file_count))
        else:
            self.setTitle("Import Road Line Shapefile: {} of {}".format(self.file_index + 1, shp_file_count))

    def show_add_network_dlg(self):
        graip_db_file = self.wizard.line_edit_mdb_file.text()
        conn = pyodbc.connect(MS_ACCESS_CONNECTION % graip_db_file)
        cursor = conn.cursor()
        dlg = AddRoadNetworkDefinitionsDialog(graip_db_file)
        dlg.show()
        dlg.exec_()
        if dlg.result() == QDialog.Accepted:
            self.rd_network_combo_box.addItem(dlg.rd_network_name)
            sql_insert = "INSERT INTO RoadNetworkDefinitions(RoadNetwork, BaseRate, Description) VALUES (?, ?, ?)"
            data = (dlg.rd_network_name, dlg.rd_base_rate, dlg.rd_description)
            cursor.execute(sql_insert, data)
            conn.commit()
            # find the number of records
            network_def_row = cursor.execute("SELECT COUNT(*) AS def_count FROM RoadNetworkDefinitions").fetchone()
            # set the network combobox to the just entered new definition
            index = network_def_row.def_count - 1
            self.rd_network_combo_box.setCurrentIndex(index)
            conn.close()

    def update_rd_network_gui_elements(self):
        graip_db_file = self.wizard.line_edit_mdb_file.text()
        conn = pyodbc.connect(MS_ACCESS_CONNECTION % graip_db_file)
        cursor = conn.cursor()
        network_combo_index = self.rd_network_combo_box.currentIndex()
        road_network_id = network_combo_index + 1
        rd_network_def_row = cursor.execute("SELECT * FROM RoadNetworkDefinitions WHERE RoadNetworkID=?",
                                            road_network_id).fetchone()
        if rd_network_def_row is not None:
            self.rd_ber_line_edit.setText(str(rd_network_def_row.BaseRate))
            self.rd_des_line_edit.setText(rd_network_def_row.Description)
        conn.close()

    def delete_rn_def_record(self):
        graip_db_file = self.wizard.line_edit_mdb_file.text()
        conn = pyodbc.connect(MS_ACCESS_CONNECTION % graip_db_file)
        cursor = conn.cursor()
        msg_box = GraipMessageBox()
        msg_box.setWindowTitle("Not allowed")
        msg_box.setText("Network 'Default' can't be deleted")
        selected_network_name = self.rd_network_combo_box.currentText()
        if selected_network_name.lower() == 'default':
            msg_box.exec_()
        else:
            cursor.execute("DELETE FROM RoadNetworkDefinitions WHERE RoadNetwork=?", selected_network_name)
            conn.commit()
            conn.close()
            current_index = self.rd_network_combo_box.currentIndex()
            self.rd_network_combo_box.removeItem(current_index)


def delete_shapefile(shp_file_to_delete):
    if os.path.isfile(shp_file_to_delete):
        # prompt to delete the file
        file_name = os.path.basename(shp_file_to_delete)
        file_delete_msgbox = FileDeleteMessageBox(file_name)
        user_input = file_delete_msgbox.exec_()
        if user_input == QMessageBox.Yes:
            gdal_driver = ogr.GetDriverByName(GDALFileDriver.ShapeFile())
            gdal_driver.DeleteDataSource(shp_file_to_delete)


def create_log_file(graip_db_file, log_file, log_type):
    with open(log_file, 'w') as file_obj:
        file_obj.write("GRAIP Database File:{}".format(graip_db_file) + '\n')
        file_obj.write(time.strftime("%m-%d-%Y %H:%M:%S") + '\n')
        if log_type == 'DP':
            file_obj.write("GRAIPDID, Drain Type, Error Message, Action Taken \n")
        else:
            file_obj.write("GRAIPDID, Road Type, Error Message, Action Taken \n")


def clear_data_tables(graip_db_file):
    conn = pyodbc.connect(MS_ACCESS_CONNECTION % graip_db_file)
    cursor = conn.cursor()
    cursor.execute("DELETE * FROM DrainPoints")
    cursor.execute("DELETE * FROM RoadLines")
    drain_type_def_rows = cursor.execute("SELECT TableName FROM DrainTypeDefinitions").fetchall()
    for row in drain_type_def_rows:
        cursor.execute("DELETE * FROM {}".format(row.TableName))
    conn.commit()
    conn.close()


def get_shapefile_attribute_column_names(shp_file):
    gdal_driver = ogr.GetDriverByName(GDALFileDriver.ShapeFile())
    data_source = gdal_driver.Open(shp_file, 1)
    layer = data_source.GetLayer(0)
    layer_definition = layer.GetLayerDefn()

    attribute_names = []
    for i in range(layer_definition.GetFieldCount()):
        att_name = layer_definition.GetFieldDefn(i).GetName()
        if att_name != 'FID' and att_name != 'Shape':
            attribute_names.append(att_name)

    data_source.Destroy()
    return attribute_names


def populate_drain_type_combobox(graip_db_file, dp_type_combo_box):
    conn = pyodbc.connect(MS_ACCESS_CONNECTION % graip_db_file)
    cursor = conn.cursor()
    drain_point_def_rows = cursor.execute("SELECT DrainTypeName FROM DrainTypeDefinitions").fetchall()
    drain_point_types = [row.DrainTypeName for row in drain_point_def_rows]
    dp_type_combo_box.addItems(drain_point_types)
    conn.close()
    return dp_type_combo_box


def set_index_dp_type_combo_box(dp_shp_file, dp_type_combo_box):
    dp_shp_file_name = os.path.basename(dp_shp_file)

    matching_index = 0
    for index in range(dp_type_combo_box.count()):
        drain_type_name = dp_type_combo_box.itemText(index)
        no_char_match = 0
        for i in range(1, len(drain_type_name)):
            if drain_type_name[0:i] == dp_shp_file_name[0:i]:
                no_char_match += 1

        if no_char_match > 2:
            dp_type_combo_box.setCurrentIndex(index)
            return dp_type_combo_box
        if no_char_match == 1:
            matching_index = index

    dp_type_combo_box.setCurrentIndex(matching_index)
    return dp_type_combo_box


def add_entry_to_error_table(graip_db_file, err_table_name, graipid, ftype, err_msg, action_taken):
    if err_table_name not in (DP_ERROR_LOG_TABLE_NAME, RD_ERROR_LOG_TABLE_NAME):
        raise Exception("{} is not a valid table name for logging error.".format(err_table_name))

    conn = pyodbc.connect(MS_ACCESS_CONNECTION % graip_db_file)
    cursor = conn.cursor()
    if err_table_name == DP_ERROR_LOG_TABLE_NAME:
        sql_insert = "INSERT INTO {} (GRAIPDID, DrainType, ErrorMessage, ActionTaken) VALUES(?, ?, ?, ?)"
        sql_insert = sql_insert.format(DP_ERROR_LOG_TABLE_NAME)
    else:
        sql_insert = "INSERT INTO {} (GRAIPRID, RoadType, ErrorMessage, ActionTaken) VALUES(?, ?, ?, ?)"
        sql_insert = sql_insert.format(RD_ERROR_LOG_TABLE_NAME)

    data = (graipid, ftype, err_msg, action_taken)
    cursor.execute(sql_insert, data)
    conn.commit()
    conn.close()


def add_entry_to_log_file(log_file, graipid, ftype, message, action_taken):
    with open(log_file, 'a') as file_obj:
        text_to_write = "{}, {}, {}, {} \n"
        text_to_write = text_to_write.format(graipid, ftype, message, action_taken)
        file_obj.write(text_to_write)


def is_data_type_match(graip_db_file, table_name, col_name, data_value):
    conn = pyodbc.connect(MS_ACCESS_CONNECTION % graip_db_file)
    cursor = conn.cursor()
    for row in cursor.columns(table=table_name):
        if row.column_name == col_name:
            if row.data_type in (pyodbc.SQL_CHAR, pyodbc.SQL_VARCHAR, pyodbc.SQL_WCHAR, pyodbc.SQL_WLONGVARCHAR,
                                 pyodbc.SQL_WVARCHAR, pyodbc.SQL_LONGVARCHAR):
                return type(data_value) is str
            elif row.data_type in (pyodbc.SQL_TINYINT, pyodbc.SQL_SMALLINT, pyodbc.SQL_INTEGER, pyodbc.SQL_BIGINT,
                                   pyodbc.SQL_NUMERIC):
                return type(data_value) is int or type(data_value) is float
            elif row.data_type in (pyodbc.SQL_DECIMAL, pyodbc.SQL_FLOAT, pyodbc.SQL_DOUBLE, pyodbc.SQL_REAL):
                return type(data_value) is float or type(data_value) is int
            elif row.data_type in (pyodbc.SQL_TYPE_DATE, pyodbc.SQL_TYPE_TIMESTAMP):
                if type(data_value) is datetime:
                    return True
                if isinstance(data_value, basestring):
                    if len(data_value) < 8:
                        return False
                    else:
                        # here we are checking only 2 date format (e.g., 01/12/2016 and 2016/12/01)
                        try:
                            datetime.strptime(data_value, "%m/%d/%Y")
                            return True
                        except:
                            try:
                                datetime.strptime(data_value, "%Y/%m/%d")
                                return True
                            except:
                                return False

                return False

    conn.close()
    return False


def get_drain_id(time_in, date_in, vehicle_id):
    # remove am/pm
    if isinstance(time_in, basestring):
        time = time_in[:-2]
        time_hour, time_min, time_sec = time.split(":")
        time_hour = int(time_hour)
        time_min = int(time_min)
        time_sec = int(time_sec)
        if time_sec > 30:
            time_min += 1
        if "p" in time_in and time_hour != 12:
            time_hour += 12

        if "a" in time_in and time_hour == 12:
            time_hour = 0
        time_24_format = (time_hour * 100) + time_min
    else:
        time_24_format = time_in
    dt_obj = date_in
    # get 2 digit year
    year = int(str(dt_obj.year)[2:])
    drain_id = (year * 1000000000) + (dt_obj.month * 10000000) + (dt_obj.day * 100000) + \
               (time_24_format * 10) + vehicle_id
    return drain_id


def get_table_column_data_type(graip_db_file, table_name, col_name):
    conn = pyodbc.connect(MS_ACCESS_CONNECTION % graip_db_file)
    cursor = conn.cursor()
    for row in cursor.columns(table=table_name):
        if row.column_name == col_name:
            return row.data_type

    return None


def get_items_from_list_box(list_box):
    item_list = []
    for i in range(list_box.count()):
        item = list_box.item(i).text()
        item_list.append(item)
    return item_list