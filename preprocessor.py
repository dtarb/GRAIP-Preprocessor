import sys
import os
import shutil
from datetime import datetime
from dateutil import parser

import pyodbc
from osgeo import ogr, gdal, osr
from gdalconst import *

from PySide.QtGui import *
from PySide.QtCore import *

import utils


"""
Ref http://srinikom.github.io/pyside-docs/PySide/QtGui/QWizard.html
"""
qt_app = QApplication(sys.argv)


class Preprocessor(QWizard):
    def __init__(self, parent=None):
        super(Preprocessor, self).__init__(parent)
        self.dp_log_file = None
        self.rd_log_file = None
        self.is_uninterrupted = False
        self.options_dlg = utils.OptionsDialog()

        self.setWindowTitle("GRAIP Preprocessor (Version 2.0)")
        # add a custom button
        self.btn_options = QPushButton('Options')
        #self.btn_options.clicked.connect(self.show_options_dialog)
        self.setButton(self.CustomButton1, self.btn_options)
        self.setOptions(self.HaveCustomButton1)

        # add pages
        self.addPage(FileSetupPage(parent=self))
        self.addPage(DrainPointPage(shp_file_index=0, shp_file="", parent=self))

        self.currentIdChanged.connect(self.show_options_button)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.resize(1800, 1000)

    def show_options_button(self):
        if self.currentId() == 0:
            # show the Options button if it is the FileSetup page
            self.btn_options.show()

    # def show_options_dialog(self):
    #     self.dp_log_file, self.rd_log_file, self.is_uninterrupted = self.options_dlg.get_data_from_dialog(
    #         self.dp_log_file, self.rd_log_file)

    def run(self):
        # Show the form
        self.show()

        # Run the qt application
        qt_app.exec_()


class FileSetupPage(QWizardPage):
    def __init__(self, parent=None):
        super(FileSetupPage, self).__init__(parent)
        self.wizard = parent
        self.wizard.btn_options.clicked.connect(self.show_options_dialog)
        self.working_directory = None
        self.dp_shp_file_processing_track_dict = {}
        self.form_layout = QFormLayout()
        self.msg_label = QLabel()
        self.msg_label.setText("The GRAIP Preprocessor is a tool to import USDA Forest Service road inventory "
                               "information into a MS Access database for use by GRAIP Analysis Tools in ArcGIS")
        self.msg_label.setWordWrap(True)
        self.form_layout.addRow("", self.msg_label)

        self.group_box_mdb_file = QGroupBox("GRAIP Database (*.mdb)")
        self.line_edit_mdb_file = QLineEdit()
        self.btn_browse_mdb_file = QPushButton('....')
        self.btn_browse_mdb_file.clicked.connect(self.browse_db_file)

        layout_mdb = QHBoxLayout()
        layout_mdb.addWidget(self.line_edit_mdb_file)
        layout_mdb.addWidget(self.btn_browse_mdb_file)
        self.group_box_mdb_file.setLayout(layout_mdb)
        self.form_layout.addRow(self.group_box_mdb_file)

        # self.group_box_prj_file = QGroupBox("Project File")
        # self.line_edit_prj_file = QLineEdit()
        # self.btn_browse_prg_file = QPushButton('....')
        # # connect the browse button to the function
        # self.btn_browse_prg_file.clicked.connect(self.browse_project_file)
        #
        # layout_prj = QHBoxLayout()
        # layout_prj.addWidget(self.line_edit_prj_file)
        # layout_prj.addWidget(self.btn_browse_prg_file)
        # self.group_box_prj_file.setLayout(layout_prj)
        # self.form_layout.addRow(self.group_box_prj_file)

        layout_input_files = QVBoxLayout()
        self.group_box_input_files = QGroupBox("Input Files")

        self.dem_file_label = QLabel()
        self.dem_file_label.setText("DEM File('sta.adf' inside the DEM folder)")
        self.dem_file_label.setWordWrap(True)
        #self.form_layout.addRow("", self.dem_file_label)
        v_layout_dem_file = QVBoxLayout()
        v_layout_dem_file.addWidget(self.dem_file_label)
        h_layout_dem_files = QHBoxLayout()
        self.line_edit_dem_file = QLineEdit()
        self.btn_browse_dem_file = QPushButton('....')
        # connect the browse button to the function
        self.btn_browse_dem_file.clicked.connect(self.browse_dem_file)

        h_layout_dem_files.addWidget(self.line_edit_dem_file)
        h_layout_dem_files.addWidget(self.btn_browse_dem_file)
        v_layout_dem_file.addLayout(h_layout_dem_files)
        layout_input_files.addLayout(v_layout_dem_file)

        v_layout_rd_shp_files = QVBoxLayout()
        self.rd_shp_file_label = QLabel()
        self.rd_shp_file_label.setText("Road Shapefiles")
        self.rd_shp_file_label.setWordWrap(True)
        v_layout_rd_shp_files.addWidget(self.rd_shp_file_label)

        self.lst_widget_road_shp_files = QListWidget()
        # To add an item to the above listWidget,
        # ref: http://srinikom.github.io/pyside-docs/PySide/QtGui/QListWidget.html
        v_layout_rd_shp_buttons = QVBoxLayout()
        self.btn_add_rd_shp_file = QPushButton('Add')
        self.btn_add_rd_shp_file.clicked.connect(self.browse_rd_shp_files)
        self.btn_remove_rd_shp_file = QPushButton('Remove')
        self.btn_remove_rd_shp_file.clicked.connect(self.remove_rd_shp_files)

        v_layout_rd_shp_buttons.addWidget(self.btn_add_rd_shp_file)
        v_layout_rd_shp_buttons.addWidget(self.btn_remove_rd_shp_file)
        h_layout_rd_shp_files = QHBoxLayout()
        h_layout_rd_shp_files.addWidget(self.lst_widget_road_shp_files)
        h_layout_rd_shp_files.addLayout(v_layout_rd_shp_buttons)
        v_layout_rd_shp_files.addLayout(h_layout_rd_shp_files)
        layout_input_files.addLayout(v_layout_rd_shp_files)

        v_layout_dp_shp_files = QVBoxLayout()
        self.dp_shp_file_label = QLabel()
        self.dp_shp_file_label.setText("Drain Points Shapefiles")
        self.dp_shp_file_label.setWordWrap(True)
        v_layout_dp_shp_files.addWidget(self.dp_shp_file_label)

        self.lst_widget_dp_shp_files = QListWidget()
        # To add an item to the above listWidget,
        # ref: http://srinikom.github.io/pyside-docs/PySide/QtGui/QListWidget.html
        v_layout_dp_shp_buttons = QVBoxLayout()
        self.btn_add_dp_shp_file = QPushButton('Add')
        self.btn_add_dp_shp_file.clicked.connect(self.browse_dp_shp_files)
        self.btn_remove_dp_shp_file = QPushButton('Remove')
        self.btn_remove_dp_shp_file.clicked.connect(self.remove_dp_shp_files)

        v_layout_dp_shp_buttons.addWidget(self.btn_add_dp_shp_file)
        v_layout_dp_shp_buttons.addWidget(self.btn_remove_dp_shp_file)
        h_layout_dp_shp_files = QHBoxLayout()
        h_layout_dp_shp_files.addWidget(self.lst_widget_dp_shp_files)
        h_layout_dp_shp_files.addLayout(v_layout_dp_shp_buttons)
        v_layout_dp_shp_files.addLayout(h_layout_dp_shp_files)
        layout_input_files.addLayout(v_layout_dp_shp_files)

        self.group_box_input_files.setLayout(layout_input_files)
        self.form_layout.addRow(self.group_box_input_files)
        self.dp_log_file = self.wizard.dp_log_file
        self.rd_log_file = self.wizard.rd_log_file
        self.is_uninterrupted = self.wizard.is_uninterrupted
        # self.group_box_mdb_file = QGroupBox("GRAIP Database (*.mdb)")
        # self.line_edit_mdb_file = QLineEdit()
        # self.btn_browse_mdb_file = QPushButton('....')
        # self.btn_browse_mdb_file.clicked.connect(self.browse_db_file)
        #
        # layout_mdb = QHBoxLayout()
        # layout_mdb.addWidget(self.line_edit_mdb_file)
        # layout_mdb.addWidget(self.btn_browse_mdb_file)
        # self.group_box_mdb_file.setLayout(layout_mdb)
        # self.form_layout.addRow(self.group_box_mdb_file)

        # make some of the page fields available to other pages
        # make the list widget for drain point shapefiles available to other pages of this wizard
        self.current_imported_dp_file = QLineEdit()
        self.registerField("current_imported_dp_file", self.current_imported_dp_file, 'text')

        self.setLayout(self.form_layout)
        self.setTitle("File Setup")

    def show_options_dialog(self):
        if len(self.line_edit_mdb_file.text().strip()) == 0:
            msg_box = QMessageBox()
            msg_box.setWindowTitle("GRAIP Database File")
            msg_box.setText("Before you can set options, select GRAIP database file.")
            msg_box.exec_()
        else:
            self.dp_log_file, self.rd_log_file, self.is_uninterrupted = self.wizard.options_dlg.get_data_from_dialog(
                self.dp_log_file, self.rd_log_file, self.is_uninterrupted)

    def validatePage(self, *args, **kwargs):
        # This function gets called when the next button is clicked for the FileSetup page

        # remove if there are any drainpoint wizard pages
        for page_id in self.wizard.pageIds():
            if page_id > 0:
                self.wizard.removePage(page_id)

        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        # check project file has been selected
        # if len(self.line_edit_prj_file.text().strip()) == 0:
        #     msg_box.setText("Project file is required")
        #     msg_box.exec_()
        #     return False

        if len(self.line_edit_dem_file.text().strip()) == 0:
            msg_box.setText("DEM file is required")
            msg_box.exec_()
            return False

        if len(self.line_edit_mdb_file.text().strip()) == 0:
            msg_box.setText("GRAIP database file is required")
            msg_box.exec_()
            return False

        if self.lst_widget_road_shp_files.count() == 0:
            msg_box.setText("At least one road shapefile is required")
            msg_box.exec_()
            return False

        if self.lst_widget_dp_shp_files.count() == 0:
            msg_box.setText("At least one drain points shapefile is required")
            msg_box.exec_()
            return False

        dp_shp_file_list = []
        shp_file_count = self.lst_widget_dp_shp_files.count()
        for i in range(self.lst_widget_dp_shp_files.count()):
            shp_file = self.lst_widget_dp_shp_files.item(i).text()
            dp_shp_file_list.append(shp_file)
            self.wizard.addPage(DrainPointPage(shp_file_index=i, shp_file=shp_file, shp_file_count=shp_file_count,
                                               parent=self))

        rd_shp_file_list = []
        for i in range(self.lst_widget_road_shp_files.count()):
            shp_file = self.lst_widget_road_shp_files.item(i).text()
            rd_shp_file_list.append(shp_file)
            # TODO: implement RoadLinesPage
            # self.wizard.addPage(RoadLinesPage(shp_file_index=i, shp_file=shp_file, shp_file_count=shp_file_count,
            #                                   parent=self))

        if self.lst_widget_dp_shp_files.count() > 0:
            self.current_imported_dp_file.setText(self.lst_widget_dp_shp_files.item(0).text())
            self.wizard.setStartId(0)
            self.wizard.restart()
            total_wizard_pages = len(self.wizard.pageIds())
            if total_wizard_pages > self.lst_widget_dp_shp_files.count() + 1:   # 1 is for the FileSetup page
                # here we are removing the DrainPoint page we added in the init of the wizard
                self.wizard.removePage(1)

        # write file setup data to the FileSetup table
        graip_db_file = self.line_edit_mdb_file.text()
        conn = pyodbc.connect(utils.MS_ACCESS_CONNECTION % graip_db_file)
        cursor = conn.cursor()
        # delete all data from FileSetup table
        cursor.execute("DELETE * FROM FileSetup")
        conn.commit()

        # insert data to FileSetup table
        db_file = self.line_edit_mdb_file.text()
        dem_file_path = self.line_edit_dem_file.text()
        dp_shp_files = ','.join(dp_shp_file_list)
        rd_shp_files = ','.join(rd_shp_file_list)
        cursor.execute("INSERT INTO FileSetup(GRAIP_DB_File, DEM_Path, Road_Shapefiles, DrainPoints_Shapefiles) "
                       "VALUES (?, ?, ?, ?)", db_file, dem_file_path, rd_shp_files, dp_shp_files)
        conn.commit()
        conn.close()

        # check if there exists 'DrainPoints.shp' or 'RoadLines.shp file at the same location as the db file
        # and prompt user if it needs to be deleted and then delete it if user says so
        db_directory = QFileInfo(graip_db_file).path()
        shp_file = os.path.join(db_directory, 'DrainPoints.shp')
        if os.path.isfile(shp_file):
            utils.delete_shapefile(shp_file)

        shp_file = os.path.join(db_directory, 'RoadLines.shp')
        if os.path.isfile(shp_file):
            utils.delete_shapefile(shp_file)

        # create 2 log files
        if self.dp_log_file is None:
            db_file_name = os.path.basename(graip_db_file)
            db_file_name_wo_ext = db_file_name.split('.')[0]
            self.dp_log_file = os.path.join(self.working_directory, db_file_name_wo_ext + 'DP.log')
            self.rd_log_file = os.path.join(self.working_directory, db_file_name_wo_ext + 'RD.log')

        utils.create_log_file(graip_db_file=graip_db_file, log_file=self.dp_log_file, log_type='DP')
        utils.create_log_file(graip_db_file=graip_db_file, log_file=self.rd_log_file, log_type='RD')

        # cleanup graip database relevant tables in preparation for loading new data
        utils.clear_data_tables(graip_db_file)
        # hide the Options button
        self.wizard.btn_options.hide()
        return True

    # TODO: This method is no more used - so delete it
    def browse_project_file(self):
        graip_prj_file, _ = QFileDialog.getSaveFileName(None, 'Enter GRAIP Project Filename', os.getcwd(),
                                                        filter="GRAIP (*.graip)")
        # check if the cancel was clicked on the file dialog
        if len(graip_prj_file) == 0:
            return
        # change the directory separator to windows
        graip_prj_file = os.path.abspath(graip_prj_file)

        self.line_edit_prj_file.setText(graip_prj_file)
        self.working_directory = os.path.abspath(QFileInfo(graip_prj_file).path())

        # set the path of the GRAIP database file if it is not already set
        if len(self.line_edit_mdb_file.text()) == 0:
            # get the file name of the project
            graip_prj_file_name = QFileInfo(self.line_edit_prj_file.text()).fileName()
            name, ext = graip_prj_file_name.split(".")
            db_file_name = name + ".mdb"
            db_file = os.path.join(self.working_directory, db_file_name)
            self.line_edit_mdb_file.setText(db_file)

    def browse_dem_file(self):
        working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        graip_dem_file, _ = QFileDialog.getOpenFileName(None, "Select file 'sta.adf'", working_dir,
                                                        filter="DEM (*.adf)")
        # check if the cancel was clicked on the file dialog
        if len(graip_dem_file) == 0:
            return

        self.line_edit_dem_file.setText(QFileInfo(graip_dem_file).path())

    def browse_rd_shp_files(self):
        working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        rd_shp_files, _ = QFileDialog.getOpenFileNames(None, "Select Road Shapefiles", working_dir,
                                                       filter="Shapefiles (*.shp)")

        # TODO: If the file already in the list widget do not add that
        self.lst_widget_road_shp_files.addItems(rd_shp_files)

    def remove_rd_shp_files(self):
        for shp_file_item_index in self.lst_widget_road_shp_files.selectedIndexes():
            item = self.lst_widget_road_shp_files.takeItem(shp_file_item_index.row())
            item = None

    def browse_dp_shp_files(self):
        working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        dp_shp_files, _ = QFileDialog.getOpenFileNames(None, "Select Drain Points Shapefiles", working_dir,
                                                       filter="Shapefiles (*.shp)")

        # TODO: If the file already in the list widget do not add that
        self.lst_widget_dp_shp_files.addItems(dp_shp_files)

    def remove_dp_shp_files(self):
        for shp_file_item_index in self.lst_widget_dp_shp_files.selectedIndexes():
            item = self.lst_widget_dp_shp_files.takeItem(shp_file_item_index.row())
            item = None

    def browse_db_file(self):
        working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        graip_db_file, _ = QFileDialog.getSaveFileName(None, 'Enter GRAIP Database Filename', working_dir,
                                                       filter="GRAIP (*.mdb)", options=QFileDialog.DontConfirmOverwrite)
        # check if the cancel was clicked on the file dialog
        if len(graip_db_file) == 0:
            return

        # change the directory separator to windows
        graip_db_file = os.path.abspath(graip_db_file)

        self.working_directory = os.path.abspath(QFileInfo(graip_db_file).path())
        # if the database file does not exists then copy the empty database
        # as the filename provided by the user
        if not os.path.isfile(graip_db_file):
            this_script_dir = os.path.dirname(os.path.realpath(__file__))
            empty_graip_db_file = os.path.join(this_script_dir, 'GRAIP_DB', "GRAIP.mdb")
            shutil.copyfile(empty_graip_db_file, graip_db_file)
        else:
            # prompt the user if the existing db file to be overwritten
            msg_box = QMessageBox()
            msg = "File {} exits. Do you want to overwrite it?".format(graip_db_file)
            msg_box.setText(msg)
            msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg_box.setDefaultButton(QMessageBox.No)
            response = msg_box.exec_()
            if response == QMessageBox.Yes:
                os.remove(graip_db_file)
                this_script_dir = os.path.dirname(os.path.realpath(__file__))
                empty_graip_db_file = os.path.join(this_script_dir, 'GRAIP_DB', "GRAIP.mdb")
                shutil.copyfile(empty_graip_db_file, graip_db_file)

            else:
                # read file setup data from the FileSetup table and populate various file inputs
                # in this wizard page
                conn = pyodbc.connect(utils.MS_ACCESS_CONNECTION % graip_db_file)
                cursor = conn.cursor()
                file_setup_row = cursor.execute("SELECT * FROM FileSetup").fetchone()
                conn.close()
                if file_setup_row:
                    self.line_edit_dem_file.setText(file_setup_row.DEM_Path)
                    road_shp_files = file_setup_row.Road_Shapefiles.split(',')
                    self.lst_widget_road_shp_files.addItems(road_shp_files)
                    dp_shp_files = file_setup_row.DrainPoints_Shapefiles.split(',')
                    self.lst_widget_dp_shp_files.addItems(dp_shp_files)

        db_file_name = os.path.basename(graip_db_file)
        db_file_name_wo_ext = db_file_name.split('.')[0]
        self.dp_log_file = os.path.join(self.working_directory, db_file_name_wo_ext + 'DP.log')
        self.rd_log_file = os.path.join(self.working_directory, db_file_name_wo_ext + 'RD.log')
        self.line_edit_mdb_file.setText(graip_db_file)


class DrainPointPage(QWizardPage):
    def __init__(self, shp_file_index=0, shp_file="", shp_file_count=0, parent=None):
        super(DrainPointPage, self).__init__(parent=parent)
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

        self.v_dp_layout = QVBoxLayout()
        self.group_box_set_field_names = QGroupBox("Set Field Names")
        self.dp_label = QLabel("Drain Point Type")
        self.dp_type_combo_box = QComboBox()
        self.v_dp_layout.addWidget(self.dp_label)
        self.v_dp_layout.addWidget(self.dp_type_combo_box)

        self.table_label = QLabel("For each target field, select the source field that should be loaded into it")
        # TODO: need to pass data for the table to the TableWidget()
        self.field_match_table_wizard = None
        self.v_dp_layout.addWidget(self.table_label)

        self.group_box_set_field_names.setLayout(self.v_dp_layout)
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
        self.setTitle("Import Drain Point Shapefile: {} of {}".format(self.file_index + 1, shp_file_count))

    def initializePage(self, *args, **kwargs):
        graip_db_file = self.wizard.line_edit_mdb_file.text()
        conn = pyodbc.connect(utils.MS_ACCESS_CONNECTION % graip_db_file)
        cursor = conn.cursor()
        # populate the field match table based on the drain point shapefile being imported
        if self.field_match_table_wizard is None:
            dp_shapefile = self.line_edit_imported_file.text()
            shp_file_attribute_names = utils.get_shapefile_attribute_column_names(dp_shapefile)
            self.no_match_use_default = '<No Match Use Default>'
            shp_file_attribute_names = [self.no_match_use_default] + shp_file_attribute_names
            table_headers = ['Target Field', 'Matching Source Field']
            self.dp_type_combo_box = utils.populate_drain_type_combobox(graip_db_file, self.dp_type_combo_box)
            # set the current index for the drain type combo box
            self.dp_type_combo_box = utils.set_index_dp_type_combo_box(dp_shapefile, self.dp_type_combo_box)
            drain_type_name = self.dp_type_combo_box.currentText()
            drain_type_def_row = cursor.execute("SELECT DrainTypeID FROM DrainTypeDefinitions "
                                                "WHERE DrainTypeName = ?", drain_type_name).fetchone()

            # find the target field names in the database corresponding to the shapefile being imported
            field_name_rows = cursor.execute("SELECT DBField FROM FieldMatches "
                                             "WHERE AttTableID = ?", drain_type_def_row.DrainTypeID).fetchall()
            target_field_col_data = [row.DBField for row in field_name_rows]

            source_field_col_data = []
            for target_fld in target_field_col_data:
                source_field_row = cursor.execute("SELECT DBFField FROM FieldMatches "
                                                  "WHERE DBField = ?", target_fld).fetchone()
                if source_field_row:
                    # check if the DBFField value matches (if at least first 3 chars need to match)
                    # with any of the values in the combobox used for the 2nd column of the table
                    found_match = False
                    for shp_att_name in shp_file_attribute_names:
                        if shp_att_name.lower() == source_field_row.DBFField.lower():
                            source_field_col_data.append(source_field_row.DBFField)
                            found_match = True
                            break

                    if not found_match:
                        for shp_att_name in shp_file_attribute_names:
                            if len(shp_att_name) > 2 and len(source_field_row.DBFField) > 2:
                                # match first 3 characters
                                if shp_att_name[0:2].lower() == source_field_row.DBFField[0:2].lower():
                                    #source_field_col_data.append(source_field_row.DBFField)
                                    source_field_col_data.append(shp_att_name)
                                    found_match = True
                                    break
                    if not found_match:
                        source_field_col_data.append(self.no_match_use_default)

            target_source_combined = zip(target_field_col_data, source_field_col_data)
            table_data = [[item[0], item[1]] for item in target_source_combined]
            #table_data = [['CDate', 'CDATE'], ['CTime', 'CTIME'], ['VehicleID', 'VEHICLE']]
            cmb_data = shp_file_attribute_names
            #cmb_data = ['CDATE', 'CTIME', 'STREAM_CON', 'SLOPE_SHP']
            self.field_match_table_wizard = utils.TableWidget(table_data=table_data, table_header=table_headers,
                                                              cmb_data=cmb_data)

            self.v_dp_layout.addWidget(self.field_match_table_wizard)
            # drain_point_def_rows = cursor.execute("SELECT DrainTypeName FROM DrainTypeDefinitions").fetchall()
            # drain_point_types = [row.DrainTypeName for row in drain_point_def_rows]
            # #drain_point_types = ['Broad base type', 'Diffuse drain', 'Ditch relief', 'Lead off']
            # self.dp_type_combo_box.addItems(drain_point_types)
            conn.close

    def validatePage(self, *args, **kwargs):
        # this function is executed when next button is selected
        # here we should be processing the currently imported drain shape file
        # during the processing if we find a value that's not in the definitions table we need
        # to show the DefineValueDialog
        #get general drainpoint table joined with specific drainpoint table from the access database
        #selectStmt = "SELECT * FROM DrainPoints," & impVar.attTableNames & _
        # " WHERE DrainPoints.GRAIPDID=" & impVar.attTableNames & _
        # ".GRAIPDID ORDER BY DrainPoints.GRAIPDID ASC"
        is_error = False
        try:
            graip_db_file = self.wizard.line_edit_mdb_file.text()
            conn = pyodbc.connect(utils.MS_ACCESS_CONNECTION % graip_db_file)
            cursor = conn.cursor()
            dp_table_field_names = ['GRAIPDID', 'DrainTypeID', 'CDate', 'CTime', 'VehicleID', 'DrainID',
                                    'StreamConnectID', 'OrphanID', 'DischargeToID', 'Comments']
            # Testing strategy for inserting records to DrainPoints table
            # cursor.execute("INSERT INTO DrainPoints (GRAIPDID, CTime) VALUES(?, ?)", -1, "")
            # conn.commit()
            # # get the record just added
            # dp_new_row = cursor.execute("SELECT * FROM DrainPoints WHERE GRAIPDID=?", -1).fetchone()
            # if dp_new_row:
            #     print "working"
            # else:
            #     print "failed"
            #
            # return

            drain_type_name = self.dp_type_combo_box.currentText()
            # find drain type attribute table name - this is the table to which we will be writing data in addition
            # to writing data to the DrainPoints table
            drain_type_def_row = cursor.execute("SELECT DrainTypeID, TableName FROM DrainTypeDefinitions WHERE "
                                                "DrainTypeName=?", drain_type_name).fetchone()
            if drain_type_def_row:
                # data for the fields (GRAIPDID, DrainTypeID, CDate, CTime, VehicleID, DrainID, StreamConnectID, OrphanID,
                # DischargeToID) needs to be written to the DrainPoints table, Data for other fields need to be written to
                # the corresponding matching drainpoint attribute table. Data first need to be written to the DrainPoints
                # table then to the matching attribute table

                # add
                # cursor.execute("SELECT * FROM DrainPoints, ? WHERE DrainPoints.GRAIPDID=?", drain_type_def_row.TableName,
                #                drain_type_def_row.TableName + '.GRAIPDID')

                # check if the drainpoints file has been processed before
                update_main_dp_table = False
                dp_shapefile = self.line_edit_imported_file.text()
                if dp_shapefile in self.wizard.dp_shp_file_processing_track_dict:
                    graipid = self.wizard.dp_shp_file_processing_track_dict[dp_shapefile]
                    # delete all records from the matching attribute table
                    cursor.execute("DELETE * FROM {}".format(drain_type_def_row.TableName))
                    conn.commit()
                    # In this case we will be updating records in the DrainPoints table
                    update_main_dp_table = True
                else:
                    # find out the next GRAIPDID from the DrainPoints table
                    dp_row = cursor.execute("SELECT MAX(GRAIPDID)AS Max_GRAIPDID FROM DrainPoints").fetchone()
                    if dp_row[0] is not None:
                        graipid = dp_row.Max_GRAIPDID + 1
                    else:
                        graipid = 0

                    self.wizard.dp_shp_file_processing_track_dict[dp_shapefile] = graipid
            else:
                raise Exception("No matching drain type attribute table was found")

            # open drainpoint shapefile
            gdal_driver = ogr.GetDriverByName(utils.GDALFileDriver.ShapeFile())
            data_source = gdal_driver.Open(dp_shapefile, 1)
            layer = data_source.GetLayer(0)

            # set progress bar
            progress_bar_max = len(layer)
            self.progress_bar.setMaximum(progress_bar_max)
            progress_bar_counter = 0
            track_field_mismatch = []
            # for each drain point in shapefile
            for dp in layer:
                #if progress_bar_counter % 10 == 0:
                # update_pct = int(progress_bar_max/progress_bar_counter)
                # self.progress_bar.setValue(update_pct)
                # Fill StreamConnectID field for stream crossing or sump
                if drain_type_def_row.TableName == "StrXingAtt":
                    stream_connect_id = 2
                elif drain_type_def_row.TableName == "SumpAtt":
                    stream_connect_id = 1
                else:
                    # this seems to be the database default value for the StreamConnectID field
                    stream_connect_id = 2

                fld_match_table_model = self.field_match_table_wizard.table_model
                dp_target_column_names = [str(fld_match_table_model.index(row, 0).data()) for row
                                          in range(fld_match_table_model.rowCount())]
                # collect data to be inserted to both the DrainPoints table and shapefile specific attribute table

                dp_row_data = {'GRAIPDID': graipid, 'DrainTypeID': drain_type_def_row.DrainTypeID,
                               'StreamConnectID': stream_connect_id, 'Comments': ''}

                # dp_row_data = {'GRAIPDID': graipid}

                # if 'DrainTypeID' in dp_target_column_names:
                #     dp_row_data['DrainTypeID'] = drain_type_def_row.DrainTypeID
                # if 'StreamConnectID' in dp_target_column_names:
                #     dp_row_data['StreamConnectID'] = stream_connect_id
                # if 'Comments' in dp_target_column_names:
                #     dp_row_data['Comments'] = ""

                dp_att_row_data = {'GRAIPDID': graipid}

                # for each field name imported from the shape file (table grid source column (col#2))
                for row in range(fld_match_table_model.rowCount()):
                    # find the source column name (2nd column of the table)
                    dp_src_field_name = str(fld_match_table_model.index(row, 1).data())
                    dp_target_field_name = str(fld_match_table_model.index(row, 0).data())
                    # if a matching source field name exists
                    if dp_src_field_name != self.no_match_use_default:
                        # get the value for that field from the shapefile
                        dp_att_value = dp.GetField(dp_src_field_name)
                        # if a value/data exists in the shapefile
                        if dp_att_value:
                            # replace any single quotes or a double quote in the value
                            if isinstance(dp_att_value, basestring):
                                dp_att_value = dp_att_value.replace("'", "")
                                dp_att_value = dp_att_value.replace('"', "")

                            # if this is the PipeDimID(oval) field, then merge with PipeDimID field
                            if drain_type_def_row.TableName == "StrXingAtt" \
                                    and fld_match_table_model.index(row, 0).data() == "PipeDimID(oval)":
                                prev_row = row - 1
                                field_name = fld_match_table_model.index(prev_row, 0).data()
                                # TODO: change the FieldMatch table so that it has FillDepth in place of FillDepthID
                                # so that the following elif is not needed
                            elif drain_type_def_row.TableName == "StrXingAtt" \
                                    and fld_match_table_model.index(row, 0).data() == "FillDepthID":
                                field_name = "FillDepth"
                            else:
                                field_name = fld_match_table_model.index(row, 0).data()

                            # check if there is a DefinitionTable in MetaData table matching the field_name
                            metadata_row = cursor.execute("SELECT DefinitionTable FROM MetaData WHERE IDFieldName=?",
                                                          field_name).fetchone()

                            # if no matching definition table was found then write the original data value
                            if metadata_row is None:
                                is_type_match = False
                                if field_name in dp_table_field_names:
                                    # check datatype of the column with the value being assigned
                                    is_type_match = utils.is_data_type_match(graip_db_file,
                                                                             "DrainPoints", field_name,
                                                                             dp_att_value)
                                    if is_type_match:
                                        dp_row_data[field_name] = dp_att_value

                                else:
                                    is_type_match = utils.is_data_type_match(graip_db_file,
                                                                             drain_type_def_row.TableName, field_name,
                                                                             dp_att_value)
                                    if is_type_match:
                                        dp_att_row_data[field_name] = dp_att_value

                                if not is_type_match:
                                    # show message describing the mismatch
                                    dp_shapefile_basename = os.path.basename(dp_shapefile)
                                    if field_name not in track_field_mismatch:
                                        track_field_mismatch.append(field_name)
                                        action_taken_msg = "A default value will be used"
                                        log_message = "Type mismatch between field '{target_fld_name}' in database and " \
                                                      "value '{value}' in field '{source_fld_name}' in shapefile." \
                                                      "'{shp_file}'.".format(target_fld_name=field_name, value=dp_att_value,
                                                                             source_fld_name=dp_src_field_name,
                                                                             shp_file=dp_shapefile_basename)
                                        msg = log_message + "{}.".format(action_taken_msg)

                                        if not self.wizard.is_uninterrupted:
                                            msg_box = QMessageBox()
                                            msg_box.setIcon(QMessageBox.Information)
                                            msg_box.setText(msg)
                                            msg_box.setWindowTitle("Mismatch")
                                            msg_box.exec_()

                                        # TODO: need to write to the log file and other bits of processing
                                        # Ref to getIDFromDefinitionTable function in modGeneralFunction
                                        utils.add_entry_to_log_file(self.wizard.dp_log_file, graipid, drain_type_name,
                                                                    log_message, action_taken_msg)

                            else:
                                # found a matching Definition Table
                                definition_table_name = metadata_row.DefinitionTable
                                # definition_table_name = definition_table_name.replace("'", "''")
                                sql_select = "SELECT * FROM {}".format(definition_table_name)
                                definition_table_rows = cursor.execute(sql_select).fetchall()
                                value_matching_id = None
                                for table_row in definition_table_rows:
                                    # second column has the definition values
                                    if table_row[1] == dp_att_value:
                                        # first column has the definition ID
                                        value_matching_id = table_row[0]
                                        break

                                if value_matching_id is None:
                                    # check to see if already done Define Value dialog for this value, If so, use that
                                    # value - no need to show Define Value dialog
                                    sql_select = "SELECT * FROM ValueReassigns WHERE FromField=? AND DefinitionTable=?"
                                    params = (dp_att_value, definition_table_name)
                                    value_assigns_row = cursor.execute(sql_select, params).fetchone()
                                    if value_assigns_row:
                                        value_matching_id = value_assigns_row.DefinitionID
                                    else:
                                        # need to show Define Value dialog
                                        # for details on how to setup the DefineValueDialog
                                        # refer to vb module modGeneralFunctions and function getIDFromDefinitionTable
                                        if field_name in ('FlowPathVeg1ID', 'FlowPathVeg2ID', 'SurfaceTypeID'):
                                            define_value_dlg = utils.DefineValueDialog(missing_field_value=dp_att_value,
                                                                                       missing_field_name=field_name,
                                                                                       graip_db_file=graip_db_file,
                                                                                       def_table_name=definition_table_name,
                                                                                       is_multiplier=True)
                                        else:
                                            define_value_dlg = utils.DefineValueDialog(missing_field_value=dp_att_value,
                                                                                       missing_field_name=field_name,
                                                                                       graip_db_file=graip_db_file,
                                                                                       def_table_name=definition_table_name)
                                        define_value_dlg.show()
                                        define_value_dlg.exec_()
                                        if define_value_dlg.is_cancel:
                                            raise Exception("Aborting processing of this shapefile")

                                        value_matching_id = define_value_dlg.definition_id
                                        action_taken_msg = define_value_dlg.action_taken_msg
                                        # write to the log table
                                        log_message = "Value '{}' in field '{}' is not in the '{}' definitions table."
                                        log_message = log_message.format(dp_att_value, field_name, definition_table_name)
                                        utils.add_entry_to_error_table(graip_db_file, utils.DP_ERROR_LOG_TABLE_NAME,
                                                                       graipid, drain_type_name, log_message,
                                                                       action_taken_msg)
                                        # write to the log text file
                                        utils.add_entry_to_log_file(self.wizard.dp_log_file, graipid, drain_type_name,
                                                                    log_message, action_taken_msg)

                                    if field_name in dp_table_field_names:
                                        dp_row_data[field_name] = value_matching_id
                                    else:
                                        dp_att_row_data[field_name] = value_matching_id
                                else:
                                    if field_name in dp_table_field_names:
                                        dp_row_data[field_name] = value_matching_id
                                    else:
                                        dp_att_row_data[field_name] = value_matching_id
                    else:
                        if dp_target_field_name in ('CDate', 'CTime', 'VehicleID'):
                            msg = dp_target_field_name + " can't take default values"
                            raise Exception(msg)

                        # TODO: probably we have to display the define value dialog here too

                # insert/update data to DrainPoints table
                if not update_main_dp_table:
                    col_names = ",".join(k for k in dp_row_data.keys())
                    col_values = tuple(dp_row_data.values())
                    insert_sql = "INSERT INTO DrainPoints({col_names}) VALUES {col_values}"
                    insert_sql = insert_sql.format(col_names=col_names, col_values=col_values)
                    cursor.execute(insert_sql)
                    # cursor.execute("INSERT INTO DrainPoints(GRAIPDID, DrainTypeID, CDate, CTime, VehicleID, "
                    #                "StreamConnectID, DischargeToID) "
                    #                "VALUES (?, ?, ?, ?, ?, ?, ?)", dp_row_data['GRAIPDID'],
                    #                dp_row_data['DrainTypeID'], dp_row_data['CDate'], dp_row_data['CTime'],
                    #                dp_row_data['VehicleID'], dp_row_data['StreamConnectID'],
                    #                dp_row_data['DischargeToID'])
                else:
                    if "GRAIPDID" in dp_row_data:
                        del dp_row_data['GRAIPDID']

                    col_names = ",".join(k + "=?" for k in dp_row_data.keys())
                    params = tuple(dp_row_data.values()) + tuple(graipid,)
                    update_sql = "UPDATE DrainPoints SET {}  WHERE GRAIPDID=?".format(col_names)
                    cursor.execute(update_sql, params)

                    # update_sql = "UPDATE DrainPoints SET DrainTypeID=?, CDate=?, CTime=?, VehicleID=?, " \
                    #              "StreamConnectID=?, DischargeToID=?  WHERE GRAIPDID=?"
                    # data = (dp_row_data['DrainTypeID'], dp_row_data['CDate'], dp_row_data['CTime'],
                    #         dp_row_data['VehicleID'], dp_row_data['StreamConnectID'],
                    #         dp_row_data['DischargeToID'], graipid)
                    #cursor.execute(update_sql, data)

                # insert data to matching attribute table
                col_names = ",".join(k for k in dp_att_row_data.keys())
                col_values = ",".join(str(v) for v in dp_att_row_data.values())
                insert_sql = "INSERT INTO {table_name}({col_names}) VALUES ({col_values})"
                insert_sql = insert_sql.format(table_name=drain_type_def_row.TableName, col_names=col_names,
                                               col_values=col_values)
                cursor.execute(insert_sql)
                conn.commit()
                # set data for the DrainID field in DrainPoints table
                dp_row = cursor.execute("SELECT * FROM DrainPoints WHERE GRAIPDID=?", graipid).fetchone()
                if dp_row.CTime != 999:
                    # remove am/pm
                    time = dp_row.CTime[:-2]
                    time_hour, time_min, time_sec = time.split(":")
                    time_hour = int(time_hour)
                    time_min = int(time_min)
                    time_sec = int(time_sec)
                    if time_sec > 30:
                        time_min += 1
                    if "p" in dp_row.CTime and time_hour != 12:
                        time_hour += 12

                    if "a" in dp_row.CTime and time_hour == 12:
                        time_hour = 0
                    time_24_format = (time_hour * 100) + time_min

                    dt_obj = dp_row.CDate
                    # get 2 digit year
                    year = int(str(dt_obj.year)[2:])
                    drain_id = (year * 1000000000) + (dt_obj.month * 10000000) + (dt_obj.day * 100000) + \
                               (time_24_format * 10) + dp_row.VehicleID
                    update_sql = "UPDATE DrainPoints SET DrainID=? WHERE GRAIPDID=?"
                    # DrainID is of data type double in graip database
                    drain_id = float(drain_id)
                    data = (drain_id, graipid)
                    cursor.execute(update_sql, data)
                    conn.commit()

                graipid += 1
                # update progress bar
                progress_bar_counter += 1
                self.progress_bar.setValue(progress_bar_counter)
            conn.close()
        except Exception as ex:
            # TODO: write the error to the log file
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setText(ex.message)
            msg_box.exec_()
            print(ex.message)
            is_error = True
        finally:
            if conn and is_error:
                conn.close()

            # return True will take to the next page. return False will keep on the same page
            return not is_error

app = Preprocessor()
app.run()