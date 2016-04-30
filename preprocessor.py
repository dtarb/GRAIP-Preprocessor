import sys
import os
import shutil

import pyodbc
from osgeo import ogr

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
        self.working_directory = None
        self.dp_log_file = None
        self.rd_log_file = None
        self.is_uninterrupted = False
        self.options_dlg = utils.OptionsDialog()

        self.setWindowTitle("GRAIP Preprocessor (Version 1.0)")
        # add a custom button
        self.btn_options = QPushButton('Options')
        self.setButton(self.CustomButton1, self.btn_options)
        self.setOptions(self.HaveCustomButton1)

        # add pages
        self.addPage(FileSetupPage(parent=self))
        self.addPage(DrainPointPage(shp_file_index=0, shp_file="", parent=self))

        self.currentIdChanged.connect(self.show_options_button)
        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        graip_icon = QIcon(utils.GRAIP_ICON_FILE)
        self.setWindowIcon(graip_icon)
        self.resize(800, 600)

    def show_options_button(self):
        if self.currentId() == 0:
            # show the Options button if it is the FileSetup page
            self.btn_options.show()

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
        self.rd_shp_file_processing_track_dict = {}
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

        layout_input_files = QVBoxLayout()
        self.group_box_input_files = QGroupBox("Input Files")

        # Not using the selection of dem file for now
        self.dem_file_label = QLabel()
        self.dem_file_label.setText("DEM File('sta.adf' inside the DEM folder)")
        self.dem_file_label.setWordWrap(True)

        v_layout_dem_file = QVBoxLayout()
        #v_layout_dem_file.addWidget(self.dem_file_label)
        h_layout_dem_files = QHBoxLayout()
        self.line_edit_dem_file = QLineEdit()
        self.btn_browse_dem_file = QPushButton('....')
        # connect the browse button to the function
        self.btn_browse_dem_file.clicked.connect(self.browse_dem_file)

        #h_layout_dem_files.addWidget(self.line_edit_dem_file)
        #h_layout_dem_files.addWidget(self.btn_browse_dem_file)
        v_layout_dem_file.addLayout(h_layout_dem_files)
        layout_input_files.addLayout(v_layout_dem_file)

        v_layout_rd_shp_files = QVBoxLayout()
        self.rd_shp_file_label = QLabel()
        self.rd_shp_file_label.setText("Road Shapefiles")
        self.rd_shp_file_label.setWordWrap(True)
        v_layout_rd_shp_files.addWidget(self.rd_shp_file_label)

        self.lst_widget_rd_shp_files = QListWidget()
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
        h_layout_rd_shp_files.addWidget(self.lst_widget_rd_shp_files)
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

        # make some of the page fields available to other pages
        # make the list widget for drain point shapefiles available to other pages of this wizard
        self.current_imported_dp_file = QLineEdit()
        self.registerField("current_imported_dp_file", self.current_imported_dp_file, 'text')

        self.setLayout(self.form_layout)
        self.setTitle("File Setup")

    def show_options_dialog(self):
        if len(self.line_edit_mdb_file.text().strip()) == 0:
            msg_box = utils.GraipMessageBox()
            msg_box.setWindowTitle("GRAIP Database File")
            msg_box.setText("Before you can set options, select GRAIP database file.")
            msg_box.exec_()
        else:
            self.dp_log_file, self.rd_log_file, self.is_uninterrupted = self.wizard.options_dlg.get_data_from_dialog(
                self.dp_log_file, self.rd_log_file, self.is_uninterrupted)

    def isComplete(self, *args, **kwargs):
        # NOTE: in this function DO NOT display any error message box

        self.lst_widget_dp_shp_files.setStyleSheet("background-color:white;")

        if self.wizard.currentId() == 0:
            if len(self.line_edit_mdb_file.text().strip()) == 0:
                return False

            # elif len(self.line_edit_dem_file.text().strip()) == 0:
            #     return False
            elif self.lst_widget_rd_shp_files.count() == 0:
                return False
            elif self.lst_widget_dp_shp_files.count() == 0:
                return False
            else:
                # check that not same file have been selected for both as DrainPoints as well as RoadLines
                dp_files = utils.get_items_from_list_box(self.lst_widget_dp_shp_files)
                rd_files = utils.get_items_from_list_box(self.lst_widget_rd_shp_files)
                common_shp_files = [shp_file for shp_file in rd_files if shp_file in dp_files]
                if common_shp_files:
                    self.lst_widget_dp_shp_files.setStyleSheet("background-color:red;")
                    return False

        return True

    def validatePage(self, *args, **kwargs):
        # This function gets called when the next button is clicked for the FileSetup page

        # remove if there are any drainpoint wizard pages
        for page_id in self.wizard.pageIds():
            if page_id > 0:
                self.wizard.removePage(page_id)

        msg_box = utils.GraipMessageBox()
        msg_box.setWindowTitle("Input errors")
        msg_box.setIcon(QMessageBox.Critical)

        if len(self.line_edit_mdb_file.text().strip()) == 0:
            msg_box.setText("GRAIP database file is required")
            msg_box.exec_()
            return False

        # if len(self.line_edit_dem_file.text().strip()) == 0:
        #     msg_box.setText("DEM file is required")
        #     msg_box.exec_()
        #     return False

        if self.lst_widget_rd_shp_files.count() == 0:
            msg_box.setText("At least one road shapefile is required")
            msg_box.exec_()
            return False

        if self.lst_widget_dp_shp_files.count() == 0:
            msg_box.setText("At least one drain points shapefile is required")
            msg_box.exec_()
            return False

        # check that not same file have been selected for both as DrainPoints as well as RoadLines
        # dp_files = utils.get_items_from_list_box(self.lst_widget_dp_shp_files)
        # rd_files = utils.get_items_from_list_box(self.lst_widget_rd_shp_files)
        # common_shp_files = [shp_file for shp_file in rd_files if shp_file in dp_files]
        # if common_shp_files:
        #     msg_box.setText("One or more file selected for RodLines also have been selected for DrainPoints")
        #     msg_box.exec_()
        #     self.wizard.setStartId(0)
        #     self.wizard.restart()
        #     return False

        dp_shp_file_list = []
        shp_file_count = self.lst_widget_dp_shp_files.count()
        for i in range(self.lst_widget_dp_shp_files.count()):
            shp_file = self.lst_widget_dp_shp_files.item(i).text()
            dp_shp_file_list.append(shp_file)
            self.wizard.addPage(DrainPointPage(shp_type='DP', shp_file_index=i, shp_file=shp_file,
                                               shp_file_count=shp_file_count, parent=self))

        rd_shp_file_list = []
        shp_file_count = self.lst_widget_rd_shp_files.count()
        for i in range(self.lst_widget_rd_shp_files.count()):
            shp_file = self.lst_widget_rd_shp_files.item(i).text()
            rd_shp_file_list.append(shp_file)
            self.wizard.addPage(RoadLinePage(shp_type='RD', shp_file_index=i, shp_file=shp_file,
                                             shp_file_count=shp_file_count, parent=self))

        if self.lst_widget_dp_shp_files.count() > 0:
            self.current_imported_dp_file.setText(self.lst_widget_dp_shp_files.item(0).text())
            self.wizard.setStartId(0)
            self.wizard.restart()
            total_wizard_pages = len(self.wizard.pageIds())
            # 1 is for the FileSetup page
            if total_wizard_pages > (self.lst_widget_dp_shp_files.count() + self.lst_widget_rd_shp_files.count() + 1):
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
        dem_file_path = "path will be set in future version"
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

    def browse_dem_file(self):
        working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        graip_dem_file, _ = QFileDialog.getOpenFileName(None, "Select file 'sta.adf'", working_dir,
                                                        filter="DEM (*.adf)")
        # check if the cancel was clicked on the file dialog
        if len(graip_dem_file) == 0:
            return

        self.line_edit_dem_file.setText(QFileInfo(graip_dem_file).path())
        self.completeChanged.emit()

    def browse_rd_shp_files(self):
        working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        rd_shp_files, _ = QFileDialog.getOpenFileNames(None, "Select Road Shapefiles", working_dir,
                                                       filter="Shapefiles (*.shp)")

        # TODO: If the file already in the list widget do not add that
        existing_rd_files = utils.get_items_from_list_box(self.lst_widget_rd_shp_files)
        if len(existing_rd_files) > 0:
            for shp_file in rd_shp_files:
                if shp_file not in existing_rd_files:
                    self.lst_widget_rd_shp_files.addItem(shp_file)
        else:
            self.lst_widget_rd_shp_files.addItems(rd_shp_files)
        self.completeChanged.emit()

    def remove_rd_shp_files(self):
        for shp_file_item_index in self.lst_widget_rd_shp_files.selectedIndexes():
            item = self.lst_widget_rd_shp_files.takeItem(shp_file_item_index.row())
            item = None
        self.completeChanged.emit()

    def browse_dp_shp_files(self):
        working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        dp_shp_files, _ = QFileDialog.getOpenFileNames(None, "Select Drain Points Shapefiles", working_dir,
                                                       filter="Shapefiles (*.shp)")

        # If the file already in the list widget do not add that
        existing_dp_files = utils.get_items_from_list_box(self.lst_widget_dp_shp_files)
        if len(existing_dp_files) > 0:
            for shp_file in dp_shp_files:
                if shp_file not in existing_dp_files:
                    self.lst_widget_dp_shp_files.addItem(shp_file)
        else:
            self.lst_widget_dp_shp_files.addItems(dp_shp_files)
        self.completeChanged.emit()

    def remove_dp_shp_files(self):
        for shp_file_item_index in self.lst_widget_dp_shp_files.selectedIndexes():
            item = self.lst_widget_dp_shp_files.takeItem(shp_file_item_index.row())
            item = None
        self.completeChanged.emit()

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
        self.wizard.working_directory = self.working_directory

        # if the database file does not exists then copy the empty database
        # as the filename provided by the user
        if not os.path.isfile(graip_db_file):
            this_script_dir = os.path.dirname(os.path.realpath(__file__))
            empty_graip_db_file = os.path.join(this_script_dir, 'GRAIP_DB', "GRAIP.mdb")
            shutil.copyfile(empty_graip_db_file, graip_db_file)
        else:
            # prompt the user if the existing db file to be overwritten
            msg_box = utils.GraipMessageBox()
            msg_box.setWindowTitle("File action")
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
                    self.lst_widget_rd_shp_files.addItems(road_shp_files)
                    dp_shp_files = file_setup_row.DrainPoints_Shapefiles.split(',')
                    self.lst_widget_dp_shp_files.addItems(dp_shp_files)

        db_file_name = os.path.basename(graip_db_file)
        db_file_name_wo_ext = db_file_name.split('.')[0]
        self.dp_log_file = os.path.join(self.working_directory, db_file_name_wo_ext + 'DP.log')
        self.rd_log_file = os.path.join(self.working_directory, db_file_name_wo_ext + 'RD.log')
        self.line_edit_mdb_file.setText(graip_db_file)
        self.completeChanged.emit()


class DrainPointPage(utils.ImportWizardPage):

    def initializePage(self, *args, **kwargs):
        self.progress_bar.setValue(0)
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
                        matching_field = None
                        for shp_att_name in shp_file_attribute_names:
                            if len(shp_att_name) > 2 and len(source_field_row.DBFField) > 2:
                                # match first 3 characters
                                match_count = 0
                                for i in range(len(shp_att_name)):
                                    if shp_att_name[0:i+1].lower() == source_field_row.DBFField[0:i+1].lower():
                                        match_count += 1
                                if match_count > 2:
                                    matching_field = shp_att_name
                                    break

                        if matching_field is not None:
                            source_field_col_data.append(matching_field)
                        else:
                            source_field_col_data.append(self.no_match_use_default)

            target_source_combined = zip(target_field_col_data, source_field_col_data)
            table_data = [[item[0], item[1]] for item in target_source_combined]
            cmb_data = shp_file_attribute_names
            self.field_match_table_wizard = utils.TableWidget(table_data=table_data, table_header=table_headers,
                                                              cmb_data=cmb_data)

            self.v_set_fields_layout.addWidget(self.field_match_table_wizard)
            conn.close()

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
                                            msg_box = utils.GraipMessageBox()
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
                                        if not self.wizard.is_uninterrupted:
                                            define_value_dlg.show()
                                            define_value_dlg.exec_()
                                            if define_value_dlg.is_cancel:
                                                raise Exception("Aborting processing of this shapefile")
                                        else:
                                            define_value_dlg.accept()

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
                    params = tuple(dp_row_data.values()) + (graipid,)
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
                    # TODO: Instead of the following code use the function in utils.py
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
                qApp.processEvents()
            conn.close()
        except (pyodbc.DatabaseError, pyodbc.DataError) as ex:
            raise Exception(ex.message)
        except Exception as ex:
            # TODO: write the error to the log file
            msg_box = utils.GraipMessageBox()
            msg_box.setWindowTitle("Error")
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

    def cleanupPage(self, *args, **kwargs):
        """
        This function is called when the wizard back button is clicked
        Here we are resting the progress bar for the page to which the back button will take
        """
        if self.wizard.wizard.currentId() > 0:
            prev_page_id = self.wizard.wizard.currentId() - 1
            prev_page = self.wizard.wizard.page(prev_page_id)

            if not isinstance(prev_page, FileSetupPage):
                prev_page.progress_bar.setValue(0)

class RoadLinePage(utils.ImportWizardPage):

    def initializePage(self, *args, **kwargs):
        graip_db_file = self.wizard.line_edit_mdb_file.text()
        conn = pyodbc.connect(utils.MS_ACCESS_CONNECTION % graip_db_file)
        cursor = conn.cursor()
        # populate the field match table based on the road line shapefile being imported
        if self.field_match_table_wizard is None:
            rd_shapefile = self.line_edit_imported_file.text()
            shp_file_attribute_names = utils.get_shapefile_attribute_column_names(rd_shapefile)
            self.no_match_use_default = '<No Match Use Default>'
            shp_file_attribute_names = [self.no_match_use_default] + shp_file_attribute_names
            table_headers = ['Target Field', 'Matching Source Field']
            # TODO: populate the Road Network combobox by loading data from the RoadNetworkDefinitions table
            rd_network_def_rows = cursor.execute("SELECT * FROM RoadNetworkDefinitions").fetchall()
            rd_network_values = [row.RoadNetwork for row in rd_network_def_rows]
            self.rd_network_combo_box.addItems(rd_network_values)
            self.update_rd_network_gui_elements()

            # self.dp_type_combo_box = utils.populate_drain_type_combobox(graip_db_file, self.dp_type_combo_box)
            # # set the current index for the drain type combo box
            # self.dp_type_combo_box = utils.set_index_dp_type_combo_box(rd_shapefile, self.dp_type_combo_box)
            # drain_type_name = self.dp_type_combo_box.currentText()
            # drain_type_def_row = cursor.execute("SELECT DrainTypeID FROM DrainTypeDefinitions "
            #                                     "WHERE DrainTypeName = ?", drain_type_name).fetchone()

            # find the target field names in the database corresponding to the shapefile being imported
            field_name_rows = cursor.execute("SELECT DBField FROM FieldMatches "
                                             "WHERE AttTableID =0").fetchall()
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
                        match_field = None
                        for shp_att_name in shp_file_attribute_names:
                            if len(shp_att_name) > 2 and len(source_field_row.DBFField) > 2:
                                # match at least any first 3 characters
                                match_count = 0
                                for i in range(len(shp_att_name)):
                                    if i < len(source_field_row.DBFField):
                                        if shp_att_name[0:i+1].lower() == source_field_row.DBFField[0:i+1].lower():
                                            match_count += 1

                                if match_count > 2:
                                    match_field = shp_att_name
                                    break

                        if match_field is not None:
                            source_field_col_data.append(match_field)
                        else:
                            source_field_col_data.append(self.no_match_use_default)
            target_source_combined = zip(target_field_col_data, source_field_col_data)
            table_data = [[item[0], item[1]] for item in target_source_combined]
            cmb_data = shp_file_attribute_names
            self.field_match_table_wizard = utils.TableWidget(table_data=table_data, table_header=table_headers,
                                                              cmb_data=cmb_data)

            self.v_set_fields_layout.addWidget(self.field_match_table_wizard)
            conn.close()

    def validatePage(self, *args, **kwargs):
        # this function is executed when next button is selected
        # here we should be processing the currently imported road line shape file
        # during the processing if we find a value that's not in the definitions table we need
        # to show the DefineValueDialog

        is_error = False
        try:
            graip_db_file = self.wizard.line_edit_mdb_file.text()
            conn = pyodbc.connect(utils.MS_ACCESS_CONNECTION % graip_db_file)
            cursor = conn.cursor()
            # get column names of the RoadLines table
            rd_table_field_names = [row.column_name for row in cursor.columns(table="RoadLines")]

            # open roadlines shapefile
            update_main_rd_table = False
            rd_shapefile = self.line_edit_imported_file.text()
            if rd_shapefile in self.wizard.rd_shp_file_processing_track_dict:
                graipid = self.wizard.rd_shp_file_processing_track_dict[rd_shapefile]
                # delete all records from RoadLines table for this graipid
                cursor.execute("DELETE * FROM RoadLines WHERE GRAIPRID=?", graipid)
                conn.commit()
                update_main_rd_table = True
            else:
                # find out the next GRAIPRID from the RoadLines table
                rd_row = cursor.execute("SELECT MAX(GRAIPRID)AS Max_GRAIPRID FROM RoadLines").fetchone()
                if rd_row[0] is not None:
                    graipid = rd_row.Max_GRAIPRID + 1
                else:
                    graipid = 0
                self.wizard.rd_shp_file_processing_track_dict[rd_shapefile] = graipid

            gdal_driver = ogr.GetDriverByName(utils.GDALFileDriver.ShapeFile())
            data_source = gdal_driver.Open(rd_shapefile, 1)
            layer = data_source.GetLayer(0)

            # set progress bar
            progress_bar_max = len(layer)
            self.progress_bar.setMaximum(progress_bar_max)
            progress_bar_counter = 0
            track_field_mismatch = []
            # for each drain point in shapefile
            for dp in layer:
                fld_match_table_model = self.field_match_table_wizard.table_model
                dp_target_column_names = [str(fld_match_table_model.index(row, 0).data()) for row
                                          in range(fld_match_table_model.rowCount())]

                # collect data to be inserted to RoadLines table
                rd_row_data = {'GRAIPRID': graipid, 'Comments': ''}

                # for each field name imported from the shape file (table grid source column (col#2))
                for row in range(fld_match_table_model.rowCount()):
                    # find the source column name (2nd column of the table)
                    rd_src_field_name = str(fld_match_table_model.index(row, 1).data())
                    rd_target_field_name = str(fld_match_table_model.index(row, 0).data())
                    # if a matching source field name exists
                    if rd_src_field_name != self.no_match_use_default:
                        # get the value for that field from the shapefile
                        rd_att_value = dp.GetField(rd_src_field_name)
                        # if a value/data exists in the shapefile
                        if rd_att_value:
                            field_name = fld_match_table_model.index(row, 0).data()

                            # check if there is a DefinitionTable in MetaData table matching the field_name
                            metadata_row = cursor.execute("SELECT DefinitionTable FROM MetaData WHERE IDFieldName=?",
                                                          field_name).fetchone()

                            # if no matching definition table was found then write the original data value
                            if metadata_row is None:
                                is_type_match = False
                                if field_name in rd_table_field_names:
                                    # check datatype of the column with the value being assigned
                                    is_type_match = utils.is_data_type_match(graip_db_file,
                                                                             "RoadLines", field_name,
                                                                             rd_att_value)
                                    if is_type_match:
                                        rd_row_data[field_name] = rd_att_value

                                # else:
                                #     is_type_match = utils.is_data_type_match(graip_db_file,
                                #                                              drain_type_def_row.TableName, field_name,
                                #                                              dp_att_value)
                                #     if is_type_match:
                                #         dp_att_row_data[field_name] = dp_att_value

                                if not is_type_match:
                                    # show message describing the mismatch
                                    rd_shapefile_basename = os.path.basename(rd_shapefile)
                                    if field_name not in track_field_mismatch:
                                        track_field_mismatch.append(field_name)
                                        action_taken_msg = "A default value will be used"
                                        log_message = "Type mismatch between field '{target_fld_name}' in database and " \
                                                      "value '{value}' in field '{source_fld_name}' in shapefile." \
                                                      "'{shp_file}'.".format(target_fld_name=field_name, value=rd_att_value,
                                                                             source_fld_name=rd_src_field_name,
                                                                             shp_file=rd_shapefile_basename)
                                        msg = log_message + "{}.".format(action_taken_msg)

                                        if not self.wizard.is_uninterrupted:
                                            msg_box = utils.GraipMessageBox()
                                            msg_box.setIcon(QMessageBox.Information)
                                            msg_box.setText(msg)
                                            msg_box.setWindowTitle("Mismatch")
                                            msg_box.exec_()

                                        # TODO: need to write to the log file and other bits of processing
                                        # Ref to getIDFromDefinitionTable function in modGeneralFunction
                                        rd_network_type = self.rd_network_combo_box.currentText()
                                        utils.add_entry_to_log_file(self.wizard.dp_log_file, graipid, rd_network_type,
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
                                    if table_row[1] == rd_att_value:
                                        # first column has the definition ID
                                        value_matching_id = table_row[0]
                                        break

                                if value_matching_id is None:
                                    # check to see if already done Define Value dialog for this value, If so, use that
                                    # value - no need to show Define Value dialog
                                    sql_select = "SELECT * FROM ValueReassigns WHERE FromField=? AND DefinitionTable=?"
                                    params = (rd_att_value, definition_table_name)
                                    value_assigns_row = cursor.execute(sql_select, params).fetchone()
                                    if value_assigns_row:
                                        value_matching_id = value_assigns_row.DefinitionID
                                    else:
                                        # need to show Define Value dialog
                                        # for details on how to setup the DefineValueDialog
                                        # refer to vb module modGeneralFunctions and function getIDFromDefinitionTable
                                        if field_name in ('FlowPathVeg1ID', 'FlowPathVeg2ID', 'SurfaceTypeID'):
                                            define_value_dlg = utils.DefineValueDialog(missing_field_value=rd_att_value,
                                                                                       missing_field_name=field_name,
                                                                                       graip_db_file=graip_db_file,
                                                                                       def_table_name=definition_table_name,
                                                                                       is_multiplier=True)
                                        else:
                                            define_value_dlg = utils.DefineValueDialog(missing_field_value=rd_att_value,
                                                                                       missing_field_name=field_name,
                                                                                       graip_db_file=graip_db_file,
                                                                                       def_table_name=definition_table_name)
                                        if not self.wizard.is_uninterrupted:
                                            define_value_dlg.show()
                                            define_value_dlg.exec_()
                                            if define_value_dlg.is_cancel:
                                                raise Exception("Aborting processing of this shapefile")
                                        else:
                                            define_value_dlg.accept()

                                        value_matching_id = define_value_dlg.definition_id
                                        action_taken_msg = define_value_dlg.action_taken_msg
                                        # write to the log table
                                        log_message = "Value '{}' in field '{}' is not in the '{}' definitions table."
                                        log_message = log_message.format(rd_att_value, field_name, definition_table_name)
                                        rd_network_type = self.rd_network_combo_box.currentText()
                                        utils.add_entry_to_error_table(graip_db_file, utils.DP_ERROR_LOG_TABLE_NAME,
                                                                       graipid, rd_network_type, log_message,
                                                                       action_taken_msg)
                                        # write to the log text file
                                        utils.add_entry_to_log_file(self.wizard.dp_log_file, graipid, rd_network_type,
                                                                    log_message, action_taken_msg)

                                    if field_name in rd_table_field_names:
                                        rd_row_data[field_name] = value_matching_id
                                    else:
                                        raise Exception("'{}' field name is not in RoadLines table")
                                else:
                                    if field_name in rd_table_field_names:
                                        rd_row_data[field_name] = value_matching_id
                                    else:
                                        raise Exception("'{}' field name is not in RoadLines table")
                    else:
                        if rd_target_field_name in ('CDate', 'CTime1', 'CTime2' 'VehicleID'):
                            msg = rd_target_field_name + " can't take default values"
                            raise Exception(msg)

                        # TODO: probably we have to display the define value dialog here too

                # insert/update data to RoadLines table
                if not update_main_rd_table:
                    col_names = ",".join(k for k in rd_row_data.keys())
                    col_values = tuple(rd_row_data.values())
                    insert_sql = "INSERT INTO RoadLines({col_names}) VALUES {col_values}"
                    insert_sql = insert_sql.format(col_names=col_names, col_values=col_values)
                    cursor.execute(insert_sql)
                else:
                    if "GRAIPRID" in rd_row_data:
                        del rd_row_data['GRAIPRID']

                    col_names = ",".join(k + "=?" for k in rd_row_data.keys())
                    params = tuple(rd_row_data.values()) + (graipid,)
                    update_sql = "UPDATE RoadLines SET {}  WHERE GRAIPRID=?".format(col_names)
                    cursor.execute(update_sql, params)

                conn.commit()

                # update the RoadNetworkID field value in RoadLines table
                rd_network_type = self.rd_network_combo_box.currentText()
                rd_network_type_id = self.rd_network_combo_box.findText(rd_network_type)
                if rd_network_type_id != -1:
                    rd_network_type_id += 1
                    update_sql = "UPDATE RoadLines SET RoadNetworkID=? WHERE GRAIPRID=?"
                    data = (rd_network_type_id, graipid)
                    cursor.execute(update_sql, data)
                    conn.commit()

                # set data for the DrainID fields in RoadLines table
                rd_row = cursor.execute("SELECT * FROM RoadLines WHERE GRAIPRID=?", graipid).fetchone()
                if rd_row.CTime1 != 999:
                    drain_id = utils.get_drain_id(rd_row.CTime1, rd_row.CDate, rd_row.VehicleID)
                    update_sql = "UPDATE RoadLines SET OrigDrainID1=? WHERE GRAIPRID=?"
                    # DrainID is of data type double in graip database
                    drain_id = float(drain_id)
                    data = (drain_id, graipid)
                    cursor.execute(update_sql, data)
                    conn.commit()

                if rd_row.CTime2 != 999:
                    drain_id = utils.get_drain_id(rd_row.CTime2, rd_row.CDate, rd_row.VehicleID)
                    update_sql = "UPDATE RoadLines SET OrigDrainID2=? WHERE GRAIPRID=?"
                    # DrainID is of data type double in graip database
                    drain_id = float(drain_id)
                    data = (drain_id, graipid)
                    cursor.execute(update_sql, data)
                    conn.commit()

                # TODO: populate the GRAIPDID1, GRAIPDID2, StreamConnect1ID, StreamConnect2ID
                # Ref: frmPPWizard3 (importValuesToDatabase) here is the copied code from old graip
                has_dp1 = False
                has_dp2 = False
                # check to see if the 1st drainpoint record exists
                rd_row = cursor.execute("SELECT * FROM RoadLines WHERE GRAIPRID=?", graipid).fetchone()
                dp_row = cursor.execute("SELECT * FROM DrainPoints WHERE DrainID=?", rd_row.OrigDrainID1).fetchone()
                if dp_row is not None:
                    update_sql = "UPDATE RoadLines SET GRAIPDID1=? WHERE GRAIPRID=?"
                    data = (dp_row.GRAIPDID, graipid)
                    cursor.execute(update_sql, data)
                    update_sql = "UPDATE RoadLines SET StreamConnect1ID=? WHERE GRAIPRID=?"
                    data = (dp_row.StreamConnectID, graipid)
                    cursor.execute(update_sql, data)
                    conn.commit()
                    has_dp1 = True

                # check to see if the 2nd drainpoint record exists
                dp_row = cursor.execute("SELECT * FROM DrainPoints WHERE DrainID=?", rd_row.OrigDrainID2).fetchone()
                if dp_row is not None:
                    update_sql = "UPDATE RoadLines SET GRAIPDID2=? WHERE GRAIPRID=?"
                    data = (dp_row.GRAIPDID, graipid)
                    cursor.execute(update_sql, data)
                    update_sql = "UPDATE RoadLines SET StreamConnect2ID=? WHERE GRAIPRID=?"
                    data = (dp_row.StreamConnectID, graipid)
                    cursor.execute(update_sql, data)
                    conn.commit()
                    has_dp2 = True

                # if no drainpoints write to the log
                if not has_dp1 and not has_dp2:
                    utils.add_entry_to_log_file(self.wizard.dp_log_file, graipid, rd_network_type, "Doesn't drain",
                                                "Nothing")
                    # TODO: check why we not writing the database log table

                graipid += 1
                # update progress bar
                progress_bar_counter += 1
                self.progress_bar.setValue(progress_bar_counter)
                qApp.processEvents()
            conn.close()
            # show consolidate shapefiles dialog
            dp_shp_files = utils.get_items_from_list_box(self.wizard.lst_widget_dp_shp_files)
            rd_shp_files = utils.get_items_from_list_box(self.wizard.lst_widget_rd_shp_files)
            dlg = utils.ConsolidateShapeFiles(graip_db_file=graip_db_file, dp_shp_files=dp_shp_files,
                                              rd_shp_files=rd_shp_files,
                                              dp_log_file=self.wizard.dp_log_file,
                                              rd_log_file=self.wizard.rd_log_file,
                                              working_directory=self.wizard.working_directory,
                                              parent=self)
            self.wizard.hide()
            dlg.show()
            dlg.do_process()
            dlg.exec_()

        except (pyodbc.DatabaseError, pyodbc.DataError) as ex:
            raise Exception(ex.message)

        except Exception as ex:
            # TODO: write the error to the log file
            msg_box = utils.GraipMessageBox()
            msg_box.setWindowTitle("Error")
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

    def cleanupPage(self, *args, **kwargs):
        """
        This function is called when the wizard back button is clicked
        Here we are resting the progress bar for the page to which the back button will take
        """
        if self.wizard.wizard.currentId() > 0:
            prev_page_id = self.wizard.wizard.currentId() - 1
            prev_page = self.wizard.wizard.page(prev_page_id)
            prev_page.progress_bar.setValue(0)

app = Preprocessor()
app.run()
