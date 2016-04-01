import sys
import os
import shutil

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
        self.options_dlg = utils.OptionsDialog()

        self.addPage(FileSetupPage(parent=self))
        self.addPage(DrainPointPage(shp_file_index=0, shp_file="", parent=self))

        self.setWindowTitle("GRAIP Preprocessor (Version 2.0)")
        # add a custom button
        self.btn_options = QPushButton('Options')
        self.btn_options.clicked.connect(self.show_options_dialog)
        self.setButton(self.CustomButton1, self.btn_options)
        self.setOptions(self.HaveCustomButton1)

        self.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.resize(1000, 600)

    def show_options_dialog(self):
        self.dp_log_file, self.rd_log_file = self.options_dlg.get_data_from_dialog(self.dp_log_file, self.rd_log_file)

    def run(self):
        # Show the form
        self.show()

        # Run the qt application
        qt_app.exec_()


class FileSetupPage(QWizardPage):
    def __init__(self, parent=None):
        super(FileSetupPage, self).__init__(parent)
        self.wizard = parent
        self.working_directory = None
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
        return True

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
                                                       filter="GRAIP (*.mdb)")
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

        self.line_edit_mdb_file.setText(graip_db_file)


class DrainPointPage(QWizardPage):
    def __init__(self, shp_file_index=0, shp_file="", shp_file_count=0, parent=None):
        super(DrainPointPage, self).__init__(parent=parent)
        self.wizard = parent
        self.file_index = shp_file_index
        self.working_directory = None
        self.lst_widget_dp_shp_files = None
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
        self.field_match_table = None
        self.v_dp_layout.addWidget(self.table_label)

        self.group_box_set_field_names.setLayout(self.v_dp_layout)
        self.form_layout.addRow(self.group_box_set_field_names)

        self.setLayout(self.form_layout)
        self.setTitle("Import Drain Point Shapefile: {} of {}".format(self.file_index + 1, shp_file_count))

    def initializePage(self, *args, **kwargs):
        graip_db_file = self.wizard.line_edit_mdb_file.text()
        conn = pyodbc.connect(utils.MS_ACCESS_CONNECTION % graip_db_file)
        cursor = conn.cursor()
        # populate the field match table based on the drain point shapefile being imported
        if self.field_match_table is None:
            # TODO: replace hard coded data with real data from the graip database
            dp_shapefile = self.line_edit_imported_file.text()
            shp_file_attribute_names = utils.get_shapefile_attribute_column_names(dp_shapefile)
            no_match_default = '<No Match Use Default>'
            shp_file_attribute_names.append(no_match_default)

            table_headers = ['Target Field', 'Matching Source Field']
            self.dp_type_combo_box = utils.populate_drain_type_combobox(graip_db_file, self.dp_type_combo_box)
            # TODO: set the current index for the drain type combo box
            self.dp_type_combo_box = utils.set_index_dp_type_combo_box(dp_shapefile, self.dp_type_combo_box)
            drain_type_name = self.dp_type_combo_box.currentText()
            drain_type_def_row = cursor.execute("SELECT DrainTypeID FROM DrainTypeDefinitions "
                                                "WHERE DrainTypeName = ?", drain_type_name).fetchone()

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
                            if len(shp_att_name) > 2 and len(source_field_row.DBFField):
                                # match first 3 characters
                                if shp_att_name[0:2].lower() == source_field_row.DBFField[0:2].lower():
                                    source_field_col_data.append(source_field_row.DBFField)
                                    found_match = True
                                    break
                    if not found_match:
                        source_field_col_data.append(no_match_default)

            target_source_combined = zip(target_field_col_data, source_field_col_data)
            table_data = [[item[0], item[1]] for item in target_source_combined]
            #table_data = [['CDate', 'CDATE'], ['CTime', 'CTIME'], ['VehicleID', 'VEHICLE']]
            cmb_data = shp_file_attribute_names
            #cmb_data = ['CDATE', 'CTIME', 'STREAM_CON', 'SLOPE_SHP']
            self.field_match_table = utils.TableWidget(table_data=table_data, table_header=table_headers,
                                                       cmb_data=cmb_data)

            self.v_dp_layout.addWidget(self.field_match_table)
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
        define_value_dlg = utils.DefineValueDialog(missing_field_value='<25 degrees',
                                                   missing_field_name='ChannelAngleID')
        define_value_dlg.show()
        define_value_dlg.exec_()
        return True

app = Preprocessor()
app.run()