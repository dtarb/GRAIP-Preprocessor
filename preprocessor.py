import sys
import os

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
        self.options_dlg = utils.OptionsDialog()

        self.addPage(FileSetupPage(parent=self))
        self.addPage(DrainPointPage(shp_file_index=0, shp_file="", parent=self))

        self.setWindowTitle("GRAIP Preprocessor (Version 2.0)")
        # add a custom button
        self.btn_options = QPushButton('Options')
        self.btn_options.clicked.connect(self.show_options_dialog)
        self.setButton(self.CustomButton1, self.btn_options)
        self.setOptions(self.HaveCustomButton1)
        self.dp_log_file = None
        self.rd_log_file = None

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

        self.group_box_prj_file = QGroupBox("Project File")
        self.line_edit_prj_file = QLineEdit()
        self.btn_browse_prg_file = QPushButton('....')
        # connect the browse button to the function
        self.btn_browse_prg_file.clicked.connect(self.browse_project_file)

        layout_prj = QHBoxLayout()
        layout_prj.addWidget(self.line_edit_prj_file)
        layout_prj.addWidget(self.btn_browse_prg_file)
        self.group_box_prj_file.setLayout(layout_prj)
        self.form_layout.addRow(self.group_box_prj_file)

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
        # To add an item to the above listWidget, ref: http://srinikom.github.io/pyside-docs/PySide/QtGui/QListWidget.html
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

        self.group_box_mdb_file = QGroupBox("GRAIP Database (*.mdb)")
        self.line_edit_mdb_file = QLineEdit()
        self.btn_browse_mdb_file = QPushButton('....')
        self.btn_browse_mdb_file.clicked.connect(self.browse_db_file)

        layout_mdb = QHBoxLayout()
        layout_mdb.addWidget(self.line_edit_mdb_file)
        layout_mdb.addWidget(self.btn_browse_mdb_file)
        self.group_box_mdb_file.setLayout(layout_mdb)
        self.form_layout.addRow(self.group_box_mdb_file)

        # make some of the page fields available to other pages
        # make the list widget for drain point shapefiles available to other pages of this wizard
        self.current_imported_dp_file = QLineEdit()

        self.registerField("current_imported_dp_file", self.current_imported_dp_file, 'text')

        self.setLayout(self.form_layout)
        self.setTitle("File Setup")

    def validatePage(self, *args, **kwargs):

        # remove if there are any drainpoint wizard pages
        for page_id in self.wizard.pageIds():
            if page_id > 0:
                self.wizard.removePage(page_id)

        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Critical)
        # check project file has been selected
        if len(self.line_edit_prj_file.text().strip()) == 0:
            msg_box.setText("Project file is required")
            msg_box.exec_()
            return False

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

        shp_file_count = self.lst_widget_dp_shp_files.count()
        for i in range(self.lst_widget_dp_shp_files.count()):
            shp_file = self.lst_widget_dp_shp_files.item(i).text()
            self.wizard.addPage(DrainPointPage(shp_file_index=i, shp_file=shp_file, shp_file_count=shp_file_count,
                                               parent=self))

        if self.lst_widget_dp_shp_files.count() > 0:
            self.current_imported_dp_file.setText(self.lst_widget_dp_shp_files.item(0).text())
            self.wizard.setStartId(0)
            self.wizard.restart()
            total_wizard_pages = len(self.wizard.pageIds())
            if total_wizard_pages > self.lst_widget_dp_shp_files.count() + 1:   # 1 is for the FileSetup page
                # here we are removing the DrainPoint page we added in the init of the wizard
                self.wizard.removePage(1)

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
        self.line_edit_mdb_file.setText(graip_db_file)


class DrainPointPage(QWizardPage):
    def __init__(self, shp_file_index=0, shp_file="", shp_file_count=0, parent=None):
        super(DrainPointPage, self).__init__(parent)
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
        # populate the field match table based on the drain point shapefile being imported
        if self.field_match_table is None:
            # TODO: replace hard coded data with real data from the graip database
            table_headers = ['Target Field', 'Matching Source Field']
            table_data = [['CDate', 'CDATE'], ['CTime', 'CTIME'], ['VehicleID', 'VEHICLE']]
            cmb_data = ['CDATE', 'CTIME', 'STREAM_CON', 'SLOPE_SHP']
            self.field_match_table = utils.TableWidget(table_data=table_data, table_header=table_headers,
                                                       cmb_data=cmb_data)

            self.v_dp_layout.addWidget(self.field_match_table)
            #TODO: drain point types need to be raed from the DrainTypeDefinitions table
            drain_point_types = ['Broad base type', 'Diffuse drain', 'Ditch relief', 'Lead off']
            self.dp_type_combo_box.addItems(drain_point_types)

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