import os
import time

import pyodbc
from osgeo import ogr, gdal, osr
from gdalconst import *

from PySide.QtGui import *
from PySide.QtCore import *

MS_ACCESS_CONNECTION = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;"


class GDALFileDriver(object):
    @classmethod
    def ShapeFile(cls):
        return "ESRI Shapefile"

    @classmethod
    def TifFile(cls):
        return "GTiff"


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

    def openClicked(self):
        inds = self.tree.selectionModel().selectedIndexes()
        files = []
        for i in inds:
            if i.column() == 0:
                #print("i.data():" + str(i.data()))
                files.append(os.path.join(str(self.directory().absolutePath()),str(i.data())))
        self.selectedFiles = files
        self.hide()

    def filesSelected(self):
        return self.selectedFiles


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
        self.resize(600, 300)
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

    def get_log_files(self):
        return self.line_edit_dp_log_file.text(), self.line_edit_rd_log_file.text()

    @staticmethod
    def get_data_from_dialog(dp_log_file, rd_log_file):
        dialog = OptionsDialog()
        dialog.line_edit_dp_log_file.setText(dp_log_file)
        dialog.line_edit_rd_log_file.setText(rd_log_file)
        dialog.exec_()
        return dialog.get_log_files()


class DefineValueDialog(QDialog):
    use_default = False
    reassign_value = False
    add_new = False
    is_cancel = False
    definition_id = 0

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
            if len(self.missing_field_value) > 2:
                for index in range(self.cmb_definitions.count()):
                    definition = self.cmb_definitions.itemText(index)
                    if len(definition) > 2:
                        # match first 3 characters
                        if self.missing_field_value[0:2].lower() == definition[0:2].lower():
                            initial_combobox_index = index
                            field_match_found = True
                            break
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
            elif self.radio_btn_use_default.isChecked():
                self.use_default = True
                self.definition_id = int(self.line_edit_id.text())
                sql_insert = "INSERT INTO ValueReassigns (FromField, ToField, DefinitionID, DefinitionTable) " \
                             "VALUES (?, ?, ?, ?)"
                data = (self.line_edit_def.text(), self.line_edit_default.text(), self.definition_id,
                        self.def_table_name)
            elif self.radio_btn_add_new.isChecked():
                self.add_new = True
                self.definition_id = int(self.line_edit_id.text())
                sql_insert = "INSERT INTO {} VALUES (?, ?, ?)".format(self.def_table_name)
                data = (self.definition_id, self.line_edit_def.text(), self.line_edit_description.text())

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

class FileDeleteMessageBox(QMessageBox):
    def __init__(self, file_to_delete, parent=None):
        super(FileDeleteMessageBox, self).__init__(parent)
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
        file_obj.write(time.strftime("%m-%d-%Y %H:%M:%S"))
        if log_type == 'DP':
            file_obj.write("GRAIPDID, Drain Type, Error Message, Action Taken")
        else:
            file_obj.write("GRAIPDID, Type, Error Message, Action Taken")


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


