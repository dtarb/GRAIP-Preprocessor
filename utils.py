import os

from PySide.QtGui import *
from PySide.QtCore import *


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

        #v_layout_log_files = QVBoxLayout()
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
        #self.line_edit_dp_log_file.setText("abc-123")
        if len(dp_log_file) == 0:
            return

        self.line_edit_dp_log_file.setText(dp_log_file)

    def browse_rd_log_file(self):
        #working_dir = self.working_directory if self.working_directory is not None else os.getcwd()
        rd_log_file, _ = QFileDialog.getSaveFileName(None, 'Enter Road Log Filename', os.getcwd(),
                                                     filter="Road Logfile (*.log)")
         # check if the cancel was clicked on the file dialog
        #self.line_edit_rd_log_file.setText("abc-123")
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
    def __init__(self, missing_field_value, missing_field_name, parent=None):
        super(DefineValueDialog, self).__init__(parent)
        self.missing_field_value= missing_field_value
        self.missing_field_name = missing_field_name
        self.form_layout = QFormLayout()
        v_main_layout = QVBoxLayout()
        self.radio_btn_use_default = QRadioButton("Use defualt value")
        # set this radio button as checked
        self.radio_btn_use_default.toggle()
        self.radio_btn_reassign = QRadioButton("Reassign this value to an existing value in definitions table")
        self.radio_btn_add_new = QRadioButton("Add new entry to definitions table")
        missing_value_msg = "Value '{}' in field '{}' is not in the definitions " \
                            "table".format(self.missing_field_value, self.missing_field_name)
        self.lbl_missing_value = QLabel(missing_value_msg)
        v_main_layout.addWidget(self.lbl_missing_value)
        h_layout_default = QHBoxLayout()
        self.line_edit_default = QLineEdit()
        h_layout_default.addWidget(self.radio_btn_use_default)
        h_layout_default.addWidget(self.line_edit_default)
        v_main_layout.addLayout(h_layout_default)
        v_main_layout.addWidget(self.radio_btn_reassign)
        h_layout_definitions = QHBoxLayout()
        self.label_definitions = QLabel("Definitions")
        self.cmb_definitions = QComboBox()
        h_layout_definitions.addWidget(self.label_definitions)
        h_layout_definitions.addWidget((self.cmb_definitions))
        v_main_layout.addLayout(h_layout_definitions)
        v_main_layout.addWidget(self.radio_btn_add_new)

        self.grp_box_new_entry = QGroupBox()
        grid_layout_new_entry = QGridLayout()
        v_layout_new_entry = QVBoxLayout()
        h_layout_table_name = QHBoxLayout()
        self.label_table_name = QLabel("Table Name")
        self.line_edit_table_name = QLineEdit()
        grid_layout_new_entry.addWidget(self.label_table_name, 0, 0)
        grid_layout_new_entry.addWidget(self.line_edit_table_name, 0, 1)

        # h_layout_table_name.addWidget(self.label_table_name)
        # h_layout_table_name.addWidget(self.line_edit_table_name)
        # v_layout_new_entry.addLayout(h_layout_table_name)

        h_layout_id = QHBoxLayout()
        self.label_id = QLabel("ID")
        self.line_edit_id = QLineEdit()
        grid_layout_new_entry.addWidget(self.label_id, 1, 0)
        grid_layout_new_entry.addWidget(self.line_edit_id, 1, 1)

        # h_layout_id.addWidget(self.label_id)
        # h_layout_id.addWidget(self.line_edit_id)
        # v_layout_new_entry.addLayout(h_layout_id)

        h_layout_def = QHBoxLayout()
        self.label_def = QLabel("Definition")
        self.line_edit_def = QLineEdit()
        grid_layout_new_entry.addWidget(self.label_def, 2, 0)
        grid_layout_new_entry.addWidget(self.line_edit_def, 2, 1)

        # h_layout_def.addWidget(self.label_def)
        # h_layout_def.addWidget(self.line_edit_def)
        # v_layout_new_entry.addLayout(h_layout_def)

        h_layout_description = QHBoxLayout()
        self.label_description = QLabel("Description")
        self.line_edit_description = QLineEdit()
        grid_layout_new_entry.addWidget(self.label_description, 3, 0)
        grid_layout_new_entry.addWidget(self.line_edit_description, 3, 1)

        # h_layout_description.addWidget(self.label_description)
        # h_layout_description.addWidget(self.line_edit_description)
        # v_layout_new_entry.addLayout(h_layout_description)

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

        self.setWindowTitle("Define Value")
        self.resize(600, 300)
        self.setLayout(self.form_layout)
        self.setModal(True)

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
        editor.setCurrentIndex(index.row())
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

        # Set the delegate for column 1 of our table
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

        def closeEvent(self, *args, **kwargs):
            # This is test that we can retrieve the updated data bound to the tableview object
            for row in range(self._tm.rowCount()):
                for col in range(self._tm.columnCount()):
                    #index = QtCore.QModelIndex(row, col)
                    value = str(self._tm.index(row, col).data())
                    print("row:" + str(row) + " col:" + str(col) + " value:" + value)