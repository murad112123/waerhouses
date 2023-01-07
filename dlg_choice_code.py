#!/usr/bin/env python3

import PyQt5.QtWidgets as QtWidgets
import PyQt5.uic as uic
import database

FormPrintDialog, _ = uic.loadUiType("dlg_choice_code.ui")


class PrintDialog(QtWidgets.QDialog, FormPrintDialog):  # type: ignore
    table_view: QtWidgets.QTableWidget

    def __init__(self, table, code: str) -> None:
        QtWidgets.QDialog.__init__(self)
        FormPrintDialog.__init__(self)
        self.setupUi(self)

        self.result_value: dict[str, str] | None = None
        self.material: dict[str, str] | None = None

        if table == "material":
            self.table_view_columns = ['تسلسلي', 'كود المادة', 'اسم المادة', 'السعر']
        else:
            self.table_view_columns = ['تسلسلي', 'كود المنتج', 'اسم المنتج', 'السعر']

        self.setup_controls()
        self.update_table(table, code)

    def setup_controls(self):
        # setup table
        self.table_view.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.table_view.setColumnCount(len(self.table_view_columns))
        self.table_view.setHorizontalHeaderLabels(self.table_view_columns)

        self.table_view.clicked.connect(lambda mi: self.one_click(self.table_view.item(mi.row(), 0).mid))
        self.table_view.doubleClicked.connect(lambda mi: self.double_click(self.table_view.item(mi.row(), 0).mid))

        # connect slots
        self.btn_choice.clicked.connect(lambda: self.double_click(self.material))

    def one_click(self, a: dict) -> None:
        self.material = a
        self.btn_choice.setEnabled(True)

    def double_click(self, a: dict[str, str]) -> None:
        self.result_value = a
        self.accept()

    def update_table(self, table: str, code: str) -> None:
        rows = database.db.get_material_product_by_code(table, code)
        self.table_view.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            self.table_view.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(row_idx + 1)))
            self.table_view.item(row_idx, 0).mid = row
            self.table_view.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.table_view.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(row['name'])))
            self.table_view.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(row['price'])))
        self.table_view.resizeColumnsToContents()
