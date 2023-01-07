#!/usr/bin/env python3

import glob
import json
import locale
import math
import os
import shutil
import sys
import tempfile
import time
from uuid6 import uuid8
from shutil import copy2

import PyQt5.QtCore as QtCore
import PyQt5.QtGui as QtGui
import PyQt5.QtWidgets as QtWidgets
import PyQt5.uic as uic

from jinja2 import Template
import xlwt
import hashlib

import toaster_Notify
from QDate import QDate
from QImageSelect import QImageSelect
import aes
import assets
import database
from dlg_choice_code import PrintDialog

Form_Main, _ = uic.loadUiType('warehouses.ui')
Form_Requests, _ = uic.loadUiType('requests.ui')
Form_InternalImex, _ = uic.loadUiType('internal_imex.ui')
PAGE_SIZE = 10
USER = ''
PASS = ''
PERMISSION = ''


class ReadOnlyDelegate(QtWidgets.QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        # print('createEditor event fired')
        return


class Requests(QtWidgets.QDialog, Form_Requests):
    def __init__(self, id):
        QtWidgets.QDialog.__init__(self)
        Form_Requests.__init__(self)
        self.setupUi(self)

        self.validator_money = QtGui.QRegExpValidator(
            QtCore.QRegExp('^(\$)?(([1-9]\d{0,2}(\,\d{3})*)|([1-9]\d*)|(0))(\.\d{1,2})?$'))

        self.r_id = id
        self.code = None
        self.setup_control()

    def setup_control(self):
        self.req_date.setDate(QDate.currentDate())

        self.req_brunch.clear()
        self.req_brunch.addItem('')
        self.req_brunch.addItems(database.db.query_req())

        self.req_table: QtWidgets.QTableWidget
        delegate = ReadOnlyDelegate(self.req_table)
        self.req_table.setItemDelegateForColumn(1, delegate)
        self.req_table.setItemDelegateForColumn(4, delegate)
        self.req_table.setItemDelegateForColumn(6, delegate)
        self.req_table.setItemDelegateForColumn(12, delegate)
        self.req_table.setRowCount(1)
        self.req_table.keyReleaseEvent = self.table_key_press_event

        self.fill_request(self.r_id)

        self.btn_save.clicked.connect(self.save_request)
        self.btn_cancel.clicked.connect(self.reject)
        # self.btn_print_bill.clicked.connect(self.print_bill)

        self.btn_save.setAutoDefault(False)
        self.btn_cancel.setAutoDefault(False)
        # self.btn_print_bill.setAutoDefault(False)

    def fill_request(self, id):
        if id == 0:
            self.r_id = database.db.get_next_id('requests')
            self.code = int(self.r_id) + 10000
        else:
            req = database.db.query_row('requests', id)
            self.r_id = req['id']
            self.code = req['code']
            self.req_date.setDate(QDate(req['date']))
            # self.req_brunch.setCurrentText(database.db.get_code_by_id('branches', req['r_id']))
            self.total_req.setText(str(req['total_req']))
            self.total_buy.setText(str(req['total_buy']))

        self.req_code.setText(str(self.code))

        orders = database.db.get_order_requests('req_order_v', self.r_id)
        self.req_table.setRowCount(len(orders) + 1)
        for row_idx, row in enumerate(orders):
            self.req_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.req_table.item(row_idx, 0).id = row['id']
            self.req_table.item(row_idx, 0).mid = row['m_id']
            self.req_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['name'])))
            self.req_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(row['quantity'])))
            self.req_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(row['requester'])))
            self.req_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(str(row['description'])))
            self.req_table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(str(row['quantity'])))
            self.req_table.setItem(row_idx, 6, QtWidgets.QTableWidgetItem(str(row['requester'])))
            self.req_table.setItem(row_idx, 7, QtWidgets.QTableWidgetItem(str(row['price'])))
            self.req_table.setItem(row_idx, 8, QtWidgets.QTableWidgetItem(str(row['total'])))
            self.req_table.setItem(row_idx, 9, QtWidgets.QTableWidgetItem(str(row['link'])))
            self.req_table.setItem(row_idx, 10, QtWidgets.QTableWidgetItem(str(row['project'])))
            self.req_table.setItem(row_idx, 11, QtWidgets.QTableWidgetItem(str(row['priority'])))
            self.req_table.setItem(row_idx, 12, QtWidgets.QTableWidgetItem(str(row['quantity_receive'])))
            self.req_table.setItem(row_idx, 13, QtWidgets.QTableWidgetItem(str(row['price_receive'])))
            self.req_table.setItem(row_idx, 14, QtWidgets.QTableWidgetItem(str(row['total_receive'])))
            self.req_table.setItem(row_idx, 15, QtWidgets.QTableWidgetItem(str(row['seller_note'])))
            btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
            btn_delete.clicked.connect(lambda: self.delete_order(self.req_table.currentRow()))
            self.req_table.setCellWidget(row_idx, 16, btn_delete)

    def delete_order(self, current_row):
        self.req_table.removeRow(current_row)
        self.calculate_total()

    def table_key_press_event(self, event: QtGui.QKeyEvent):
        self.req_table: QtWidgets.QTableWidget
        if event.key() == QtCore.Qt.Key_Return:
            if self.req_table.currentColumn() == 0 and self.req_table.currentRow() + 1 == self.req_table.rowCount():
                self.update_table(self.req_table.currentRow())
            else:
                self.enter_event(self.req_table.currentRow())

    def update_table(self, current_row):
        code = self.req_table.item(current_row, 0).text()
        material = database.db.get_all_by_code("material", code)
        if material:
            for idx in range(self.req_table.rowCount() - 1):
                if self.req_table.item(idx, 0).text() == code:
                    new = int(self.req_table.item(idx, 2).text()) + 1
                    self.req_table.setItem(idx, 2, QtWidgets.QTableWidgetItem(str(new)))
                    self.req_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))
                    return

            self.req_table.item(current_row, 0).pid = material['id']
            self.req_table.setItem(current_row, 1, QtWidgets.QTableWidgetItem(material['name']))
            self.req_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))
            self.req_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem(str(material['description'])))
            self.req_table.setItem(current_row, 5, QtWidgets.QTableWidgetItem(str(material['price'])))
            self.req_table.setItem(current_row, 7, QtWidgets.QTableWidgetItem(str(material['link'])))
            self.req_table.setRowCount(self.req_table.rowCount() + 1)
            btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
            btn_delete.clicked.connect(lambda: self.delete_order(self.req_table.currentRow()))
            self.req_table.setCellWidget(current_row, 14, btn_delete)
            self.calculate_total()
        else:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
            self.req_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))

    def enter_event(self, current_row):
        code = self.req_table.item(current_row, 0).text()
        material = database.db.get_all_by_code("material", code)
        if self.req_table.item(current_row, 6).text() == '':
            self.req_table.setItem(current_row, 6, QtWidgets.QTableWidgetItem('0'))

        if self.req_table.item(current_row, 12).text() == '':
            self.req_table.setItem(current_row, 12, QtWidgets.QTableWidgetItem('0'))

        if self.req_table.item(current_row, 2).text() == '':
            self.req_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))
        quantity_req = int(self.req_table.item(current_row, 2).text())

        total_req = quantity_req * float(self.req_table.item(current_row, 5).text())
        self.req_table.setItem(current_row, 6, QtWidgets.QTableWidgetItem(str(total_req)))

        if self.req_table.item(current_row, 10).text() == '':
            self.req_table.setItem(current_row, 10, QtWidgets.QTableWidgetItem('1'))
        quantity_rec = int(self.req_table.item(current_row, 10).text())

        total_req = quantity_rec * float(self.req_table.item(current_row, 11).text())
        self.req_table.setItem(current_row, 12, QtWidgets.QTableWidgetItem(str(total_req)))
        self.calculate_total()

    def calculate_total(self):
        total_req = 0
        total_buy = 0
        for i in range(0, self.req_table.rowCount()):
            if self.req_table.item(i, 6) is not None:
                total_req += float(self.req_table.item(i, 6).text())
            if self.req_table.item(i, 12) is not None:
                total_buy += float(self.req_table.item(i, 12).text())
        self.total_req.setText(str(total_req))
        self.total_buy.setText(str(total_buy))

    def save_request(self):
        request = dict()
        request['id'] = self.r_id
        request['code'] = self.req_code.text()
        request['brunch'] = self.req_brunch.text()
        request['date'] = QDate.toString(self.req_date.date())
        request['total'] = self.total.text()
        request['discount'] = self.discount.text()

        orders = []
        for idx in range(self.req_table.rowCount()):
            order = dict()
            order['r_id'] = self.r_id
            if self.req_table.item(idx, 0) and self.req_table.item(idx, 0).text():
                if hasattr(self.req_table.item(idx, 0), 'id'):
                    order['id'] = self.req_table.item(idx, 0).id
                    order['m_id'] = self.req_table.item(idx, 0).mid
                else:
                    order['id'] = int(database.db.get_next_id('req_order')) + idx
                    order['m_id'] = database.db.get_id_by_code('material', self.req_order.item(idx, 0).text())
                if self.req_order.item(idx, 2) and self.req_order.item(idx, 2).text():
                    order['quantity'] = self.req_order.item(idx, 2).text()

                if self.req_order.item(idx, 3) and self.req_order.item(idx, 3).text():
                    order['requester'] = self.req_order.item(idx, 3).text()

                if self.req_order.item(idx, 4) and self.req_order.item(idx, 4).text():
                    order['description'] = self.req_order.item(idx, 4).text()

                if self.req_order.item(idx, 5) and self.req_order.item(idx, 5).text():
                    order['price'] = self.req_order.item(idx, 5).text()

                if self.req_order.item(idx, 6) and self.req_order.item(idx, 6).text():
                    order['total'] = self.req_order.item(idx, 6).text()

                if self.req_order.item(idx, 7) and self.req_order.item(idx, 7).text():
                    order['link'] = self.req_order.item(idx, 7).text()

                if self.req_order.item(idx, 8) and self.req_order.item(idx, 8).text():
                    order['project'] = self.req_order.item(idx, 8).text()

                if self.req_order.item(idx, 9) and self.req_order.item(idx, 9).text():
                    order['priority'] = self.req_order.item(idx, 9).text()

                if self.req_order.item(idx, 10) and self.req_order.item(idx, 10).text():
                    order['quantity_receive'] = self.req_order.item(idx, 10).text()

                if self.req_order.item(idx, 11) and self.req_order.item(idx, 11).text():
                    order['price_receive'] = self.req_order.item(idx, 11).text()

                if self.req_order.item(idx, 12) and self.req_order.item(idx, 12).text():
                    order['total_receive'] = self.req_order.item(idx, 12).text()

                if self.req_order.item(idx, 13) and self.req_order.item(idx, 13).text():
                    order['seller_note'] = self.req_order.item(idx, 13).text()

                orders.append(order)
        if int(database.db.count_row("requests", request['code'])) == 0:
            database.db.insert_row("requests", request)
        else:
            database.db.update_row("requests", request)

        database.db.insert_table('req_order', orders)
        self.accept()

    def print_bill(self):
        pass


class InternalImex(QtWidgets.QDialog, Form_InternalImex):
    def __init__(self, id):
        QtWidgets.QDialog.__init__(self)
        Form_InternalImex.__init__(self)
        self.setupUi(self)

        self.validator_money = QtGui.QRegExpValidator(
            QtCore.QRegExp('^(\$)?(([1-9]\d{0,2}(\,\d{3})*)|([1-9]\d*)|(0))(\.\d{1,2})?$'))

        self.code = None
        self.b_id = id
        self.setup_control()

    def setup_control(self):
        self.b_date.setDate(QDate.currentDate())
        self.branch_codes = database.db.query_csp("branches")
        self.bill_type.addItem('')
        self.sender_brunch.addItem('')
        self.receiver_brunch.addItem('')
        self.sender_brunch.addItems(self.branch_codes)
        self.receiver_brunch.addItems(self.branch_codes)

        self.internal_table: QtWidgets.QTableWidget
        delegate = ReadOnlyDelegate(self.internal_table)
        self.internal_table.setItemDelegateForColumn(1, delegate)
        self.internal_table.setItemDelegateForColumn(4, delegate)
        self.internal_table.setRowCount(1)
        self.internal_table.keyReleaseEvent = self.table_key_press_event

        self.fill_bill(self.b_id)

        self.btn_save.clicked.connect(self.save_bill)
        self.btn_cancel.clicked.connect(self.reject)
        # self.btn_print_bill.clicked.connect(self.print_bill)

        self.btn_save.setAutoDefault(False)
        self.btn_cancel.setAutoDefault(False)
        # self.btn_print_bill.setAutoDefault(False)

    def fill_bill(self, id):
        if id == 0:
            self.b_id = database.db.get_next_id('internal_imex')
            self.code = int(self.b_id) + 10000
        else:
            bill = database.db.query_row('internal_imex', id)
            self.b_id = bill['id']
            self.code = bill['code']
            self.b_date.setDate(QDate(bill['date']))
            self.bill_type.setCurrentText(bill['type'])
            self.sender_brunch.setCurrentText(bill['sender'])
            self.receiver_brunch.setCurrentText(bill['receiver'])
            self.total.setText(str(bill['total']))

        self.bill_code.setText(str(self.code))

        orders = database.db.get_order_bill('int_order_v', self.b_id)
        self.internal_table.setRowCount(len(orders) + 1)
        for row_idx, row in enumerate(orders):
            self.internal_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.internal_table.item(row_idx, 0).id = row['id']
            self.internal_table.item(row_idx, 0).mid = row['m_id']
            self.internal_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['name'])))
            self.internal_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(row['quantity'])))
            self.internal_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(row['price'])))
            self.internal_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(str(row['total'])))
            btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
            btn_delete.clicked.connect(lambda: self.delete_order(self.internal_table.currentRow()))
            self.internal_table.setCellWidget(row_idx, 5, btn_delete)

    def delete_order(self, current_row):
        self.internal_table.removeRow(current_row)
        self.calculate_total()

    def table_key_press_event(self, event: QtGui.QKeyEvent):
        self.internal_table: QtWidgets.QTableWidget
        if event.key() == QtCore.Qt.Key_Return:
            if self.internal_table.currentColumn() == 0 and self.internal_table.currentRow() + 1 == self.internal_table.rowCount():
                self.update_table(self.internal_table.currentRow())
                self.internal_table.setRowCount(self.internal_table.rowCount() + 1)
            elif self.internal_table.currentColumn() == 0 and self.internal_table.currentRow() + 1 != self.internal_table.rowCount():
                self.update_table(self.internal_table.currentRow())
            else:
                self.enter_event(self.internal_table.currentRow())

    def update_table(self, current_row):
        code = self.internal_table.item(current_row, 0).text()
        material = database.db.get_material_by_code(code)
        if self.bill_type.currentIndex() == 1:
            self.receiver_brunch.setCurrentText = USER
            if material:
                for idx in range(self.internal_table.rowCount() - 1):
                    if self.internal_table.item(idx, 0).text() == code:
                        new = int(self.internal_table.item(idx, 2).text()) + 1
                        self.internal_table.setItem(idx, 2, QtWidgets.QTableWidgetItem(str(new)))
                        self.internal_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))
                        self.delete_order(current_row)
                        return

                self.internal_table.item(current_row, 0).pid = material['id']
                self.internal_table.setItem(current_row, 1, QtWidgets.QTableWidgetItem(material['name']))
                self.internal_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))
                self.internal_table.setItem(current_row, 3, QtWidgets.QTableWidgetItem(str(material['price'])))
                self.internal_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem(str(material['price'])))
                self.internal_table.setRowCount(self.internal_table.rowCount() + 1)
                btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
                btn_delete.clicked.connect(lambda: self.delete_order(self.internal_table.currentRow()))
                self.internal_table.setCellWidget(current_row, 5, btn_delete)
                self.calculate_total()
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
                self.internal_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))

        else:
            self.sender_brunch.setCurrentText = USER
            if material:
                if int(material['quantity']) >= 1:
                    for idx in range(self.internal_table.rowCount() - 1):
                        if self.internal_table.item(idx, 0).text() == code:
                            new = int(self.internal_table.item(idx, 2).text()) + 1
                            self.internal_table.setItem(idx, 2, QtWidgets.QTableWidgetItem(str(new)))
                            self.internal_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))
                            self.delete_order(current_row)
                            return

                    self.internal_table.item(current_row, 0).pid = material['id']
                    self.internal_table.setItem(current_row, 1, QtWidgets.QTableWidgetItem(material['name']))
                    self.internal_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))
                    self.internal_table.setItem(current_row, 3, QtWidgets.QTableWidgetItem(str(material['price'])))
                    self.internal_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem(str(material['price'])))
                    self.internal_table.setRowCount(self.internal_table.rowCount() + 1)
                    btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
                    btn_delete.clicked.connect(lambda: self.delete_order(self.internal_table.currentRow()))
                    self.internal_table.setCellWidget(current_row, 5, btn_delete)
                    self.calculate_total()
                else:
                    self.internal_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))
                    QtWidgets.QMessageBox.warning(None, 'خطأ', 'غير متوفر\n لا يوجد لديك من هذه المادة')
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
                self.internal_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))

    def enter_event(self, current_row):
        code = self.internal_table.item(current_row, 0).text()
        material = database.db.get_material_by_code(code)
        if self.internal_table.item(current_row, 4).text() == '':
            self.internal_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem('0'))

        if self.internal_table.item(current_row, 2).text() == '':
            self.internal_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))
        quantity = int(self.internal_table.item(current_row, 2).text())
        if self.bill_type.currentIndex() == 2:
            if quantity > int(material['quantity']):
                quantity = int(material['quantity'])
                self.internal_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem(material['quantity']))
                toaster_Notify.QToaster.show_message(parent=self,
                                                     message=f"غير متوفر\n لقد بقي من هذه المادة {material['quantity']} قطعة فقط ")

        total = quantity * float(self.internal_table.item(current_row, 3).text())
        self.internal_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem(str(total)))
        self.calculate_total()

    def calculate_total(self):
        total = 0
        for i in range(0, self.internal_table.rowCount()):
            if self.internal_table.item(i, 4) is not None:
                total += float(self.internal_table.item(i, 4).text())
        self.total.setText(str(total))

    def save_bill(self):
        bill = dict()
        bill['id'] = self.b_id
        bill['code'] = self.bill_code.text()
        bill['type'] = self.bill_type.text()
        bill['sender'] = self.sender_brunch.text()
        bill['receiver'] = self.receiver_brunch.text()
        bill['date'] = QDate.toString(self.b_date.date())
        bill['total'] = self.total.text()

        orders = []
        for idx in range(self.internal_table.rowCount()):
            order = dict()
            order['b_id'] = self.b_id
            if self.internal_table.item(idx, 0) and self.internal_table.item(idx, 0).text():
                if hasattr(self.internal_table.item(idx, 0), 'id'):
                    order['id'] = self.internal_table.item(idx, 0).id
                    order['m_id'] = self.internal_table.item(idx, 0).mid
                else:
                    order['id'] = int(database.db.get_next_id('int_order')) + idx
                    order['m_id'] = database.db.get_id_by_code('material', self.internal_table.item(idx, 0).text())
                if self.internal_table.item(idx, 2) and self.internal_table.item(idx, 2).text():
                    order['quantity'] = self.internal_table.item(idx, 2).text()

                if self.internal_table.item(idx, 3) and self.internal_table.item(idx, 3).text():
                    order['price'] = self.internal_table.item(idx, 3).text()

                if self.internal_table.item(idx, 4) and self.internal_table.item(idx, 4).text():
                    order['total'] = self.internal_table.item(idx, 4).text()

                orders.append(order)

        if int(database.db.count_row("internal_imex", bill['code'])) == 0:
            database.db.insert_row("internal_imex", bill)
        else:
            database.db.update_row("internal_imex", bill)

        database.db.insert_table('int_order', orders, self.b_id)
        self.accept()

    def print_bill(self):
        pass


class AppMainWindow(QtWidgets.QMainWindow, Form_Main):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Form_Main.__init__(self)
        self.setupUi(self)

        self.config = None

        self.validator_code = QtGui.QRegExpValidator(QtCore.QRegExp('[\u0621-\u064A0-9a-zA-Z][0-9]*'))
        self.validator_int = QtGui.QRegExpValidator(QtCore.QRegExp('[0-9]+'))
        self.validator_money = QtGui.QRegExpValidator(
            QtCore.QRegExp('^(\$)?(([1-9]\d{0,2}(\,\d{3})*)|([1-9]\d*)|(0))(\.\d{1,2})?$'))

        self.branch_id = 0
        self.branch_codes = None

        self.m_img.attach = None
        self.m_img.asset = None

        self._typing_timer_m = QtCore.QTimer()
        self.material_id = 0
        self.material_co = 0
        self.material_codes = None

        self._typing_timer_m_b = QtCore.QTimer()
        self.material_branch_id = 0

        self._typing_timer_p = QtCore.QTimer()
        self.product_id = 0
        self.product_co = 0
        self.product_codes = None

        self._typing_timer_p_b = QtCore.QTimer()
        self.product_branch_id = 0

        self._typing_timer_req = QtCore.QTimer()
        self.requests_id = 0
        self.page_size_requests = PAGE_SIZE

        self.setup_login()

    def setup_login(self):
        self.menubar.setVisible(False)
        self.stackedWidget.setCurrentIndex(1)
        self.txt_username.setFocus()
        self.btn_in.clicked.connect(self.enter_app)
        self.btn_exit.clicked.connect(lambda: sys.exit(1))

    def enter_app(self):
        global PASS
        global USER
        global PERMISSION
        password = self.txt_password.text()
        PASS = hashlib.sha256(password.encode()).digest()
        USER = self.txt_username.text()
        if password != '':
            self.config = dict()
            with open('config.dat', 'r') as config_file:
                aes_cipher = aes.AESCipher(password)
                try:
                    self.config = json.loads(aes_cipher.decrypt(config_file.read()))
                except Exception:
                    self.lbl_wrong.setText('* كلمة المرور غير صحيحة !!!')
                else:
                    database.Database.open_database(self.config['password'])
                    p = database.db.is_user(USER)
                    if p is not None:
                        self.branch_id = p['id']
                        PERMISSION = p['permission']
                        self.setup_controls()
                        self.stackedWidget.setCurrentIndex(0)
                    else:
                        self.lbl_wrong.setText('* اسم المستخدم غير صحيح !!!')
        else:
            self.lbl_wrong.setText('* يجب ادخال كلمة المرور !!!')

    def change_pass_(self):
        self.stackedWidget.setCurrentIndex(2)
        self.menubar.setVisible(False)
        self.old_pass.setFocus()
        self.btn_save_pass.clicked.connect(self.save_new_pass)
        self.btn_cancel_pass.clicked.connect(
            lambda: self.stackedWidget.setCurrentIndex(0) or self.menubar.setVisible(True))

    def save_new_pass(self):
        global PASS
        if self.old_pass.text() == PASS:
            if self.new_pass.text() == self.new_pass_confirm.text():
                if self.new_pass.text() != '':
                    PASS = hashlib.sha256(self.new_pass.text().encode()).digest()
                    database.db.change_user_pass(USER, PASS)
                    self.stackedWidget.setCurrentIndex(0)
                    toaster_Notify.QToaster.show_message(parent=self,
                                                         message="تغيير كلمة المرور\nتم تغيير كلمة المرور بنجاح")
                else:
                    self.lbl_wrong.setText('* كلمة المرور الجديدة لا يمكن أن تكون فارغة !!!')
            else:
                self.lbl_wrong.setText('* كلمة المرور الجديدة غير متطابقة !!!')
        else:
            self.lbl_wrong_e.setText('* كلمة المرور القديمة غير صحيحة !!!')

    def setup_controls(self):
        self.menubar.setVisible(True)
        self.tabWidget.tabBar().setVisible(False)
        self.tabWidget.setCurrentIndex(0)

        self.branch_codes = database.db.query_csp("branches")

        self.setup_controls_material()
        self.setup_controls_product()
        self.setup_controls_requests()

    # export tables to excel
    def to_excel(self, table):
        file_name, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File', '', ".dot(*.dot)")
        wbk = xlwt.Workbook()
        sheet = wbk.add_sheet("sheet", cell_overwrite_ok=True)
        sheet.cols_right_to_left = True
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold = True
        style.font = font
        model = table.model()
        for c in range(1, model.columnCount()):
            text = model.headerData(c, QtCore.Qt.Horizontal)
            sheet.write(0, c - 1, text, style=style)

        for c in range(1, model.columnCount()):
            for r in range(model.rowCount()):
                text = model.data(model.index(r, c))
                sheet.write(r + 1, c - 1, text)
        try:
            wbk.save(file_name)
        except IOError:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'يوجد خطأ في حفظ الملف')

    def change_page_size(self, table):
        if table == 'material':
            self.m_page_num.setRange(1, math.ceil(int(database.db.count_row("material", 1)) / self.m_page_size.value()))
            self._typing_timer_m.start(1000)
        elif table == 'material_branch':
            self.material_branch_page_num.setRange(1, math.ceil(
                int(database.db.count_row("available_m", 1)) / self.material_branch_page_size.value()))
            self._typing_timer_m_b.start(1000)
        elif table == 'product':
            self.p_page_num.setRange(1, math.ceil(int(database.db.count_row("product", 1)) / self.p_page_size.value()))
            self._typing_timer_p.start(1000)
        elif table == 'product_branch':
            self.product_branch_page_num.setRange(1, math.ceil(
                int(database.db.count_row("available_p", 1)) / self.product_branch_page_size.value()))
            self._typing_timer_p_b.start(1000)
        elif table == 'requests':
            self.page_size_requests = self.req_page_siz.value()
            self.req_page_num.setRange(1,
                                       math.ceil(int(database.db.count_row("requests", 1)) / self.page_size_requests))
            self._typing_timer_req.start(1000)

    def check_date_from(self, x):
        if x == 'requests':
            self._typing_timer_req.start(1000)
            if self.ch_req_date_from.isChecked():
                self.req_date_from.setEnabled(True)
                self.req_date_from.dateChanged.connect(lambda: self._typing_timer_req.start(1000))
                self.ch_req_date_to.setEnabled(True)
            else:
                self.req_date_from.setEnabled(False)
                self.req_date_from.setDate(QDate.currentDate())
                self.ch_req_date_to.setEnabled(False)
                self.req_date_to.setEnabled(False)
                self.req_date_to.setDate(QDate.currentDate())
                self.ch_req_date_to.setChecked(False)

    def check_date_to(self, x):
        if x == 'requests':
            self._typing_timer_req.start(1000)
            if self.ch_req_date_to.isChecked():
                self.req_date_to.setEnabled(True)
                self.req_date_to.dateChanged.connect(lambda: self._typing_timer_req.start(1000))
            else:
                self.req_date_to.setEnabled(False)
                self.req_date_to.setDate(QDate.currentDate())

    # #################################################################################
    # material
    def setup_controls_material(self):
        self.tabWidget_3.setCurrentIndex(0)
        self.m_code.setValidator(self.validator_code)
        self.m_code_search.setValidator(self.validator_code)
        self.m_less_quantity.setValidator(self.validator_int)
        self.m_price.setValidator(self.validator_money)

        self.m_img.mousePressEvent = lambda *args, **kwargs: self.pick_image(self.m_img)

        self._typing_timer_m.setSingleShot(True)
        self._typing_timer_m.timeout.connect(self.update_material_table)

        # search
        self.m_code_search.textChanged.connect(lambda text: self._typing_timer_m.start(1000))
        self.m_name_search.textChanged.connect(lambda text: self._typing_timer_m.start(1000))
        self.m_type_search.textChanged.connect(lambda text: self._typing_timer_m.start(1000))

        self.m_page_num.setRange(1, math.ceil(int(database.db.count_row("material", 1)) / self.m_page_size.value()))
        self.m_page_num.valueChanged.connect(lambda text: self._typing_timer_m.start(1000))
        self.m_page_size.valueChanged.connect(lambda: self.change_page_size('material'))

        # table
        self.m_table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.m_table.doubleClicked.connect(lambda mi: self.fill_material_info(self.m_table.item(mi.row(), 0).id))
        self.m_table.clicked.connect(lambda mi: self.one_click_m(self.m_table.item(mi.row(), 0).id))

        # btn
        self.btn_add_material.clicked.connect(self.create_new_material)
        self.btn_edit_material.clicked.connect(self.update_material)
        self.btn_clear_material.clicked.connect(self.clear_material_inputs)
        self.btn_delete_material.clicked.connect(self.delete_material)

        # print and to exel
        self.btn_print_table_m.clicked.connect(self.print_table_material)
        self.btn_to_exel_m.clicked.connect(lambda: self.to_excel(self.m_table))

        # pages
        self.m_post.clicked.connect(lambda: self.m_page_num.setValue(self.m_page_num.value() + 1))
        self.m_previous.clicked.connect(lambda: self.m_page_num.setValue(self.m_page_num.value() - 1))
        self.m_last.clicked.connect(lambda: self.m_page_num.setValue(
            math.ceil(int(database.db.count_row("material", 1)) / self.m_page_size.value())))
        self.m_first.clicked.connect(lambda: self.m_page_num.setValue(1))

        # ############################################################

        self.m_quantity.setValidator(self.validator_int)
        self.m_b_code_search.setValidator(self.validator_code)

        self.material_codes = database.db.query_csp("material")

        if PERMISSION == '1':
            self.m_branches.setEnabled(True)
            self.m_branches.addItem('')
            self.m_branches.addItems(self.branch_codes.values())
        else:
            self.m_branches.addItem(USER)
            self.m_branches.setEnabled(False)

        self.m_branches_search.addItem('')
        self.m_branches_search.addItems(self.branch_codes.values())

        self.m_b_code.returnPressed.connect(self.material_code_key_press_event)

        self._typing_timer_m_b.setSingleShot(True)
        self._typing_timer_m_b.timeout.connect(self.update_material_branch_table)

        self.m_b_code_search.textChanged.connect(lambda text: self._typing_timer_m_b.start(1000))
        self.m_branches_search.currentTextChanged.connect(lambda text: self._typing_timer_m_b.start(1000))

        self.material_branch_page_num.setRange(1, math.ceil(
            int(database.db.count_row("available_m", 1)) / self.material_branch_page_size.value()))
        self.material_branch_page_num.valueChanged.connect(lambda text: self._typing_timer_m_b.start(1000))
        self.material_branch_page_size.valueChanged.connect(lambda: self.change_page_size('material_branch'))

        self.material_branch_table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.material_branch_table.doubleClicked.connect(
            lambda mi: self.fill_material_branch_info(self.material_branch_table.item(mi.row(), 0).id))
        self.material_branch_table.clicked.connect(
            lambda mi: self.one_click_m_b(self.material_branch_table.item(mi.row(), 0).id))

        # btn
        self.btn_add_material_branch.clicked.connect(self.create_new_material_branch)
        self.btn_edit_material_branch.clicked.connect(self.update_material_branch)
        self.btn_delete_material_branch.clicked.connect(self.delete_material_branch)
        self.btn_clear_material_branch.clicked.connect(self.clear_material_branch_inputs)

        # print and to exel
        self.btn_print_table_material_branch.clicked.connect(self.print_table_material_branch)
        self.btn_to_exel_material_branch.clicked.connect(lambda: self.to_excel(self.m_table))

        # pages
        self.material_branch_post.clicked.connect(
            lambda: self.material_branch_page_num.setValue(self.material_branch_page_num.value() + 1))
        self.material_branch_previous.clicked.connect(
            lambda: self.material_branch_page_num.setValue(self.material_branch_page_num.value() - 1))
        self.material_branch_last.clicked.connect(lambda: self.material_branch_page_num.setValue(
            math.ceil(int(database.db.count_row("available_m", 1)) / self.material_branch_page_size.value())))
        self.material_branch_first.clicked.connect(lambda: self.material_branch_page_num.setValue(1))

        self.update_material_table()
        self.clear_material_inputs()

        self.update_material_branch_table()
        self.clear_material_branch_inputs()

    def save_material_info(self):
        global USER
        material = dict()
        material['code'] = self.m_code.text()
        material['name'] = self.m_name.text()
        material['description'] = self.m_desciption.text()
        material['type'] = self.m_type.text()
        if self.ch_pr.isChecked():
            material['pu_pr'] = USER
        else:
            material['pu_pr'] = 0
        material['source'] = self.m_source.text()
        material['less_quantity'] = self.m_less_quantity.text()
        material['price'] = self.m_price.text()
        material['link'] = self.m_link.text()
        material['note'] = self.m_note.toPlainText()

        material['pic'] = self.m_img.asset

        return material

    def create_new_material(self):
        global USER
        material = self.save_material_info()
        material['id'] = str(uuid8())
        if material['code'] and material['name']:
            if int(database.db.count_row("material", material['code'])) == 0:
                if self.m_img.attach:
                    material['pic'] = assets.create_asset(material['id'], self.m_img.attach, self.config['password'])
                database.db.insert_row("material", material)
                toaster_Notify.QToaster.show_message(parent=self, message=f"إضافة مادة\nتم إضافة المادة {material['name']} بنجاح")
                self.update_material_table()
                self.clear_material_inputs()
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'إن الكود مكرر')
        else:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'يجب أن تدخل الكود واسم المادة')

    def update_material(self):
        material = self.save_material_info()
        material['id'] = self.material_id
        if material['code'] and material['name']:
            if material['code'] == self.material_co:
                if self.m_img.attach:
                    if self.m_img.asset:
                        assets.asset_delete(material['id'], self.m_img.asset)
                    material['pic'] = assets.create_asset(material['id'], self.m_img.attach, self.config['password'])
                else:
                    material['pic'] = self.m_img.asset
                database.db.update_row("material", material)
                self.update_material_table()
                self.clear_material_inputs()
                toaster_Notify.QToaster.show_message(parent=self,
                                                     message=f"تعديل مادة\nتم تعديل المادة {material['name']} بنجاح")
            elif int(database.db.count_row("material", material['code'])) == 0:
                if self.m_img.attach:
                    if self.m_img.asset:
                        assets.asset_delete(material['id'], self.m_img.asset)
                    material['pic'] = assets.create_asset(material['id'], self.m_img.attach, self.config['password'])
                else:
                    material['pic'] = self.m_img.asset
                database.db.update_row("material", material)
                self.update_material_table()
                self.clear_material_inputs()
                toaster_Notify.QToaster.show_message(parent=self,
                                                     message=f"تعديل مادة\nتم تعديل المادة {material['name']} بنجاح")
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'إن الكود مكرر')
        else:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'يجب أن تدخل الكود واسم المادة')

    def delete_material(self):
        material = database.db.query_row("material", self.material_id)
        msg = QtWidgets.QMessageBox()
        button_reply = msg.question(self, 'تأكيد', f"هل أنت متأكد من حذف {material['name']} ؟ ",
                                    msg.Yes | msg.No,
                                    msg.No)
        if button_reply == msg.Yes:
            database.db.delete_row("material", self.material_id)
            try:
                shutil.rmtree(f"assets/{self.material_id}")
            except OSError as e:
                print("Error: %s - %s." % (e.filename, e.strerror))
            self.update_material_table()
            self.clear_material_inputs()
            toaster_Notify.QToaster.show_message(parent=self, message=f"حذف مادة\nتم حذف المادة{material['name']} بنجاح")

        self.btn_delete_material.setEnabled(False)

    def search_material_save(self):
        fil = {}
        if self.m_code_search.text():
            fil['code'] = self.m_code_search.text()
        if self.m_name_search.text():
            fil['name'] = self.m_name_search.text()
        if self.m_type_search.text():
            fil['type'] = self.m_type_search.text()

        return fil

    def update_material_table(self):
        fil = self.search_material_save()
        rows = database.db.query_all_material(fil, self.m_page_size.value() * (self.m_page_num.value() - 1),
                                              self.m_page_size.value())
        self.m_table.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            self.m_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(
                str(row_idx + 1 + (self.m_page_size.value() * (self.m_page_num.value() - 1)))))
            self.m_table.item(row_idx, 0).id = row['id']
            self.m_table.item(row_idx, 0).setTextAlignment(QtCore.Qt.AlignCenter)
            self.m_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.m_table.item(row_idx, 1).setTextAlignment(QtCore.Qt.AlignCenter)
            self.m_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(row['name']))
            self.m_table.item(row_idx, 2).setTextAlignment(QtCore.Qt.AlignCenter)
            self.m_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(row['description']))
            self.m_table.item(row_idx, 3).setTextAlignment(QtCore.Qt.AlignCenter)
            self.m_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(row['type']))
            self.m_table.item(row_idx, 4).setTextAlignment(QtCore.Qt.AlignCenter)
            self.m_table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(row['price']))
            self.m_table.item(row_idx, 5).setTextAlignment(QtCore.Qt.AlignCenter)
        self.m_table.resizeColumnsToContents()

    def clear_material_inputs(self):
        self.material_id = 0
        self.material_co = 0
        self.m_code.clear()
        self.m_code.setFocus()
        self.m_name.clear()
        self.m_desciption.clear()
        self.m_type.clear()
        self.ch_pr.setChecked(False)
        self.m_less_quantity.setText('0')
        self.m_price.setText('0')
        self.m_link.clear()
        self.m_source.clear()
        self.m_note.clear()

        pixmap = QtGui.QPixmap("icons/empty.png")
        self.m_img.setPixmap(pixmap.scaled(self.m_img.size(), QtCore.Qt.KeepAspectRatio))
        self.m_img.attach = None
        self.m_img.asset = None

        self.btn_edit_material.setEnabled(False)
        self.btn_delete_material.setEnabled(False)
        self.btn_add_material.setEnabled(True)

    def one_click_m(self, id):
        self.material_id = id
        self.btn_delete_material.setEnabled(True)

    def fill_material_info(self, id):
        self.btn_edit_material.setEnabled(True)
        self.btn_delete_material.setEnabled(True)
        self.btn_add_material.setEnabled(False)
        self.material_id = id
        material = database.db.query_row("material", id)
        if material:
            self.material_co = material['code']
            self.m_code.setText(material['code'])
            self.m_name.setText(material['name'])
            self.m_desciption.setText(material['description'])
            self.m_type.setText(material['type'])
            if material['pu_pr'] == '0':
                self.ch_pr.setChecked(False)
            else:
                self.ch_pr.setChecked(True)
            self.m_source.setText(material['source'])
            self.m_less_quantity.setText(material['less_quantity'])
            self.m_price.setText(material['price'])
            self.m_link.setText(material['link'])
            self.m_note.setText(material['note'])

            self.m_img.asset = material['pic']
            if material['pic']:
                pixmap = QtGui.QPixmap(assets.decrypt_asset(self.material_id, material['pic'], self.config['password']))
            else:
                pixmap = QtGui.QPixmap("icons/empty.png")

            self.m_img.setPixmap(pixmap.scaled(self.m_img.size(), QtCore.Qt.KeepAspectRatio))

    def print_table_material(self):
        fil = self.search_material_save()
        materials = database.db.query_all_material(fil, 0, database.db.count_row("material", 1))

        with open('./html/material_template.html', 'r') as f:
            template = Template(f.read())
            fp = tempfile.NamedTemporaryFile(mode='w', delete=False, dir='./html/tmp/', suffix='.html')
            for idx, material in enumerate(materials):
                material['idx'] = idx + 1

            html = template.render(materials=materials, date=time.strftime("%A, %d %B %Y %I:%M %p"))
            html = html.replace('style.css', '../style.css').replace('ph1.png', '../ph1.png')
            fp.write(html)
            fp.close()
            os.system('setsid firefox ' + fp.name + ' &')

    def pick_image(self, label: QtWidgets.QLabel):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setLayoutDirection(QtCore.Qt.RightToLeft)
        msg.setText('هل تريد عرض الصورة الشخصية ام اختيار صورة جديدة ؟')
        btn_export = msg.addButton('تصدير الصورة', QtWidgets.QMessageBox.AcceptRole)
        btn_show = msg.addButton('عرض الصورة', QtWidgets.QMessageBox.AcceptRole)
        btn_new = msg.addButton('اختيار صورة جديدة', QtWidgets.QMessageBox.AcceptRole)
        msg.addButton('الغاء', QtWidgets.QMessageBox.RejectRole)
        msg.exec_()
        # new entry
        if msg.clickedButton() == btn_new:
            opt = QtWidgets.QFileDialog.Options()
            opt |= QtWidgets.QFileDialog.DontUseNativeDialog
            ext = 'Image files(*.jpg *.jpeg *.png *.bmp);;All file types(*.*)'
            file, _ = QtWidgets.QFileDialog.getOpenFileName(self, caption="اختيار صورة", filter=ext, options=opt)
            if file:
                pixmap = QImageSelect.spawn(title="select image", image=file, maximum_size=QtCore.QSize(600, 400))
                if pixmap:
                    label.clear()
                    pixmap.save(f"/tmp/prisons/{os.path.basename(file)}", "PNG")
                    label.attach = f"/tmp/prisons/{os.path.basename(file)}"
                    label.setPixmap(pixmap.scaled(label.size(), QtCore.Qt.KeepAspectRatio))
        elif msg.clickedButton() == btn_show:
            if label.asset:
                os.system('xdg-open "' + assets.decrypt_asset(self.material_id, label.asset, self.config['password']) + '"')
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'لا يوجد صورة لعرضها')
        elif msg.clickedButton() == btn_export:
            if label.asset:
                path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Directory")
                if path:
                    asset_path = assets.decrypt_asset(self.material_id, label.asset, self.config['password'])
                    stern = os.path.basename(assets.get_asset_path(self.material_id, label.asset))
                    stern = stern[:len(stern) - 16]
                    copy2(asset_path, path + "/" + stern)
                    os.remove(asset_path)
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'لا يوجد صورة لتصديرها')

    # #### #### #### #### material_branch #### #### #### ####

    def material_code_key_press_event(self):
        mat_result = database.db.get_material_product_by_code("material", self.m_b_code.text())

        if len(mat_result) == 0:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
            return
        elif len(mat_result) == 1:
            material = mat_result[0]
        else:
            dlg = PrintDialog("material", self.m_b_code.text())
            dlg.exec()
            if dlg.result_value:
                material = dlg.result_value
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
                return

        self.m_b_code.setText(material['code'])
        self.m_b_name.setText(material['name'])

    def save_material_branch_info(self):
        material_branch = dict()
        material_branch['b_id'] = [k for k, v in self.branch_codes.items() if v == self.m_branches.currentText()][0]
        material_branch['m_id'] = [k for k, v in self.material_codes.items() if v == self.m_b_code.text()][0]
        material_branch['quantity'] = self.m_quantity.text()
        material_branch['place'] = self.m_place.text()

        return material_branch

    def create_new_material_branch(self):
        material_branch = self.save_material_branch_info()

        if material_branch['b_id'] and material_branch['m_id']:
            if int(database.db.count_quantity_branch("available_m", material_branch['b_id'], material_branch['m_id'])) == 0:
                material_branch['id'] = str(uuid8())
                database.db.insert_row("available_m", material_branch)
                self.update_material_branch_table()
                self.clear_material_branch_inputs()
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'المادة موجودة في المستودع')
        else:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'يجب أن تختار كود مادة')

    def update_material_branch(self):
        material_branch = self.save_material_branch_info()
        material_branch['id'] = self.material_branch_id
        database.db.update_row("available_m", material_branch)
        self.update_material_branch_table()
        self.clear_material_branch_inputs()

    def delete_material_branch(self):
        msg = QtWidgets.QMessageBox()
        button_reply = msg.question(self, 'تأكيد', f"هل أنت متأكد من حذف هذا القيد ؟ ", msg.Yes | msg.No, msg.No)

        if button_reply == msg.Yes:
            database.db.delete_row("available_m", self.material_branch_id)
            self.update_material_branch_table()
            self.clear_material_branch_inputs()
            toaster_Notify.QToaster.show_message(parent=self, message=f"حذف قيد تواجد مادة \nتم حذف تواجد المادة من الفرع بنجاح")

        self.btn_delete_material_branch.setEnabled(False)

    def search_material_branch_save(self):
        fil = {}
        if self.m_b_code_search.text():
            fil['m_code'] = self.m_b_code_search.text()
        if self.m_branches_search.currentIndex() != 0:
            fil['b_id'] = [k for k, v in self.branch_codes.items() if v == self.m_branches_search.currentText()][0]

        return fil

    def update_material_branch_table(self):
        fil = self.search_material_branch_save()
        rows = database.db.query_all_material_branch("available_m", fil, self.material_branch_page_size.value() * (self.material_branch_page_num.value() - 1), self.material_branch_page_size.value())

        self.material_branch_table.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            self.material_branch_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(
                str(row_idx + 1 + (self.material_branch_page_size.value() * (self.material_branch_page_num.value() - 1)))))
            self.material_branch_table.item(row_idx, 0).id = row['id']
            self.material_branch_table.item(row_idx, 0).setTextAlignment(QtCore.Qt.AlignCenter)
            self.material_branch_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(self.branch_codes[row['b_id']])))
            self.material_branch_table.item(row_idx, 1).setTextAlignment(QtCore.Qt.AlignCenter)
            self.material_branch_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(self.material_codes[row['m_id']]))
            self.material_branch_table.item(row_idx, 2).setTextAlignment(QtCore.Qt.AlignCenter)
            name = database.db.get_all_by_code("material", self.material_codes[row['m_id']])['name']
            self.material_branch_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(name))
            self.material_branch_table.item(row_idx, 3).setTextAlignment(QtCore.Qt.AlignCenter)
            self.material_branch_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(row['quantity']))
            self.material_branch_table.item(row_idx, 4).setTextAlignment(QtCore.Qt.AlignCenter)
            self.material_branch_table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(row['place']))
            self.material_branch_table.item(row_idx, 5).setTextAlignment(QtCore.Qt.AlignCenter)
        self.material_branch_table.resizeColumnsToContents()

    def clear_material_branch_inputs(self):
        self.m_branches.setCurrentIndex(0)
        self.m_b_code.clear()
        self.m_b_code.setEnabled(True)
        self.m_b_name.clear()
        self.m_quantity.setText('0')
        self.m_place.clear()

        self.btn_edit_material_branch.setEnabled(False)
        self.btn_delete_material_branch.setEnabled(False)
        self.btn_add_material_branch.setEnabled(True)

    def one_click_m_b(self, id):
        self.material_branch_id = id
        self.btn_delete_material_branch.setEnabled(True)

    def fill_material_branch_info(self, id):
        self.material_branch_id = id
        material_branch = database.db.query_row("available_m", id)
        if self.branch_codes[material_branch['b_id']] != USER:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'لا يمكن تعديل تواجد المادة لفرع آخر')
            return

        self.m_b_code.setText(self.material_codes[material_branch['m_id']])
        self.m_b_code.setEnabled(False)
        self.m_b_name.setText(database.db.get_all_by_code("material", self.material_codes[material_branch['m_id']])['name'])
        self.m_quantity.setText(material_branch['quantity'])
        self.m_place.setText(material_branch['place'])

        self.btn_add_material_branch.setEnabled(False)
        self.btn_edit_material_branch.setEnabled(True)
        self.btn_delete_material_branch.setEnabled(True)

    def print_table_material_branch(self):
        pass

    # ##############################################################################
    # product
    def setup_controls_product(self):
        self.tabWidget_2.setCurrentIndex(0)
        self.p_code.setValidator(self.validator_code)
        self.p_code_search.setValidator(self.validator_code)
        self.p_price.setValidator(self.validator_money)
        self.p_cost.setValidator(self.validator_money)

        self._typing_timer_p.setSingleShot(True)
        self._typing_timer_p.timeout.connect(self.update_product_table)

        # search
        self.p_code_search.textChanged.connect(lambda text: self._typing_timer_p.start(1000))
        self.p_name_search.textChanged.connect(lambda text: self._typing_timer_p.start(1000))

        self.p_page_num.setRange(1, math.ceil(int(database.db.count_row("product", 1)) / self.p_page_size.value()))
        self.p_page_num.valueChanged.connect(lambda text: self._typing_timer_p.start(1000))
        self.p_page_size.valueChanged.connect(lambda: self.change_page_size('product'))

        # table
        self.p_table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.p_table.doubleClicked.connect(lambda mi: self.double_click(self.p_table.item(mi.row(), 0).id))
        self.p_table.clicked.connect(lambda mi: self.one_click_p(self.p_table.item(mi.row(), 0).id))

        # btn
        self.btn_edit_show_product.clicked.connect(self.edit_show_product)
        self.btn_delete_product.clicked.connect(self.delete_product)

        self.btn_save_product.clicked.connect(self.save_product_info)
        self.btn_clear_product.clicked.connect(self.clear_product_inputs)

        # print and to exel
        self.btn_print_table_p.clicked.connect(self.print_table_product)
        self.btn_to_exel_p.clicked.connect(lambda: self.to_excel(self.p_table))

        # pages
        self.p_post.clicked.connect(lambda: self.p_page_num.setValue(self.p_page_num.value() + 1))
        self.p_previous.clicked.connect(lambda: self.p_page_num.setValue(self.p_page_num.value() - 1))
        self.p_last.clicked.connect(lambda: self.p_page_num.setValue(
            math.ceil(int(database.db.count_row("product", 1)) / self.p_page_size.value())))
        self.p_first.clicked.connect(lambda: self.p_page_num.setValue(1))

        delegate = ReadOnlyDelegate(self.p_mat_table)
        self.p_mat_table.setItemDelegateForColumn(1, delegate)
        self.p_mat_table.setItemDelegateForColumn(3, delegate)
        self.p_mat_table.setItemDelegateForColumn(4, delegate)
        self.p_mat_table.setItemDelegateForColumn(5, delegate)
        self.p_mat_table.setRowCount(1)

        self.p_mat_table.keyReleaseEvent = self.table_key_press_event

        self.update_product_table()
        self.clear_product_inputs()

        # ###############################################################
        # product_available

        self.p_quantity.setValidator(self.validator_int)
        self.p_b_code_search.setValidator(self.validator_code)

        self.product_codes = database.db.query_csp("product")

        if PERMISSION == '1':
            self.p_branches.setEnabled(True)
            self.p_branches.addItem('')
            self.p_branches.addItems(self.branch_codes.values())
        else:
            self.p_branches.addItem(USER)
            self.p_branches.setEnabled(False)

        self.p_branches_search.addItem('')
        self.p_branches_search.addItems(self.branch_codes.values())

        self.p_b_code.returnPressed.connect(self.product_code_key_press_event)

        self._typing_timer_p_b.setSingleShot(True)
        self._typing_timer_p_b.timeout.connect(self.update_product_branch_table)

        self.p_b_code_search.textChanged.connect(lambda text: self._typing_timer_p_b.start(1000))
        self.p_branches_search.currentTextChanged.connect(lambda text: self._typing_timer_p_b.start(1000))

        self.product_branch_page_num.setRange(1, math.ceil(int(database.db.count_row("available_p", 1)) / self.product_branch_page_size.value()))
        self.product_branch_page_num.valueChanged.connect(lambda text: self._typing_timer_p_b.start(1000))
        self.product_branch_page_size.valueChanged.connect(lambda: self.change_page_size('product_branch'))

        self.product_branch_table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.product_branch_table.doubleClicked.connect(lambda mi: self.fill_product_branch_info(self.product_branch_table.item(mi.row(), 0).id))
        self.product_branch_table.clicked.connect(lambda mi: self.one_click_p_b(self.product_branch_table.item(mi.row(), 0).id))

        # btn
        self.btn_add_product_branch.clicked.connect(self.create_new_product_branch)
        self.btn_edit_product_branch.clicked.connect(self.update_product_branch)
        self.btn_delete_product_branch.clicked.connect(self.delete_product_branch)
        self.btn_clear_product_branch.clicked.connect(self.clear_product_branch_inputs)

        # print and to exel
        self.btn_print_table_product_branch.clicked.connect(self.print_table_product_branch)
        self.btn_to_exel_product_branch.clicked.connect(lambda: self.to_excel(self.product_branch_table))

        # pages
        self.product_branch_post.clicked.connect(lambda: self.product_branch_page_num.setValue(self.product_branch_page_num.value() + 1))
        self.product_branch_previous.clicked.connect(lambda: self.product_branch_page_num.setValue(self.product_branch_page_num.value() - 1))
        self.product_branch_last.clicked.connect(lambda: self.product_branch_page_num.setValue(math.ceil(int(database.db.count_row("available_p", 1)) / self.product_branch_page_size.value())))
        self.product_branch_first.clicked.connect(lambda: self.product_branch_page_num.setValue(1))

        self.update_product_branch_table()
        self.clear_product_branch_inputs()

        # ###############################################################
        # product_manufacturing
        self.p_code_manufact.setValidator(self.validator_code)
        self.p_code_manufact.returnPressed.connect(self.manufact_key_press_event)

        self.quantity_manufact.setValidator(self.validator_int)
        self.quantity_manufact.textChanged.connect(self.manufact_change_quantity)

        self.btn_manufact.clicked.connect(self.manufact_new_product)
        self.btn_clear_manufact.clicked.connect(self.clear_product_manufact)

    def search_product_save(self):
        fil = {}
        if self.p_code_search.text():
            fil['code'] = self.p_code_search.text()
        if self.p_name_search.text():
            fil['name'] = self.p_name_search.text()

        return fil

    def save_product_info(self):
        product = dict()
        if self.product_id == 0:
            product['id'] = str(uuid8())
        else:
            product['id'] = self.product_id

        product['code'] = self.p_code.text()
        product['name'] = self.p_name.text()
        product['description'] = self.p_description.text()
        product['price'] = self.m_price.text()
        product['cost'] = self.p_cost.text()

        orders = []
        for idx in range(self.p_mat_table.rowCount()):
            order = dict()
            order['p_id'] = self.product_id
            if self.p_mat_table.item(idx, 0) and self.p_mat_table.item(idx, 0).text():
                if hasattr(self.p_mat_table.item(idx, 0), 'id'):
                    order['id'] = self.p_mat_table.item(idx, 0).id
                    order['m_id'] = self.p_mat_table.item(idx, 0).mid
                else:
                    order['id'] = str(uuid8())
                    order['m_id'] = database.db.get_id_by_code('material', self.p_mat_table.item(idx, 0).text())
                if self.p_mat_table.item(idx, 2) and self.p_mat_table.item(idx, 2).text():
                    order['quantity'] = self.p_mat_table.item(idx, 2).text()

                orders.append(order)

        if int(database.db.count_row("product", product['code'])) == 0:
            database.db.insert_row("product", product)
        else:
            database.db.update_row("product", product)

        database.db.insert_table('product_material', orders, self.product_id)

        self.update_product_table()
        self.clear_product_inputs()
        self.tabWidget_2.setCurrentIndex(0)

    def update_product_table(self):
        fil = self.search_product_save()
        rows = database.db.query_all_product(fil, self.p_page_size.value() * (self.p_page_num.value() - 1),
                                             self.p_page_size.value())
        self.p_table.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            self.p_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(
                str(row_idx + 1 + (self.p_page_size.value() * (self.p_page_num.value() - 1)))))
            self.p_table.item(row_idx, 0).id = row['id']
            self.p_table.item(row_idx, 0).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.p_table.item(row_idx, 1).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(row['name']))
            self.p_table.item(row_idx, 2).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(row['description']))
            self.p_table.item(row_idx, 3).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(row['price']))
            self.p_table.item(row_idx, 4).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(row['cost']))
            self.p_table.item(row_idx, 5).setTextAlignment(QtCore.Qt.AlignCenter)
        self.p_table.resizeColumnsToContents()

    def table_key_press_event(self, event: QtGui.QKeyEvent):
        self.p_mat_table: QtWidgets.QTableWidget
        if event.key() == QtCore.Qt.Key_Return:
            if self.p_mat_table.currentColumn() == 0 and self.p_mat_table.currentRow() + 1 == self.p_mat_table.rowCount():
                self.update_product_material_table(self.p_mat_table.currentRow())
                self.p_mat_table.setRowCount(self.p_mat_table.rowCount() + 1)
            elif self.p_mat_table.currentColumn() == 0 and self.p_mat_table.currentRow() + 1 != self.p_mat_table.rowCount():
                self.update_product_material_table(self.p_mat_table.currentRow())
            else:
                self.enter_event_product_material_table(self.p_mat_table.currentRow())

    def update_product_material_table(self, current_row):
        code = self.p_mat_table.item(current_row, 0).text()
        material: dict[str,str]
        mat_result = database.db.get_material_product_by_code("material", code)
        if len(mat_result) == 0:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
            self.delete_order_product(current_row)
            return
        elif len(mat_result) == 1:
            material = mat_result[0]
        else:
            dlg = PrintDialog("material", code)
            dlg.exec()
            if dlg.result_value:
                material = dlg.result_value
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
                self.delete_order_product(current_row)
                return

        for idx in range(self.p_mat_table.rowCount() - 1):
            if self.p_mat_table.item(idx, 0).text() == material['code'] and current_row != idx:
                new = int(self.p_mat_table.item(idx, 2).text()) + 1
                self.p_mat_table.setItem(idx, 2, QtWidgets.QTableWidgetItem(str(new)))
                total = new * float(self.p_mat_table.item(idx, 3).text())
                total = round(total, 2)
                self.p_mat_table.setItem(idx, 4, QtWidgets.QTableWidgetItem(str(total)))
                self.delete_order_product(current_row)
                self.calculate_total_product()
                return

        self.p_mat_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(material['code']))
        self.p_mat_table.item(current_row, 0).pid = material['id']
        self.p_mat_table.setItem(current_row, 1, QtWidgets.QTableWidgetItem(material['name']))
        self.p_mat_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))
        self.p_mat_table.setItem(current_row, 3, QtWidgets.QTableWidgetItem(str(material['price'])))
        total = float(self.p_mat_table.item(current_row, 2).text()) * float(self.p_mat_table.item(current_row, 3).text())
        total = round(total, 2)
        self.p_mat_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem(str(total)))

        btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
        btn_delete.clicked.connect(lambda: self.delete_order_product(self.p_mat_table.currentRow()))
        self.p_mat_table.setCellWidget(current_row, 5, btn_delete)
        self.calculate_total_product()

    def enter_event_product_material_table(self, current_row):
        if self.p_mat_table.item(current_row, 2).text() == '':
            self.p_mat_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))

        quantity = int(self.p_mat_table.item(current_row, 2).text())

        total = quantity * float(self.p_mat_table.item(current_row, 3).text())
        total = round(total, 2)
        self.p_mat_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem(str(total)))
        self.calculate_total_product()

    def clear_product_inputs(self):
        self.product_id = 0
        self.product_co = 0
        self.p_code.clear()
        self.p_code.setFocus()
        self.p_name.clear()
        self.p_description.clear()
        self.p_price.setText('0')
        self.p_cost.setText('0')

        self.p_mat_table.clearContents()
        self.p_mat_table.setRowCount(1)

        self.btn_edit_show_product.setEnabled(False)
        self.btn_delete_product.setEnabled(False)

    def delete_product(self):
        product = database.db.query_row("product", self.product_id)
        msg = QtWidgets.QMessageBox()
        button_reply = msg.question(self, 'تأكيد', f"هل أنت متأكد من حذف {product['name']} ؟ ",
                                    msg.Yes | msg.No,
                                    msg.No)
        if button_reply == msg.Yes:
            database.db.delete_row("product", self.product_id)
            self.update_product_table()
            toaster_Notify.QToaster.show_message(parent=self, message=f"حذف منتج\nتم حذف المنتج{product['name']} بنجاح")

    def edit_show_product(self):
        self.tabWidget_2.setCurrentIndex(1)
        self.fill_product_info(self.product_id)

    def one_click_p(self, id):
        self.clear_product_inputs()
        self.product_id = id
        self.btn_edit_show_product.setEnabled(True)
        self.btn_delete_product.setEnabled(True)

    def double_click(self, id):
        self.clear_product_inputs()
        self.product_id = id
        self.edit_show_product()

    def fill_product_info(self, id):
        product = database.db.query_row('product', id)

        self.p_code.setText(product['code'])
        self.p_name.setText(product['name'])
        self.p_description.setText(product['description'])
        self.p_price.setText(str(product['price']))
        self.p_cost.setText(str(product['cost']))

        orders = database.db.get_product_material('pro_mat_v', id)
        self.p_mat_table.setRowCount(len(orders) + 1)
        for row_idx, row in enumerate(orders):
            self.p_mat_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.p_mat_table.item(row_idx, 0).id = row['id']
            self.p_mat_table.item(row_idx, 0).mid = row['m_id']
            self.p_mat_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['name'])))
            self.p_mat_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(row['quantity'])))
            self.p_mat_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(row['price'])))
            total = int(row['quantity']) * float(row['price'])
            total = round(total, 2)
            self.p_mat_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(str(total)))
            btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
            btn_delete.clicked.connect(lambda: self.delete_order_product(self.p_mat_table.currentRow()))
            self.p_mat_table.setCellWidget(row_idx, 5, btn_delete)

    def delete_order_product(self, current_row):
        self.p_mat_table.removeRow(current_row)
        self.calculate_total_product()

    def calculate_total_product(self):
        total = 0
        for i in range(0, self.p_mat_table.rowCount()):
            if self.p_mat_table.item(i, 4) is not None:
                total += float(self.p_mat_table.item(i, 4).text())
        self.p_cost.setText(str(total))

    def print_table_product(self):
        fil = self.search_product_save()
        product = database.db.query_all_product(fil, 0, database.db.count_row("product", 1))
        with open('./html/product_template.html', 'r') as f:
            template = Template(f.read())
            fp = tempfile.NamedTemporaryFile(mode='w', delete=False, dir='./html/tmp/', suffix='.html')
            for idx, product in enumerate(product):
                product['idx'] = idx + 1

            html = template.render(products=product, date=time.strftime("%A, %d %B %Y %I:%M %p"))
            html = html.replace('style.css', '../style.css').replace('ph1.png', '../ph1.png')
            fp.write(html)
            fp.close()
            os.system('setsid firefox ' + fp.name + ' &')

    # #### #### #### #### product_branch #### #### #### ####

    def product_code_key_press_event(self):
        pro_result = database.db.get_material_product_by_code("product", self.p_b_code.text())
        if len(pro_result) == 0:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
            return
        elif len(pro_result) == 1:
            product = pro_result[0]
        else:
            dlg = PrintDialog("product", self.p_b_code.text())
            dlg.exec()
            if dlg.result_value:
                product = dlg.result_value
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
                return

        self.p_b_code.setText(product['code'])
        self.p_b_name.setText(product['name'])

    def save_product_branch_info(self):
        product_branch = dict()
        product_branch['b_id'] = [k for k, v in self.branch_codes.items() if v == self.p_branches.currentText()][0]
        product_branch['p_id'] = [k for k, v in self.product_codes.items() if v == self.p_b_code.text()][0]
        product_branch['quantity'] = self.p_quantity.text()
        product_branch['place'] = self.p_place.text()

        return product_branch

    def create_new_product_branch(self):
        product_branch = self.save_product_branch_info()

        if product_branch['b_id'] and product_branch['p_id']:
            if int(database.db.count_quantity_branch("available_p", product_branch['b_id'], product_branch['p_id'])) == 0:
                product_branch['id'] = str(uuid8())
                database.db.insert_row("available_p", product_branch)
                self.update_product_branch_table()
                self.clear_product_branch_inputs()
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'المنتج موجود في المستودع')
        else:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'يجب أن تختار كود منتج')

    def update_product_branch(self):
        product_branch = self.save_product_branch_info()
        product_branch['id'] = self.product_branch_id
        database.db.update_row("available_p", product_branch)
        self.update_product_branch_table()
        self.clear_product_branch_inputs()

    def delete_product_branch(self):
        msg = QtWidgets.QMessageBox()
        button_reply = msg.question(self, 'تأكيد', f"هل أنت متأكد من حذف هذا القيد ؟ ", msg.Yes | msg.No, msg.No)

        if button_reply == msg.Yes:
            database.db.delete_row("available_p", self.product_branch_id)
            self.update_product_branch_table()
            self.clear_product_branch_inputs()
            toaster_Notify.QToaster.show_message(parent=self, message=f"حذف قيد تواجد منتج \nتم حذف تواجد المنتج من الفرع بنجاح")

        self.btn_delete_product_branch.setEnabled(False)

    def search_product_branch_save(self):
        fil = {}
        if self.p_b_code_search.text():
            fil['m_code'] = self.p_b_code_search.text()
        if self.p_branches_search.currentIndex() != 0:
            fil['b_id'] = [k for k, v in self.branch_codes.items() if v == self.p_branches_search.currentText()][0]

        return fil

    def update_product_branch_table(self):
        fil = self.search_product_branch_save()
        rows = database.db.query_all_material_branch("available_p", fil, self.material_branch_page_size.value() * (self.material_branch_page_num.value() - 1), self.material_branch_page_size.value())
        self.product_branch_table.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            self.product_branch_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(
                str(row_idx + 1 + (self.material_branch_page_size.value() * (self.material_branch_page_num.value() - 1)))))
            self.product_branch_table.item(row_idx, 0).id = row['id']
            self.product_branch_table.item(row_idx, 0).setTextAlignment(QtCore.Qt.AlignCenter)
            self.product_branch_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(self.branch_codes[row['b_id']])))
            self.product_branch_table.item(row_idx, 1).setTextAlignment(QtCore.Qt.AlignCenter)
            self.product_branch_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(self.product_codes[row['p_id']]))
            self.product_branch_table.item(row_idx, 2).setTextAlignment(QtCore.Qt.AlignCenter)
            name = database.db.get_all_by_code("product", self.product_codes[row['p_id']])['name']
            self.product_branch_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(name))
            self.product_branch_table.item(row_idx, 3).setTextAlignment(QtCore.Qt.AlignCenter)
            self.product_branch_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(row['quantity']))
            self.product_branch_table.item(row_idx, 4).setTextAlignment(QtCore.Qt.AlignCenter)
            self.product_branch_table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(row['place']))
            self.product_branch_table.item(row_idx, 5).setTextAlignment(QtCore.Qt.AlignCenter)
        self.product_branch_table.resizeColumnsToContents()

    def clear_product_branch_inputs(self):
        self.p_branches.setCurrentIndex(0)
        self.p_b_code.clear()
        self.p_b_code.setEnabled(True)
        self.p_b_name.clear()
        self.p_quantity.setText('0')
        self.p_place.clear()

        self.btn_edit_product_branch.setEnabled(False)
        self.btn_delete_product_branch.setEnabled(False)
        self.btn_add_product_branch.setEnabled(True)

    def one_click_p_b(self, id):
        self.product_branch_id = id
        self.btn_delete_product_branch.setEnabled(True)

    def fill_product_branch_info(self, id):
        self.product_branch_id = id
        product_branch = database.db.query_row("available_p", id)
        if self.branch_codes[product_branch['b_id']] != USER:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'لا يمكن تعديل تواجد المنتج لفرع آخر')
            return

        self.p_b_code.setText(self.product_codes[product_branch['p_id']])
        self.p_b_code.setEnabled(False)
        self.p_b_name.setText(database.db.get_all_by_code("product", self.product_codes[product_branch['p_id']])['name'])
        self.p_quantity.setText(product_branch['quantity'])
        self.p_place.setText(product_branch['place'])

        self.btn_add_product_branch.setEnabled(False)
        self.btn_edit_product_branch.setEnabled(True)
        self.btn_delete_product_branch.setEnabled(True)

    def print_table_product_branch(self):
        pass

    # #### #### #### #### product_manufact #### #### #### ####

    def manufact_key_press_event(self):
        code = self.p_code_manufact.text()
        product: dict
        product_result = database.db.get_material_product_by_code("product", code)
        if len(product_result) == 0:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
            self.clear_product_manufact()
            return
        elif len(product_result) == 1:
            product = product_result[0]
        else:
            dlg = PrintDialog("product", code)
            dlg.exec()
            if dlg.result_value:
                product = dlg.result_value
            else:
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')
                self.clear_product_manufact()
                return

        self.quantity_manufact.setEnabled(True)
        self.quantity_manufact.setText('1')
        manufact = True

        self.p_code_manufact.setText(product['code'])
        self.p_name_manufact.setText(product['name'])
        self.p_description_manufact.setText(product['description'])
        self.p_price_manufact.setText(product['price'])
        self.p_cost_manufact.setText(product['cost'])
        self.p_price_manufact_2.setText(product['price'])
        self.p_cost_manufact_2.setText(product['cost'])

        orders = database.db.get_product_material('pro_mat_v', product['id'])
        self.p_mat_table_manufact.setRowCount(len(orders))
        for row_idx, row in enumerate(orders):
            self.p_mat_table_manufact.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.p_mat_table_manufact.item(row_idx, 0).id = row['id']
            self.p_mat_table_manufact.item(row_idx, 0).mid = row['m_id']
            self.p_mat_table_manufact.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['name'])))
            self.p_mat_table_manufact.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(row['quantity'])))
            self.p_mat_table_manufact.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(row['price'])))
            total = int(row['quantity']) * float(row['price'])
            total = round(total, 2)
            self.p_mat_table_manufact.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(str(total)))
            self.p_mat_table_manufact.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(str(row['quantity'])))
            self.p_mat_table_manufact.setItem(row_idx, 6, QtWidgets.QTableWidgetItem(str(total)))
            qu = database.db.get_id_by_mid("available_m", row['m_id'], self.branch_id)
            if qu:
                if qu['quantity'] != '0':
                    if int(qu['quantity']) >= int(row['quantity']):
                        quantity = int(qu['quantity']) - int(row['quantity'])
                        self.p_mat_table_manufact.setItem(row_idx, 7, QtWidgets.QTableWidgetItem(str(quantity)))
                    else:
                        self.p_mat_table_manufact.setItem(row_idx, 7, QtWidgets.QTableWidgetItem("الكمية الموجودة لا تكفي"))
                        manufact = False
                else:
                    self.p_mat_table_manufact.setItem(row_idx, 7, QtWidgets.QTableWidgetItem("لا توجد هذه المادة في المستودع"))
                    manufact = False
            else:
                self.p_mat_table_manufact.setItem(row_idx, 7, QtWidgets.QTableWidgetItem("لا توجد هذه المادة في المستودع"))
                manufact = False

        self.btn_manufact.setEnabled(manufact)

    def manufact_change_quantity(self):
        if self.quantity_manufact.text() == '':
            product_quantity = 1
        else:
            product_quantity = int(self.quantity_manufact.text())
        manufact = True
        price = float(self.p_price_manufact.text())
        price = round(price, 2)
        cost = float(self.p_price_manufact.text())
        cost = round(cost, 2)
        self.p_price_manufact_2.setText(str(price * product_quantity))
        self.p_cost_manufact_2.setText(str(cost * product_quantity))

        for row_idx in range(self.p_mat_table_manufact.rowCount()):
            quantity = int(self.p_mat_table_manufact.item(row_idx, 2).text()) * product_quantity
            self.p_mat_table_manufact.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(str(quantity)))
            qu = database.db.get_id_by_mid("available_m", self.p_mat_table_manufact.item(row_idx, 0).mid, self.branch_id)
            if qu:
                if qu['quantity'] != '0':
                    if int(qu['quantity']) >= quantity:
                        self.p_mat_table_manufact.setItem(row_idx, 7, QtWidgets.QTableWidgetItem(str(int(qu['quantity']) - quantity)))
                    else:
                        self.p_mat_table_manufact.setItem(row_idx, 7, QtWidgets.QTableWidgetItem("الكمية الموجودة لا تكفي"))
                        manufact = False
                else:
                    self.p_mat_table_manufact.setItem(row_idx, 7, QtWidgets.QTableWidgetItem("لا توجد هذه المادة في المستودع"))
                    manufact = False
            else:
                self.p_mat_table_manufact.setItem(row_idx, 7, QtWidgets.QTableWidgetItem("لا توجد هذه المادة في المستودع"))
                manufact = False
            total = float(self.p_mat_table_manufact.item(row_idx, 3).text()) * float(self.p_mat_table_manufact.item(row_idx, 5).text())
            total = round(total, 2)
            self.p_mat_table_manufact.setItem(row_idx, 6, QtWidgets.QTableWidgetItem(str(total)))

        self.btn_manufact.setEnabled(manufact)

    def manufact_new_product(self):
        product = dict()
        product['p_id'] = [k for k, v in self.product_codes.items() if v == self.p_code_manufact.text()][0]
        product['quantity'] = self.quantity_manufact.text()
        product['b_id'] = self.branch_id
        if int(database.db.count_quantity_branch("available_p", product['b_id'], product['p_id'])) == 0:
            product['id'] = str(uuid8())
            database.db.insert_row("available_p", product)
        else:
            database.db.update_product_quantity("available_p", product)

        material = dict()
        material['b_id'] = self.branch_id
        for row_idx in range(self.p_mat_table_manufact.rowCount()):
            material['m_id'] = self.p_mat_table_manufact.item(row_idx, 0).mid
            material['quantity'] = self.p_mat_table_manufact.item(row_idx, 5).text()
            database.db.update_product_quantity("available_m", material)

        self.clear_product_manufact()
        self.update_product_branch_table()
        self.update_material_branch_table()

    def clear_product_manufact(self):
        self.p_code_manufact.clear()
        self.quantity_manufact.setText('1')
        self.p_name_manufact.clear()
        self.p_description_manufact.clear()
        self.p_price_manufact.setText('0')
        self.p_cost_manufact.setText('0')

        self.p_price_manufact_2.setText('0')
        self.p_cost_manufact_2.setText('0')

        self.p_mat_table_manufact.clearContents()
        self.p_mat_table_manufact.setRowCount(0)

        self.btn_manufact.setEnabled(False)

    # ###################################################################################
    # requests methods
    def setup_controls_requests(self):
        self.req_code.setValidator(self.validator_code)
        self._typing_timer_req.timeout.connect(self.update_requests_table)

        self.req_table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.req_table.doubleClicked.connect(lambda mi: self.double_click_req(self.req_table.item(mi.row(), 0).id))
        self.req_table.clicked.connect(lambda mi: self.one_click_req(self.req_table.item(mi.row(), 0).id))
        self.req_page_num.setRange(1, math.ceil(int(database.db.count_row("requests", 1)) / self.req_page_size.value()))

        self.req_code.textChanged.connect(lambda text: self._typing_timer_req.start(1000))
        self.req_brunch.currentTextChanged.connect(lambda text: self._typing_timer_req.start(1000))

        self.ch_req_date_from.toggled.connect(lambda: self.check_date_from('requests'))
        self.ch_req_date_to.toggled.connect(lambda: self.check_date_to('requests'))

        self.req_page_num.valueChanged.connect(lambda text: self._typing_timer_req.start(1000))
        self.req_page_size.valueChanged.connect(lambda: self.change_page_size('requests'))

        # print and to exel bill
        # self.btn_print_table_req.clicked.connect(self.print_table_req)
        # self.btn_to_exel_req.clicked.connect(lambda: self.to_excel(self.req_table))

        # pages
        self.req_post.clicked.connect(lambda: self.req_page_num.setValue(self.req_page_num.value() + 1))
        self.req_previous.clicked.connect(lambda: self.req_page_num.setValue(self.req_page_num.value() - 1))
        self.req_last.clicked.connect(lambda: self.req_page_num.setValue(
            math.ceil(int(database.db.count_row("requests", 1)) / self.page_size_requests)))
        self.req_first.clicked.connect(lambda: self.req_page_num.setValue(1))

        self.btn_add_req.clicked.connect(lambda: self.open_requests(0))
        self.btn_edit_req.clicked.connect(lambda: self.open_requests(self.requests_id))

        self.req_date_from.setSpecialValueText(' ')
        self.req_date_to.setSpecialValueText(' ')

        self.req_date_from.setDate(QDate.currentDate())
        self.req_date_to.setDate(QDate.currentDate())

        self.btn_edit_req.setEnabled(False)

        self.update_requests_table()

    def one_click_req(self, id):
        self.requests_id = id
        self.btn_edit_req.setEnabled(True)

    def double_click_req(self, id):
        self.requests_id = id
        self.open_requests(id)

    def open_requests(self, id):
        req = Requests(id)
        req.setWindowIcon(QtGui.QIcon('emp.png'))
        req.exec()
        self.btn_edit_req.setEnabled(False)
        self.update_requests_table()
        self.update_material_table()

    def search_requests_save(self):
        fil = {}
        if self.req_code.text():
            fil['code'] = self.req_code.text()
        if self.req_brunch.currentText() != '':
            fil['branch_id'] = [k for k, v in self.branches.items() if v == self.req_brunch.currentText()][0]
        if self.ch_req_date_from.isChecked():
            fil['date_from'] = QDate.toString(self.req_date_from.date())
            if self.ch_req_date_to.isChecked():
                fil['date_to'] = QDate.toString(self.req_date_to.date())

        return fil

    def update_requests_table(self):
        fil = self.search_requests_save()
        rows = database.db.query_all_bill("requests", fil, self.req_page_size.value() * (self.req_page_num.value() - 1),
                                          self.req_page_size.value())
        self.req_table.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            self.req_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(
                str(row_idx + 1 + (self.page_size_requests * (self.req_page_num.value() - 1)))))
            self.req_table.item(row_idx, 0).id = row['id']
            self.req_table.item(row_idx, 0).setTextAlignment(QtCore.Qt.AlignCenter)
            self.req_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.req_table.item(row_idx, 1).setTextAlignment(QtCore.Qt.AlignCenter)
            row['branch_id'] = self.branches[row['branch_id']]
            self.req_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(row['branch_id']))
            self.req_table.item(row_idx, 2).setTextAlignment(QtCore.Qt.AlignCenter)
            self.req_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(row['total_req'])))
            self.req_table.item(row_idx, 3).setTextAlignment(QtCore.Qt.AlignCenter)
            self.req_table.setItem(row_idx, 4, QtWidgets.QTableWidgetItem(str(row['total_buy'])))
            self.req_table.item(row_idx, 4).setTextAlignment(QtCore.Qt.AlignCenter)
            self.req_table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(row['date']))
            self.req_table.item(row_idx, 5).setTextAlignment(QtCore.Qt.AlignCenter)
        self.req_table.resizeColumnsToContents()
        # self.update_notification()

    # def print_table_bill_buy(self):
    #     fil = self.search_bill_sell_save()
    #     bills = database.db.query_all_bill("bill_buy", fil, 0, database.db.count_row("bill_buy", 1))
    #     with open('./html/bill_buy_template.html', 'r') as f:
    #         template = Template(f.read())
    #         fp = tempfile.NamedTemporaryFile(mode='w', delete=False, dir='./html/tmp/', suffix='.html')
    #         for idx, bill in enumerate(bills):
    #             bill['idx'] = idx + 1
    #             bill['s_id'] = self.suppliers[bill['s_id']]
    #             bill['total'] = str(float(bill['total']) - float(bill['discount']))
    #             if bill['ispaid'] == '1':
    #                 bill['ispaid'] = 'مدفوعة'
    #             else:
    #                 bill['ispaid'] = 'غير مدفوعة'
    #         html = template.render(bills=bills, date=time.strftime("%A, %d %B %Y %I:%M %p"))
    #         html = html.replace('style.css', '../style.css').replace('ph1.png', '../ph1.png')
    #         fp.write(html)
    #         fp.close()
    #         os.system('setsid firefox ' + fp.name + ' &')

    ####################################################################


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    locale.setlocale(locale.LC_ALL, "en_US.utf8")
    QtCore.QLocale.setDefault(QtCore.QLocale(QtCore.QLocale.English))

    mainWindow = AppMainWindow()
    mainWindow.show()
    mainWindow.setWindowIcon(QtGui.QIcon('icons/ph1.png'))
    os.system("mkdir -p /tmp/prisons")
    exit_code = app.exec_()
    os.system("rm -r /tmp/prisons")
    for filename in glob.glob("html/tmp/*"):
        os.remove(filename)
