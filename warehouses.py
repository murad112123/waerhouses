#!/usr/bin/env python3

import glob
import json
import locale
import math
import os
import sys
import tempfile
import time
from uuid6 import uuid8
import PyQt5.QtCore as QtCore
import PyQt5.QtGui as QtGui
import PyQt5.QtWidgets as QtWidgets
from jinja2 import Template
import xlwt
import PyQt5.uic as uic
import hashlib
import toaster_Notify
from QDate import QDate

import aes
import database

Form_Main, _ = uic.loadUiType('warehouses.ui')
Form_Requests, _ = uic.loadUiType('requests.ui')
PAGE_SIZE = 10
USER = ''
PASS = ''


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
        material = database.db.get_material_by_code(code)
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
        material = database.db.get_material_by_code(code)
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


class AppMainWindow(QtWidgets.QMainWindow, Form_Main):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Form_Main.__init__(self)
        self.setupUi(self)

        self.p_id = None

        self.config = None

        self.validator_code = QtGui.QRegExpValidator(QtCore.QRegExp('[\u0621-\u064A0-9a-zA-Z][0-9]*'))
        self.validator_int = QtGui.QRegExpValidator(QtCore.QRegExp('[0-9]+'))
        self.validator_money = QtGui.QRegExpValidator(
            QtCore.QRegExp('^(\$)?(([1-9]\d{0,2}(\,\d{3})*)|([1-9]\d*)|(0))(\.\d{1,2})?$'))

        self.branch_id = 0
        self.branch_codes = None

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

        self.page_size_product = PAGE_SIZE

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
            self.material_branch_page_num.setRange(1, math.ceil(int(database.db.count_row("available_m", 1)) / self.material_branch_page_size.value()))
            self._typing_timer_m.start(1000)
        elif table == 'product':
            self.page_size_product = self.p_page_size.value()
            self.p_page_num.setRange(1, math.ceil(int(database.db.count_row("product", 1)) / self.page_size_product))
            self._typing_timer_p.start(1000)
        elif table == 'requests':
            self.page_size_requests = self.req_page_siz.value()
            self.req_page_num.setRange(1, math.ceil(int(database.db.count_row("requests", 1)) / self.page_size_requests))
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

        self.m_code_b.setText(USER)

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

        self.m_branches.addItem('')
        self.m_branches.addItems(self.branch_codes.values())

        self.m_branches_search.addItem('')
        self.m_branches_search.addItem('all')
        self.m_branches_search.addItems(self.branch_codes.values())

        self.m_b_code.addItem('')
        self.m_b_code.addItems(self.material_codes.values())

        self.m_b_code.currentTextChanged.connect(self.m_b_code_change)

        self._typing_timer_m_b.setSingleShot(True)
        self._typing_timer_m_b.timeout.connect(self.update_material_branch_table)

        self.m_b_code_search.textChanged.connect(lambda text: self._typing_timer_m_b.start(1000))
        self.m_branches_search.currentTextChanged.connect(lambda text: self._typing_timer_m_b.start(1000))

        self.material_branch_page_num.setRange(1, math.ceil(int(database.db.count_row("available_m", 1)) / self.material_branch_page_size.value()))
        self.material_branch_page_num.valueChanged.connect(lambda text: self._typing_timer_m_b.start(1000))
        self.material_branch_page_size.valueChanged.connect(lambda: self.change_page_size('material_branch'))

        self.material_branch_table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.material_branch_table.doubleClicked.connect(lambda mi: self.fill_material_branch_info(self.material_branch_table.item(mi.row(), 0).id))
        self.material_branch_table.clicked.connect(lambda mi: self.one_click_m_b(self.material_branch_table.item(mi.row(), 0).id))

        # btn
        self.btn_add_material_branch.clicked.connect(self.create_new_material_branch)
        self.btn_edit_material_branch.clicked.connect(self.update_material_branch)
        self.btn_delete_material_branch.clicked.connect(self.delete_material_branch)
        self.btn_clear_material_branch.clicked.connect(self.clear_material_branch_inputs)

        # print and to exel
        self.btn_print_table_material_branch.clicked.connect(self.print_table_material_branch)
        self.btn_to_exel_material_branch.clicked.connect(lambda: self.to_excel(self.m_table))

        # pages
        self.material_branch_post.clicked.connect(lambda: self.material_branch_page_num.setValue(self.material_branch_page_num.value() + 1))
        self.material_branch_previous.clicked.connect(lambda: self.material_branch_page_num.setValue(self.material_branch_page_num.value() - 1))
        self.material_branch_last.clicked.connect(lambda: self.material_branch_page_num.setValue(math.ceil(int(database.db.count_row("available_m", 1)) / self.material_branch_page_size.value())))
        self.material_branch_first.clicked.connect(lambda: self.material_branch_page_num.setValue(1))

        self.update_material_table()
        self.clear_material_inputs()

    def m_b_code_change(self):
        if self.m_b_code.currentIndex() == 0:
            self.m_b_name.clear()
        else:
            self.m_b_name.setText(database.db.get_material_by_code(self.m_b_code.currentText())['name'])

    def save_material_info(self):
        global USER
        material = dict()
        material['id'] = str(uuid8())
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

        return material

    def create_new_material(self):
        global USER
        material = self.save_material_info()
        material['code'] = self.m_code_b.text() + '.' + material['code']
        if material['code'] and material['name']:
            if int(database.db.count_row("material", material['code'])) == 0:
                database.db.insert_row("material", material)
                toaster_Notify.QToaster.show_message(parent=self,
                                                     message=f"إضافة مادة\nتم إضافة المادة {material['name']} بنجاح")
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
                database.db.update_row("material", material)
                self.update_material_table()
                self.clear_material_inputs()
                toaster_Notify.QToaster.show_message(parent=self,
                                                     message=f"تعديل مادة\nتم تعديل المادة {material['name']} بنجاح")
            elif int(database.db.count_row("material", material['code'])) == 0:
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
            self.update_material_table()
            self.clear_material_inputs()
            toaster_Notify.QToaster.show_message(parent=self,
                                                 message=f"حذف مادة\nتم حذف المادة{material['name']} بنجاح")

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

    def update_material_branch_table(self):
        pass

    def one_click_m_b(self, id):
        pass

    def fill_material_branch_info(self, id):
        pass

    def print_table_material_branch(self):
        pass
    
    # ##############################################################################
    # product
    def setup_controls_product(self):
        self.p_code.setValidator(self.validator_code)
        self.p_code_search.setValidator(self.validator_code)
        self.p_price.setValidator(self.validator_money)

        self._typing_timer_p.setSingleShot(True)
        self._typing_timer_p.timeout.connect(self.update_product_table)

        self.p_page_num.valueChanged.connect(lambda text: self._typing_timer_p.start(1000))
        self.p_page_size.valueChanged.connect(lambda: self.change_page_size('product'))

        # table
        self.m_table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.p_table.doubleClicked.connect(lambda mi: self.double_Click(self.p_table.item(mi.row(), 0).id))
        self.p_table.clicked.connect(lambda mi: self.one_click_p(self.p_table.item(mi.row(), 0).id))

        # btn
        self.btn_save_product.clicked.connect(self.save_product_info)
        self.btn_add_new_product.clicked.connect(self.add_new_product)
        self.btn_clear_product.clicked.connect(self.clear_product_inputs)
        self.btn_delete_product.clicked.connect(self.delete_product)
        self.btn_cancel_product.clicked.connect(self.cancel_product)
        self.btn_edit_show_product.clicked.connect(self.edit_show_product)

        # print and to exel
        self.btn_print_table_p.clicked.connect(self.print_table_product)
        self.btn_to_exel_p.clicked.connect(lambda: self.to_excel(self.p_table))

        # pages
        self.p_post.clicked.connect(lambda: self.p_page_num.setValue(self.p_page_num.value() + 1))
        self.p_previous.clicked.connect(lambda: self.p_page_num.setValue(self.p_page_num.value() - 1))
        self.p_last.clicked.connect(lambda: self.p_page_num.setValue(
            math.ceil(int(database.db.count_row("product", 1)) / self.page_size_product)))
        self.p_first.clicked.connect(lambda: self.p_page_num.setValue(1))

        self.p_mat_table.keyReleaseEvent = self.table_key_press_event

        self.update_product_table()
        self.clear_product_inputs()

    def search_product_save(self):
        fil = {}
        if self.p_code_search.text():
            fil['code'] = self.p_code_search.text()
        if self.p_name_search.text():
            fil['name'] = self.p_name_search.text()

        return fil

    def save_product_info(self):
        product = dict()
        product['code'] = self.p_code.text()
        product['name'] = self.p_name.text()
        product['description'] = self.p_description.text()
        product['price'] = self.m_price.text()
        product['cost'] = self.p_cost.text()

        orders = []
        for idx in range(self.p_mat_table.rowCount()):
            order = dict()
            order['p_id'] = self.p_id
            if self.p_mat_table.item(idx, 0) and self.p_mat_table.item(idx, 0).text():
                if hasattr(self.p_mat_table.item(idx, 0), 'id'):
                    order['id'] = self.p_mat_table.item(idx, 0).id
                    order['m_id'] = self.p_mat_table.item(idx, 0).pid
                else:
                    order['id'] = int(database.db.get_next_id('product_material')) + idx
                    order['m_id'] = database.db.get_id_by_code('material', self.p_mat_table.item(idx, 0).text())
                if self.p_mat_table.item(idx, 2) and self.p_mat_table.item(idx, 2).text():
                    order['quantity'] = self.p_mat_table.item(idx, 2).text()

                orders.append(order)

        if int(database.db.count_row("product", product['code'])) == 0:
            database.db.insert_row("product", product)
        else:
            database.db.update_row("product", product)

        database.db.insert_table('product', orders)
        self.accept()

    def update_product_table(self):
        fil = self.search_product_save()
        rows = database.db.query_all_product(fil, self.page_size_product * (self.p_page_num.value() - 1),
                                             self.page_size_product)
        self.p_table.setRowCount(len(rows))
        for row_idx, row in enumerate(rows):
            self.p_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(
                str(row_idx + 1 + (self.page_size_product * (self.p_page_num.value() - 1)))))
            self.p_table.item(row_idx, 0).id = row['id']
            self.p_table.item(row_idx, 0).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.p_table.item(row_idx, 1).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(row['name']))
            self.p_table.item(row_idx, 2).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(row['description']))
            self.p_table.item(row_idx, 3).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 5, QtWidgets.QTableWidgetItem(row['price']))
            self.p_table.item(row_idx, 5).setTextAlignment(QtCore.Qt.AlignCenter)
            self.p_table.setItem(row_idx, 6, QtWidgets.QTableWidgetItem(row['cost']))
            self.p_table.item(row_idx, 6).setTextAlignment(QtCore.Qt.AlignCenter)
        self.p_table.resizeColumnsToContents()

    def table_key_press_event(self, event: QtGui.QKeyEvent):
        self.p_mat_table: QtWidgets.QTableWidget
        if event.key() == QtCore.Qt.Key_Return:
            if self.p_mat_table.currentColumn() == 0 and self.p_mat_table.currentRow() + 1 == self.p_mat_table.rowCount():
                self.update_table(self.p_mat_table.currentRow())
            else:
                self.enter_event(self.p_mat_table.currentRow())

    def enter_event(self, current_row):
        code = self.p_mat_table.item(current_row, 0).text()
        material = database.db.get_material_by_code(code)

        if self.p_mat_table.item(current_row, 2).text() == '':
            self.p_mat_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))
        quantity = int(self.p_mat_table.item(current_row, 2).text())

        # total = quantity * float(self.p_mat_table.item(current_row, 3).text())
        # self.p_mat_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem(str(total)))
        # self.calculate_total()

    def update_table(self, current_row):
        code = self.p_mat_table.item(current_row, 0).text()
        material = database.db.get_material_by_code(code)
        if material:
            if int(material['quantity']) >= 1:
                for idx in range(self.p_mat_table.rowCount() - 1):
                    if self.p_mat_table.item(idx, 0).text() == code:
                        new = int(self.p_mat_table.item(idx, 2).text()) + 1
                        self.p_mat_table.setItem(idx, 2, QtWidgets.QTableWidgetItem(str(new)))
                        self.p_mat_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))
                        return

                self.p_mat_table.item(current_row, 0).pid = material['id']
                self.p_mat_table.setItem(current_row, 1, QtWidgets.QTableWidgetItem(material['name']))
                self.p_mat_table.setItem(current_row, 2, QtWidgets.QTableWidgetItem('1'))
                self.p_mat_table.setItem(current_row, 3, QtWidgets.QTableWidgetItem(str(material['price'])))

                self.p_mat_table.setItem(current_row, 4, QtWidgets.QTableWidgetItem(
                    str(float(self.p_mat_table.item(current_row, 2).text()) * float(
                        self.p_mat_table.item(current_row, 3).text()))))

                btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
                btn_delete.clicked.connect(lambda: self.delete_order(self.p_mat_table.currentRow()))
                self.p_mat_table.setCellWidget(current_row, 5, btn_delete)
                # self.calculate_total()
            else:
                self.p_mat_table.setItem(current_row, 0, QtWidgets.QTableWidgetItem(''))
                QtWidgets.QMessageBox.warning(None, 'خطأ', 'غير متوفر\n لقد انتهت كمية هذه المادة بالفعل')
        else:
            QtWidgets.QMessageBox.warning(None, 'خطأ', 'الرقم غير موجود\n أعد ادخال رقم صحيح')

    # def create_new_product(self):
    #     product = self.save_product_info()
    #     if product['code'] and product['name']:
    #         if int(database.db.count_row("product", product['code'])) == 0:
    #             database.db.insert_row("product", product)
    #             toaster_Notify.QToaster.show_message(parent=self,
    #                                                  message=f"إضافة المنتج\nتم إضافة المنتج {product['name']} بنجاح")
    #             self.tabWidget_2.setCurrentIndex(0)
    #             self.update_product_table()
    #             self.clear_product_inputs()
    #         else:
    #             QtWidgets.QMessageBox.warning(None, 'خطأ', 'إن الكود مكرر')
    #     else:
    #         QtWidgets.QMessageBox.warning(None, 'خطأ', 'يجب أن تدخل الكود واسم المنتج')

    # def update_product(self):
    #     product = self.save_product_info()
    #     product['id'] = self.product_id
    #     if product['code'] and product['name']:
    #         if product['code'] == self.product_co:
    #             database.db.update_row("product", product)
    #             self.tabWidget_2.setCurrentIndex(0)
    #             self.update_product_table()
    #             self.clear_product_inputs()
    #             toaster_Notify.QToaster.show_message(parent=self,
    #                                                  message=f"تعديل منتج\nتم تعديل المنتج {product['name']} بنجاح")
    #         elif int(database.db.count_row("product", product['code'])) == 0:
    #             database.db.update_row("product", product)
    #             self.tabWidget_2.setCurrentIndex(0)
    #             self.update_product_table()
    #             self.clear_product_inputs()
    #             toaster_Notify.QToaster.show_message(parent=self,
    #                                                  message=f"تعديل منتج\nتم تعديل المنتج {product['name']} بنجاح")
    #         else:
    #             QtWidgets.QMessageBox.warning(None, 'خطأ', 'إن الكود مكرر')
    #     else:
    #         QtWidgets.QMessageBox.warning(None, 'خطأ', 'يجب أن تدخل الكود واسم المنتج')

    def clear_product_inputs(self):
        self.product_id = 0
        self.product_co = 0
        self.p_code.clear()
        self.p_code.setFocus()
        self.p_name.clear()
        self.p_description.clear()
        self.p_price.setText('0')
        self.p_cost.setText('0')

        self.btn_delete_product.setEnabled(False)

    def delete_product(self):
        product = self.save_product_info()
        msg = QtWidgets.QMessageBox()
        button_reply = msg.question(self, 'تأكيد', f"هل أنت متأكد من حذف {product['name']} ؟ ",
                                    msg.Yes | msg.No,
                                    msg.No)
        if button_reply == msg.Yes:
            database.db.delete_row("product", self.product_id)
            self.update_product_table()
            self.clear_product_inputs()
            toaster_Notify.QToaster.show_message(parent=self,
                                                 message=f"حذف منتج\nتم حذف المنتج{product['name']} بنجاح")

    def add_new_product(self):
        self.tabWidget_2.setCurrentIndex(1)
        self.btn_clear_product.setEnabled(True)
        self.btn_cancel_product.setEnabled(True)

    def cancel_product(self):
        self.tabWidget_2.setCurrentIndex(0)
        self.clear_product_inputs()

    def edit_show_product(self):
        self.tabWidget_2.setCurrentIndex(1)
        self.fill_product_info(self.product_id)

    def one_click_p(self, id):
        self.product_id = id
        self.btn_edit_show_product.setEnabled(True)
        self.btn_delete_product.setEnabled(True)

    def double_Click(self, id):
        self.product_id = id
        self.fill_product(self.product_id)

    def fill_product(self, id):
        if id == 0:
            self.p_id = database.db.get_next_id('product')
        else:
            product = database.db.query_row('product', id)
            self.p_id = product['id']
            self.p_code.setText(product['code'])
            self.p_name.setText(product['name'])
            self.p_description.setText(product['description'])
            self.p_quantity.setText(str(product['quantity']))
            self.p_price.setText(str(product['price']))
            self.p_cost.setText(str(product['cost']))

        orders = database.db.get_product_material('pro_mat_v', self.p_id)
        self.p_mat_table.setRowCount(len(orders) + 1)
        for row_idx, row in enumerate(orders):
            self.p_mat_table.setItem(row_idx, 0, QtWidgets.QTableWidgetItem(str(row['code'])))
            self.p_mat_table.item(row_idx, 0).id = row['id']
            self.p_mat_table.item(row_idx, 0).mid = row['m_id']
            self.p_mat_table.setItem(row_idx, 1, QtWidgets.QTableWidgetItem(str(row['name'])))
            self.p_mat_table.setItem(row_idx, 2, QtWidgets.QTableWidgetItem(str(row['quantity'])))
            self.p_mat_table.setItem(row_idx, 3, QtWidgets.QTableWidgetItem(str(row['price'])))
            btn_delete = QtWidgets.QPushButton(QtGui.QIcon.fromTheme('delete'), '')
            btn_delete.clicked.connect(lambda: self.delete_order(self.p_mat_table.currentRow()))
            self.bb_table.setCellWidget(row_idx, 5, btn_delete)

    def delete_order(self, current_row):
        self.bb_table.removeRow(current_row)
        self.calculate_total()

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

    mainWindow = AppMainWindow()
    mainWindow.show()
    mainWindow.setWindowIcon(QtGui.QIcon('icons/ph1.png'))
    for filename in glob.glob("html/tmp/*"):
        os.remove(filename)
    exit_code = app.exec_()
