# -*- coding: utf-8 -*-
# ---
# @File: MergePdfTools.py
# @Author: Bob chen
# @CreateTime: Nov 18, 2021
# @Version : 1.0
# ---
import datetime

from PySide2.QtCore import *
from PySide2.QtWidgets import *
from PIL import Image
from PySide2.QtWidgets import QWidget, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, \
    QGroupBox, QListWidget, QVBoxLayout, QSpacerItem
from reportlab.lib.pagesizes import portrait
from reportlab.pdfgen import canvas
import PyPDF2
import sys
import fitz
import os

if not sys.warnoptions:
    import warnings

    warnings.simplefilter("ignore")


# pdf转换pdf操作
def get_img_files(img_path, file_type):
    dir_path = os.listdir(img_path)
    img_list = list()
    for i in dir_path:
        if file_type in i:
            img_list.append(i)
    return img_list


def convert_pdf(pdf_file):
    pdf = fitz.open(pdf_file)
    for pg in range(pdf.page_count):
        page = pdf[pg]
        zoom_x, zoom_y = 1.0, 1.0
        rotate = int(0)
        trans = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        pm.save(r'./tu' + '{:02}.png'.format(pg))

    img_list = get_img_files(img_path='./', file_type='.png')
    out_path = r'./1.pdf'
    (max_w, max_h) = Image.open(img_list[0]).size
    c = canvas.Canvas(out_path, pagesize=portrait((max_w, max_h)))
    for i in range(len(img_list)):
        c.drawImage(img_list[i], 0, 0, max_w, max_h)
        c.showPage()
        os.remove(img_list[i])
    c.save()

    return out_path


def merge_page(writer, merge_file):
    # 获取 PdfFileReader 对象
    pdf_file_reader = PyPDF2.PdfFileReader(merge_file)
    if pdf_file_reader.getIsEncrypted():
        new_pdf = convert_pdf(merge_file)
        new_inputs = PyPDF2.PdfFileReader(new_pdf)
        num_pages = new_inputs.getNumPages()
        for index in range(0, num_pages):
            page_obj = new_inputs.getPage(index)
            writer.addPage(page_obj)  # 根据每页返回的 PageObject,写入到文件
        os.remove(new_pdf)
    else:
        num_pages = pdf_file_reader.getNumPages()
        for index in range(0, num_pages):
            page_obj = pdf_file_reader.getPage(index)
            writer.addPage(page_obj)  # 根据每页返回的 PageObject,写入到文件


class UiWindows(object):
    # 设置界面
    def setupUi(self, main_windows):
        # 如果不是主窗口对象
        main_windows.resize(400, 300)

        self.central_widget = QWidget(main_windows)
        self.central_widget.setObjectName(u"中央部件")
        self.grid_layout1 = QGridLayout(self.central_widget)
        self.grid_layout1.setObjectName(u"网格布局1")
        self.horizontal_layout = QHBoxLayout()
        self.horizontal_layout.setObjectName(u"水平布局")
        self.label = QLabel(self.central_widget)
        self.label.setObjectName(u"标签")

        self.horizontal_layout.addWidget(self.label)

        self.line_edit = QLineEdit(self.central_widget)
        self.line_edit.setObjectName(u"线编辑")

        self.horizontal_layout.addWidget(self.line_edit)

        self.select_button = QPushButton(self.central_widget)
        self.select_button.setObjectName(u"选择按钮")

        self.horizontal_layout.addWidget(self.select_button)

        self.grid_layout1.addLayout(self.horizontal_layout, 0, 0, 1, 1)

        self.group_box = QGroupBox(self.central_widget)
        self.group_box.setObjectName(u"分组框")
        self.grid_layout = QGridLayout(self.group_box)
        self.grid_layout.setObjectName(u"网格布局")
        self.list_control_file = QListWidget(self.group_box)
        self.list_control_file.setObjectName(u"列表控件文件")

        self.grid_layout.addWidget(self.list_control_file, 0, 0, 1, 1)

        self.vertical_layout = QVBoxLayout()
        self.vertical_layout.setObjectName(u"垂直布局")
        self.add_button = QPushButton(self.group_box)
        self.add_button.setObjectName(u"添加按钮")

        self.vertical_layout.addWidget(self.add_button)

        self.delete_button = QPushButton(self.group_box)
        self.delete_button.setObjectName(u"删除按钮")

        self.vertical_layout.addWidget(self.delete_button)

        self.up_button = QPushButton(self.group_box)
        self.up_button.setObjectName(u"上移按钮")

        self.vertical_layout.addWidget(self.up_button)

        self.down_button = QPushButton(self.group_box)
        self.down_button.setObjectName(u"下移按钮")

        self.vertical_layout.addWidget(self.down_button)

        self.clear_button = QPushButton(self.group_box)
        self.clear_button.setObjectName(u"清空按钮")

        self.vertical_layout.addWidget(self.clear_button)
        # 垂直间距器
        self.vertical_spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)

        self.vertical_layout.addItem(self.vertical_spacer)

        self.grid_layout.addLayout(self.vertical_layout, 0, 1, 1, 1)

        self.grid_layout1.addWidget(self.group_box, 1, 0, 1, 1)
        # 水平布局2
        self.horizontal_layout2 = QHBoxLayout()
        self.horizontal_layout2.setObjectName(u"水平布局2")
        self.action_button = QPushButton(self.central_widget)
        self.action_button.setObjectName(u"执行按钮")

        self.horizontal_layout2.addWidget(self.action_button)

        self.horizontal_spacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)

        self.horizontal_layout2.addItem(self.horizontal_spacer)

        self.grid_layout1.addLayout(self.horizontal_layout2, 2, 0, 1, 1)

        self.grid_layout1.setRowStretch(0, 1)
        self.grid_layout1.setRowStretch(1, 15)
        self.grid_layout1.setRowStretch(2, 1)
        main_windows.setCentralWidget(self.central_widget)

        self.reset_ui(main_windows)

        main_windows.setWindowTitle(u"PDF合并工具")
        QMetaObject.connectSlotsByName(main_windows)

        self.select_button.clicked.connect(self.selectOutputFile)
        self.add_button.clicked.connect(self.addFile)
        self.delete_button.clicked.connect(self.deleteFile)
        self.clear_button.clicked.connect(self.clearFiles)
        self.action_button.clicked.connect(self.generatePdf)
        self.up_button.clicked.connect(self.pushUpFile)
        self.down_button.clicked.connect(self.pushDownFile)

    # 重译用户界面
    def reset_ui(self, main_windows):
        main_windows.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"输出文件路径：", None))
        self.select_button.setText(QCoreApplication.translate("MainWindow", u"浏览", None))
        self.group_box.setTitle(QCoreApplication.translate("MainWindow", u"合并文件列表", None))
        self.add_button.setText(QCoreApplication.translate("MainWindow", u"添加", None))
        self.delete_button.setText(QCoreApplication.translate("MainWindow", u"删除", None))
        self.up_button.setText(QCoreApplication.translate("MainWindow", u"上移", None))
        self.down_button.setText(QCoreApplication.translate("MainWindow", u"下移", None))
        self.clear_button.setText(QCoreApplication.translate("MainWindow", u"清空", None))
        self.action_button.setText(QCoreApplication.translate("MainWindow", u"执行", None))

    def selectOutputFile(self):
        dialog = QFileDialog(None, '请选择输出路径', '', '所有pdf文件 (*.pdf *.PDF)')
        dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        if dialog.exec_():
            file_names = dialog.selectedFiles()
            self.line_edit.setText(file_names[0])

    def addFile(self):
        dialog = QFileDialog(None, '请选择PDF文件', '', '所有pdf文件 (*.pdf *.PDF)')
        if dialog.exec_():
            file_names = dialog.selectedFiles()
            self.list_control_file.addItem(file_names[0])

    def deleteFile(self):
        del_item = self.list_control_file.currentItem()
        self.list_control_file.takeItem(self.list_control_file.row(del_item))

    def clearFiles(self):
        self.list_control_file.clear()

    def pushUpFile(self):
        cur_item = self.list_control_file.currentItem()
        cur_row = self.list_control_file.row(cur_item)
        if cur_row <= 0:
            return
        else:
            txt1 = self.list_control_file.item(cur_row - 1).text()
            txt2 = self.list_control_file.item(cur_row).text()
            self.list_control_file.item(cur_row - 1).setText(txt2)
            self.list_control_file.item(cur_row).setText(txt1)
            self.list_control_file.setCurrentRow(cur_row - 1)

    def pushDownFile(self):
        cur_item = self.list_control_file.currentItem()
        cur_row = self.list_control_file.row(cur_item)
        if cur_row >= self.list_control_file.count() - 1 or cur_row < 0:
            return
        else:
            txt1 = self.list_control_file.item(cur_row + 1).text()
            txt2 = self.list_control_file.item(cur_row).text()
            self.list_control_file.item(cur_row + 1).setText(txt2)
            self.list_control_file.item(cur_row).setText(txt1)
            self.list_control_file.setCurrentRow(cur_row + 1)

    def generatePdf(self):
        start_time = datetime.datetime.now()  # 开始时间
        output_file = self.line_edit.text()
        if not output_file:
            warn = QMessageBox()
            warn.warning(None, '提示', '请输入输出路径！')
            return
        merge_files = [self.list_control_file.item(index).text() for index in range(self.list_control_file.count())]
        if not merge_files:
            warn = QMessageBox()
            warn.warning(None, '提示', '请选择需要合并的pdf文件！', button0=None, button1=None)
            return
        pdf_file_writer = PyPDF2.PdfFileWriter()
        for merge_file in merge_files:
            merge_page(pdf_file_writer, merge_file)

        pdf_file_writer.write(open(output_file, 'wb'))
        end_time = datetime.datetime.now()  # 结束时间
        print('合并时间=', (end_time - start_time).seconds, 's')
        tip = QMessageBox()
        tip.information(None, '提示', '合并完成！', button0=None)


if __name__ == '__main__':
    app = QApplication()
    main_window = QMainWindow()

    ui = UiWindows()
    ui.setupUi(main_window)
    main_window.show()
    sys.exit(app.exec_())
