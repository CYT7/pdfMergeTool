# -*- coding: utf-8 -*-
# ---
# @File: MergePdfTools.py
# @Author: Bob chen
# @CreateTime: Nov 18, 2021
# @Version : 1.0
# ---
from PySide2.QtCore import *
from PySide2.QtWidgets import *
from PIL import Image
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
        print(pdf_file_reader.getIsEncrypted())
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


class ui_main(object):
    # 设置界面
    def setupUi(self, main_windows):
        # if not MainWindow.objectName():

        main_windows.resize(866, 663)
        self.centralwidget = QWidget(main_windows)
        self.centralwidget.setObjectName(u"centralwidget")
        self.gridLayout_2 = QGridLayout(self.centralwidget)
        self.gridLayout_2.setObjectName(u"gridLayout_2")
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.label = QLabel(self.centralwidget)
        self.label.setObjectName(u"label")

        self.horizontalLayout.addWidget(self.label)

        self.lineEdit_out = QLineEdit(self.centralwidget)
        self.lineEdit_out.setObjectName(u"lineEdit_out")

        self.horizontalLayout.addWidget(self.lineEdit_out)

        self.pushButton_select = QPushButton(self.centralwidget)
        self.pushButton_select.setObjectName(u"pushButton_select")

        self.horizontalLayout.addWidget(self.pushButton_select)

        self.gridLayout_2.addLayout(self.horizontalLayout, 0, 0, 1, 1)

        self.groupBox = QGroupBox(self.centralwidget)
        self.groupBox.setObjectName(u"groupBox")
        self.gridLayout = QGridLayout(self.groupBox)
        self.gridLayout.setObjectName(u"gridLayout")
        self.listWidget_file = QListWidget(self.groupBox)
        self.listWidget_file.setObjectName(u"listWidget_file")

        self.gridLayout.addWidget(self.listWidget_file, 0, 0, 1, 1)

        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.pushButton_add = QPushButton(self.groupBox)
        self.pushButton_add.setObjectName(u"pushButton_add")

        self.verticalLayout.addWidget(self.pushButton_add)

        self.pushButton_delete = QPushButton(self.groupBox)
        self.pushButton_delete.setObjectName(u"pushButton_delete")

        self.verticalLayout.addWidget(self.pushButton_delete)

        self.pushButton_up = QPushButton(self.groupBox)
        self.pushButton_up.setObjectName(u"pushButton_up")

        self.verticalLayout.addWidget(self.pushButton_up)

        self.pushButton_down = QPushButton(self.groupBox)
        self.pushButton_down.setObjectName(u"pushButton_down")

        self.verticalLayout.addWidget(self.pushButton_down)

        self.pushButton_clear = QPushButton(self.groupBox)
        self.pushButton_clear.setObjectName(u"pushButton_clear")

        self.verticalLayout.addWidget(self.pushButton_clear)

        self.verticalSpacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)

        self.verticalLayout.addItem(self.verticalSpacer)

        self.gridLayout.addLayout(self.verticalLayout, 0, 1, 1, 1)

        self.gridLayout_2.addWidget(self.groupBox, 1, 0, 1, 1)

        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.pushButton_do = QPushButton(self.centralwidget)
        self.pushButton_do.setObjectName(u"pushButton_do")

        self.horizontalLayout_2.addWidget(self.pushButton_do)

        self.horizontalSpacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)

        self.horizontalLayout_2.addItem(self.horizontalSpacer)

        self.gridLayout_2.addLayout(self.horizontalLayout_2, 2, 0, 1, 1)

        self.gridLayout_2.setRowStretch(0, 1)
        self.gridLayout_2.setRowStretch(1, 15)
        self.gridLayout_2.setRowStretch(2, 1)
        main_windows.setCentralWidget(self.centralwidget)

        self.reset_ui(main_windows)

        main_windows.setWindowTitle(u"PDF合并工具")
        QMetaObject.connectSlotsByName(main_windows)

        self.pushButton_select.clicked.connect(self.selectOutputFile)
        self.pushButton_add.clicked.connect(self.addFile)
        self.pushButton_delete.clicked.connect(self.deleteFile)
        self.pushButton_clear.clicked.connect(self.clearFiles)
        self.pushButton_do.clicked.connect(self.generatePdf)
        self.pushButton_up.clicked.connect(self.pushUpFile)
        self.pushButton_down.clicked.connect(self.pushDownFile)

    # 重译用户界面
    def reset_ui(self, main_windows):
        main_windows.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.label.setText(
            QCoreApplication.translate("MainWindow", u"输出文件路径：", None))
        self.pushButton_select.setText(QCoreApplication.translate("MainWindow", u"浏览", None))
        self.groupBox.setTitle(QCoreApplication.translate("MainWindow", u"合并文件列表", None))
        self.pushButton_add.setText(QCoreApplication.translate("MainWindow", u"添加", None))
        self.pushButton_delete.setText(QCoreApplication.translate("MainWindow", u"删除", None))
        self.pushButton_up.setText(QCoreApplication.translate("MainWindow", u"上移", None))
        self.pushButton_down.setText(QCoreApplication.translate("MainWindow", u"下移", None))
        self.pushButton_clear.setText(QCoreApplication.translate("MainWindow", u"清空", None))
        self.pushButton_do.setText(QCoreApplication.translate("MainWindow", u"执行", None))

    def selectOutputFile(self):
        dialog = QFileDialog(None, '请选择输出路径', '', '所有pdf文件 (*.pdf *.PDF)')
        dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        if dialog.exec_():
            file_names = dialog.selectedFiles()
            self.lineEdit_out.setText(file_names[0])

    def addFile(self):
        dialog = QFileDialog(None, '请选择PDF文件', '', '所有pdf文件 (*.pdf *.PDF)')
        if dialog.exec_():
            file_names = dialog.selectedFiles()
            self.listWidget_file.addItem(file_names[0])

    def deleteFile(self):
        del_item = self.listWidget_file.currentItem()
        self.listWidget_file.takeItem(self.listWidget_file.row(del_item))

    def clearFiles(self):
        self.listWidget_file.clear()

    def pushUpFile(self):
        cur_item = self.listWidget_file.currentItem()
        cur_row = self.listWidget_file.row(cur_item)
        if cur_row <= 0:
            return
        else:
            txt1 = self.listWidget_file.item(cur_row - 1).text()
            txt2 = self.listWidget_file.item(cur_row).text()
            self.listWidget_file.item(cur_row - 1).setText(txt2)
            self.listWidget_file.item(cur_row).setText(txt1)
            self.listWidget_file.setCurrentRow(cur_row - 1)

    def pushDownFile(self):
        cur_item = self.listWidget_file.currentItem()
        cur_row = self.listWidget_file.row(cur_item)
        if cur_row >= self.listWidget_file.count() - 1 or cur_row < 0:
            return
        else:
            txt1 = self.listWidget_file.item(cur_row + 1).text()
            txt2 = self.listWidget_file.item(cur_row).text()
            self.listWidget_file.item(cur_row + 1).setText(txt2)
            self.listWidget_file.item(cur_row).setText(txt1)
            self.listWidget_file.setCurrentRow(cur_row + 1)

    def generatePdf(self):
        output_file = self.lineEdit_out.text()
        if not output_file:
            warn = QMessageBox()
            warn.warning(None, '提示', '请输入输出路径！')
            return
        merge_files = [self.listWidget_file.item(index).text() for index in range(self.listWidget_file.count())]
        if not merge_files:
            warn = QMessageBox()
            warn.warning(None, '提示', '请选择需要合并的pdf文件！')
            return
        pdf_file_writer = PyPDF2.PdfFileWriter()
        for merge_file in merge_files:
            merge_page(pdf_file_writer, merge_file)

        pdf_file_writer.write(open(output_file, 'wb'))
        tip = QMessageBox()
        tip.information(None, '提示', '合并完成！')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = QMainWindow()

    ui = ui_main()
    ui.setupUi(main_window)
    main_window.show()
    sys.exit(app.exec_())
