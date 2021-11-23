#!/usr/bin/env python3
# @Author : Tsurumi Nazo

from PySide2.QtWidgets import QApplication, QMessageBox, QFileDialog
from PySide2.QtUiTools import QUiLoader
import pdfplumber
import pandas as pd
import os


def get_all_pages(txt, symbol):

    numbers = []
    while txt.find(symbol) != -1:
        symbol_pos = txt.find(symbol)
        num_text = txt[0:symbol_pos]
        numbers.append(int(num_text))
        txt = txt[symbol_pos + 1:]
    numbers.append(int(txt))

    print(numbers)
    return numbers


class MainWidget:

    def __init__(self):
        self.ui = QUiLoader().load('MainFrame.ui')
        self.ui.le_CSVPos.setText(os.path.join(os.path.expanduser('~'), 'Desktop'))
        self.ui.tb_OpenPdfPos.clicked.connect(self.choose_pdf_pos)
        self.ui.tb_OpenCSVPos.clicked.connect(self.choose_csv_pos)
        self.ui.buttonGroup.buttonClicked.connect(self.change_state)
        self.ui.pb_Extract.clicked.connect(self.extract_csv)
        self.ui.pb_OpenFile.clicked.connect(self.open_csv)

    def choose_pdf_pos(self):
        filepath = QFileDialog.getOpenFileName(self.ui, "选择待解析的PDF")
        if os.path.splitext(filepath[0])[-1][1:] == 'pdf':
            self.ui.le_PdfPos.setText(filepath[0])
        else:
            QMessageBox.critical(
                self.ui,
                '错误',
                '选择的文件类型应为pdf'
            )

    def choose_csv_pos(self):
        filepath = QFileDialog.getExistingDirectory(self.ui, "选择保存CSV的位置")
        if filepath != '':
            self.ui.le_CSVPos.setText(filepath[0])

    def change_state(self):
        if self.ui.buttonGroup.checkedButton().text() == '全部':
            self.ui.le_CustomPages.setText('')
            self.ui.le_CustomPages.setReadOnly(True)
        else:
            self.ui.le_CustomPages.setReadOnly(False)

    def extract_csv(self):
        if self.ui.le_PdfPos.text() == '':
            QMessageBox.critical(
                self.ui,
                '错误',
                '请选择要解析的PDF'
            )
        elif self.ui.buttonGroup.checkedButton().text() != '全部' and self.ui.le_CustomPages.text() == '':
            QMessageBox.critical(
                self.ui,
                '错误',
                '请输入待解析的页数'
            )
        elif self.ui.le_CSVName.text() == '':
            QMessageBox.critical(
                self.ui,
                '错误',
                '待输出的CSV文件名不能为空'
            )
        else:
            pdf = pdfplumber.open(self.ui.le_PdfPos.text())
            csv_pos = self.ui.le_CSVPos.text() + '/' + self.ui.le_CSVName.text() + '.csv'

            if os.path.exists(csv_pos):
                os.remove(csv_pos)

            if self.ui.buttonGroup.checkedButton().text() == '全部':

                for page in pdf.pages:
                    table = page.extract_tables()
                    for t in table:
                        table_df = pd.DataFrame(t[1:], columns=t[0])
                        table_df.to_csv(csv_pos, mode='a', encoding='utf_8_sig')

                QMessageBox.information(
                    self.ui,
                    '操作成功',
                    '完成PDF中全部表格解析'
                )
            else:
                numbers_text = self.ui.le_CustomPages.text()

                if numbers_text.find(',') != -1 or numbers_text.find('，') != -1:
                    if numbers_text.find(',') != -1:
                        numbers = get_all_pages(numbers_text, ',')
                    else:
                        numbers = get_all_pages(numbers_text, '，')

                    for i in numbers:
                        page = pdf.pages[i]
                        table = page.extract_tables()
                        for t in table:
                            table_df = pd.DataFrame(t[1:], columns=t[0])
                            table_df.to_csv(csv_pos, mode='a', encoding='utf_8_sig')

                    QMessageBox.information(
                        self.ui,
                        '操作成功',
                        '完成PDF中全部表格解析'
                    )
                else:
                    QMessageBox.critical(
                        self.ui,
                        '错误',
                        '页码格式错误'
                    )

    def open_csv(self):
        csv_pos = self.ui.le_CSVPos.text() + '/' + self.ui.le_CSVName.text() + '.csv'
        os.startfile(csv_pos)


if __name__ == '__main__':
    app = QApplication([])
    widget = MainWidget()
    widget.ui.show()
    app.exec_()
