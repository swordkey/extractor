import sys
from PyQt5 import QtWidgets
from PyQt5 import uic
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QErrorMessage
import pandas as pd
import os


class merge_vars:
    current_df = []
    result_df = pd.DataFrame()
    add_list = []
    fileName = ""
    sheetlist = []

class Form(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        self.ui = uic.loadUi("EPI_extract.ui", self)
        self.ui.show()
        self.key_listA.addItems(["EPI", "RID"])
        self.key_listB.addItems(["EPI", "RID"])


    def importbutton_clicked(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*)", options=options)
        if fileName:
            merge_vars.current_df = pd.read_csv(fileName, encoding='utf-8-sig', index_col=None, header=0)
            self.ui.df_state.setText(os.path.splitext(os.path.basename(fileName))[0])
            self.current_list.clear()
            for cols in merge_vars.current_df.columns:
                self.current_list.addItem(cols)
        self.add_list.clear()

    def xlsximport_clicked(self):
        self.sheetname_combo.clear()
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*)", options=options)
        self.sheetname_combo.addItems(pd.ExcelFile(fileName).sheet_names)
        merge_vars.fileName =fileName



    def combo_select(self):
        fileName = merge_vars.fileName
        if fileName:
            merge_vars.current_df = pd.read_excel(fileName, self.sheetname_combo.currentText())
            self.ui.df_state.setText(os.path.splitext(os.path.basename(fileName))[0])
            self.current_list.clear()
            for cols in merge_vars.current_df.columns:
                self.current_list.addItem(cols)
        self.add_list.clear()
        print(merge_vars.current_df)


    def list_add_clicked(self):
        if self.current_list.currentItem():
            self.add_list.addItem(self.current_list.currentItem().text())
        else:
            msgBox = QMessageBox()
            msgBox.setText("선택된 리스트 없음")
            msgBox.exec()

    def list_delete_clicked(self):
        self.add_list.takeItem(self.add_list.currentRow())

    def list_clear_clicked(self):
        self.add_list.clear()
    def current_doubleclicked(self):
        self.add_list.addItem(self.current_list.currentItem().text())


    def key_A(self):
        if self.add_list.currentItem():
            self.key_listA.addItem(self.add_list.currentItem().text())
        else:
            msgBox = QMessageBox()
            msgBox.setText("선택된 리스트 없음")
            msgBox.exec()
    def key_B(self):
        if self.selected_list.currentItem():
            self.key_listB.addItem(self.selected_list.currentItem().text())
        else:
            msgBox = QMessageBox()
            msgBox.setText("선택된 리스트 없음")
            msgBox.exec()
    def key_AD(self):
        self.key_listA.takeItem(self.key_listA.currentRow())
    def key_BD(self):
        self.key_listB.takeItem(self.key_listB.currentRow())

    def merge_clicked(self):


        if merge_vars.result_df.empty==True:

            add_list = [self.add_list.item(x).text() for x in range(self.add_list.count())]
            merge_vars.result_df = merge_vars.current_df[add_list]
            print(merge_vars.result_df)

        else:
            add_list = [self.add_list.item(x).text() for x in range(self.add_list.count())]
            current_df = merge_vars.current_df[add_list]
            merge_vars.result_df = pd.merge(current_df, merge_vars.result_df, left_on = [self.key_listA.item(x).text() for x in range(self.key_listA.count())], right_on = [self.key_listB.item(x).text() for x in range(self.key_listB.count())], how = 'outer')
            print(merge_vars.result_df.columns)

        self.selected_list.clear()
        self.selected_list.addItems(merge_vars.result_df.columns)
        self.add_list.clear()

    def export_clicked(self):
        merge_vars.result_df.to_csv(self.export_name.text()+".csv", encoding = 'utf-8-sig')
        self.selected_list.clear()
        merge_vars.result_df = pd.DataFrame()
        merge_vars.add_list = []



    def test_clicked(self):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setText("Message box pop up window")
        msgBox.setWindowTitle("QMessageBox Example")
        msgBox.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        returnValue = msgBox.exec()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    w = Form()
    sys.exit(app.exec())