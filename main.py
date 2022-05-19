import sys  # sys нужен для передачи argv в QApplication
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import form  # Это наш конвертированный файл дизайна
import pandas as pd
import random
import csv
import os
from threading import Timer


#https://tproger.ru/translations/python-gui-pyqt/ Example

class ExampleApp(QtWidgets.QMainWindow, form.Ui_Dialog):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле form.py
        super().__init__()
        self.setupUi(self)  # Это нужно для инициализации нашего дизайна
        self.setFixedSize(self.size())
        
        self.cb_encoding.addItem("CP1251")
        self.cb_encoding.addItem("UTF-8")
        
        #
        self.btnConvert.setEnabled(False)
        self.label_complit.setHidden(True) 
        self.btnSelectFile.clicked.connect(self.browse_folder)
        self.btnConvert.clicked.connect(self.file_convert)

    def file_convert(self):
    
        self.label_complit.setHidden(True) 
        self.btnConvert.setEnabled(False)
        self.btnConvert.setText('Подождите...')
        fileName = self.lineFileName.text()
        
        #print(fileName)
        
        if (fileName):
            
            base = os.path.basename(fileName)
        
            self.progressBar.setValue(random.randint(5, 20))
            Timer(0.1, self.set_progress_bar).start()
        
            
            model = self.cbModel.currentText()
            status_variant = self.cbStatusVariant.currentText()
          
            xls_file = pd.read_excel(fileName)
            xls_file = xls_file.dropna(how='all')
            xls_file.to_csv(os.path.splitext(base)[0] + '.csv', columns=[model[3:], status_variant[3:]], header=[self.leModel.text(), self.leStatusVariant.text()], index = False, sep = ';', quoting = csv.QUOTE_NONNUMERIC, encoding = self.cb_encoding.currentText(), float_format='%g')
            
            
            self.progressBar.reset()
            
            self.label_complit.setHidden(False) 
            self.label_complit.setText('Готово!')
            self.label_complit.setStyleSheet("color: green")
            
            
        self.btnConvert.setText('Конвертировать')
        self.btnConvert.setDisabled(False)
        
            
        
    def browse_folder(self):

        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","Exel Files (*.xlsx)", options=options)
        if fileName:
            self.lineFileName.setText(fileName)

            self.progressBar.setValue(random.randint(5, 20))
            Timer(0.1, self.set_progress_bar).start()
            
            xls_file = pd.read_excel(fileName)
            #print(xls_file.columns)
            for num, name in enumerate(xls_file.columns, start=1):
                self.cbModel.addItem("{}: {}".format(chr(64+num), name))
                self.cbStatusVariant.addItem("{}: {}".format(chr(64+num), name))
        
        
            self.cbModel.setCurrentIndex(0)             # col A
            self.cbStatusVariant.setCurrentIndex(19)    # col T
        
            self.progressBar.reset()
            
            self.btnConvert.setDisabled(False) # файл есть можно конвертировать
            
    
    def set_progress_bar(self):
        self.progressBar.setValue(self.progressBar.value() + 1)
        if (self.progressBar.value() != 0):
            Timer(0.1, self.set_progress_bar).start()
        else:
            self.progressBar.reset()

def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = ExampleApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение

if __name__ == '__main__':  # Если мы запускаем файл напрямую, а не импортируем
    main()  # то запускаем функцию main()