from PySide6.QtGui import QIcon, QRegularExpressionValidator
from PySide6 import QtCore, QtWidgets
from PySide6.QtCore import Qt

from qt_material import apply_stylesheet
import pandas as pd
import sys, os

basedir = os.path.dirname(__file__)

class MyApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Счётчик калорий")
        self.setWindowIcon(QIcon(os.path.join(basedir, "icon.ico")))
        self.resize(900, 400)
        
        self.genderLabel = QtWidgets.QLabel("Ваш пол: ", alignment = Qt.AlignCenter, objectName = "boldedBlackText")
        self.menGender = QtWidgets.QRadioButton("Мужской", objectName="menGender", minimumWidth=128)
        self.menGender.setLayoutDirection(Qt.RightToLeft)
        self.menGender.setChecked(True)
        self.womenGender = QtWidgets.QRadioButton("Женский", objectName="womenGender", minimumWidth=128)
        self.womenGender.toggled.connect(self.updateNormalAndNeededCalories)

        self.globalValidator = QRegularExpressionValidator(QtCore.QRegularExpression("[0-9]{6}"))
        validator = QRegularExpressionValidator(QtCore.QRegularExpression("[0-9]{3}"))

        self.ageLabel = QtWidgets.QLabel("Ваш возраст: ")
        self.age = QtWidgets.QLineEdit()
        self.age.setValidator(validator)
        self.age.textEdited.connect(self.updateNormalAndNeededCalories)
        self.age.setPlaceholderText("25 (лет)")

        self.heightLabel = QtWidgets.QLabel("Ваш рост: ")
        self.height = QtWidgets.QLineEdit()
        self.height.setValidator(validator)
        self.height.textEdited.connect(self.updateNormalAndNeededCalories)
        self.height.setPlaceholderText("175 (см)")
        
        self.weightLabel = QtWidgets.QLabel("Ваш вес: ")
        self.weight = QtWidgets.QLineEdit()
        self.weight.setValidator(validator)
        self.weight.textEdited.connect(self.updateNormalAndNeededCalories)
        self.weight.setPlaceholderText("70 (кг)")

        self.table = QtWidgets.QTableWidget(minimumHeight=256)
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["Продукт", "Вес, г", "Ккал", "Белки", "Жиры", "Углеводы", "X"])
        self.table.setItemDelegate(AlignDelegate())
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.currentCellChanged.connect(self.changeSelectedData)
        header = self.table.horizontalHeader()
        for i in range(self.table.columnCount()-1):
            header.setSectionResizeMode(i, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(6, QtWidgets.QHeaderView.ResizeToContents)   
        self.normalLabel = QtWidgets.QLabel("Ваша норма", objectName = "boldedOrangeText")
        self.normalValue = QtWidgets.QLabel("0 ккал", objectName = "boldedBlackText")

        self.eatedLabel = QtWidgets.QLabel("Вы употребили", objectName = "boldedOrangeText")
        self.eatedValue = QtWidgets.QLabel("0 ккал", objectName = "boldedBlackText")

        self.neededLabel = QtWidgets.QLabel("Ещё необходимо", objectName = "boldedOrangeText")
        self.neededValue = QtWidgets.QLabel("0 ккал", objectName = "boldedBlackText")

        self.db = Database("блюда.xlsx")
        # print(self.db.dishesNames)
        # self.set_dark_mode()
        self.setupUi()
        self.loadCss()
        self.addAppendBtn()
        self.addResults()
        self.addItem()
        self.addItem()
        self.addItem()

    def loadCss(self):
        with open("style.css", "r") as f:
            _style = f.read()
            self.setStyleSheet(_style)

    def setupUi(self):

        self.layout = QtWidgets.QGridLayout(self)
        self.layout.setContentsMargins(10,10,10,10)

        self.layout.addWidget(self.genderLabel, 0, 0, 1, 1)

        radioBtnLayout = QtWidgets.QHBoxLayout(objectName="radioBtnLayout", spacing=0)
        radioBtnLayout.setObjectName("radioBtnLayout")
        radioBtnLayout.addWidget(self.menGender)
        radioBtnLayout.addWidget(self.womenGender)
        self.layout.addLayout(radioBtnLayout, 0, 1, 1, 2)

        verticalSpacer = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)

        self.layout.addItem(verticalSpacer, 1, 0, 1, 3)

        ageLayout = QtWidgets.QVBoxLayout()
        ageLayout.addWidget(self.ageLabel)
        ageLayout.addWidget(self.age)
        self.layout.addLayout(ageLayout, 2, 0, 1, 1)

        heightLayout = QtWidgets.QVBoxLayout()
        heightLayout.addWidget(self.heightLabel)
        heightLayout.addWidget(self.height)
        heightLayout.setAlignment(Qt.AlignCenter)
        self.layout.addLayout(heightLayout, 2, 1, 1, 1)

        weightLayout = QtWidgets.QVBoxLayout()
        weightLayout.addWidget(self.weightLabel)
        weightLayout.addWidget(self.weight)
        self.layout.addLayout(weightLayout, 2, 2, 1, 1)

        self.layout.addItem(verticalSpacer, 3, 0, 1, 3)
        self.layout.addWidget(self.table, 4, 0, 1, 3)
        self.layout.addItem(verticalSpacer, 5, 0, 1, 3) 

        normalLayout = QtWidgets.QVBoxLayout()
        normalLayout.addWidget(self.normalLabel, alignment=Qt.AlignCenter)
        normalLayout.addWidget(self.normalValue, alignment=Qt.AlignCenter)
        self.layout.addLayout(normalLayout, 6, 0, 1, 1)   

        eatedLayout = QtWidgets.QVBoxLayout()
        eatedLayout.addWidget(self.eatedLabel, alignment=Qt.AlignCenter)
        eatedLayout.addWidget(self.eatedValue, alignment=Qt.AlignCenter)
        self.layout.addLayout(eatedLayout, 6, 1, 1, 1)     

        neededLayout = QtWidgets.QVBoxLayout()
        neededLayout.addWidget(self.neededLabel, alignment=Qt.AlignCenter)
        neededLayout.addWidget(self.neededValue, alignment=Qt.AlignCenter)
        self.layout.addLayout(neededLayout, 6, 2, 1, 1)   

        self.layout.addItem(verticalSpacer, 7, 0, 1, 3)
    
    @QtCore.Slot()
    def addItem(self):
        row = self.table.rowCount()-3
        self.table.insertRow(row)
        comboBox = QtWidgets.QComboBox()
        comboBox.addItems(self.db.dishesNames)
        comboBox.currentIndexChanged.connect(self.updateRow)
        self.table.setCellWidget(row, 0, comboBox)

        dataRow = self.db.dishes[self.db.dishesNames[0]]
        weight = QtWidgets.QLineEdit(str(dataRow[0]), alignment=Qt.AlignCenter)
        weight.setValidator(self.globalValidator)
        weight.editingFinished.connect(self.updateResults)
        self.table.setCellWidget(row, 1, weight)

        for column in range(2, self.table.columnCount() - 1):
            self.table.setItem(row, column, QtWidgets.QTableWidgetItem(str(dataRow[column-1])))

        removeBtn = QtWidgets.QPushButton("X")
        removeBtn.clicked.connect(self.removeRow)
        self.table.setCellWidget(row, 6, removeBtn)
        self.updateResults()

    def addAppendBtn(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        addBtn = QtWidgets.QPushButton(" + ", objectName="addBtn")
        addBtn.clicked.connect(self.addItem)
        self.table.setCellWidget(row, 3, addBtn)

    def addResults(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        res = QtWidgets.QLabel("Итог: ", objectName="resultText")
        self.table.setCellWidget(row, 0, res)        
        row = self.table.rowCount()
        self.table.insertRow(row)
        res = QtWidgets.QLabel("Итог (на 100гр.): ", objectName="resultText")
        self.table.setCellWidget(row, 0, res)      
    
    @QtCore.Slot()
    def changeSelectedData(self, *args):
        self.curRow = args[0]
        self.curCol = args[1]
        self.prewRow = args[2]
        self.prewCol = args[3]

    @QtCore.Slot()
    def updateRow(self, dishIndex):
        comboBox = QtWidgets.QComboBox()
        comboBox.addItems(self.db.dishesNames)
        comboBox.setCurrentIndex(dishIndex)
        comboBox.currentIndexChanged.connect(self.updateRow)
        self.table.setCellWidget(self.curRow, 0, comboBox)

        dataRow = self.db.dishes[self.db.dishesNames[dishIndex]]
        weight = QtWidgets.QLineEdit(str(dataRow[0]), alignment=Qt.AlignCenter)
        weight.editingFinished.connect(self.updateResults)
        self.table.setCellWidget(self.curRow, 1, weight)

        for column in range(2, self.table.columnCount() - 1):
            self.table.setItem(self.curRow, column, QtWidgets.QTableWidgetItem(str(dataRow[column-1])))

        removeBtn = QtWidgets.QPushButton("X")
        removeBtn.clicked.connect(self.removeRow)
        self.table.setCellWidget(self.curRow, 6, removeBtn)
        self.updateResults()

    @QtCore.Slot()    
    def removeRow(self):
        self.table.removeRow(self.curRow)
        self.updateResults()

    @QtCore.Slot()
    def updateResults(self):
        print('update')
        resultsRow = self.table.rowCount() - 2
        resultsRow100 = self.table.rowCount() - 1

        for column in range(1, self.table.columnCount()-1):
            columnSum = self.getSumOfColum(column)
            if column == 2: self.eatedValue.setText(f"{columnSum:.0f} ккал") 
            self.table.setItem(resultsRow, column, QtWidgets.QTableWidgetItem(f"{columnSum:.2f}"))
            self.table.setItem(resultsRow100, column, QtWidgets.QTableWidgetItem(f"{columnSum / self.table.rowCount() - 3:.2f}"))
        
        self.neededValue.setText(f"{int(self.normalValue.text()[:-5]) - int(self.eatedValue.text()[:-5])} ккал")

    def getSumOfColum(self, columIndex) -> int:
        columnSum = 0
        try:
            for row in range(self.table.rowCount() - 3):
                item = self.table.item(row, columIndex)
                columnSum += float(item.text())
        except AttributeError:
            for row in range(self.table.rowCount() - 3):
                item = self.table.cellWidget(row, columIndex)
                columnSum += float(item.text())
        return columnSum
    
    @QtCore.Slot()
    def updateNormalAndNeededCalories(self):
        try:
            age = int(self.age.text())
            height = int(self.height.text())
            weight = int(self.weight.text())
        except:
            return
        
        if self.menGender.isChecked():
            self.normalValue.setText(f"{9.99 * weight + 6.25 * height - 4.92 * age + 5:.0f} ккал")
        elif self.womenGender.isChecked():
            self.normalValue.setText(f"{9.99 * weight + 6.25 * height - 4.92 * age - 161:.0f} ккал")
        else:
            print('error')

        self.neededValue.setText(f"{int(self.normalValue.text()[:-5]) - int(self.eatedValue.text()[:-5])} ккал")

    @QtCore.Slot()
    def check_path(self):
        path = self.readFileName.text()
        if os.path.isfile(path) and path.split(".")[-1] in ("xlsx", "xls"):
            self.add_read_file(path)
            self.update_tree_widget(self.get_complete_dict())
            self.tree.itemChanged.connect(self.check_all)

    @QtCore.Slot()
    def open_file(self):
        path = QtWidgets.QFileDialog.getOpenFileName(self, 'Открыть Excel файл', '',
                                        'Excel files (*.xlsx *.xls)')
        if path != ('', ''):
            self.readFileName.setText(path[0])
            self.check_path()

class AlignDelegate(QtWidgets.QItemDelegate):
    def paint(self, painter, option, index):
        option.displayAlignment = QtCore.Qt.AlignCenter
        QtWidgets.QItemDelegate.paint(self, painter, option, index)

class Database:
    dishes = {}
    dishesNames = []

    def __init__(self, excelPath : str) -> None:
        sheet = pd.read_excel(excelPath)
        for i in sheet.index:
            self.dishes.update({sheet.values[i][0]:sheet.values[i][1:]})
        self.dishesNames = list(self.dishes.keys())

if __name__ == "__main__":
    application = QtWidgets.QApplication([])
    MyApp = MyApp()
    MyApp.show()
    apply_stylesheet(application, theme='light_orange.xml')
    sys.exit(application.exec())