# Импортирование необходимых модулей из библиотеки PySide6 для разработки графического интерфейса
from PySide6.QtGui import QIcon, QRegularExpressionValidator
from PySide6 import QtCore, QtWidgets
from PySide6.QtCore import Qt

# Импортирование функции apply_stylesheet из модуля qt_material для применения пользовательских стилей к графическому интерфейсу
from qt_material import apply_stylesheet

# Импортирование библиотеки pandas для манипулирования и анализа данных
import pandas as pd

# Импортирование модулей sys и os для взаимодействия с операционной системой
import sys
import os


basedir = os.path.dirname(__file__)  # Получаем путь к директории, где находится исполняемый файл скрипта

try:
    from ctypes import windll  # Импортируем модуль ctypes для взаимодействия с библиотеками Windows
    myappid = 'school.calories.calculator.1.0'  # Устанавливаем идентификатор приложения для Windows
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)  # Устанавливаем идентификатор приложения для текущего процесса
except ImportError:
    pass  # В случае ошибки импорта (например, на других операционных системах) пропускаем это действие

'''
Главный класс приложения, в котором формируется весь интерфейс и логика работы
'''
class MyApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # Устанавливаем заголовок окна
        self.setWindowTitle("Счётчик калорий")
        
        # Устанавливаем иконку окна
        self.setWindowIcon(QIcon(os.path.join(basedir, "icon.ico")))
        
        # Устанавливаем размеры окна
        self.resize(900, 400)
        
        # Создаем и настраиваем виджеты для ввода данных пользователя и отображения результатов
        self.genderLabel = QtWidgets.QLabel("Ваш пол: ", alignment=Qt.AlignCenter, objectName="boldedBlackText")
        self.menGender = QtWidgets.QRadioButton("Мужской", objectName="menGender", minimumWidth=128)
        self.menGender.setLayoutDirection(Qt.RightToLeft)
        self.menGender.setChecked(True)
        self.womenGender = QtWidgets.QRadioButton("Женский", objectName="womenGender", minimumWidth=128)
        self.womenGender.toggled.connect(self.updateNormalAndNeededCalories)

        # Создаем и настраиваем валидатор для текстового поля
        validator = QRegularExpressionValidator(QtCore.QRegularExpression("[0-9]{3}"))

        # Добавляем поля для ввода возраста, роста и веса пользователя
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

        # Создаем и настраиваем таблицу для отображения данных о продуктах
        self.table = QtWidgets.QTableWidget(minimumHeight=256)
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["Продукт", "Вес, г", "Ккал", "Белки", "Жиры", "Углеводы", "X"])
        self.table.setItemDelegate(AlignDelegate())
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.currentCellChanged.connect(self.changeSelectedData)
        
        # Настраиваем размеры столбцов таблицы
        header = self.table.horizontalHeader()
        for i in range(self.table.columnCount() - 1):
            header.setSectionResizeMode(i, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(6, QtWidgets.QHeaderView.ResizeToContents)   

        # Создаем и настраиваем виджеты для отображения нормы калорий, употребленных калорий и необходимых калорий
        self.normalLabel = QtWidgets.QLabel("Ваша норма", objectName="boldedOrangeText")
        self.normalValue = QtWidgets.QLabel("0 ккал", objectName="boldedBlackText")

        self.eatedLabel = QtWidgets.QLabel("Вы употребили", objectName="boldedOrangeText")
        self.eatedValue = QtWidgets.QLabel("0 ккал", objectName="boldedBlackText")

        self.neededLabel = QtWidgets.QLabel("Ещё необходимо", objectName="boldedOrangeText")
        self.neededValue = QtWidgets.QLabel("0 ккал", objectName="boldedBlackText")

        # Инициализируем базу данных (считываем поля из excel таблицы для загрузки данных в память и последующей работы с ними)
        self.db = Database("блюда.xlsx")

        # Вызываем методы для настройки интерфейса
        self.setupUi()
        self.loadCss()
        self.addAppendBtn()
        self.addResults()


    def loadCss(self):
        # Открываем файл style.css, используя его путь, объединенный с путем каталога basedir, в режиме чтения (оператор "r")
        with open(os.path.join(basedir, "style.css"), "r") as f:
            # Считываем содержимое файла в переменную _style
            _style = f.read()
            # Применяем содержимое файла style.css к виджету, вызывающему этот метод, изменяя его стиль
            self.setStyleSheet(_style)


    def setupUi(self):
        # Создаем сетку для размещения виджетов
        self.layout = QtWidgets.QGridLayout(self)
        self.layout.setContentsMargins(10, 10, 10, 10)  # Устанавливаем отступы сетки

        # Добавляем метку с полем для ввода пола на первую строку сетки
        self.layout.addWidget(self.genderLabel, 0, 0, 1, 1)

        # Создаем горизонтальный слой для радиокнопок с выбором пола на первую строку сетки
        radioBtnLayout = QtWidgets.QHBoxLayout(objectName="radioBtnLayout", spacing=0)
        radioBtnLayout.setObjectName("radioBtnLayout")
        radioBtnLayout.addWidget(self.menGender)
        radioBtnLayout.addWidget(self.womenGender)
        self.layout.addLayout(radioBtnLayout, 0, 1, 1, 2)

        # Создаем вертикальный отступ
        verticalSpacer = QtWidgets.QSpacerItem(20, 10, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)

        # Добавляем вертикальный отступ на вторую строку сетки для разделения содержимого
        self.layout.addItem(verticalSpacer, 1, 0, 1, 3)

        # Создаем вертикальный слой для возраста
        ageLayout = QtWidgets.QVBoxLayout()
        ageLayout.addWidget(self.ageLabel)
        ageLayout.addWidget(self.age)
        self.layout.addLayout(ageLayout, 2, 0, 1, 1)  # Добавляем на третью строку сетки

        # Создаем вертикальный слой для роста
        heightLayout = QtWidgets.QVBoxLayout()
        heightLayout.addWidget(self.heightLabel)
        heightLayout.addWidget(self.height)
        heightLayout.setAlignment(Qt.AlignCenter)
        self.layout.addLayout(heightLayout, 2, 1, 1, 1)  # Добавляем на третью строку сетки

        # Создаем вертикальный слой для веса
        weightLayout = QtWidgets.QVBoxLayout()
        weightLayout.addWidget(self.weightLabel)
        weightLayout.addWidget(self.weight)
        self.layout.addLayout(weightLayout, 2, 2, 1, 1)  # Добавляем на третью строку сетки

        # Добавляем вертикальный отступ на четвертую строку сетки
        self.layout.addItem(verticalSpacer, 3, 0, 1, 3)

        # Добавляем таблицу на пятую строку сетки
        self.layout.addWidget(self.table, 4, 0, 1, 3)

        # Добавляем вертикальный отступ на шестую строку сетки
        self.layout.addItem(verticalSpacer, 5, 0, 1, 3)

        # Создаем вертикальный слой для отображения нормы
        normalLayout = QtWidgets.QVBoxLayout()
        normalLayout.addWidget(self.normalLabel, alignment=Qt.AlignCenter)
        normalLayout.addWidget(self.normalValue, alignment=Qt.AlignCenter)
        self.layout.addLayout(normalLayout, 6, 0, 1, 1)  # Добавляем на седьмую строку сетки

        # Создаем вертикальный слой для отображения употребленных калорий
        eatedLayout = QtWidgets.QVBoxLayout()
        eatedLayout.addWidget(self.eatedLabel, alignment=Qt.AlignCenter)
        eatedLayout.addWidget(self.eatedValue, alignment=Qt.AlignCenter)
        self.layout.addLayout(eatedLayout, 6, 1, 1, 1)  # Добавляем на седьмую строку сетки

        # Создаем вертикальный слой для отображения необходимых калорий
        neededLayout = QtWidgets.QVBoxLayout()
        neededLayout.addWidget(self.neededLabel, alignment=Qt.AlignCenter)
        neededLayout.addWidget(self.neededValue, alignment=Qt.AlignCenter)
        self.layout.addLayout(neededLayout, 6, 2, 1, 1)  # Добавляем на седьмую строку сетки

        # Добавляем вертикальный отступ на восьмую строку сетки
        self.layout.addItem(verticalSpacer, 7, 0, 1, 3)

    
    @QtCore.Slot() # декоратор нужен для того, чтобы иметь возможность вызвать функцию при нажатии на кнопку в интерфейсе
    def addItem(self):
        # Получаем номер строки для вставки нового элемента
        row = self.table.rowCount() - 2
        # Вставляем новую строку в таблицу
        self.table.insertRow(row)

        # Создаем и настраиваем комбо-бокс для выбора блюда
        comboBox = QtWidgets.QComboBox()
        comboBox.addItems(self.db.dishesNames)
        comboBox.currentIndexChanged.connect(self.updateRow)
        self.table.setCellWidget(row, 0, comboBox)

        # Получаем данные о первом блюде из базы данных
        dataRow = self.db.dishes[self.db.dishesNames[0]]
        # Создаем и настраиваем поле для ввода веса
        weight = QtWidgets.QLineEdit(str(dataRow[0]), alignment=Qt.AlignCenter)
        validator = QRegularExpressionValidator(QtCore.QRegularExpression("[0-9.]{10}"))
        weight.setValidator(validator)
        weight.textEdited.connect(self.updateResults)
        self.table.setCellWidget(row, 1, weight)

        # Заполняем ячейки таблицы данными о блюде
        for column in range(2, self.table.columnCount() - 1):
            self.table.setItem(row, column, QtWidgets.QTableWidgetItem(str(dataRow[column - 1])))

        # Создаем и настраиваем кнопку удаления строки
        removeBtn = QtWidgets.QPushButton("X")
        removeBtn.clicked.connect(self.removeRow)
        self.table.setCellWidget(row, 6, removeBtn)

        # Вызываем функцию обновления результатов
        self.updateResults()


    def addAppendBtn(self):
        # Функция для добавления кнопки "Добавить" в таблицу
        row = self.table.rowCount()  # Получаем текущее количество строк в таблице
        self.table.insertRow(row)  # Вставляем новую строку в таблицу
        addBtn = QtWidgets.QPushButton(" + ", objectName="addBtn")  # Создаем кнопку " + "
        addBtn.clicked.connect(self.addItem)  # Привязываем к кнопке метод добавления элемента
        self.table.setCellWidget(row, 3, addBtn)  # Устанавливаем кнопку в ячейку таблицы

    def addResults(self):
        # Функция для добавления текстового поля с результатом в таблицу
        row = self.table.rowCount()  # Получаем текущее количество строк в таблице
        self.table.insertRow(row)  # Вставляем новую строку в таблицу
        res = QtWidgets.QLabel("Итог: ", objectName="resultText")  # Создаем текстовое поле с надписью "Итог: "
        self.table.setCellWidget(row, 0, res)  # Устанавливаем текстовое поле в ячейку таблицы

    @QtCore.Slot()
    def changeSelectedData(self, *args):
        # Слот для обработки изменения выделенных данных в таблице
        self.curRow = args[0]  # Устанавливаем текущую строку
        self.curCol = args[1]  # Устанавливаем текущий столбец
        self.prewRow = args[2]  # Устанавливаем предыдущую строку
        self.prewCol = args[3]  # Устанавливаем предыдущий столбец


    @QtCore.Slot()
    def updateRow(self, dishIndex):
        # Создаем комбо-бокс
        comboBox = QtWidgets.QComboBox()
        # Добавляем элементы из базы данных в комбо-бокс
        comboBox.addItems(self.db.dishesNames)
        # Устанавливаем текущий индекс выбранного блюда
        comboBox.setCurrentIndex(dishIndex)
        # Привязываем сигнал изменения индекса к слоту обновления строки
        comboBox.currentIndexChanged.connect(self.updateRow)
        # Устанавливаем комбо-бокс в таблицу в указанную ячейку
        self.table.setCellWidget(self.curRow, 0, comboBox)

        # Получаем данные о блюде из базы данных
        dataRow = self.db.dishes[self.db.dishesNames[dishIndex]]
        # Создаем поле веса и заполняем данными из базы данных
        weight = QtWidgets.QLineEdit(str(dataRow[0]), alignment=Qt.AlignCenter)
        # Привязываем сигнал завершения редактирования к обновлению результатов
        weight.editingFinished.connect(self.updateResults)
        # Устанавливаем поле веса в таблицу в указанную ячейку
        self.table.setCellWidget(self.curRow, 1, weight)

        # Заполняем остальные ячейки таблицы данными о блюде
        for column in range(2, self.table.columnCount() - 1):
            self.table.setItem(self.curRow, column, QtWidgets.QTableWidgetItem(str(dataRow[column-1])))

        # Создаем кнопку удаления строки
        removeBtn = QtWidgets.QPushButton("X")
        # Привязываем сигнал нажатия кнопки к методу удаления строки
        removeBtn.clicked.connect(self.removeRow)
        # Устанавливаем кнопку удаления в таблицу в указанную ячейку
        self.table.setCellWidget(self.curRow, 6, removeBtn)
        # Обновляем результаты
        self.updateResults()

    @QtCore.Slot()    
    def removeRow(self):
        # Удаляем текущую строку из таблицы
        self.table.removeRow(self.curRow)
        # Обновляем результаты
        self.updateResults()


    @QtCore.Slot()
    def updateResults(self):
        # Обновление результатов
        resultsRow = self.table.rowCount() - 1
        # Нахождение номера строки для вывода результатов

        for column in range(1, self.table.columnCount()-1):
            # Перебор всех столбцов с данными (кроме первого и последнего)
            columnSum = self.getSumOfColum(column)
            # Вычисление суммы значений в столбце
            if column == 2: self.eatedValue.setText(f"{columnSum:.0f} ккал") 
            # Если столбец равен 2, обновляем значение в соответствующем виджете
            self.table.setItem(resultsRow, column, QtWidgets.QTableWidgetItem(f"{columnSum:.2f}"))
            # Устанавливаем значение суммы в таблицу

        self.neededValue.setText(f"{int(self.normalValue.text()[:-5]) - int(self.eatedValue.text()[:-5])} ккал")
        # Вычисляем и устанавливаем разницу между нормой и употребленными калориями

    def getSumOfColum(self, columIndex) -> int:
        # Функция для вычисления суммы значений в столбце
        columnSum = 0
        try:
            # Попытка найти значения в ячейках таблицы
            for row in range(self.table.rowCount() - 2):
                item = self.table.item(row, columIndex)
                # Получаем элемент из таблицы
                columnSum += float(item.text() if item.text() != '' else 0)
                # Суммируем значения, преобразуя их в числовой формат
        except AttributeError:
            # Обработка ошибки, если элемент не найден
            for row in range(self.table.rowCount() - 2):
                item = self.table.cellWidget(row, columIndex)
                # Получаем виджет из таблицы
                columnSum += float(item.text() if item.text() != '' else 0)
                # Суммируем значения виджетов, преобразуя их в числовой формат
        return columnSum
        # Возвращаем сумму значений в столбце

    # Обновление нормы и необходимых калорий в зависимости от введенных параметров
    @QtCore.Slot()
    def updateNormalAndNeededCalories(self):
        try:
            # Получаем возраст, рост и вес из введенных значений
            age = int(self.age.text())
            height = int(self.height.text())
            weight = int(self.weight.text())
        except:
            return  # Прерываем выполнение функции в случае возникновения исключения

        # Вычисление нормы калорий в зависимости от пола
        if self.menGender.isChecked():
            self.normalValue.setText(f"{9.99 * weight + 6.25 * height - 4.92 * age + 5:.0f} ккал")
        elif self.womenGender.isChecked():
            self.normalValue.setText(f"{9.99 * weight + 6.25 * height - 4.92 * age - 161:.0f} ккал")
        else:
            print('error')  # Выводим сообщение об ошибке, если пол не выбран корректно

        # Вычисление необходимых калорий для достижения нормы
        self.neededValue.setText(f"{int(self.normalValue.text()[:-5]) - int(self.eatedValue.text()[:-5])} ккал")

    # Проверка пути к файлу и добавление его к списку для чтения
    @QtCore.Slot()
    def check_path(self):
        path = self.readFileName.text()
        # Проверяем, является ли указанный путь файлом и имеет ли расширение .xlsx или .xls
        if os.path.isfile(path) and path.split(".")[-1] in ("xlsx", "xls"):
            self.add_read_file(path)  # Добавляем файл к списку для чтения
            self.update_tree_widget(self.get_complete_dict())  # Обновляем виджет с данными из файла
            self.tree.itemChanged.connect(self.check_all)  # Подключаем функцию check_all к событию изменения элемента

    # Открытие файла через диалоговое окно
    @QtCore.Slot()
    def open_file(self):
        # Запрашиваем у пользователя путь к файлу через диалоговое окно
        path = QtWidgets.QFileDialog.getOpenFileName(self, 'Открыть Excel файл', '',
                                        'Excel files (*.xlsx *.xls)')
        if path != ('', ''):  # Если пользователь выбрал файл
            self.readFileName.setText(path[0])  # Устанавливаем путь к файлу в соответствующее поле
            self.check_path()  # Проверяем путь и обрабатываем файл

'''
Определение пользовательского делегата AlignDelegate, который является подклассом QItemDelegate
Просто чтобы ячейки в таблице были отцентрованы
'''
class AlignDelegate(QtWidgets.QItemDelegate):
    # Переопределение метода paint для кастомной отрисовки элементов делегата
    def paint(self, painter, option, index):
        # Установка горизонтального выравнивания текста в ячейке по центру
        option.displayAlignment = QtCore.Qt.AlignCenter
        # Вызов родительского метода paint для отображения содержимого элемента делегата
        QtWidgets.QItemDelegate.paint(self, painter, option, index)

'''
Простая реализации базы данных для получения блюд из таблицы и хранения их в памяти.
'''
class Database:
    # Инициализация атрибутов dishes и dishesNames
    dishes = {}  # Словарь для хранения блюд и их состава
    dishesNames = []  # Список для хранения названий блюд

    # Метод инициализации (конструктор) класса, принимающий путь к файлу excelPath
    def __init__(self, excelPath: str) -> None:
        # Чтение данных из файла excelPath и сохранение их в переменную sheet
        sheet = pd.read_excel(excelPath)
        
        # Перебор строк (индексов) в таблице данных
        for i in sheet.index:
            # Добавление в словарь dishes новой записи в виде пары ключ-значение,
            # где ключ - название блюда, значение - состав блюда (в виде списка)
            self.dishes.update({sheet.values[i][0]: sheet.values[i][1:]})
        
        # Получение ключей (названий блюд) из словаря dishes и сохранение их в списке dishesNames
        self.dishesNames = list(self.dishes.keys())


# Если данный скрипт запускается как основной (а не, например, импортируется из другого скрипта),
# то выполняется следующий блок кода
if __name__ == "__main__":
    # Создаем экземпляр приложения QtWidgets.QApplication([])
    application = QtWidgets.QApplication([])
    
    # Создаем экземпляр класса MyApp
    MyApp = MyApp()
    
    # Показываем приложение, чтобы пользователь мог видеть его интерфейс
    MyApp.show()
    
    # Применяем стилевую тему с помощью функции apply_stylesheet,
    # передавая в нее приложение и название темы 'light_orange.xml'
    apply_stylesheet(application, theme='light_orange.xml')
    
    # Выходим из цикла выполнения приложения и завершаем программу с кодом возврата application.exec()
    sys.exit(application.exec())
