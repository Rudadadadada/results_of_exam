import sys
import xlrd
import sqlite3
from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QListWidgetItem, QMainWindow
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QInputDialog
import matplotlib.pyplot as plt


MONTHS = {'Сентябрь': 1, 'Октябрь': 2, "Ноябрь": 3,
          "Декабрь": 4, "Январь": 5, "Февраль": 6,
          "Март": 7, "Апрель": 8, "Май": 9, "Июнь": 10, "Июль": 11, "Август": 12}


def clear_db():
    """Удаляю все таблицы из предыдущей базы данных"""
    con = sqlite3.connect('results.db')
    cur = con.cursor()
    table_names = [''.join(i)
                   for i in list(cur.execute("select name"
                                             " from sqlite_master"
                                             " where type='table'"))]
    for i in range(len(table_names)):
        cur.execute(f'drop table {table_names[i]}')
        con.commit()


class IntroducingWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('ui/intro.ui', self)
        self.setFixedSize(self.geometry().width(), self.geometry().height())
        self.btn_choose_table.clicked.connect(self.choose_table)
        self.button_skip.clicked.connect(self.skip)

    def choose_table(self):
        """Функция вызова начального окна с добавлением начальных данных в БД."""
        self.ex = MainProjectsWindow()
        self.ex.open_table()
        self.ex.show()
        self.hide()

    def skip(self):
        """Функция вызова начального окна без добавления начальных данных в БД."""
        self.ex = MainProjectsWindow()
        self.ex.show()
        self.hide()


class MainProjectsWindow(QMainWindow):
    """Главное окно"""

    def __init__(self):
        super().__init__()
        uic.loadUi('ui/results_qt.ui', self)
        self.setFixedSize(self.geometry().width(), self.geometry().height())
        self.window_button.clicked.connect(self.open_dialog_sorts)
        self.window_button1.clicked.connect(self.open_dialog_student)
        self.window_button2.clicked.connect(self.open_dialog_table)
        self.table_b.clicked.connect(self.open_table)

    def open_dialog_sorts(self):
        """Октрываю окно сортировки и графиков"""
        self.dialog = GraphsOptions()
        self.dialog.show()
        self.hide()

    def open_dialog_student(self):
        """Открываю окно добавления ученика"""
        self.dialog = AddStudent()
        self.dialog.show()
        self.hide()

    def open_dialog_table(self):
        """Открываю окно вывода таблицы"""
        self.dialog = Data()
        self.dialog.show()
        self.hide()

    def open_table(self):
        """Ввод exel таблицы"""
        table = QFileDialog.getOpenFileName(self, 'Выбрать картинку', '',
                                            "Таблица(*.xlsx)")[0]
        """Делаю через try, во избежание ошибки, если я ничего не ввел"""
        try:
            count = 1
            file = xlrd.open_workbook(table)  # открываю таблицу
            sheet = file.sheet_by_index(0)  # беру первый лист
            """В переменной vals храню всю информацию о таблице"""
            vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
            data = []
            letter = list(vals[0][0])[1]
            month = vals[0][1]
            subject = vals[0][2]
            table_name = f'{subject}_{month}_{letter}'
            """Беру с третьего элемента, потому что первые два не нужны"""
            for i in vals[2:]:
                data.append([int(i[0]), i[1], i[-3], i[-2], i[-1] * 100])
            con = sqlite3.connect('results.db')
            cur = con.cursor()
            """Делаю список таблиц"""
            table_names = [''.join(i) for i in
                           list(cur.execute("select name from "
                                            "sqlite_master where type='table'"))]
            """Проверяю есть ли такая таблица в списке таблиц"""
            if table_name not in table_names:
                cur.execute(f"CREATE TABLE {table_name} (Id INTEGER, ФИ TEXT, "
                            f"Количество_баллов INTEGER, Оценка INTEGER, "
                            f"Процент_выполнения DOUBLE, PRIMARY KEY (id));")
            """Создаю список имен"""
            student_names = [''.join(i) for i in
                             list(cur.execute(f"select ФИ from {table_name}"))]
            count += len(student_names)  # для подсчета ID беру длину списка имен
            for i in data:
                student_name = i[1].strip()
                """Проверяю, есть ли имя в таблице и добавляю, если нет"""
                if student_name not in student_names:
                    cur.execute(f"insert into {table_name} "
                                f"(Id, ФИ, Количество_баллов, "
                                f"Оценка, Процент_выполнения) values "
                                f"({count}, '{student_name}',"
                                f" {int(i[2])}, {int(i[3])}, {float(i[4])})")
                count += 1
            con.commit()
            '''Вывожу диалоговое окно, сведительствующее об успехе операции'''
            self.success_dialog_add = QMessageBox.information(self, 'Успех',
                                                              'Предмет успешно'
                                                              ' добавлен в '
                                                              'базу данных',
                                                              buttons=
                                                              QMessageBox.Ok)
        except:
            pass


class AddStudent(QMainWindow):
    """Окно добавления ученика"""

    def __init__(self):
        super().__init__()
        uic.loadUi('ui/AddStudent.ui', self)
        self.setFixedSize(self.geometry().width(), self.geometry().height())
        self.back_b1.clicked.connect(self.return_main_window)
        self.add_b.clicked.connect(self.add_student)

    def return_main_window(self):
        """Открывает главное окно"""
        self.main_window = MainProjectsWindow()
        self.main_window.show()
        self.hide()

    def add_student(self):
        """Функция добавления ученика"""
        count = 1
        letter = self.letter.text()
        subject = self.subject.text()
        month = self.month.text()
        student_name = self.name.text()
        points = self.points.text()
        mark = self.mark.text()
        percent = self.percent.text()
        check = False
        check1 = False
        check2 = False
        if letter.isalpha():
            check = True
        if mark and percent:
            if int(mark) in [2, 3, 4, 5] and 0 <= int(percent) <= 100:
                check1 = True
        if all([letter, subject, month, student_name, points, mark, percent]):
            check2 = True
        if all([check, check1, check2]):
            self.letter.clear()
            self.subject.clear()
            self.month.clear()
            self.name.clear()
            self.points.clear()
            self.mark.clear()
            self.percent.clear()
            table_name = f'{subject.strip(" ")}_{month}_{letter}'
            con = sqlite3.connect('results.db')
            cur = con.cursor()
            table_names = [''.join(i) for i in
                           list(cur.execute("select name from "
                                            "sqlite_master where type='table'"))]
            if table_name not in table_names:
                cur.execute(f"CREATE TABLE {table_name} (id INTEGER, ФИ TEXT, "
                            f"Количество_баллов INTEGER, Оценка INTEGER, "
                            f"Процент_выполнения DOUBLE, PRIMARY KEY (id));")
            student_names = [''.join(i) for i in
                             list(cur.execute(f"select ФИ from {table_name}"))]
            length = len(student_names)
            count += length
            """Добавляю ученика в таблицу"""
            if student_name not in student_names:
                cur.execute(f"insert into {table_name} "
                            f"(Id, ФИ, Количество_баллов, Оценка, "
                            f"Процент_выполнения) values "
                            f"({count}, '{student_name}', "
                            f"{int(points)}, {int(mark)}, {float(percent)})")
                con.commit()
            self.msg = QMessageBox.information(self, 'Успех',
                                               'Ученник успешно добавлен'
                                               ' в базу данных',
                                               buttons=QMessageBox.Ok)
        else:
            """Если что-то введенно неправильно"""
            self.error = QMessageBox.critical(self, 'Ошибка',
                                              'Ошибка при добавлении '
                                              'ученика в базу данных',
                                              buttons=QMessageBox.Ok)


class GraphsOptions(QMainWindow):
    """Окно для ввода нужны данных для построения графика"""
    def __init__(self):
        super().__init__()
        uic.loadUi('ui/graphs_options.ui', self)
        self.setFixedSize(self.geometry().width(), self.geometry().height())
        self.btn_add.clicked.connect(self.add_qlistitem)
        self.btn_edit.clicked.connect(self.edit_qlistitem)
        self.btn_del.clicked.connect(self.delete_qlistitem)
        self.btn_build.clicked.connect(self.create_graph)
        self.btn_exit.clicked.connect(self.exit)
        self.graph_view.itemDoubleClicked.connect(self.parameter_action)

    def dialog_student(self):
        """Диалоговое окно для того, чтобы узнать имя студента"""
        return QInputDialog.getText(self, 'Введите имя', 'Введите имя ученика')

    def dialog_lesson(self):
        """Диалоговое окно для того, чтобы узнать название предмета"""
        return QInputDialog.getText(self, 'Название предмета',
                                    'Введите название предета')

    def dialog_class(self):
        """Диалоговое окно для того, чтобы узнать литра класса"""
        return QInputDialog.getText(self, 'Введите литру',
                                    'Введите литру класса/группы')

    def error_student(self, row):
        """Диалоговое окно, свидельствующее о том, что ученик не найден"""
        return QMessageBox.critical(self, 'Ошибка',
                                    f'Ошибка в строке{row}: ученик не найден',
                                    buttons=QMessageBox.Ok)

    def error_lesson(self, row):
        """Диалоговое окно, свидельствующее о том, что предмет не найден"""
        return QMessageBox.critical(self, 'Ошибка',
                                    f'Ошибка в строке {row}: предмет не найден',
                                    buttons=QMessageBox.Ok)

    def error_class(self, row):
        """Диалоговое окно, свидельствующее о том, что класс не найден"""
        return QMessageBox.critical(self, 'Ошибка',
                                    f'Ошибка в строке{row}: класс не найден',
                                    buttons=QMessageBox.Ok)

    def parameter_action(self):
        """Функция определяющая действие, которое нужно проводить над элементом"""
        self.action, isPressed = QInputDialog.getItem(self, "Выбор",
                                                      "Выберите действие",
                                                      ("Редактировать элемент",
                                                       "Удалить элемент"),
                                                      0, False)
        if isPressed:
            if self.action == 'Редактировать элемент':
                self.edit_qlistitem()  # Вызываем функцию редактирования элемента
            else:
                self.delete_qlistitem()  # Вызываем функцию удаления элемента

    def get_parameter_with_dialogues(self):
        '''Функция, создающая QListWidgetItem и взаимодействует с пользователем'''
        self.input_graphs_param1, okBtnPressed = \
            QInputDialog.getItem(self, "Выбор",
                                 "Выберите параметр 1",
                                 ("Ученик", "Класс", "Параллель"),
                                 0, False)  # Запрашиваем 1-й параметр
        if okBtnPressed:
            if self.input_graphs_param1 == 'Ученик':
                self.function_call = self.dialog_student
                # Назначаем перемнной self.function_call функцию выбора ученика
            elif self.input_graphs_param1 == 'Класс':
                self.function_call = self.dialog_class
                # Назначаем перемнной self.function_call функцию выбора класса
            else:
                self.function_call = lambda a=None: ('%', True)
                # Не назначаем перемнной self.function_call функцию,
                # т.к. параллель - совокупность всех классов
            self.param1, isPressed = self.function_call()
            # Вызываем переменную с функцией
            if isPressed:
                self.param1_string = self.input_graphs_param1[::]
                self.param1_string += '\t' + self.param1
                # Создаём строку с данными о 1-ом параметре
                self.input_graphs_param2, okBtnPressed = \
                    QInputDialog.getItem(self, "Выбор",
                                         "Выберите параметр 2",
                                         ("Предмет", "Предметы"),
                                         0, False) # Запрашиваем 2-й параметр
                if okBtnPressed:
                    if self.input_graphs_param2 == 'Предмет':
                        self.function_call = self.dialog_lesson
                        # Назначаем self.function_call функцию выбора предмета
                    else:
                        self.function_call = lambda a=None: ('%', True)
                        # Не назначаем self.function_call функцию,
                        # т.к. считываем все предметы
                    self.param2, isPressed = self.function_call()
                    # Вызываем переменную с функцией
                    if isPressed:
                        self.param2_string = self.input_graphs_param2[::]
                        self.param2_string += '\t' + self.param2
                        # Создаём строку с данными о 2-ом параметре
                        return QListWidgetItem(self.param1_string + '\t'
                                               + self.param2_string)
                        # Соединяем параметры и возвращаем объект класса QListWidget
                        # с содержимым 1 + 2 параметра.

    def add_qlistitem(self):
        '''Функция для добавления элемента в таблицу'''
        adding_string = self.get_parameter_with_dialogues()
        self.graph_view.addItem(adding_string)

    def edit_qlistitem(self):
        '''Функция для редактирования параметра в таблице'''
        self.item_manager = self.graph_view.selectedItems()[0]
        self.item_number = self.graph_view.row(self.item_manager)
        self.graph_view.takeItem(self.item_number)
        # Сначала убираем нужный нам элемент,
        self.graph_view.insertItem(self.item_number,
                                   self.get_parameter_with_dialogues())
        # а потом вставляем на то же место элемент, созданный с помощью
        # get_parameter_with_dialogues()
        self.update()

    def delete_qlistitem(self):
        '''Функция для удаления элемента из таблицы'''
        self.item_manager = self.graph_view.selectedItems()[0]
        self.item_number = self.graph_view.row(self.item_manager)
        self.graph_view.takeItem(self.item_number)  # Убираем нужный нам элемент
        self.update()

    def exit(self):
        """Открывает главное окно"""
        self.main_window = MainProjectsWindow()
        self.main_window.show()
        self.hide()

    def draw_graphs(self, x_label, *axis):
        '''Функция для непосредственного рисования графиков'''
        fig, axs = plt.subplots(1, 1, squeeze=False)  # Создаём холст
        for i, elem in enumerate(axis):
            axs[0, 0].plot(x_label, list(elem.values())[0],
                           label=list(elem.keys())[0], marker='o')
            # Рисуем кривую, соответствующую заданным точкам,
            # отмеченными жирными точками
            for j in range(len(list(elem.values())[0])):
                x, y = x_label[j], list(elem.values())[0][j]
                plt.annotate(str(y) + '%',
                             xy=(x, y),
                             xytext=(x, y + 1))  # Пишем значение рядом с точками
        fig.canvas.set_window_title('Результаты экзаменов')  # Меняем название окна
        plt.xlabel('Месяцы')  # Создаём надпись по оси Х
        plt.ylabel('% выполнения')  # Создаём надпись по оси Y
        plt.title("Статистика")  # Создаём название холста
        plt.legend()  # Включаем условные обозначения
        plt.show()  # Показываем таблицу

    def create_graph(self):
        """Функция преобразует данные из таблицы в данные для функции draw_graphs"""
        self.graph_options = [self.graph_view.item(i).text()
                              for i in range(self.graph_view.model().rowCount())]
        # Получаем строки всех элементов в таблице
        label_x = []
        label_y = []
        for counter, i in enumerate(self.graph_options):
            class1, param1, class2, param2 = i.split('\t')
            if class1 == "Ученик":
                con = sqlite3.connect('results.db')  # Подключаемся к таблице
                cur = con.cursor()
                table_results = []
                table_names = [''.join(i) for i in
                               list(cur.execute(f"select name from "
                                                "sqlite_master where type='table' "
                                                "and name like "
                                                f"'{param2}%'"))]
                # Получаем название таблицы, содержащей
                # данного ученика
                for i in table_names:
                    res = cur.execute(f'select Процент_выполнения from {i} where '
                                      f'(ФИ = "{param1}")').fetchall()
                    # Получаем проценты выполнения ученика
                    if res:
                        table_results.append(res[0][0])
                        # Добавляем в таблицу с данными данные из таблицы

                if not table_names or not table_results:
                    if not table_names:  # Если нет таблиц с данным уроком
                        self.err = self.error_lesson(counter + 1)
                    elif not table_results:  # Если нет таблиц с данным учеником
                        self.err = self.error_student(counter + 1)
                else:
                    label_x += [i.split('_')[1] for i in table_names]
                    label_y.append({param1: table_results})
            elif class1 == 'Класс':
                con = sqlite3.connect('results.db')  # Подключаемся к таблице
                cur = con.cursor()
                table_results = []
                table_class = cur.execute(f"select name from "
                                          "sqlite_master where type='table' "
                                          "and name like "
                                          f"'%{param1}'").fetchall()
                # Получаем название таблицы, содержащей данный класс
                table_lesson = cur.execute(f"select name from "
                                           "sqlite_master where type='table' "
                                           "and name like "
                                           f"'{param2}%'").fetchall()
                # Получаем название таблицы, содержащей данный урок
                table_names = [''.join(i) for i in
                               list(cur.execute(f"select name from "
                                                "sqlite_master where type='table' "
                                                "and name like "
                                                f"'{param2}%{param1}'"))]
                # Получаем название таблицы, содержащей нужные данные
                for i in range(len(table_names)):
                    """Проверяю есть ли такая таблица в списке таблиц"""
                    summa = cur.execute(f'select Процент_выполнения from '
                                        f'{table_names[i]}').fetchall()
                    if summa:
                        table_results.append(sum([i[0] for i in summa])
                                             / len(summa))
                        # Добавляем значения
                        # в список значений

                if not table_names:
                    if not table_lesson:  # Если нет таблиц с данным уроком
                        self.err = self.error_lesson(counter + 1)
                    elif not table_class:  # Если нет таблиц с данным классом
                        self.err = self.error_class(counter + 1)
                else:
                    label_x += [i.split('_')[1] for i in table_names]
                    label_y.append({param1: table_results})
            else:
                con = sqlite3.connect('results.db')
                cur = con.cursor()
                """Делаю список таблиц"""
                table_results = []
                table_names = [''.join(i) for i in
                               list(cur.execute(f"select name from "
                                                "sqlite_master where type='table' "
                                                "and name like "
                                                f"'{param2}%'"))]
                counter = 0
                count = 0
                month_cur = ''
                for i in range(len(table_names)):
                    month_temp = table_names[i].split('_')[1]
                    """Проверяю есть ли такая таблица в списке таблиц"""
                    coun = [i[0] for i in cur.execute(f'select Процент_выполнения '
                                                      f'from {table_names[i]}'
                                                      f'').fetchall()]
                    if coun:
                        if month_temp == month_cur or month_cur == '':
                            counter += sum(coun)
                            count += len(coun)
                        if month_temp != month_cur:
                            table_results.append(counter / count)
                            month_cur = month_temp[::]
                if not table_names:
                    self.err = self.error_lesson(counter + 1)
                else:
                    label_x += [i.split('_')[1] for i in table_names]
                    label_y.append({param1: table_results})
        label_x = sorted(set(label_x), key=lambda a: MONTHS[a])
        if label_y:
            self.draw_graphs(label_x, *label_y)  # Строим график


class Data(QMainWindow):
    """Окно вывода таблицы"""

    def __init__(self):
        super().__init__()
        uic.loadUi('ui/Data.ui', self)
        self.setFixedSize(self.geometry().width(), self.geometry().height())
        self.back_b2.clicked.connect(self.return_main_window)
        self.print_table.clicked.connect(self.return_table)

    def return_main_window(self):
        """Открывает главное окно"""
        self.main_window = MainProjectsWindow()
        self.main_window.show()
        self.hide()

    def return_table(self):
        """Вывод таблицы"""
        letter = self.letter1.text()
        subject = self.subject.text()
        month = self.month.text()
        table_name = f'{subject}_{month}_{letter}'
        con = sqlite3.connect('results.db')
        cur = con.cursor()
        table_names = [''.join(i) for i in
                       list(cur.execute("select name from "
                                        "sqlite_master where type='table'"))]
        if table_name not in table_names:
            """Если такой таблицы не существует"""
            self.list.clear()
            error = QListWidgetItem('Ошибка при выводе таблицы')
            self.list.addItem(error)
        else:
            """Вывожу таблицу"""
            data = [[str(j) for j in list(i)]
                    for i in list(cur.execute(f'select * from {table_name}'))]
            self.list.clear()
            self.list.addItem(f' ID              '
                              f'Фамилия_Имя\t   '
                              f'Количество_баллов\tОценка\tПроцент_выполнения')
            for i in data:
                main_str = f"{i[0].rjust(3, ' ')}        " \
                           f"{i[1].rjust(20, ' ')}\t\t      " \
                           f"  {i[2].rjust(2, ' ')}\t           {i[3]}" \
                           f"\t\t      {i[4].rjust(4, ' ')}"
                self.list.addItem(main_str)

    """Это надо удалить, когда будем сдавать"""
    def error(self):
        raise Exception('Неожиданная ошибкочка')


def my_excepthook(type, value, tback):
    a = QMessageBox.critical(windows, "Упс... Ошибка", str(value) +
                             "\nЗаскринь и отправь мне",
                             QMessageBox.Cancel)

    sys.__excepthook__(type, value, tback)


sys.excepthook = my_excepthook

clear_db()
app = QApplication(sys.argv)
windows = IntroducingWindow()
windows.show()
sys.exit(app.exec_())
