from dll_fetch import *
from global_import import *
from graphs_window import GraphsWindow
from additional_classes import *

temperature_massive = []
root_for_timer = None
text_for_timer = "Процессор"
cpu_cores = None
db = getcwd() + '\\aipcdb.db'

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        fetch_stats(HardwareHandle, "CPU", "Temperature", temperature_massive)
        global cpu_cores
        cpu_cores = len(temperature_massive)
        self.setWindowTitle("AInPC")
        self.setFixedSize(QSize(830, 500))    
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(192, 192, 192))
        
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Вид")
        menubar_height = self.menuBar().height()
        open_graphs = QAction("Графики", self)
        create_report = QAction("Отчет", self)
        open_cmd = QAction("Командная строка", self)
        open_powershell = QAction("Windows PowerShell", self)
        open_graphs.triggered.connect(self.open_graphs_window)
        open_cmd.triggered.connect(self.open_cmd)
        open_powershell.triggered.connect(self.open_powershell)
        create_report.triggered.connect(self.create_report)
        file_menu.addAction(open_graphs)
        file_menu.addAction(open_cmd)
        file_menu.addAction(open_powershell)
        file_menu.addAction(create_report)

        self.my_thread = WorkerThread()
        self.my_thread.mysignal.connect(self.on_change, Qt.ConnectionType.QueuedConnection)

        self.tree_widget = QTreeWidget(self)
        self.tree_widget.setGeometry(10, menubar_height, 150, self.height() - menubar_height - 10)
        self.tree_widget.setHeaderHidden(True)
        self.setup_tree()
        self.tree_widget.itemSelectionChanged.connect(self.on_item_selected)
        self.tree_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

        self.table_widget = QTableWidget(self)
        self.table_widget.setGeometry(self.tree_widget.width() + 20 + 10, menubar_height, self.width() - self.tree_widget.width() - 20 - 10 - 10, self.height() - menubar_height - 10)
        self.table_widget.verticalHeader().setVisible(False)
        self.table_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table_widget.verticalHeader().setVisible(False)
        self.table_widget.setShowGrid(False)
        self.table_widget.horizontalHeader().setStretchLastSection(True)
        self.table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table_widget.verticalHeader().setDefaultSectionSize(1)
        self.table_widget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_widget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.setup_table(9)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_table_with_timer)

        asd = "Krappa123"
        #self.table_widget.setItem(5, 1, QTableWidgetItem(f"\u00A0{asd}")) # - важно для дальнейшей красоты
        self.show()

    def fix_table(self):
        for row in range(self.table_widget.rowCount()):
            for col in range(self.table_widget.columnCount()):
                if (self.table_widget.rowSpan(row, col) > 1) or (self.table_widget.columnSpan(row, col) > 1):
                    self.table_widget.setSpan(row, col, 1, 1)
        self.table_widget.clearContents()

    def setup_table(self, row_count):
        column_names = ["Поле", "Описание"]
        self.table_widget.setRowCount(row_count)
        self.table_widget.setColumnCount(2)
        self.table_widget.setHorizontalHeaderLabels(column_names)

    def setup_tree(self):
        root_item_cpu = QTreeWidgetItem(self.tree_widget, ["Процессор"])               
        child_item1_cpu = QTreeWidgetItem(root_item_cpu, ["Температура"])
        child_item2_cpu = QTreeWidgetItem(root_item_cpu, ["Загрузка"])
        сhild_item3_cpu = QTreeWidgetItem(root_item_cpu, ["Частота"])
        child_item4_cpu = QTreeWidgetItem(root_item_cpu, ["Напряжение"])

        root_item_gpu = QTreeWidgetItem(self.tree_widget, ["Видеокарта"])
        child_item1_gpu = QTreeWidgetItem(root_item_gpu, ["Температура"])
        child_item2_gpu = QTreeWidgetItem(root_item_gpu, ["Загрузка"])
        сhild_item3_gpu = QTreeWidgetItem(root_item_gpu, ["Частота"])
        child_item4_gpu = QTreeWidgetItem(root_item_gpu, ["Напряжение"])
        child_item5_gpu = QTreeWidgetItem(root_item_gpu, ["Память"])

        root_item_ram = QTreeWidgetItem(self.tree_widget, ["Оперативная память"])
        child_item1_ram = QTreeWidgetItem(root_item_ram, ["Числовая информация"])

        root_item_mb = QTreeWidgetItem(self.tree_widget, ["Материнская плата"])
        mb_manufacturer_cmd = str(os.popen("wmic baseboard get manufacturer").read().encode()).split("\\n\\n")
        if mb_manufacturer_cmd[1].rstrip() != "Notebook":
            child_item1_mb = QTreeWidgetItem(root_item_mb, ["Температура"])
            child_item2_mb = QTreeWidgetItem(root_item_mb, ["Вольтаж"])
            child_item3_mb = QTreeWidgetItem(root_item_mb, ["Вентиляторы"])
        else:
            child_item1_mb = QTreeWidgetItem(root_item_mb, ["Вольтаж"])
        
        root_item_logical_disks = QTreeWidgetItem(self.tree_widget, ["Логические диски"])
        root_item_physical_disks = QTreeWidgetItem(self.tree_widget, ["Физические диски"])
        
        self.tree_widget.expandAll()                #Открывает все элементы дерева

    def update_table_with_timer(self):
        #self.my_thread.start()
        fetch_stats(HardwareHandle, root_for_timer, text_for_timer, temperature_massive)
        if root_for_timer == "CPU":
            if text_for_timer == "Temperature":                               
                self.fill_table_cpu_temp()
            elif text_for_timer == "Load":
                self.fill_table_cpu_load()
            elif text_for_timer == "Clock":
                self.fill_table_cpu_clock()
            elif text_for_timer == "Power":
                self.fill_table_cpu_power()
        elif root_for_timer == "GpuNvidia":
            if text_for_timer == "Temperature":
                self.fill_table_gpu_temp()
            elif text_for_timer == "Load":
                self.fill_table_gpu_load()
            elif text_for_timer == "Clock":
                self.fill_table_gpu_clock()
            elif text_for_timer == "Power":
                self.fill_table_gpu_power()
            elif text_for_timer == "SmallData":
                self.fill_table_gpu_memory()
        elif root_for_timer == "RAM":
            if text_for_timer == "Load":
                self.fill_table_ram_nums()
        elif root_for_timer == "SuperIO":
            if text_for_timer == "Temperature":
                self.fill_table_mb_temp()
            elif text_for_timer == "Voltage":
                self.fill_table_mb_voltage()
            elif text_for_timer == "Control":
                self.fill_table_mb_fans()

    def on_change(self, result):
        global temperature_massive
        temperature_massive = result

    def get_cpu_info(self, name):
        try:
            conn = sqlite3.connect(db)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM cpu_info WHERE cpu_name = ?", (name,))
            return cursor.fetchone()
        except sqlite3.Error as error:
            message_box = QMessageBox()
            message_box.setWindowTitle("Ошибка")
            message_box.setText(f"Возникла ошибка в ходе получения информации о процессоре {name}, пожалуйста, попробуйте открыть вкладку повторно или перезапустите программу.")
            message_box.setIcon(QMessageBox.Icon.Warning)
            message_box.exec()
        finally:
            if conn:
                conn.close()

    def fill_table_cpu_temp(self):
        for row in range(cpu_cores - 1):
            for item in temperature_massive:
                if row == item[0]:
                    item[1] = item[1].replace("CPU Core", "Ядро процессора")
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(item[2]) + '\u00B0C'))
                    break
        
        self.table_widget.setSpan(cpu_cores - 1 , 0, 1, 2)

        self.table_widget.setItem(cpu_cores, 0, QTableWidgetItem(str(temperature_massive[cpu_cores - 1][1]) + '\t'))
        self.table_widget.setItem(cpu_cores, 1, QTableWidgetItem(str(temperature_massive[cpu_cores - 1][2]) + '\u00B0C'))

    def fill_table_cpu_load(self):
        for row in range(cpu_cores):
            for item in temperature_massive:
                if item[0] == row or cpu_cores == len(temperature_massive):
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + "%"))
                    break
                else:
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(0) + "%"))
                    continue
    
    def fill_table_cpu_clock(self):
        for row in range(cpu_cores - 1):
            for item in temperature_massive:
                if row + 1 == item[0]:
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + ' GHz'))
                    break

        self.table_widget.setSpan(cpu_cores - 1 , 0, 1, 2)

        self.table_widget.setItem(cpu_cores, 0, QTableWidgetItem(str(temperature_massive[cpu_cores - 1][1]) + "\t"))
        self.table_widget.setItem(cpu_cores, 1, QTableWidgetItem(str(round(temperature_massive[cpu_cores - 1][2], 2)) + ' MHz'))

    def fill_table_cpu_power(self):
        i = 0
        for item in temperature_massive:
            self.table_widget.setItem(i, 0, QTableWidgetItem(str(item[1]) + '\t'))
            self.table_widget.setItem(i, 1, QTableWidgetItem(str(round(item[2], 2)) + ' W'))
            i = i + 1
    
    def get_gpu_info(self, name):
        try:
            conn = sqlite3.connect(db)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM gpu_info WHERE gpu_name = ?", (name,))
            return cursor.fetchone()
        except sqlite3.Error as error:
            message_box = QMessageBox()
            message_box.setWindowTitle("Ошибка")
            message_box.setText(f"Возникла ошибка в ходе получения информации о видеокарте {name}, пожалуйста, попробуйте открыть вкладку повторно или перезапустите программу.")
            message_box.setIcon(QMessageBox.Icon.Warning)
            message_box.exec()
        finally:
            if conn:
                conn.close()

    def fill_table_gpu_temp(self):
        self.table_widget.setItem(0, 0, QTableWidgetItem(str(temperature_massive[0][1]) + '\t'))
        self.table_widget.setItem(0, 1, QTableWidgetItem(str(temperature_massive[0][2]) + '\u00B0C'))
        if len(temperature_massive) == 1: 
            self.table_widget.setItem(1, 0, QTableWidgetItem("GPU Memory\t"))
            gpu_memory_temp = QTableWidgetItem(str("Отсутствует датчик температуры на видеокарте"))
            gpu_memory_temp.setFlags(gpu_memory_temp.flags() & ~Qt.ItemFlag.ItemIsEditable)
            gpu_memory_temp.setTextAlignment(Qt.AlignmentFlag.AlignHCenter)
            self.table_widget.setItem(1, 1, gpu_memory_temp)
        else:
            self.table_widget.setItem(1, 0, QTableWidgetItem(str(temperature_massive[1][1]) + '\t'))
            self.table_widget.setItem(1, 1, QTableWidgetItem(str(temperature_massive[1][2]) + '\u00B0C'))

    def fill_table_gpu_load(self):
        for row in range(3):
            for item in temperature_massive:
                table_gpu_name = self.table_widget.item(row, 0).text().strip()
                if item[1] == table_gpu_name:
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + "%"))
                    break
                else:
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(0) + "%"))
                    continue

    def fill_table_gpu_clock(self):
        for row in range(len(temperature_massive)):
            for item in temperature_massive:
                if row == item[0]:
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + ' MHz'))
                    break
    
    def fill_table_gpu_power(self):
        for row in range(len(temperature_massive)):
            for item in temperature_massive:
                self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + ' W'))

    def fill_table_gpu_memory(self):
        for row in range(len(temperature_massive)):
            for item in temperature_massive:
                if row == item[0] - 1:
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + ' MB'))

    def fill_table_ram_nums(self):
        if temperature_massive[0][1] == "Memory":
            temperature_massive[0][1] = "Загруженность памяти"
        self.table_widget.setItem(0, 0, QTableWidgetItem(str(temperature_massive[0][1]) + '\t'))
        self.table_widget.setItem(0, 1, QTableWidgetItem(str(round(temperature_massive[0][2], 2)) + ' %'))
        self.table_widget.setSpan(1, 0, 1, 2)
        fetch_stats(HardwareHandle, "RAM", "Data", temperature_massive)
        for row in range(1, len(temperature_massive) + 2):
            for item in temperature_massive:
                if row - 2 == item[0]:
                    if item[1] == "Used Memory":
                        item[1] = "Используемая память"
                    elif item[1] == "Available Memory":
                        item[1] = "Доступная память"
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + ' Gb'))
                    break

    def get_mb_info(self):
        mb_name = str(os.popen("wmic baseboard get product").read().encode()).split("\\n\\n")
        mb_name.pop(0)
        try:
            conn = sqlite3.connect(db)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM mb_info WHERE mb_name = ?", (mb_name[0].rstrip(),))
            return cursor.fetchone()
        except sqlite3.Error as error:
            message_box = QMessageBox()
            message_box.setWindowTitle("Ошибка")
            message_box.setText("Возникла ошибка в ходе получения информации о материнской плате, пожалуйста, попробуйте открыть вкладку повторно или перезапустите программу.")
            message_box.setIcon(QMessageBox.Icon.Warning)
            message_box.exec()
        finally:
            if conn:
                conn.close()

    def fill_table_mb_temp(self):
       for row in range(len(temperature_massive)):
            for item in temperature_massive:
                if row + 1 == item[0]:
                    item[1] = item[1].replace("Temperature", "Сенсор")
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(item[2]) + '\u00B0C'))
                    break

    def fill_table_mb_voltage(self):
        for row in range(len(temperature_massive)):
            for item in temperature_massive:
                if row == item[0]:
                    if "Voltage" in item[1]:
                        item[1] = item[1].replace("Voltage", "Вольтаж")
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + 'V'))
                    break

    def fill_table_mb_fans(self):
        lenght = len(temperature_massive)
        for row in range(len(temperature_massive)):
            for item in temperature_massive:
                if row == item[0]:
                    if "Fan Control" in item[1]:
                        item[1] = item[1].replace("Fan Control", "Контроллер вентиляторов")
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + '%'))
                    break
        self.table_widget.setSpan(lenght, 0, 1, 2)
        fetch_stats(HardwareHandle, "SuperIO", "Fan", temperature_massive)
        for row in range(lenght, len(temperature_massive) + lenght + 1):
            for item in temperature_massive:
                if row - lenght - 1 == item[0]:
                    if "Fan" in item[1]:
                        item[1] = item[1].replace("Fan", "Вентилятор")
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[1]) + '\t'))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(item[2], 2)) + ' об\мин'))
                    break
    
    def on_item_selected(self):
        global root_for_timer
        global text_for_timer
        root_for_timer = None
        text_for_timer = None
        selected_item = self.tree_widget.selectedItems()[0]
        selected_text = selected_item.text(0)
        root_item = selected_item.parent()
        if root_item is not None:
            root_text = root_item.text(0)
        else:
            root_text = None        
        if selected_item:
            self.initialize_table(root_text, selected_text)

    def initialize_table(self, root, text):
        global root_for_timer
        global text_for_timer
        self.fix_table()       
        if not root:
            self.timer.stop()
            if text == "Процессор":
                cpu_specs = ["Наименование процессора", "Сокет", "Базовая частота", "Кэш L1", "Кэш L2", "Кэш L3", "Кэш L4", "Фирма-производитель", "Информация от производителя"]
                cpu_names = str(os.popen("wmic cpu get name").read().encode()).split("\\n\\n")
                cpu_names.pop(0)
                for _ in range(2):
                    cpu_names.pop(-1)
                cpu_names = [item.rstrip() for item in cpu_names]
                cpu_count = len(cpu_names)
                self.setup_table((cpu_count * 9) + (cpu_count - 1))
                row = 0
                if cpu_count != 1:
                    for i in range(len(cpu_count)):
                        info = self.get_cpu_info(cpu_names[i])
                        self.table_widget.setItem(row, 0, QTableWidgetItem(str(cpu_specs[0]) + f"#{j + 1}\t"))
                        self.table_widget.setItem(row, 1, QTableWidgetItem(str(cpu_names[i])))
                        if info:
                            for j in range(row + 1, row + len(info) - 1):
                                self.table_widget.setItem(j, 0, QTableWidgetItem(str(cpu_specs[j - row]) + f"#{j + 1}\t"))
                                self.table_widget.setItem(j, 1, QTableWidgetItem(str(info[j - row])))
                            self.table_widget.setItem(row + 8, 0, QTableWidgetItem(str(cpu_specs[j - row]) + f"#{j + 1}\t"))
                            self.table_widget.setCellWidget(row + 8, 1, HyperlinkLabel(f"{info[j - row]}", f"{info[j - row]}"))
                        else:
                            for j in range(row + 1, row + 9):
                                self.table_widget.setItem(j, 0, QTableWidgetItem(str(cpu_specs[j - row]) + f"#{j + 1}\t"))
                                self.table_widget.setItem(j, 1, QTableWidgetItem(str("-")))
                        row += 10
                else:
                    info = self.get_cpu_info(cpu_names[0])
                    self.table_widget.setItem(0, 0, QTableWidgetItem(str(cpu_specs[0]) + "\t"))
                    self.table_widget.setItem(0, 1, QTableWidgetItem(str(cpu_names[0])))
                    if info:
                        for i in range(1, len(info) - 1):
                            self.table_widget.setItem(i, 0, QTableWidgetItem(str(cpu_specs[i]) + "\t"))
                            self.table_widget.setItem(i, 1, QTableWidgetItem(str(info[i])))
                        self.table_widget.setItem(8, 0, QTableWidgetItem(str(cpu_specs[8]) + "\t"))
                        self.table_widget.setCellWidget(8, 1, HyperlinkLabel(f"{info[8]}", f"{info[8]}"))
                    else:
                        for i in range(1, 9):
                            self.table_widget.setItem(i, 0, QTableWidgetItem(str(cpu_specs[i]) + "\t"))
                            self.table_widget.setItem(i, 1, QTableWidgetItem(str("-")))                   
            elif text == "Видеокарта":
                gpu_specs = ["Наименование графического ускорителя", "Базовая частота работы видеоядра", "Фирма-производитель видеокарты", "Информация от производителя видеокарты", "Официальный драйвер видеокарты", "Текущая версия драйвера видеокарты"] 
                gpu_names = str(os.popen("wmic path win32_VideoController get name").read().encode()).split("\\n\\n")
                gpu_names.pop(0)
                for _ in range(2):
                    gpu_names.pop(-1)
                gpu_names = [item.rstrip() for item in gpu_names]
                gpu_count = len(gpu_names)
                self.setup_table((gpu_count * 6) + (gpu_count - 1))
                row = 0
                if gpu_count != 1:
                    for i in range(len(gpu_count)):
                        info = self.get_gpu_info(gpu_names[i])
                        gpu_driver_info = str(os.popen("wmic path win32_VideoController get DriverVersion").read().encode()).split("\\n\\n")
                        gpu_driver_info.pop(0)
                        gpu_driver_info_short = gpu_driver_info[0].rstrip().replace(".", "")
                        self.table_widget.setItem(row, 0, QTableWidgetItem(str(gpu_specs[0]) + f"#{j + 1}\t"))
                        self.table_widget.setItem(row, 1, QTableWidgetItem(str(gpu_names[i])))
                        if info:
                            for j in range(row + 1, len(info) - 2):
                                self.table_widget.setItem(j, 0, QTableWidgetItem(str(gpu_specs[j - row]) + f"#{j + 1}\t"))
                                self.table_widget.setItem(j, 1, QTableWidgetItem(str(info[j - row])))
                            for j in range(row + 3, row + 5):
                                self.table_widget.setItem(j, 0, QTableWidgetItem(str(gpu_specs[j - row]) + f"#{j + 1}\t"))
                                hyperlink = HyperlinkLabel(f"{info[j - row]}", f"{info[j - row]}")
                                self.table_widget.setCellWidget(j, 1, hyperlink)
                        else:
                            for j in range(row + 1, row + 5):
                                self.table_widget.setItem(j, 0, QTableWidgetItem(str(gpu_specs[j - row]) + f"#{j + 1}\t"))
                                self.table_widget.setItem(j, 1, QTableWidgetItem(str("-")))
                        self.table_widget.setItem(j + 1, 0, QTableWidgetItem(str(gpu_specs[5]) + f"#{j + 1}\t"))
                        self.table_widget.setItem(j + 1, 1, QTableWidgetItem(str(gpu_driver_info[0].rstrip() + " \ " + gpu_driver_info_short[-5:-2] + "." + gpu_driver_info_short[-2:])))
                        row += 7
                else:
                    info = self.get_gpu_info(gpu_names[0])
                    self.table_widget.setItem(0, 0, QTableWidgetItem(str(gpu_specs[0]) + "\t"))
                    self.table_widget.setItem(0, 1, QTableWidgetItem(str(gpu_names[0])))
                    gpu_driver_info = str(os.popen("wmic path win32_VideoController get DriverVersion").read().encode()).split("\\n\\n")
                    gpu_driver_info.pop(0)
                    gpu_driver_info_short = gpu_driver_info[0].rstrip().replace(".", "")
                    self.table_widget.setItem(5, 0, QTableWidgetItem(str(gpu_specs[5]) + "\t"))
                    self.table_widget.setItem(5, 1, QTableWidgetItem(str(gpu_driver_info[0].rstrip() + " \ " + gpu_driver_info_short[-5:-2] + "." + gpu_driver_info_short[-2:])))
                    if info: 
                        for i in range(1, len(info) - 2):
                            self.table_widget.setItem(i, 0, QTableWidgetItem(str(gpu_specs[i]) + "\t"))
                            self.table_widget.setItem(i, 1, QTableWidgetItem(str(info[i])))
                        for i in range(3, 5):
                            self.table_widget.setItem(i, 0, QTableWidgetItem(str(gpu_specs[i]) + "\t"))
                            hyperlink = HyperlinkLabel(f"{info[i]}", f"{info[i]}")
                            self.table_widget.setCellWidget(i, 1, hyperlink)
                    else:
                        for i in range(1, 5):
                            self.table_widget.setItem(i, 0, QTableWidgetItem(str(gpu_specs[i]) + "\t"))
                            self.table_widget.setItem(i, 1, QTableWidgetItem(str("-")))
            elif text == "Оперативная память":
                ram_manufacturer = str(os.popen("wmic memorychip get Manufacturer").read().encode()).split("\\n\\n")
                ram_speed = str(os.popen("wmic memorychip get Speed").read().encode()).split("\\n\\n")
                pagefile = str(os.popen("wmic pagefileset get InitialSize, MaximumSize, Name").read().encode()).split("\\n\\n")
                ram_manufacturer.pop(0)
                ram_speed.pop(0)
                pagefile.pop(0)
                row = 0
                for _ in range(2):
                    ram_manufacturer.pop(-1)
                    ram_speed.pop(-1)
                    pagefile.pop(-1)
                self.setup_table((len(ram_manufacturer) * 2) + (len(pagefile) * 3) + 1)
                if len(ram_manufacturer) != 1:
                    for i in range(len(ram_manufacturer)):
                        self.table_widget.setItem(row, 0, QTableWidgetItem(str(f"Фирма-производитель модуля памяти #{i + 1}\t")))
                        if ram_manufacturer[i].isdigit():
                            ram_manufacturer[i] = ram_manufacturer[i].rstrip() + "- производитель незарегистрирован"
                        self.table_widget.setItem(row, 1, QTableWidgetItem(str(ram_manufacturer[i].rstrip())))
                        self.table_widget.setItem(row + 1, 0, QTableWidgetItem(str(f"Частота работы модуля памяти #{i + 1}\t")))
                        self.table_widget.setItem(row + 1, 1, QTableWidgetItem(str(ram_speed[i].rstrip() + " МГц")))
                        row += 2
                else:
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str("Фирма-производитель модуля памяти\t")))
                    if ram_manufacturer[0].isdigit():
                        ram_manufacturer[0] = ram_manufacturer[0].rstrip() + "- производитель незарегистрирован"
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(ram_manufacturer[0].rstrip())))
                    self.table_widget.setItem(row + 1, 0, QTableWidgetItem(str("Частота работы модуля памяти\t")))
                    self.table_widget.setItem(row + 1, 1, QTableWidgetItem(str(ram_speed[0].rstrip() + " МГц")))             
                row += 1
                for i in range(len(pagefile)):
                    pagefile[i] = pagefile[i].split()
                if len(pagefile) != 1:
                    for i in range(len(pagefile)):
                        self.table_widget.setItem(row, 0, QTableWidgetItem(str(f"Адрес файла подкачки #{i + 1}\t")))
                        self.table_widget.setItem(row, 1, QTableWidgetItem(str(pagefile[i][2].replace("\\\\", "\\"))))
                        self.table_widget.setItem(row + 1, 0, QTableWidgetItem(str(f"Начальный размер файла подкачки #{i + 1}\t")))
                        self.table_widget.setItem(row + 1, 1, QTableWidgetItem(str(round(float(pagefile[i][0]) / 1024, 2)) + " Гб"))
                        self.table_widget.setItem(row + 2, 0, QTableWidgetItem(str(f"Текущий размер файла подкачки #{i + 1}\t")))
                        self.table_widget.setItem(row + 2, 1, QTableWidgetItem(str(round(float(pagefile[i][1]) / 1024, 2)) + " Гб"))
                        row += 3
                else:
                    self.table_widget.setItem(row, 0, QTableWidgetItem(str("Адрес файла подкачки\t")))
                    self.table_widget.setItem(row, 1, QTableWidgetItem(str(pagefile[0][2].replace("\\\\", "\\"))))
                    self.table_widget.setItem(row + 1, 0, QTableWidgetItem(str("Начальный размер файла подкачки\t")))
                    self.table_widget.setItem(row + 1, 1, QTableWidgetItem(str(round(float(pagefile[0][0]) / 1024, 2)) + " Гб"))
                    self.table_widget.setItem(row + 2, 0, QTableWidgetItem(str("Текущий размер файла подкачки\t")))
                    self.table_widget.setItem(row + 2, 1, QTableWidgetItem(str(round(float(pagefile[0][1]) / 1024, 2)) + " Гб"))
            elif text == "Материнская плата":
                self.setup_table(5)
                self.table_widget.setItem(0, 0, QTableWidgetItem(str("Наименование материнской платы\t")))
                self.table_widget.setItem(1, 0, QTableWidgetItem(str("Фирма-производитель\t")))
                self.table_widget.setItem(2, 0, QTableWidgetItem(str("Информация от производителя\t")))
                self.table_widget.setItem(3, 0, QTableWidgetItem(str("Официальный драйвер\t")))
                self.table_widget.setItem(4, 0, QTableWidgetItem(str("Версия BIOS\t")))             
                info = self.get_mb_info()
                if info:
                    self.table_widget.setItem(0, 1, QTableWidgetItem(str(info[0])))
                    hyperlink_man = HyperlinkLabel(f"{info[1]}", f"{info[1]}")
                    self.table_widget.setCellWidget(2, 1, hyperlink_man)
                    hyperlink_driver = HyperlinkLabel(f"{info[2]}", f"{info[2]}")
                    self.table_widget.setCellWidget(3, 1, hyperlink_driver)
                else:
                    self.table_widget.setItem(0, 1, QTableWidgetItem(str("-")))
                    self.table_widget.setItem(2, 1, QTableWidgetItem(str("-")))
                    self.table_widget.setItem(3, 1, QTableWidgetItem(str("-")))                   
                mb_info_cmd = str(os.popen("wmic baseboard get manufacturer").read().encode()).split("\\n\\n")
                mb_bios_cmd = str(os.popen("wmic bios get smbiosbiosversion").read().encode()).split("\\n\\n")
                mb_info_cmd.pop(0)
                mb_bios_cmd.pop(0)
                self.table_widget.setItem(1, 1, QTableWidgetItem(str(mb_info_cmd[0].rstrip())))
                self.table_widget.setItem(4, 1, QTableWidgetItem(str(mb_bios_cmd[0].rstrip())))
            elif text == "Логические диски":
                column_names = ["Том системы", "Свободное место", "Размер тома", "Загруженность", "Тип файловой системы"]
                disks_info = str(os.popen("wmic logicaldisk get deviceid, freespace, size, filesystem").read().encode()).split("\\n\\n")
                disks_info.pop(0)
                for _ in range(2):
                    disks_info.pop(-1)            
                self.table_widget.setRowCount(len(disks_info))
                self.table_widget.setColumnCount(5)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                for i in range(len(disks_info)):
                    current_disk = disks_info[i].split()
                    self.table_widget.setItem(i, 0, QTableWidgetItem(str(current_disk[0])))
                    self.table_widget.setItem(i, 1, QTableWidgetItem(str(round(float(current_disk[2]) / pow(1024, 3), 2)) + " Гб"))
                    self.table_widget.setItem(i, 2, QTableWidgetItem(str(round(float(current_disk[3]) / pow(1024, 3), 2)) + " Гб"))
                    self.table_widget.setItem(i, 3, QTableWidgetItem(str(round((float(current_disk[2]) / float(current_disk[3])) * 100, 2)) + "%"))
                    self.table_widget.setItem(i, 4, QTableWidgetItem(str(current_disk[1])))                          
            elif text == "Физические диски":
                column_names = ["Имя диска", "Всего места"]
                disks_info = str(os.popen("wmic diskdrive get model, size").read().encode()).split("\\n\\n")
                disks_info.pop(0)
                row = 0
                for _ in range(2):
                    disks_info.pop(-1)  
                self.table_widget.setRowCount(len(disks_info))
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                for item in disks_info:
                    item = item.rstrip()
                    for i in range(len(item) - 1, 0, -1):
                        if not item[i].isdigit():
                            self.table_widget.setItem(row, 0, QTableWidgetItem(str(item[0: i - 1].rstrip()) + "\t"))
                            self.table_widget.setItem(row, 1, QTableWidgetItem(str(round(float(item[i + 1:]) / pow(1024, 3), 2)) + " Гб"))
                            row += 1
                            break                     
        elif root == "Процессор":
            if text == "Температура":               
                fetch_stats(HardwareHandle, "CPU", "Temperature", temperature_massive)        
                root_for_timer = "CPU"
                text_for_timer = "Temperature"
                column_names = ["Устройство", "Температура"]               
                self.table_widget.setRowCount(cpu_cores + 1)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)              
                self.fill_table_cpu_temp()
            if text == "Загрузка":
                fetch_stats(HardwareHandle, "CPU", "Load", temperature_massive)
                root_for_timer = "CPU"
                text_for_timer = "Load"
                column_names = ["Устройство", "Загрузка"]
                self.table_widget.setRowCount(cpu_cores + 1)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.table_widget.setItem(0, 0, QTableWidgetItem("CPU Total\t"))
                for row in range(1, cpu_cores):
                    self.table_widget.setItem(row, 0, QTableWidgetItem(f"CPU Core #{row}\t"))
                self.fill_table_cpu_load()
            if text == "Частота":
                fetch_stats(HardwareHandle, "CPU", "Clock", temperature_massive)
                root_for_timer = "CPU"
                text_for_timer = "Clock"
                column_names = ["Устройство", "Частота"]
                self.table_widget.setRowCount(cpu_cores + 1)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_cpu_clock()
            if text == "Напряжение":
                fetch_stats(HardwareHandle, "CPU", "Power", temperature_massive)
                root_for_timer = "CPU"
                text_for_timer = "Power"
                column_names = ["Устройство", "Напряжение"]
                self.table_widget.setRowCount(len(temperature_massive))
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_cpu_power()
        elif root == "Видеокарта":
            if text == "Температура":
                fetch_stats(HardwareHandle, "GpuNvidia", "Temperature", temperature_massive)
                root_for_timer = "GpuNvidia"
                text_for_timer = "Temperature"
                column_names = ["Устройство", "Температура"]               
                self.table_widget.setRowCount(2)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_gpu_temp()
            elif text == "Загрузка":
                fetch_stats(HardwareHandle, "GpuNvidia", "Load", temperature_massive)
                root_for_timer = "GpuNvidia"
                text_for_timer = "Load"
                column_names = ["Устройство", "Загрузка"]
                self.table_widget.setRowCount(5)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)

                self.table_widget.setItem(0, 0, QTableWidgetItem("GPU Core\t"))
                self.table_widget.setItem(1, 0, QTableWidgetItem("GPU Frame Buffer\t"))
                self.table_widget.setItem(2, 0, QTableWidgetItem("GPU Memory\t"))

                self.fill_table_gpu_load()
            elif text == "Частота":
                fetch_stats(HardwareHandle, "GpuNvidia", "Temperature", temperature_massive)
                root_for_timer = "GpuNvidia"
                text_for_timer = "Clock"
                column_names = ["Устройство", "Частота"]               
                self.table_widget.setRowCount(2)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_gpu_clock()
            elif text == "Напряжение":
                fetch_stats(HardwareHandle, "GpuNvidia", "Power", temperature_massive)
                root_for_timer = "GpuNvidia"
                text_for_timer = "Power"
                column_names = ["Поле", "Напряжение"]
                self.table_widget.setRowCount(1)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_gpu_power()
            elif text == "Память":
                fetch_stats(HardwareHandle, "GpuNvidia", "SmallData", temperature_massive)
                root_for_timer = "GpuNvidia"
                text_for_timer = "SmallData"
                column_names = ["Поле", "Значение"]
                self.table_widget.setRowCount(3)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_gpu_memory()
        elif root == "Оперативная память":
            if text == "Числовая информация":
                fetch_stats(HardwareHandle, "RAM", "Data", temperature_massive)                                    
                root_for_timer = "RAM"
                text_for_timer = "Load"
                column_names = ["Поле", "Значение"]
                count = len(temperature_massive)
                fetch_stats(HardwareHandle, "RAM", "Load", temperature_massive) 
                count += len(temperature_massive)
                self.table_widget.setRowCount(count + 1)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_ram_nums()
                count = 0
        elif root == "Материнская плата":
            if text == "Температура":
                fetch_stats(HardwareHandle, "SuperIO", "Temperature", temperature_massive)
                root_for_timer = "SuperIO"
                text_for_timer = "Temperature"
                column_names = ["Сенсор", "Значение"]
                self.table_widget.setRowCount(len(temperature_massive))
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_mb_temp()
            elif text == "Вольтаж":
                fetch_stats(HardwareHandle, "SuperIO", "Voltage", temperature_massive)
                root_for_timer = "SuperIO"
                text_for_timer = "Voltage"
                column_names = ["Сенсор", "Значение"]
                self.table_widget.setRowCount(len(temperature_massive))
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_mb_voltage()
            elif text == "Вентиляторы":
                fetch_stats(HardwareHandle, "SuperIO", "Fan", temperature_massive)                                    
                root_for_timer = "SuperIO"
                text_for_timer = "Control"
                count = len(temperature_massive)
                fetch_stats(HardwareHandle, "SuperIO", "Control", temperature_massive) 
                count += len(temperature_massive)

                column_names = ["Поле", "Значение"]
                self.table_widget.setRowCount(count + 1)
                self.table_widget.setColumnCount(2)
                self.table_widget.setHorizontalHeaderLabels(column_names)
                self.fill_table_mb_fans()
                count = 0
        self.timer.start(1000)

    def open_graphs_window(self):
        self.graphs_window = GraphsWindow(HardwareHandle, cpu_cores)
        self.graphs_window.show()
    
    def open_cmd(self):
        os.system('start cmd.exe')

    def open_powershell(self):
        os.system('start powershell.exe')
    
    def create_report(self):
        path = getcwd()
        os.system(f'start cmd.exe /C msinfo32 /report "{path}\\AInPC System report.txt"')
        
@main_requires_admin
def main(): 
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec())

if __name__ == "__main__":
    try:
        HardwareHandle = initialize_openhardwaremonitor()
        main()
    except PyWinError:
        exit()
