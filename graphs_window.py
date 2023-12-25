from dll_fetch import *
from global_import import *

HardwareHandle = None
cpu_cores = None
values_list = []
names_list = []
graphs_list = []

class GraphsWindow(QMainWindow):
    def __init__(self, handle, cores_count):
        super().__init__()

        global HardwareHandle
        global cpu_cores
        HardwareHandle = handle
        cpu_cores = cores_count

        self.setWindowTitle("AInPC Графики")
        self.setFixedSize(QSize(900, 500))
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(192, 192, 192))
        self.create_graphs_values()

        central_widget = QWidget()
        layout = QVBoxLayout()

        scroll_area = QScrollArea(self)
        scroll_widget = QWidget(scroll_area)
        scroll_layout = QVBoxLayout(scroll_widget)

        self.plot_widgets = []
        print(len(graphs_list))
        n = 0
        if len(graphs_list) // 2:
            row_count = len(graphs_list) // 2 + 1
        else:
            row_count = len(graphs_list) // 2
        for _ in range(0, row_count):
            if n == len(graphs_list):
                break
            graph_row_widget = QWidget()
            row_layout = QHBoxLayout()
            for _ in range(2):
                name = graphs_list[n][0]
                plot_widget = pyqtgraph.PlotWidget()
                plot_widget.setMouseEnabled(x = False, y = False)
                plot_widget.setBackground("w")
                plot_widget.setFixedWidth(405)
                plot_widget.plotItem.setFixedWidth(395)
                plot_widget.setFixedHeight(220)
                plot_widget.plotItem.setFixedHeight(215)
                plot_widget.plotItem.setTitle(name, color = (0, 0, 0)) 
                plot_widget.plotItem.setLabel('left', 'Значение', color = (0, 0, 0))
                plot_widget.plotItem.setLabel('bottom', 'Время', color = (0, 0, 0))
                x_axis = plot_widget.getAxis("bottom")
                x_axis.setStyle(showValues = False)
                if "Температура" in name:
                    pen = pyqtgraph.mkPen(color = (255, 0, 0))
                elif "Частота" in name:
                    pen = pyqtgraph.mkPen(color = (0, 0, 0))
                elif "Загруженность" in name:
                    pen = pyqtgraph.mkPen(color = (0, 255, 0))
                else:
                    pen = pyqtgraph.mkPen(color = (0, 0, 255))
                plotx = plot_widget.plot([], [], pen = pen)
                self.plot_widgets.append((plot_widget, plotx))
                row_layout.addWidget(plot_widget)
                n = n + 1
                if n == len(graphs_list):
                    break
            graph_row_widget.setLayout(row_layout)
            scroll_layout.addWidget(graph_row_widget)
        
        scroll_widget.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)

        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        """rows = -(-len(graphs_list) // 2)  # Определение числа строк для 2 столбцов
        for i, data in enumerate(graphs_list):
            row = i // 2
            col = i % 2
            plot_widget = pyqtgraph.PlotWidget()
            plot_widget.setMouseEnabled(x = False, y = False)
            plot_widget.setBackground("w")
            plot_widget.setFixedWidth(370)
            plot_widget.plotItem.setFixedWidth(360)
            plot_widget.setFixedHeight(200)
            plot_widget.plotItem.setFixedHeight(195)

            layout.addWidget(plot_widget, row, col)
            self.plot_widgets.append(plot_widget)
            
            self.plot_widgets[i].setTitle(data[0]) 
            self.plot_widgets[i].setLabel('left', 'Значение')
            self.plot_widgets[i].setLabel('bottom', 'Время')

            x_axis = plot_widget.getAxis("bottom")
            x_axis.setStyle(showValues = False)"""
            
            
        self.timer = pyqtgraph.QtCore.QTimer(self)
        self.timer.timeout.connect(self.update_graphs)
        self.timer.start(1000)

    def create_graphs_values(self):
        global values_list
        global names_list
        global graphs_list
        for _ in range(3):
            fetch_stats_for_graphs(HardwareHandle, values_list, names_list)
            #print("fetch is done")
        for i in range(cpu_cores - 1):
            if f'Загруженность ядра процессора #{i + 1}' not in names_list:
                values_list.append([f'Загруженность ядра процессора #{i + 1}', 0.0])
                names_list.append(f'Загруженность ядра процессора #{i + 1}')
            if f'Температура ядра процессора #{i + 1}' not in names_list:
                for j in range(len(names_list)):
                    if 'Температура ядра процессора' in names_list[i]:
                        n = j
                        break
                values_list.append([f'Температура ядра процессора #{i + 1}', values_list[n][1]])
                names_list.append(f'Температура ядра процессора #{i + 1}')
        #print("create is done")
        graphs_list = self.sort_values(values_list, names_list)
        #print("sort is done")

    def update_graphs(self):
        #print("updating")
        global values_list
        global graphs_list
        fetch_stats_for_graphs(HardwareHandle, values_list, names_list)
        self.correlate_values(graphs_list, values_list, names_list)
        """for i, data in enumerate(graphs_list):
            self.plot_widgets[i].clear()
            self.plot_widgets[i].plot(data[1:])"""
        i = 0
        for self.plot_widget, plot in self.plot_widgets:
            #print(self.plot_widget, plot)
            x_ax = [i for i in range(1, len(graphs_list[i]))]
            plot.setData(x_ax, graphs_list[i][1:])
            i += 1
        #[print(len(item)) for item in graphs_list]
        
    def sort_values(self, values_list, names_list):
        temp_list = []
        temp_list += (sorted([sublist for sublist in values_list if 'Температура ядра процессора' in sublist[0]], key=lambda x: int(x[0].split("#")[1])))
        temp_list += (sorted([sublist for sublist in values_list if 'Загруженность ядра процессора' in sublist[0]], key=lambda x: int(x[0].split("#")[1])))
        temp_list += (sorted([sublist for sublist in values_list if 'Частота работы ядра процессора' in sublist[0]], key=lambda x: int(x[0].split("#")[1])))
        temp_list += [values_list[names_list.index('Частота работы шины процессора')]]
        if names_list.count('Температура ядра видеокарты') > 1:
            indexes = [index for index, item in enumerate(names_list) if item == 'Температура ядра видеокарты']
            temp_list += [values_list[index] for index in indexes]
        else:
            temp_list += [values_list[names_list.index('Температура ядра видеокарты')]]
        if 'Температура памяти видеокарты' in names_list:
            if names_list.count('Температура памяти видеокарты') > 1:
                indexes = [index for index, item in enumerate(names_list) if item == 'Температура памяти видеокарты']
                temp_list += [values_list[index] for index in indexes]
            else:
                temp_list += [values_list[names_list.index('Температура памяти видеокарты')]]
        if names_list.count('Частота работы ядра видеокарты') > 1:
            indexes = [index for index, item in enumerate(names_list) if item == 'Частота работы ядра видеокарты']
            temp_list += [values_list[index] for index in indexes]
        else:
            temp_list += [values_list[names_list.index('Частота работы ядра видеокарты')]]
        if 'Частота работы памяти видеокарты' in names_list:
            if names_list.count('Частота работы памяти видеокарты') > 1:
                indexes = [index for index, item in enumerate(names_list) if item == 'Частота работы памяти видеокарты']
                temp_list += [values_list[index] for index in indexes]
            else:
                temp_list += [values_list[names_list.index('Частота работы памяти видеокарты')]]    
        temp_list += [values_list[names_list.index('Загруженность оперативной памяти')]]  
        temp_list += (sorted([sublist for sublist in values_list if 'Температура сенсора материнской платы' in sublist[0]], key=lambda x: int(x[0].split("#")[1])))
        temp_list += (sorted([sublist for sublist in values_list if 'Скорость вращения вентилятора' in sublist[0]], key=lambda x: int(x[0].split("#")[1])))
        return temp_list

    def correlate_values(self, graphs_list, values_list, names_list):
        '''
        print("Graphs:")
        [print(item) for item in graphs_list]
        print("*" * 50)
        print("Names:")
        [print(item) for item in names_list]
        print("*" * 50)
        '''
        while values_list:
            for graphs_item in graphs_list:
                for values_item in values_list:          
                    if graphs_item[0] not in names_list:
                        if len(graphs_item) > 60:
                            del graphs_item[1]
                        graphs_list[graphs_list.index(graphs_item)].append(graphs_item[-1])
                        break
                    if values_item[0] == graphs_item[0]:
                        if len(graphs_item) > 60:
                            del graphs_item[1]
                        graphs_list[graphs_list.index(graphs_item)].append(values_item[1])
                        del values_list[values_list.index(values_item)]
                        break               
            '''print(values_list, end = "\n" + "-" * 50 + "\n")
            input()'''
        

    def closeEvent(self, event):
        self.timer.stop()