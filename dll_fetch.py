from os import getcwd
import clr
import hwtypes
hwtypes = ['Mainboard','SuperIO','CPU','RAM','GpuNvidia','GpuAti','TBalancer','Heatmaster','HDD']

def initialize_openhardwaremonitor():
    file = getcwd() + '\\Lib'
    clr.AddReference(file)

    from OpenHardwareMonitor import Hardware

    handle = Hardware.Computer()
    handle.MainboardEnabled = True
    handle.CPUEnabled = True
    handle.RAMEnabled = True
    handle.GPUEnabled = True
    handle.HDDEnabled = True
    handle.Open()
    return handle

def fetch_stats(handle, device, type_of_sensor, massive):
    massive.clear()
    for i in handle.Hardware:
        i.Update()
        for sensor in i.Sensors:
            parse_sensor(sensor, device, type_of_sensor, massive)
        for j in i.SubHardware:
            j.Update()
            for subsensor in j.Sensors:
                parse_sensor(subsensor, device, type_of_sensor, massive)
    return massive
    
def parse_sensor(sensor, device, type_of_sensor, massive):
    if sensor.Value:
        if str(sensor.SensorType) == type_of_sensor and device == hwtypes[sensor.Hardware.HardwareType]:
            massive.append([sensor.Index, sensor.Name, sensor.Value])

def fetch_stats_for_graphs(handle, values, names):
    values.clear()
    names.clear()
    for i in handle.Hardware:
        i.Update()
        for sensor in i.Sensors:
            parse_sensor_for_graphs(sensor, values, names)
        for j in i.SubHardware:
            j.Update()
            for subsensor in j.Sensors:
                parse_sensor_for_graphs(subsensor, values, names)

def parse_sensor_for_graphs(sensor, values, names):
    sensortypes = [['Temperature', 'SuperIO'], ['Fan', 'SuperIO'], ['Load', 'CPU'], ['Temperature', 'CPU'], ['Clock', 'CPU'], ['Load', 'RAM'], ['Temperature', 'Gpu'], ['Clock', 'Gpu']]
    not_allowed = ['CPU Total', 'CPU Package']
    if sensor.Value:
        for item in sensortypes:
            if str(sensor.SensorType) in item[0] and item[1] in hwtypes[sensor.Hardware.HardwareType]:
                name = sensor.Name
                if name in not_allowed:
                    continue
                elif 'Temperature' in name and str(sensor.SensorType) == 'Temperature' and hwtypes[sensor.Hardware.HardwareType] == 'SuperIO':
                    name = name.replace('Temperature', 'Температура сенсора материнской платы')
                elif 'Fan' in name and str(sensor.SensorType) == 'Fan' and hwtypes[sensor.Hardware.HardwareType] == 'SuperIO':
                    name = name.replace('Fan', 'Скорость вращения вентилятора')
                elif 'CPU Core' in name and str(sensor.SensorType) == 'Load':
                    name = name.replace('CPU Core', 'Загруженность ядра процессора')
                elif 'CPU Core' in name and str(sensor.SensorType) == 'Temperature':
                    name = name.replace('CPU Core', 'Температура ядра процессора')
                elif 'CPU Core' in name and str(sensor.SensorType) == 'Clock':
                    name = name.replace('CPU Core', 'Частота работы ядра процессора')
                elif 'Bus Speed' in name and str(sensor.SensorType) == 'Clock':
                    name = name.replace('Bus Speed', 'Частота работы шины процессора')
                elif 'Memory' in name and str(sensor.SensorType) == 'Load' and hwtypes[sensor.Hardware.HardwareType] == 'RAM':
                    name = name.replace('Memory', 'Загруженность оперативной памяти')
                elif 'GPU Core' in name and str(sensor.SensorType) == 'Temperature':
                    name = name.replace('GPU Core', 'Температура ядра видеокарты')
                elif 'GPU Memory' in name and str(sensor.SensorType) == 'Temperature':
                    name = name.replace('GPU Memory', 'Температура памяти видеокарты')
                elif 'GPU Core' in name and str(sensor.SensorType) == 'Clock':
                    name = name.replace('GPU Core', 'Частота работы ядра видеокарты')
                elif 'GPU Memory' in name and str(sensor.SensorType) == 'Clock':
                    name = name.replace('GPU Memory', 'Частота работы памяти видеокарты')
                values.append([name, float(sensor.Value)])
                names.append(name)