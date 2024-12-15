import pandas as pd
from datetime import datetime, timedelta
import random
from xlsxwriter import Workbook

# Константы
NUM_BUSES = 8 # Число автобусов
WORK_START = datetime.strptime("06:00", "%H:%M") # Начало рабочего дня
WORK_END = datetime.strptime("03:00", "%H:%M") + timedelta(days=1) # Конец рабочего дня
ROUTE_DURATION = timedelta(minutes=60) # Время, которое занимает один маршрут
ROUTE_VARIATION = 10 # Отклонение маршрута
PEAK_HOURS = [(7, 9), (17, 19)] # Час пик
LOAD_PEAK = 0.7 # Нагрузка в час пик
LOAD_OFF_PEAK = 0.3 # Нагрузка, когда нет час пика
LOAD_SATURDAY = 0.5 # Нагрузка в выхожные
LUNCH_INTERVAL = timedelta(minutes=15)  # Интервал между обедами

# Водители
drivers_A = []  # Водители типа А (сохраняются между днями)
drivers_B = []  # Водители типа Б (сбрасываются каждые три дня)
driver_A_id = 1
driver_B_id = 1

# Лог действий
actions = []
swaps = []

# Максимальное количество водителей
def max_drivers_in_a_day(drivers, current_max):
    return max(current_max, len(drivers))

# Проверяет, является ли текущий час часом пик.
def is_peak_hour(hour):
    return any(start <= hour < end for start, end in PEAK_HOURS)

# Возвращает время маршрута с учетом случайного отклонения.
def get_route_time():
    return ROUTE_DURATION + timedelta(minutes=random.randint(-ROUTE_VARIATION, ROUTE_VARIATION))

# Назначает водителя на маршрут.
def assign_driver(bus_id, current_time, end_time, is_saturday=False, is_sunday=False):
    global driver_A_id, driver_B_id

    # Для субботы и воскресенья водители типа A не работают
    if not is_saturday and not is_sunday:
        for driver in drivers_A:
            if driver["end_time"] <= current_time and driver["total_hours"] < 9:
                # Проверка окончания рабочего дня
                work_end_limit = driver["start_time"] + timedelta(hours=9)
                if current_time + ROUTE_DURATION > work_end_limit + timedelta(minutes=10):
                    continue

                # Проверка на обед
                if current_time >= driver["next_break"] and not driver["had_break"] and not is_peak_hour(current_time.hour):
                    if can_take_lunch(current_time):
                        actions.append({
                            "driver_id": driver["id"],
                            "driver_type": "A",
                            "action": "Обед",
                            "time": current_time.strftime("%H:%M"),
                            "duration": "1 час"
                        })
                        current_time += timedelta(hours=1)
                        driver["had_break"] = True

                driver["end_time"] = end_time
                driver["total_hours"] += (end_time - current_time).total_seconds() / 3600
                return "A", driver["id"]

    # Обработка водителей типа Б
    for driver in drivers_B:
        if driver["end_time"] <= current_time and driver["total_hours"] < 24:
            if driver["last_break"] + timedelta(hours=2) <= current_time and not is_peak_hour(current_time.hour):
                actions.append({
                    "driver_id": driver["id"],
                    "driver_type": "B",
                    "action": "Мини-обед",
                    "time": current_time.strftime("%H:%M"),
                    "duration": "15 минут"
                })
                current_time += timedelta(minutes=15)
                driver["last_break"] = current_time

            driver["end_time"] = end_time
            driver["total_hours"] += (end_time - current_time).total_seconds() / 3600
            return "B", driver["id"]

    # Если подходящего водителя нет, создаем нового
    if not is_saturday and not is_sunday and random.random() < 0.5:  # Тип A
        start_time = WORK_START + timedelta(minutes=random.randint(0, 240))  # Водители типа А могут начинать с 6:00 до 10:00
        driver = {
            "id": driver_A_id,
            "type": "A",
            "start_time": start_time,
            "end_time": end_time,
            "total_hours": (end_time - current_time).total_seconds() / 3600,
            "next_break": start_time + timedelta(hours=4),
            "had_break": False,
            "last_lunch": None
        }
        drivers_A.append(driver)
        driver_A_id += 1
        return "A", driver["id"]
    else:  # Тип B
        driver = {
            "id": driver_B_id,
            "type": "B",
            "end_time": end_time,
            "total_hours": (end_time - current_time).total_seconds() / 3600,
            "last_break": current_time
        }
        drivers_B.append(driver)
        driver_B_id += 1
        return "B", driver["id"]

# Проверяет, можно ли взять обед в текущий момент времени.
def can_take_lunch(current_time):
    last_lunch_times = [driver["last_lunch"] for driver in drivers_A if driver["last_lunch"] is not None]
    if not last_lunch_times:
        return True
    last_lunch = max(last_lunch_times)
    return current_time >= last_lunch + LUNCH_INTERVAL

# Создает расписание автобусов на указанный день.
def generate_schedule(day, is_saturday=False, is_sunday=False):
    current_time = WORK_START
    schedule = []

    while current_time < WORK_END:
        for bus_id in range(1, NUM_BUSES + 1):
            # Определяем нагрузку в зависимости от дня недели
            if is_saturday or is_sunday:
                load = LOAD_SATURDAY
            else:
                load = LOAD_PEAK if is_peak_hour(current_time.hour) else LOAD_OFF_PEAK

            active_buses = max(1, int(NUM_BUSES * load))

            if bus_id <= active_buses:
                route_time = get_route_time()
                end_time = current_time + route_time

                # Назначаем водителя с учетом дня недели
                driver_type, driver_id = assign_driver(bus_id, current_time, end_time, is_saturday=is_saturday, is_sunday=is_sunday)
                schedule.append({
                    "bus_id": bus_id,
                    "route": f"Маршрут {bus_id % 2 + 1}",
                    "start_time": current_time.strftime("%H:%M"),
                    "end_time": end_time.strftime("%H:%M"),
                    "driver_type": driver_type,
                    "driver_id": driver_id,
                    "load": f"{int(load * 100)}%"
                })

                # Добавление пересменки
                if random.random() < 0.2:  # Вероятность смены водителя
                    new_driver_type, new_driver_id = assign_driver(bus_id, end_time, end_time + timedelta(minutes=15), is_saturday=is_saturday, is_sunday=is_sunday)
                    swaps.append({
                        "bus_id": bus_id,
                        "time": end_time.strftime("%H:%M"),
                        "old_driver_id": driver_id,
                        "old_driver_type": driver_type,
                        "new_driver_id": new_driver_id,
                        "new_driver_type": new_driver_type
                    })
            current_time += timedelta(minutes=5)

    return schedule

# Сохраняет расписание, действия, пересменки и итог в Excel.
def save_to_excel(schedule, actions, swaps, day, writer):
    # Названия дней недели
    days_of_week = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
    sheet_name = days_of_week[day % 7]  # Определяем название дня недели

    # Фильтруем данные по текущему дню
    df_schedule = pd.DataFrame(schedule)
    df_actions = pd.DataFrame(actions)
    df_swaps = pd.DataFrame(swaps)

    # Создаем итоговый DataFrame
    summary = {
        "Максимальный ID Водителей типа А": max([driver["id"] for driver in drivers_A], default=0),
        "Максимальный ID Водителей типа Б": max([driver["id"] for driver in drivers_B], default=0),
    }
    summary_df = pd.DataFrame([summary])

    # Сохраняем данные в разные листы с уникальными названиями для каждого дня
    df_schedule.to_excel(writer, sheet_name=f"Расписание {sheet_name}", index=False)
    df_actions.to_excel(writer, sheet_name=f"Действия {sheet_name}", index=False)
    df_swaps.to_excel(writer, sheet_name=f"Пересменка {sheet_name}", index=False)
    summary_df.to_excel(writer, sheet_name=f"Итог {sheet_name}", index=False)


# Основной код
max_drivers_A = 0
max_drivers_B = 0

# Сбрасывает список водителей типа Б каждые три дня
def reset_drivers_B():
    global drivers_B, driver_B_id
    drivers_B = []
    driver_B_id = 1

# Основной код
if __name__ == "__main__":
    writer = pd.ExcelWriter("Метод_в_лоб.xlsx", engine="xlsxwriter")

    for day in range(7):  # Генерация расписания на неделю
        is_saturday = (day % 7 == 5)  # Суббота
        is_sunday = (day % 7 == 6)  # Воскресенье

        # Сбрасываем водителей типа Б каждые три дня
        if day % 3 == 0:
            reset_drivers_B()

        # Сбрасываем список действий и пересменок для нового дня
        actions = []  # Лог действий
        swaps = []    # Пересменки

        # Сбрасываем водителей типа А в начале нового дня
        drivers_A = []
        driver_A_id = 1

        # Генерируем расписание
        schedule = generate_schedule(day, is_saturday=is_saturday, is_sunday=is_sunday)

        # Сохраняем расписание для текущего дня
        save_to_excel(schedule, actions, swaps, day, writer)

    writer.close()
    print("Расписание сохранено в файл 'Метод_в_лоб.xlsx'")

