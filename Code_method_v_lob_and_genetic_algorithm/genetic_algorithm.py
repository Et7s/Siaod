import random
import pandas as pd
from datetime import datetime, timedelta
from xlsxwriter import Workbook

# Константы
NUM_BUSES = 8 # Число автобусов
WORK_START = datetime.strptime("06:00", "%H:%M") # Начало рабочего дня
WORK_END = datetime.strptime("03:00", "%H:%M") + timedelta(days=1) # Конец рабочего дня
ROUTE_DURATION = timedelta(minutes=60) # Время, которое занимает один маршрут
ROUTE_VARIATION = 10 # Отклонение маршрута
PEAK_HOURS = [(7, 9), (17, 19)] # Час пик
LOAD_PEAK = 0.7 # Нагрузка в час пик
LOAD_OFF_PEAK = 0.3 # Нагрузка не в час пик
LOAD_SATURDAY = 0.5 # Нагрузка в выходные
LUNCH_INTERVAL = timedelta(minutes=15)  # Интервал между обедами
POPULATION_SIZE = 50 # Размер популяции
GENERATIONS = 100 # Колчество поколений
MUTATION_RATE = 0.1 # Вероятность мутации

# Проверка на час пик
def is_peak_hour(hour):
    return any(start <= hour < end for start, end in PEAK_HOURS)

# Фитнес-функция
def fitness(schedule):
    violations = 0
    drivers_A = set()
    drivers_B = set()

    for entry in schedule:
        # Конвертация строк времени в datetime
        start_time = datetime.strptime(entry["start_time"], "%H:%M")
        end_time = datetime.strptime(entry["end_time"], "%H:%M")

        if entry["driver_type"] == "A":
            drivers_A.add(entry["driver_id"])
            # Проверка на превышение рабочего времени (9 часов)
            if end_time - start_time > timedelta(hours=9):
                violations += 10
        elif entry["driver_type"] == "B":
            drivers_B.add(entry["driver_id"])

    # Штрафы за превышение допустимого количества водителей
    if len(drivers_A) > 10:
        violations += (len(drivers_A) - 10) * 10
    if len(drivers_B) > 14:
        violations += (len(drivers_B) - 14) * 10

    return violations


# Генерация случайного расписания
def generate_random_schedule(day, is_saturday=False, is_sunday=False):
    schedule = []
    current_time = WORK_START

    while current_time < WORK_END:
        for bus_id in range(1, NUM_BUSES + 1):
            load = LOAD_SATURDAY if is_saturday or is_sunday else (
                LOAD_PEAK if is_peak_hour(current_time.hour) else LOAD_OFF_PEAK
            )
            active_buses = max(1, int(NUM_BUSES * load))

            if bus_id <= active_buses:
                route_time = ROUTE_DURATION + timedelta(minutes=random.randint(-ROUTE_VARIATION, ROUTE_VARIATION))
                end_time = current_time + route_time
                driver_type = "A" if random.random() < 0.5 and not is_saturday and not is_sunday else "B"
                driver_id = random.randint(1, 20)

                schedule.append({
                    "bus_id": bus_id,
                    "route": f"Маршрут {bus_id % 2 + 1}",
                    "start_time": current_time.strftime("%H:%M"),
                    "end_time": end_time.strftime("%H:%M"),
                    "driver_type": driver_type,
                    "driver_id": driver_id,
                    "load": f"{int(load * 100)}%"
                })
        current_time += timedelta(minutes=10)
    return schedule

# Кроссовер
def crossover(parent1, parent2):
    point = random.randint(0, len(parent1) - 1)
    return parent1[:point] + parent2[point:]

# Мутация
def mutate(schedule):
    if random.random() < MUTATION_RATE:
        index = random.randint(0, len(schedule) - 1)
        schedule[index]["driver_id"] = random.randint(1, 20)
    return schedule

# Генетический алгоритм
def genetic_algorithm(day, is_saturday=False, is_sunday=False):
    population = [generate_random_schedule(day, is_saturday, is_sunday) for _ in range(POPULATION_SIZE)]

    for _ in range(GENERATIONS):
        population = sorted(population, key=fitness)
        next_generation = population[:10]

        while len(next_generation) < POPULATION_SIZE:
            parent1, parent2 = random.choices(population[:20], k=2)
            child = crossover(parent1, parent2)
            child = mutate(child)
            next_generation.append(child)

        population = next_generation

    return sorted(population, key=fitness)[0]

# Сохранение в Excel
def save_to_excel(schedule, actions, swaps, day, writer):
    days_of_week = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
    sheet_name = days_of_week[day % 7]

    df_schedule = pd.DataFrame(schedule)
    df_actions = pd.DataFrame(actions)
    df_swaps = pd.DataFrame(swaps)

    summary = {
        "Максимальный ID Водителей типа А": max([entry["driver_id"] for entry in schedule if entry["driver_type"] == "A"], default=0),
        "Максимальный ID Водителей типа Б": max([entry["driver_id"] for entry in schedule if entry["driver_type"] == "B"], default=0),
    }
    summary_df = pd.DataFrame([summary])

    df_schedule.to_excel(writer, sheet_name=f"Расписание {sheet_name}", index=False)
    summary_df.to_excel(writer, sheet_name=f"Итог {sheet_name}", index=False)

# Основной код
if __name__ == "__main__":
    writer = pd.ExcelWriter("Генетический_алгоритм.xlsx", engine="xlsxwriter")

    for day in range(7):
        is_saturday = (day % 7 == 5)
        is_sunday = (day % 7 == 6)

        best_schedule = genetic_algorithm(day, is_saturday, is_sunday)
        actions, swaps = [], []

        save_to_excel(best_schedule, actions, swaps, day, writer)

    writer.close()
    print("Расписание сохранено в файл 'Генетический_алгоритм.xlsx'")
