import pandas as pd
import numpy as np
import random
import time
import xlsxwriter
import math

def calculate_path_length(path, distance_data):
    length = distance_data[f"{path[0]}:{path[-1]}"]
    for i in range(len(path) - 1):
        length += distance_data[f"{path[i]}:{path[i + 1]}"]
    return length

def swap(path, point1, point2):
    p1_idx = path.index(point1)
    p2_idx = path.index(point2)
    path[p1_idx], path[p2_idx] = path[p2_idx], path[p1_idx]
    return path

def insert(path, point, index):
    path.remove(point)  
    path.insert(index, point)  
    return path

def reverse(path, point1, point2):
    p1_idx = path.index(point1)
    p2_idx = path.index(point2)
    if p1_idx > p2_idx:
        p1_idx, p2_idx = p2_idx, p1_idx 
    path[p1_idx:p2_idx + 1] = reversed(path[p1_idx:p2_idx + 1])
    return path


file_name = "Dane48.xlsx"  
iterations = [1000, 5000, 10000, 15000]  
neighborhoods = ["swap", "insert", "reverse"]  
temperatures = [1000, 1500, 2000, 3000]  
alphas = [0.1, 0.05, 0.025, 0.01]  

parameter_combinations = [
    {'iterations': iter_count, 'neighborhood': nb, 'temperature': temp, 'alpha': alpha}
    for iter_count in iterations
    for nb in neighborhoods
    for temp in temperatures
    for alpha in alphas
]

def load_distances(file_name):
    excel_data = pd.read_excel(file_name, header=None) 
    data_array = excel_data.values  
    distances = {
        f"{x + 1}:{y + 1}": data_array[y, x]  
        for y in range(len(data_array))
        for x in range(len(data_array[y]))
    }
    return distances

def select_best_move(moves):
    return sorted(moves, key=lambda move: move['length'])[0]  


def accept_worse_move(temp, current_length, new_length):
    diff = current_length - new_length  
    probability = math.exp(diff / temp)  
    return random.random() < probability


for run_id in range(1, 6):  
    distances = load_distances(file_name)  
    all_results = [] 
    start_time = time.time()  
    for _ in range(5):  
        for params in parameter_combinations:  
            combination_start_time = time.perf_counter()

            current_route = list(range(1, int(math.sqrt(len(distances))) + 1))
            random.shuffle(current_route)  
            current_length = calculate_path_length(current_route, distances)
            temp = params['temperature'] 

            for _ in range(params['iterations']):  
                if temp < 0.005 * params['temperature']:  
                    break

                candidate_moves = [] 
                for _ in range(5):  
                    temp_route = current_route.copy()

                    idx1, idx2 = random.choice(range(1, len(current_route))), random.choice(range(1, len(current_route)))

                    if params['neighborhood'] == "swap":
                        temp_route = swap(temp_route, idx1, idx2)
                    elif params['neighborhood'] == "insert":
                        temp_route = insert(temp_route, idx1, idx2)
                    elif params['neighborhood'] == "reverse":
                        temp_route = reverse(temp_route, idx1, idx2)

                    move = {'arg1': idx1, 'arg2': idx2, 'length': calculate_path_length(temp_route, distances)}
                    candidate_moves.append(move)

                best_move = select_best_move(candidate_moves)  

                
                if best_move['length'] < current_length or accept_worse_move(temp, current_length, best_move['length']):
                    if params['neighborhood'] == "swap":
                        current_route = swap(current_route, best_move['arg1'], best_move['arg2'])
                    elif params['neighborhood'] == "insert":
                        current_route = insert(current_route, best_move['arg1'], best_move['arg2'])
                    elif params['neighborhood'] == "reverse":
                        current_route = reverse(current_route, best_move['arg1'], best_move['arg2'])
                    current_length = calculate_path_length(current_route, distances)

                temp *= 1 - params['alpha']  

            combination_end_time = time.perf_counter()
            elapsed_ms = (combination_end_time - combination_start_time) * 1000
            all_results.append({'params': params, 'length': current_length, 'path': current_route, 'time_ms': elapsed_ms})

    all_results = sorted(all_results, key=lambda res: res['length']) 

    max_path_length = len(','.join(str(node) for node in all_results[0]['path']))
    workbook = xlsxwriter.Workbook(f"results_SA_{run_id}_{file_name}")

    worksheet = workbook.add_worksheet()
    worksheet.set_column(7, 7, int(0.85 * max_path_length))

    row = 1
    for result in all_results:
        worksheet.write(row, 0, result['params']['iterations'])
        worksheet.write(row, 1, result['params']['neighborhood'])
        worksheet.write(row, 2, result['params']['temperature'])
        worksheet.write(row, 3, result['params']['alpha'])
        worksheet.write(row, 4, result['length'])
        worksheet.write(row, 5, ','.join(str(node) for node in result['path']))
        worksheet.write(row, 6, result['time_ms'])
        row += 1

    workbook.close()
