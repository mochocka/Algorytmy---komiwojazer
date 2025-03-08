import pandas as pd
import numpy as np
import random
import math
import time
import xlsxwriter

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

def load_distance_data(file_name):
    data = pd.read_excel(file_name, header=None) 
    values = data.values 
    distance_data = {}
    for y in range(len(values)): 
        for x in range(len(values[y])):
            distance_data[f"{x + 1}:{y + 1}"] = values[y, x]
    return distance_data


file_name = "Dane48.xlsx"  
neighborhood_methods = ["swap", "insert", "reverse"]  
iteration_counts = [100, 500, 1000, 5000]  
no_improvement_limits = [50, 100, 500, 1000] 
initial_paths_count_list = [50] 

parameter_combinations = []
for method in neighborhood_methods:
    for iterations in iteration_counts:
        for no_improvement_limit in no_improvement_limits:
            if no_improvement_limit < iterations: 
                for initial_paths_count in initial_paths_count_list:
                    parameter_combinations.append({
                        'method': method,
                        'iterations': iterations,
                        'no_improvement_limit': no_improvement_limit,
                        'initial_paths_count': initial_paths_count
                    })

distance_data = load_distance_data(file_name)


for execution_num in range(1, 3):  
    results = []  
    start_time = time.time()

    for run_num in range(1, 6):  
        for params in parameter_combinations: 
            combination_start_time = time.perf_counter()

            best_overall_path = None  
            best_overall_length = float('inf') 
            for _ in range(params['initial_paths_count']):  
                current_path = list(range(1, int(math.sqrt(len(distance_data))) + 1))
                random.shuffle(current_path)  
                current_length = calculate_path_length(current_path, distance_data)

                no_improvement_counter = 0  

                while no_improvement_counter < params['no_improvement_limit']:  
                    best_neighbor_path = None  #
                    best_neighbor_length = float('inf')  

                    for _ in range(len(current_path)):  
                        path_copy = current_path.copy()

                        arg1 = random.choice(range(1, len(current_path)))
                        arg2 = random.choice(range(1, len(current_path)))

                        if params['method'] == "swap":  
                            path_copy = swap(path_copy, arg1, arg2)
                        elif params['method'] == "insert":
                            path_copy = insert(path_copy, arg1, arg2)
                        elif params['method'] == "reverse":
                            path_copy = reverse(path_copy, arg1, arg2)

                        neighbor_length = calculate_path_length(path_copy, distance_data)

                        if neighbor_length < best_neighbor_length:  
                            best_neighbor_path = path_copy
                            best_neighbor_length = neighbor_length

                    if best_neighbor_length < current_length: 
                        current_path = best_neighbor_path
                        current_length = best_neighbor_length
                        no_improvement_counter = 0  
                    else:
                        no_improvement_counter += 1  

                if current_length < best_overall_length:  
                    best_overall_path = current_path
                    best_overall_length = current_length

            combination_end_time = time.perf_counter()
            elapsed_time_ms = (combination_end_time - combination_start_time) * 1000  

           
            results.append({
                'method': params['method'],
                'iterations': params['iterations'],
                'no_improvement_limit': params['no_improvement_limit'],
                'initial_paths_count': params['initial_paths_count'],
                'length': best_overall_length,
                'path': best_overall_path,
                'time_ms': elapsed_time_ms
            })

    end_time = time.time() - start_time

    workbook = xlsxwriter.Workbook(f"HillClimb_Results_48_{execution_num}.xlsx")
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "Neighborhood Type")
    worksheet.write(0, 1, "Iterations")
    worksheet.write(0, 2, "No Improvement Limit")
    worksheet.write(0, 3, "Path Length")
    worksheet.write(0, 4, "Path")
    worksheet.write(0, 5, "Time (ms)")

    row = 1
    for res in results: 
        worksheet.write(row, 0, res['method'])
        worksheet.write(row, 1, res['iterations'])
        worksheet.write(row, 2, res['no_improvement_limit'])
        worksheet.write(row, 3, res['length'])
        worksheet.write(row, 4, ','.join(map(str, res['path'])))
        worksheet.write(row, 5, res['time_ms'])
        row += 1

    workbook.close() 
