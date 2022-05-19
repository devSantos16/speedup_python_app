import time
import os
import random

from model.excel import Excel
from pathlib import Path
from multiprocessing import Pool

_count = 2
_array_a = _array_b = []
_excel = Excel('tentativa.xlsx')


def return_serial_array(size):
    global _count

    print(f"Começando Vetor Sequencial ({size})")

    start_time = time.time()
    return_array(size)
    time_execution = time.time() - start_time

    print(f"Tempo de Execução: {time_execution}")
    _excel.write_worksheet(f'B{_count}', f"{time_execution} segundos")
    print("Excel sendo processado !")
    _count += 1


def return_array(size):
    local_sum = 0

    for increment in range(size):
        _array_a.append(1)
        _array_b.append(1)

    for increment in range(size):
        local_sum += (_array_a[increment] * _array_b[increment])

    print(f"Tamanho e soma dos vetores: {local_sum} ({size})")


def generate_process(num_process, size):
    pool = Pool(num_process)
    result_process = num_process

    print(f"Começando vetor com processo ({result_process}) do tamanho ({size})")

    start = time.time()
    pool.apply_async(return_array, size)
    pool.close()
    pool.join()
    end = time.time()

    time_execution = end - start

    print(f"Tempo de Execução: {time_execution}")
    verify_process(time_execution, size, result_process)
    print("Excel sendo processado!")


def verify_process(time_execution, size, process):
    array_size = [1, 1000, 10000, 10000000]

    if process == 1:
        write_excel(array_size, size, process, time_execution, 'C')

    if process == 2:
        write_excel(array_size, size, process, time_execution, 'D')

    if process == 5:
        write_excel(array_size, size, process, time_execution, 'E')

    if process == 10:
        write_excel(array_size, size, process, time_execution, 'F')


def write_excel(array_size, size, thread, time_execution, letter):
    _excel.write_worksheet(f'{letter}1', "{} THREAD".format(thread))
    for element_array_size in array_size:
        if size == element_array_size:
            result_index = array_size.index(element_array_size)
            _excel.write_worksheet(f'{letter}{result_index + 2}', "{}s".format(time_execution))


def move_path():
    hash_file = random.getrandbits(128)
    dir_name = 'excel'

    try:
        os.mkdir(dir_name)
        print("Diretório ", dir_name, "criado")

    except FileExistsError:
        print("Diretorio", dir_name, "já existe, inserindo o arquivo na pasta {}".format(dir_name))

    Path(_excel.file_name).rename("{}/{}.xlsx".format(dir_name, hash_file))


if __name__ == '__main__':
    array_numbers_thread = [1, 2, 5, 10]

    _excel.init_workbook()

    _excel.write_worksheet("A1", "Carga")
    _excel.write_worksheet('B1', "Sequencial")
    _excel.write_worksheet('A2', 1)
    _excel.write_worksheet('A3', 1000)
    _excel.write_worksheet('A4', 10000)
    _excel.write_worksheet('A5', 10000000)

    return_serial_array(1)
    return_serial_array(1000)
    return_serial_array(10000)
    return_serial_array(10000000)

    for i in array_numbers_thread:
        generate_process(i, 1)
        generate_process(i, 1000)
        generate_process(i, 10000)
        generate_process(i, 10000000)

    print("Fechando excel !")
    _excel.end_workbook()
    move_path()
    print("Fim do programa !")



