import threading
import time
import os
import random

from model.excel import Excel
from threading import Thread
from pathlib import Path

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
        _array_a.append(size)
        _array_b.append(size)

    for increment in range(size):
        local_sum += (_array_a[increment] * _array_b[increment])


def generate_thread(num_thread, size):
    result_thread = num_thread
    threads = []

    print(f"Começando vetor com thread ({result_thread}) do tamanho ({size})")

    # start thread
    increment = 0
    start = time.time()

    while increment < num_thread:
        thread = Thread(target=return_array, args=(size,))
        thread.start()
        threads.append(thread)
        print(f"Numeros de thread: {threading.active_count()}")
        increment += 1

    for thread in threads:
        thread.join()

    end = time.time()
    time_execution = end - start

    print(f"Fim da vetorização {result_thread} ! Numeros de thread rodando agora: {threading.active_count()}")
    print(f"Tempo de Execução: {time_execution}")
    write_excel(time_execution, size, result_thread)
    print("Excel sendo processado!")


def write_excel(time_execution, size, thread):
    if thread == 1:
        _excel.write_worksheet('C1', "{} THREAD".format(thread))
        if size == 1:
            _excel.write_worksheet('C2', "{} segundos".format(time_execution))
        if size == 1000:
            _excel.write_worksheet('C3', "{} segundos".format(time_execution))
        if size == 10000:
            _excel.write_worksheet('C4', "{} segundos".format(time_execution))
        if size == 10000000:
            _excel.write_worksheet('C5', "{} segundos".format(time_execution))
    if thread == 2:
        _excel.write_worksheet('D1', "{} THREAD".format(thread))
        if size == 1:
            _excel.write_worksheet('D2', "{} segundos".format(time_execution))
        if size == 1000:
            _excel.write_worksheet('D3', "{} segundos".format(time_execution))
        if size == 10000:
            _excel.write_worksheet('D4', "{} segundos".format(time_execution))
        if size == 10000000:
            _excel.write_worksheet('D5', "{} segundos".format(time_execution))

    if thread == 5:
        _excel.write_worksheet('E1', "{} THREAD".format(thread))
        if size == 1:
            _excel.write_worksheet('E2', "{} segundos".format(time_execution))
        if size == 1000:
            _excel.write_worksheet('E3', "{} segundos".format(time_execution))
        if size == 10000:
            _excel.write_worksheet('E4', "{} segundos".format(time_execution))
        if size == 10000000:
            _excel.write_worksheet('E5', "{} segundos".format(time_execution))

    if thread == 10:
        _excel.write_worksheet('F1', "{} THREAD".format(thread))
        if size == 1:
            _excel.write_worksheet('F2', "{} segundos".format(time_execution))
        if size == 1000:
            _excel.write_worksheet('F3', "{} segundos".format(time_execution))
        if size == 10000:
            _excel.write_worksheet('F4', "{} segundos".format(time_execution))
        if size == 10000000:
            _excel.write_worksheet('F5', "{} segundos".format(time_execution))


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
        generate_thread(i, 1)
        generate_thread(i, 1000)
        generate_thread(i, 10000)
        generate_thread(i, 10000000)

    print("Fechando excel !")
    _excel.end_workbook()
    move_path()
    print("Fim do programa !")



