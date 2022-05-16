import time
import os
import random

from model.excel import Excel
from threading import Thread
from pathlib import Path

_contador = 2
vetor_a = vetor_b = []
_excel = Excel('tentativa.xlsx')


def retornarVetorSequencial(_tamanho):
    global _contador

    print("Começando Vetor Sequencial ({})".format(_tamanho))

    start_time = time.time()
    retornarVetor(_tamanho)
    time_execution = time.time() - start_time

    print("Tempo de Execução: {}".format(time_execution))

    _excel.write_worksheet('B{}'.format(_contador), "{} segundos".format(time_execution))

    print("Excel sendo processado !")

    _contador += 1


def retornarVetor(_tamanho):
    sum_temp = 0
    for i in range(_tamanho):
        vetor_a.append(_tamanho)
        vetor_b.append(_tamanho)

    for i in range(_tamanho):
        sum_temp += (vetor_a[i] * vetor_b[i])


def gerarthread(thread, tamanho):
    resultado = thread

    print("Começando vetor com thread ({}) do tamanho ({})".format(resultado, tamanho))

    start_time = time.time()
    for i_thread in range(thread):
        thread = Thread(target=retornarVetor, args=(tamanho,))
        thread.start()
        thread.join()
    time_execution = time.time() - start_time

    print("Tempo de Execução: {}".format(time_execution))
    # metodo
    write_excel(time_execution, tamanho, resultado)

    print("Excel sendo processado!")


def write_excel(time_execution, tamanho, thread):
    if thread == 2:
        _excel.write_worksheet('C1', "{} THREAD".format(thread))
        if tamanho == 1:
            _excel.write_worksheet('C2', "{} segundos".format(time_execution))
        if tamanho == 1000:
            _excel.write_worksheet('C3', "{} segundos".format(time_execution))
        if tamanho == 10000:
            _excel.write_worksheet('C4', "{} segundos".format(time_execution))
        if tamanho == 10000000:
            _excel.write_worksheet('C5', "{} segundos".format(time_execution))

    if thread == 5:
        _excel.write_worksheet('D1', "{} THREAD".format(thread))
        if tamanho == 1:
            _excel.write_worksheet('D2', "{} segundos".format(time_execution))
        if tamanho == 1000:
            _excel.write_worksheet('D3', "{} segundos".format(time_execution))
        if tamanho == 10000:
            _excel.write_worksheet('D4', "{} segundos".format(time_execution))
        if tamanho == 10000000:
            _excel.write_worksheet('D5', "{} segundos".format(time_execution))

    if thread == 10:
        _excel.write_worksheet('E1', "{} THREAD".format(thread))
        if tamanho == 1:
            _excel.write_worksheet('E2', "{} segundos".format(time_execution))
        if tamanho == 1000:
            _excel.write_worksheet('E3', "{} segundos".format(time_execution))
        if tamanho == 10000:
            _excel.write_worksheet('E4', "{} segundos".format(time_execution))
        if tamanho == 10000000:
            _excel.write_worksheet('E5', "{} segundos".format(time_execution))


def moverpasta():
    hash = random.getrandbits(128)
    dir_name = 'excel'

    try:
        os.mkdir(dir_name)
        print("Diretório ", dir_name, "criado")

    except FileExistsError:
        print("Diretorio", dir_name, "já existe, inserindo o arquivo na pasta {}".format(dir_name))

    Path(_excel.file_name).rename("{}/{}.xlsx".format(dir_name, hash))


if __name__ == '__main__':
    vetorNumeros = [2, 5, 10]

    _excel.init_workbook()

    _excel.write_worksheet("A1", "Carga")
    _excel.write_worksheet('B1', "Sequencial")
    _excel.write_worksheet('A2', 1)
    _excel.write_worksheet('A3', 1000)
    _excel.write_worksheet('A4', 10000)
    _excel.write_worksheet('A5', 10000000)

    retornarVetorSequencial(1)
    retornarVetorSequencial(1000)
    retornarVetorSequencial(10000)
    retornarVetorSequencial(10000000)

    for i in vetorNumeros:
        gerarthread(i, 1)
        gerarthread(i, 1000)
        gerarthread(i, 10000)
        gerarthread(i, 10000000)

    print("Fechando excel !")
    _excel.end_workbook()
    moverpasta()
    print("End Program !")



