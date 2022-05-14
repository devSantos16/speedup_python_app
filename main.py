import time
import xlsxwriter
import os
import random
from threading import Thread
from threading import Lock
from pathlib import Path

_CONTADOR = 2
VETOR_A = []
VETOR_B = []
LOCK = Lock()


def retornarVetorSequencial(_tamanho):
    global _CONTADOR

    print("Começando Vetor Sequencial ({})".format(_tamanho))

    start_time = time.time()
    retornarVetorThread(_tamanho)
    time_execution = time.time() - start_time

    print("Tempo de Execução: {}".format(time_execution))

    worksheet.write('B{}'.format(_CONTADOR), "{} segundos".format(time_execution))
    print("Excel sendo processado !")

    _CONTADOR += 1


def retornarVetorThread(_tamanho):
    sum_temp = 0
    for i in range(_tamanho):
        VETOR_A.append(_tamanho)
        VETOR_B.append(_tamanho)

    for i in range(_tamanho):
        sum_temp += (VETOR_A[i] * VETOR_B[i])


def gerarthread(thread, tamanho):
    resultado = thread
    print("Começando vetor com thread ({}) do tamanho ({})".format(resultado, tamanho))
    start_time = time.time()
    for pthread in range(thread):
        thread = Thread(target=retornarVetorThread, args=(tamanho,))
        thread.start()
    time_execution = time.time() - start_time
    print("Tempo de Execução: {}".format(time_execution))
    write_excel(time_execution, tamanho, resultado)
    print("Excel sendo processado!")


def write_excel(time_execution, tamanho, thread):
    if thread == 2:
        worksheet.write('C1', "{} THREAD".format(thread))
        if tamanho == 1:
            worksheet.write('C2', "{} segundos".format(time_execution))
        if tamanho == 1000:
            worksheet.write('C3', "{} segundos".format(time_execution))
        if tamanho == 10000:
            worksheet.write('C4', "{} segundos".format(time_execution))
        if tamanho == 10000000:
            worksheet.write('C5', "{} segundos".format(time_execution))

    if thread == 5:
        worksheet.write('D1', "{} THREAD".format(thread))
        if tamanho == 1:
            worksheet.write('D2', "{} segundos".format(time_execution))
        if tamanho == 1000:
            worksheet.write('D3', "{} segundos".format(time_execution))
        if tamanho == 10000:
            worksheet.write('D4', "{} segundos".format(time_execution))
        if tamanho == 10000000:
            worksheet.write('D5', "{} segundos".format(time_execution))

    if thread == 10:
        worksheet.write('E1', "{} THREAD".format(thread))
        if tamanho == 1:
            worksheet.write('E2', "{} segundos".format(time_execution))
        if tamanho == 1000:
            worksheet.write('E3', "{} segundos".format(time_execution))
        if tamanho == 10000:
            worksheet.write('E4', "{} segundos".format(time_execution))
        if tamanho == 10000000:
            worksheet.write('E5', "{} segundos".format(time_execution))


def moverpasta():
    hash = random.getrandbits(128)
    dir_name = 'excel'
    try:
        os.mkdir(dir_name)
        print("Diretório ", dir_name, "criado")

    except FileExistsError:
        print("Diretorio", dir_name, "já existe")

    Path("tentativa.xlsx").rename("{}/{}.xlsx".format(dir_name, hash))


if __name__ == '__main__':
    vetorNumeros = [2, 5, 10]
    workbook = xlsxwriter.Workbook('tentativa.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write("A1", "Carga")
    worksheet.write('B1', "Sequencial")
    worksheet.write('A2', 1)
    worksheet.write('A3', 1000)
    worksheet.write('A4', 10000)
    worksheet.write('A5', 10000000)

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
    workbook.close()
    moverpasta()



