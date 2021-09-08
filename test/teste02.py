import xlwings as xw
import math
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sympy.physics.units as u
from pint import UnitRegistry
# @xw.func
# @xw.arg('x', np.array)
# def myfunct(x):
#     print(x.T)
#     return x @ x.T
u = UnitRegistry()
@xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"


@xw.func
def hello(name):
    return "hello {0}22222".format(name)

@xw.func
def teste(x):
    return f"testeandoafsf {x}"

@xw.func
# @xw.ret(expand='table')
# @xw.arg('diametro', doc='Diametro em mm')
def area_aco(diametro):
    '''
    diametro em mm
    area_tranversal em cm2
    '''
    diametro = diametro*u.mm
    area_tranversal = math.pi*diametro**2/4
    area_tranversal = area_tranversal.to('cm**2')
    # return [area_tranversal.args[0], str(area_tranversal.args[1])]
    return area_tranversal


@xw.func
def aco_opcoes(area_necessaria):
    bitolas = [5, 6.3, 8, 10, 12.5, 16, 20, 25]
    area_necessaria = area_necessaria*u.cm**2
    return [['Bitola']+['Quantidade']]+[[a]+[math.ceil(area_necessaria/area_aco(diametro=a))] for a in bitolas]

@xw.func
def opcoes_por_bitola(area_necessaria, bw, cobrimento, estribo):
    bitolas = [5, 6.3, 8, 10, 12.5, 16, 20, 25]
    area_necessaria = area_necessaria*u.cm**2
    bw = bw*u.cm
    cobrimento = cobrimento*u.cm
    estribo = estribo*u.mm
    esp = 2*u.cm
    bw_disponivel = bw-2*cobrimento-2*estribo
    result = []
    for bitola in bitolas:
        as_real = area_aco(diametro=bitola)
        quantidade_aco = max(math.ceil(area_necessaria/as_real), 2)
        espaco_total = (bw_disponivel-bitola*u.mm*quantidade_aco)
        espaco_entre_barras = espaco_total/(quantidade_aco-1)
        if espaco_entre_barras < esp:
            pass
        else:
            result += [[bitola]+[quantidade_aco]+[as_real.magnitude*quantidade_aco]+[espaco_entre_barras.magnitude]]
    return result

@xw.func
def combinacoes_bitola(area_necessaria, bw, cobrimento, estribo):
    bitolas = [6.3, 8, 10, 12.5, 16, 20, 25]
    area_necessaria = area_necessaria*u.cm**2
    bw = bw*u.cm
    cobrimento = cobrimento*u.cm
    estribo = estribo*u.mm
    esp = 2*u.cm
    bw_disponivel = bw-2*cobrimento-2*estribo
    result = [['Bitola', 'Quantidade', 'Area cm2', 'diff']]

    for bitola in reversed(bitolas):
        as_real = area_aco(diametro=bitola)
        # quantidade_aco = max(math.ceil(area_necessaria/as_real), 2)
        quantidade_aco = 2

        diff_aco = area_necessaria - as_real*2
        if diff_aco.magnitude < 0:
            result += [[bitola] + [quantidade_aco] + [as_real.magnitude*quantidade_aco]+[diff_aco.magnitude]]
        else:
            result += [[bitola]+[quantidade_aco]+[as_real.magnitude*quantidade_aco]+[diff_aco.magnitude]]
            for bitola2 in bitolas:
                if bitola2 > bitola:
                    pass
                else:
                    as_real2 = area_aco(diametro=bitola2)
                    quantidade_aco2 = math.ceil(diff_aco/as_real2)
                    result += [[f'{bitola2}_sup']+[quantidade_aco2]+[as_real2.magnitude*quantidade_aco2]+[diff_aco.magnitude]]



    return result




if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()
    # xw.serve()
