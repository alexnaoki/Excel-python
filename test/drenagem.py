import xlwings as xw
import math
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sympy.physics.units as u
from pint import UnitRegistry


u = UnitRegistry()
# wb = xw.Book()
@xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"

@xw.func
def vazao_tubo(diametro, limite_lamina, n, inclinacao):
    '''
    diametro(mm)
    limite_lamina (adimensional, entre 0 e 1)
    n
    inclinacao (m/m)
    '''
    diametro = diametro*u.mm
    n = n*u.m**(2/3)*u.s/u.m
    yn = diametro*limite_lamina


    theta = 2*np.arccos(1-2*yn/diametro)

    area = diametro**2*(theta-np.sin(theta))/8
    area = area.to('m**2')

    perimetro_molhado = theta*diametro/2
    perimetro_molhado = perimetro_molhado.to('m')

    rh = diametro*(1-np.sin(theta)/theta)/4
    rh = rh.to('m')

    v = 1/n*rh**(2/3)*inclinacao**(1/2)
    v = v.to('m/s')

    vazao = area*v
    vazao = vazao.to('m**3/s')

    # return [[theta.magnitude, area.magnitude, perimetro_molhado.magnitude, rh.magnitude, v.magnitude, vazao.magnitude]]
    return [[area.magnitude, rh.magnitude, v.magnitude, vazao.magnitude]]

@xw.func
def metodo_racional(coef_esc, intensidade, area):
    '''
    coeficiente de escoamento (adimensional)
    intensidade (mm/h)
    area (m2)
    '''
    coef_esc = coef_esc
    intensidade = intensidade*u.mm/u.hour
    area = area*u.m**2

    Q = coef_esc*area*intensidade
    Q = Q.to('m**3/s')
    return [[Q.magnitude]]

@xw.func(volatile=True)
def vazao_acumulada(vazaoEntrada, secaoAtual, proximaSecao, caller):
    '''
    vazaoEntrada (m3/s)
    secaoAtual (string)
    '''
    # print(range.get_adress())
    # rng = xw.Range('B1:B30')
        # soma = 0
    # for i in rng.rows:
    #     if i.value == '2A':
    #         soma += 10
    vazaoEntrada = vazaoEntrada*u.m**3/u.s
    vazaoAcumulada = vazaoEntrada

    locations = []
    valores = []
    for i, line in enumerate(proximaSecao):
        if line == secaoAtual:
            locations.append((xw.Range((i+1,5))).get_address())
            valores.append((xw.Range((i+1,5))).value)
            v = xw.Range((i+1, 5)).value*u.m**3/u.s
            # valores.append(v)
            vazaoAcumulada += v

    return [vazaoAcumulada.magnitude]

@xw.func
def get_caller_address(a,caller):
    # caller will not be exposed in Excel, so use it like so:
    # =get_caller_address()
    t = 0
    locations = []
    for i, line in enumerate(a):
        if line == '2A':
            t += 1
            locations.append((xw.Range((i+1, 2))).get_address())
    return [[t]+locations]


@xw.func
def check_dimensonamento(vazaoMax, vazaoEntrada):
    '''
    vazaoMax (m3/s) = vazão máxima do tubo
    vazaoEntrada (m3/s) = vazão de entrada (metodo racional)
    '''
    vazaoMax = vazaoMax*u.m**3/u.s
    vazaoEntrada = vazaoEntrada*u.m**3/u.s

    if vazaoEntrada <= vazaoMax:
        # caller.address.color = (255,255,0)
        return ['OK']

    else:
        return ['Vazao excedida']




if __name__ == "__main__":
    # xw.books.active.set_mock_caller()
    main()
