import xlwings as xw
import math
# import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
# import sympy.physics.units as u
from pint import UnitRegistry
from catalogo import CheckList

u = UnitRegistry()
# wb = xw.Book()
@xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"


@xw.func
def perda_carga(L, V, D, f):
    '''
    L (m)
    V (m/s)
    D (m)
    f (adimensional)
    '''

    L = L*u.m
    V = V*u.m/u.s
    D = D*u.m
    g = 9.81*u.m/u.s**2

    delta_H = f*L*V**2/(D*2*g)
    return [delta_H.magnitude] + [str(delta_H.units)]
    # a = CheckList.teste(1)
    # return a

@xw.func
def coef_atrito(velocidade, diametro, e):
    '''
    velocidade (m/s)
    diametro (m)
    e (mm) = rugosidade absoluta equivalente
    '''
    viscosidade_cinematica = 1.003*10**(-6)*u.m**2/u.s
    velocidade = velocidade*u.m/u.s
    diametro = diametro*u.m
    e = e*u.mm

    reynolds = velocidade*diametro/viscosidade_cinematica

    f = ((64/reynolds)**8+9.5*(np.log(e/(3.7*diametro)+5.74/(reynolds**0.9))-(2500/reynolds)**6)**(-16))**(0.125)
    return [f.magnitude]+[str(f.units)]

@xw.func
def perda_localizada(velocidade, K):
    '''
    velocidade (m/s)
    K (adimensional)
    '''
    velocidade = velocidade*u.m/u.s
    g = 9.81*u.m/u.s**2

    delta_H = K*velocidade**2/g

    return [delta_H.magnitude] + [str(delta_H.units)]


@xw.func
def preDimensionamento(vazao_necessaria):
    '''
    vazao_necessaria (L/h)
    '''
    catalogo = CheckList()
    vazao_necessaria = vazao_necessaria*u.l/u.hour
    vazao_necessaria = vazao_necessaria.to('m**3/s')

    velocidade_eco = 1.5*u.m/u.s
    K = (4/(np.pi*velocidade_eco))**(1/2)

    D = K*(vazao_necessaria)**(1/2)

    D = D.magnitude
    D = D*u.m

    d_catalogo = catalogo.min_diametro(D.magnitude)

    # velocidade_min = 0.5*u.m/u.s
    # velocidade_max = 3*u.m/u.s
    # g = 9.81*u.m/u.s**2


    return [K.magnitude]+[vazao_necessaria.magnitude]+[D.magnitude]+[d_catalogo.magnitude]
    # return [vazao_necessaria.magnitude] + [str(vazao_necessaria.units)]




if __name__ == "__main__":
    # xw.books.active.set_mock_caller()
    main()
