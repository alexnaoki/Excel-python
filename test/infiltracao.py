import xlwings as xw
import math
# import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
# import sympy.physics.units as u
from pint import UnitRegistry

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
def area_bacia_retangular(a, b, cota_topo, cota_base, relacao_talude):

    a = a*u.m
    b = b*u.m
    cota_topo = cota_topo*u.m
    cota_base = cota_base*u.m

    area_base = a*b

    altura = cota_topo - cota_base

    comprimento_vertical = ((altura)**2+(altura*relacao_talude)**2)**(1/2)

    area_lateral = a*comprimento_vertical*2 + b*comprimento_vertical*2

    volume_base = area_base*altura
    volume_lateral = (altura*relacao_talude)*altura/2*2*(a+b)

    return [[area_base.magnitude] + [area_lateral.magnitude]]+[[volume_base.magnitude,volume_lateral.magnitude]]

@xw.func
def idf(K, a, b, c, TR, t):

    i = (K*TR**a)/(t+b)**c
    return i

@xw.func
def convert_to_idfTabot(K, a, b, c):

    A = 0.68*K*np.e**(0.06*c**(-0.26)*b**(1.13))
    B = a
    C = 1.32*c**(-2.28)*b**(0.89)

    return [A,B,C]

@xw.func
def metodo_scs_cn(CN, P):

    S = 25400/CN - 254
    Ia = 0.2*S*u.mm

    Q = (P - Ia)**2/(P+0.8*S*u.mm)
    # return Q

    c = np.where(P < Ia, 0, Q)
    return c.to('mm')

    # if P < Ia:
    #     return 0
    #
    # else:
    #     return Q

@xw.func
def alternar_blocos(x):
    impar = [[v] for i,v in enumerate(x[::-1]) if (i)%2==1]
    par = [[v] for i,v in enumerate(x) if (i+1)%2==0]

    lista_nova = impar + par

    return lista_nova

@xw.func
def split_and_join(x):
    if len(x)%2==0:
        a1 = x[-2::-2]
        a2 = x[1::2]

    else:
        a1 = x[-2::-2]
        a2 = x[::2]

    a = np.concatenate((a1, a2))

    return a

@xw.func
@xw.ret(expand='table')
def metodo_blocos_alternados(K, a, b, c, TR, CN, dt, t, area):

    area = area*u.m*u.m

    duracao = np.arange(dt, t+dt, dt)*u.minutes
    idf_min = (idf(K, a, b, c, TR,duracao.magnitude)*u.mm/u.hours).to('mm/minutes')

    cumsum_idf = (idf_min*duracao)

    diff = np.diff(cumsum_idf, prepend=[0])

    alternar = split_and_join(diff).to('mm')

    peff = metodo_scs_cn(CN=CN, P=alternar)

    volume_por_intervalo = (peff*area).to('m**3')


    output = [[i[0].magnitude, i[1].magnitude, i[2].magnitude, i[3].magnitude, i[4].magnitude, i[5].magnitude, i[6].magnitude] for i in zip(duracao, idf_min, cumsum_idf, diff, alternar, peff, volume_por_intervalo)]


    return output

if __name__ == "__main__":
    # xw.books.active.set_mock_caller()
    main()
