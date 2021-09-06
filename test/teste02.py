import xlwings as xw
import math
import pandas as pd
import matplotlib.pyplot as plt
import sympy.physics.units as u



@xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"


@xw.func
def hello(name):
    return "hello {0}22222".format(name)

@xw.func
def teste(x):
    return f"testeando {x}"

@xw.func
@xw.ret(expand='table')
@xw.arg('diametro', doc='Diametro em mm')
def area_aco(diametro):
    '''
    diametro em mm
    area_tranversal em cm2
    '''
    diametro = diametro*u.mm
    area_tranversal = math.pi*diametro**2/4
    area_tranversal = u.convert_to(area_tranversal,u.cm**2)
    # return [area_tranversal.args[0], str(area_tranversal.args[1])]
    return area_tranversal


@xw.func
def aco_opcoes(area_necessaria):
    bitolas = [5, 6.3, 8, 10, 12.5, 16, 20, 25]

    area_necessaria = area_necessaria*u.cm**2


    return [[a]+[math.ceil(area_necessaria/area_aco(diametro=a))] for a in bitolas]


if __name__ == "__main__":
    xw.books.active.set_mock_caller()
    main()
