import xlwings as xw
import math
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
# import sympy.physics.units as u
from pint import UnitRegistry
from functools import reduce

import xml.etree.ElementTree as ET
import requests
import pathlib
# import pandas as pd
from hidroweb_downloader.download_from_api_BATCH import Hidroweb_BatchDownload

u = UnitRegistry()
# wb = xw.Book()
@xw.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xw.Book.caller()
    wb.sheets[0].range("A1").value = "Hello xlwings!"

# @xw.func
# def testando():
#     return 0

@xw.func
def download_HidrowebInventario():
    api_inventario = 'http://telemetriaws1.ana.gov.br/ServiceANA.asmx/HidroInventario'

    params = {'codEstDE':'','codEstATE':'','tpEst':'','nmEst':'','nmRio':'','codSubBacia':'',
              'codBacia':'','nmMunicipio':'','nmEstado':'','sgResp':'','sgOper':'','telemetrica':''}

    response = requests.get(api_inventario, params)


    tree = ET.ElementTree(ET.fromstring(response.content))

    root = tree.getroot()

    data = {'BaciaCodigo':[],'SubBaciaCodigo':[],'RioCodigo':[],'RioNome':[],'EstadoCodigo':[],
            'nmEstado':[],'MunicipioCodigo':[],'nmMunicipio':[],'ResponsavelCodigo':[],
            'ResponsavelSigla':[],'ResponsavelUnidade':[],'ResponsavelJurisdicao':[],
            'OperadoraCodigo':[],'OperadoraSigla':[],'OperadoraUnidade':[],'OperadoraSubUnidade':[],
            'TipoEstacao':[],'Codigo':[],'Nome':[],'CodigoAdicional':[],'Latitude':[],'Longitude':[],
            'Altitude':[],'AreaDrenagem':[],'TipoEstacaoEscala':[],'TipoEstacaoRegistradorNivel':[],
            'TipoEstacaoDescLiquida':[],'TipoEstacaoSedimentos':[],'TipoEstacaoQualAgua':[],
            'TipoEstacaoPluviometro':[],'TipoEstacaoRegistradorChuva':[],'TipoEstacaoTanqueEvapo':[],
            'TipoEstacaoClimatologica':[],'TipoEstacaoPiezometria':[],'TipoEstacaoTelemetrica':[],'PeriodoEscalaInicio':[],'PeriodoEscalaFim':[] ,
            'PeriodoRegistradorNivelInicio' :[],'PeriodoRegistradorNivelFim' :[],'PeriodoDescLiquidaInicio' :[],'PeriodoDescLiquidaFim':[] ,'PeriodoSedimentosInicio' :[],
            'PeriodoSedimentosFim':[] ,'PeriodoQualAguaInicio':[] ,'PeriodoQualAguaFim' :[],'PeriodoPluviometroInicio':[] ,'PeriodoPluviometroFim':[] ,
            'PeriodoRegistradorChuvaInicio' :[],'PeriodoRegistradorChuvaFim' :[],'PeriodoTanqueEvapoInicio':[] ,'PeriodoTanqueEvapoFim':[] ,'PeriodoClimatologicaInicio' :[],'PeriodoClimatologicaFim':[] ,
            'PeriodoPiezometriaInicio':[] ,'PeriodoPiezometriaFim' :[],'PeriodoTelemetricaInicio' :[],'PeriodoTelemetricaFim' :[],
            'TipoRedeBasica' :[],'TipoRedeEnergetica' :[],'TipoRedeNavegacao' :[],'TipoRedeCursoDagua' :[],
            'TipoRedeEstrategica':[] ,'TipoRedeCaptacao':[] ,'TipoRedeSedimentos':[] ,'TipoRedeQualAgua':[] ,
            'TipoRedeClasseVazao':[] ,'UltimaAtualizacao':[] ,'Operando':[] ,'Descricao':[] ,'NumImagens':[] ,'DataIns':[] ,'DataAlt':[]}

    # print(root.tag)
    for i in root.iter('Table'):

        for j in data.keys():
            d = i.find('{}'.format(j)).text
            if j == 'Codigo':
                data['{}'.format(j)].append('{:08}'.format(int(d)))
            else:
                data['{}'.format(j)].append(d)

    print(len(list(root.iter('Table'))))
    # print(data)
    df = pd.DataFrame(data)

    cwd = pathlib.Path(__file__).parent.absolute()/'Hidroweb_Inventario'
    cwd.mkdir(parents=True, exist_ok=True)



    df.to_csv(cwd/'Inventario.csv')


@xw.func
def find_Code(estado, min_areaDrenagem, max_areaDrenagem):
    hidrowebInventario_path = pathlib.Path(r'C:\Users\User\git\Excel-python\test\Hidroweb_Inventario\Inventario.csv')
    inventario = pd.read_csv(hidrowebInventario_path)

    df = inventario.loc[(inventario['nmEstado']==estado)&
                        (inventario['AreaDrenagem']>=min_areaDrenagem)&
                        (inventario['AreaDrenagem']<=max_areaDrenagem),
                        ['Codigo', 'AreaDrenagem']]
    a = []
    for i, row in df.iterrows():
        download_HidrowebStation(estado=estado, min_areaDrenagem=min_areaDrenagem, max_areaDrenagem=max_areaDrenagem,
                                 codigo=row['Codigo'])
    return df

@xw.func
def download_HidrowebStation(estado, min_areaDrenagem, max_areaDrenagem, codigo):
    cwd = pathlib.Path(__file__).parent.absolute()/f'Hidroweb_Stations_min{min_areaDrenagem}_max{max_areaDrenagem}'
    cwd.mkdir(parents=True, exist_ok=True)

    d = Hidroweb_BatchDownload()
    a = d.download_ANA_stations(station=int(codigo), typeData=3, folder_toDownload=cwd)
    return 0

def merge_HidrowebStation(estado, min_areaDrenagem, max_areaDrenagem):
    cwd = pathlib.Path(__file__).parent.absolute()/f'Hidroweb_Stations_min{min_areaDrenagem}_max{max_areaDrenagem}'

    dfs = []
    for data in cwd.rglob('3*.csv'):
        df = pd.read_csv(data, parse_dates=['Date'])
        dfs.append(df)

    df_merged = reduce(lambda left, right: pd.merge(left, right, on=['Date'], how='outer'), dfs)

    df_merged = df_merged.loc[:,~df_merged.columns.str.startswith('Unnamed')]
    df_merged = df_merged.loc[:,~df_merged.columns.str.startswith('Consistence')]
    df_merged = df_merged.loc[:,~df_merged.columns.str.endswith('x')]
    df_merged = df_merged.loc[:,~df_merged.columns.str.endswith('y')]

    return df_merged

@xw.func
def vazaoQuantil_HidrowebStation(estado, min_areaDrenagem, max_areaDrenagem, quantil, list_codes):
    df_merged = merge_HidrowebStation(estado=estado, min_areaDrenagem=min_areaDrenagem, max_areaDrenagem=max_areaDrenagem)

    a = []
    for station in list_codes:
        # a += [station]
        try:
            df_na = df_merged[f'Data3_{int(station)}'].dropna()
            v = df_na.sort_values(ascending=False).quantile(q=quantil/100)
            a.append([v])
        except KeyError:
            a.append(['Sem dados'])

    return a


if __name__ == "__main__":
    # xw.books.active.set_mock_caller()
    main()
