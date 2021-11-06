import xlwings as xw
import socket
import time
import xml.etree.ElementTree as ET
import requests
import pathlib
import pandas as pd

from hidroweb_downloader.download_from_api_BATCH import Hidroweb_BatchDownload

@xw.func
def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    # Check connection
    a = check_internetConection()
    if a == 'Conectado':
        xw.Range('A1').color = (0,255,0) # Green color
    else:
        xw.Range('A1').color = (255,0,0) # Red color
    sheet['A1'].value = f'{a}'

    # Find codes that fits the conditions
    Estado = xw.Range('B3').value
    AreaDrenagem_min = xw.Range('B4').value
    AreaDrenagem_max = xw.Range('B5').value

    check_input()

    # DEBUG  Temp location
    hidrowebInventario_path = pathlib.Path(r'C:\Users\User\git\Excel-python\test\Hidroweb_Inventario\Inventario.csv')
    inventario = pd.read_csv(hidrowebInventario_path)

    inventario_df = inventario.loc[(inventario['nmEstado'].str.contains(Estado))&
                                   (inventario['AreaDrenagem']>=AreaDrenagem_min)&
                                   (inventario['AreaDrenagem']<=AreaDrenagem_max),
                                   ['Codigo', 'AreaDrenagem']]

    xw.Range('B7').value = inventario_df

    count = 8
    # for i, row in inventario_df.iterrows():
    #     cell_i = f'D{count}' # Drenagem
    #     if row['AreaDrenagem'] > AreaDrenagem_max:
    #         xw.Range(cell_i).color = (255,0,0)
    #         print('aqui')
    #     else:
    #         xw.Range(cell_i).color = (0,255,0)
    #         print('aquin erro')
    #     count += 1

    for i, row in inventario_df.iterrows():
        cell_i = f'D{count}' # Drenagem

        station_download = download_HidrowebStation(estado=Estado,
                                                    min_areaDrenagem=AreaDrenagem_min,
                                                    max_areaDrenagem=AreaDrenagem_max,
                                                    codigo=row['Codigo'])
        print(station_download)

        if station_download:
            xw.Range(cell_i).color = (0,255,0)
        else:
            xw.Range(cell_i).color = (255,0,0)
        count += 1







def check_input():
    if xw.Range('B3').value is None:
        print('erro')
        xw.Range('D3').color = (255,0,0)
    else:
        xw.Range('D3').color = (0,255,0)

@xw.func
def download_HidrowebStation(estado, min_areaDrenagem, max_areaDrenagem, codigo):
    cwd = pathlib.Path(__file__).parent.absolute()/f'Hidroweb_Stations_min{min_areaDrenagem}_max{max_areaDrenagem}'
    cwd.mkdir(parents=True, exist_ok=True)

    d = Hidroweb_BatchDownload()
    a = d.download_ANA_stations(station=int(codigo), typeData=3, folder_toDownload=cwd)

    return 0




def check_internetConection(host="8.8.8.8", port=53, timeout=3):
    """
    Host: 8.8.8.8 (google-public-dns-a.google.com)
    OpenPort: 53/tcp
    Service: domain (DNS/TCP)
    """
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return 'Conectado'
    except socket.error as ex:
        print(ex)
        return 'Desconectado'

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


if __name__ == "__main__":
    xw.Book("teste_main.xlsm").set_mock_caller()
    main()
