import xlwings as xw
import socket
import time
import xml.etree.ElementTree as ET
import requests
import pathlib
import pandas as pd

from hidroweb_downloader.download_from_api_BATCH import Hidroweb_BatchDownload



#### MAIN FUNCTION TESTE ####
@xw.func
def main2():
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
    Latitude = xw.Range('B4').value
    Longitude = xw.Range('B5').value
    Buffer = xw.Range('B6').value

    AreaDrenagem_min = xw.Range('B7').value
    AreaDrenagem_max = xw.Range('B8').value
    Quantil = xw.Range('B9').value

    data_input = {'Estado': xw.Range('B3'),
                  'Latitude': xw.Range('B4'),
                  'Longitude': xw.Range('B5'),
                  'Buffer': xw.Range('B6'),
                  'AreaDrenagem_min': xw.Range('B7'),
                  'AreaDrenagem_max':xw.Range('B8'),
                  'Quantil': xw.Range('B9')}

    # Results location
    results_range = xw.Range('B10')

    # check_input()

    # DEBUG  Temp location
    hidrowebInventario_path = pathlib.Path(r'C:\Users\User\git\Excel-python\test\Hidroweb_Inventario\Inventario.csv')
    inventario = pd.read_csv(hidrowebInventario_path)

    inventario_df = inventario.loc[(inventario['nmEstado'].str.contains(Estado))&
                                   (inventario['AreaDrenagem']>=AreaDrenagem_min)&
                                   (inventario['AreaDrenagem']<=AreaDrenagem_max),
                                   ['Codigo', 'AreaDrenagem']].reset_index(drop=True).copy()

    print(inventario.columns)
    print(inventario[['Latitude', 'Longitude']])

    a = filter_inventario(inventario=inventario,
                          data_input=data_input)

    print(a)
###########################

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
    Latitude = xw.Range('B4').value
    Longitude = xw.Range('B5').value
    Buffer = xw.Range('B6').value

    AreaDrenagem_min = xw.Range('B7').value
    AreaDrenagem_max = xw.Range('B8').value
    Quantil = xw.Range('B9').value

    data_input = {'Estado': xw.Range('B3'),
                  'Latitude': xw.Range('B4'),
                  'Longitude': xw.Range('B5'),
                  'Buffer': xw.Range('B6'),
                  'AreaDrenagem_min': xw.Range('B7'),
                  'AreaDrenagem_max':xw.Range('B8'),
                  'Quantil': xw.Range('B9')}

    # Results location
    results_range = xw.Range('B11')

    # check_input()

    # DEBUG  Temp location
    hidrowebInventario_path = pathlib.Path(r'C:\Users\User\git\Excel-python\test\Hidroweb_Inventario\Inventario.csv')
    inventario = pd.read_csv(hidrowebInventario_path)
    ###

    # inventario_df = inventario.loc[(inventario['nmEstado'].str.contains(Estado))&
    #                                (inventario['AreaDrenagem']>=AreaDrenagem_min)&
    #                                (inventario['AreaDrenagem']<=AreaDrenagem_max),
    #                                ['Codigo', 'AreaDrenagem']].reset_index(drop=True).copy()

    inventario_df = filter_inventario(inventario=inventario, data_input=data_input)

    print('Clearing')
    end_range = xw.Range('D11').end('down')
    xw.Range(results_range,end_range).clear()
    # results_range.expand('down').clear()

    print('Display all results')
    results_range.options(pd.DataFrame, index=False).value = inventario_df
    xw.Range(results_range,end_range).color = None

    print('----- Initiate download -----')
    for i, row in inventario_df.iterrows():
        result_row = results_range.row + i + 1
        result_column = results_range.column + 1

        print((result_row, result_column))
        station_download = download_HidrowebStation(estado=Estado,
                                                    min_areaDrenagem=AreaDrenagem_min,
                                                    max_areaDrenagem=AreaDrenagem_max,
                                                    codigo=row['Codigo'])
        print(station_download)

        if station_download:
            print('Downloaded!')
            xw.Range((result_row, result_column)).color = (0,255,0) # Green
        else:
            print('Fail download')
            xw.Range((result_row, result_column)).color = (255,0,0) # Red

    # Calculate Permanence curve
    print('----- Calculate Permanence Curve -----')
    for i, row in inventario_df.iterrows():
        result_row = results_range.row + i + 1
        result_column = results_range.column + 2
        code = row['Codigo']

        v = vazaoQuantil_Hidroweb(code=code,
                                  quantil=Quantil,
                                  min_areaDrenagem=AreaDrenagem_min,
                                  max_areaDrenagem=AreaDrenagem_max)
        print(v)

        xw.Range((result_row, result_column)).value = v[1] # (Boolean, Vazao)
        if v[0]:
            xw.Range((result_row, result_column)).color = (0, 255, 0) # Green
        else:
            xw.Range((result_row, result_column)).color = (255, 0, 0) # Red

def vazaoQuantil_Hidroweb(code, quantil, min_areaDrenagem, max_areaDrenagem):
    # Locate folder with downloaded data
    cwd = pathlib.Path(__file__).parent.absolute()/f'Hidroweb_Stations_min{min_areaDrenagem}_max{max_areaDrenagem}'

    print(code)
    for file in cwd.rglob('3_*.csv'):
        code_int = f'{int(code):08}' # It needs to be 8 digits as an Integer
        # Check files in directory if matches
        if code_int in file.stem:
            print('Opening matched file!')
            df = pd.read_csv(file, parse_dates=['Date'])
            v = df[f'Data3_{code_int}'].dropna().sort_values(ascending=False).quantile(q=quantil/100)
            return True, v
        else:
            pass
    # If none is found, return 'NO DATA'
    return False, 'Sem dados'


def check_input(input_data):
    xw.Range('B3:B5').color = None
    if (input_data['Latitude'].value is None) or (input_data['Longitude'].value is None) or (input_data['Buffer'].value is None):
        print('Use only Estado!')
        input_data['Estado'].color = (0,255,0)
        input_data['Latitude'].color = (255,0,0)
        input_data['Longitude'].color = (255,0,0)
        input_data['Buffer'].color = (255,0,0)
        return 'Estado'
    else:
        print('Use Lat, Lon and Buffer instead of Estado!!')
        input_data['Estado'].color = (255,0,0)
        input_data['Latitude'].color = (0,255,0)
        input_data['Longitude'].color = (0,255,0)
        input_data['Buffer'].color = (0,255,0)
        return 'LatLon'



def filter_inventario(inventario, data_input):
    select_filter = check_input(input_data=data_input)
    print(select_filter)
    if select_filter == 'LatLon':
        # Using Latitude, Longitude and a Buffer value
        # Without use of Estado
        inventario_filter = inventario.loc[(inventario['Latitude']<=data_input['Latitude'].value+data_input['Buffer'].value)&
                                           (inventario['Latitude']>=data_input['Latitude'].value-data_input['Buffer'].value)&
                                           (inventario['Longitude']<=data_input['Longitude'].value+data_input['Buffer'].value)&
                                           (inventario['Longitude']>=data_input['Longitude'].value-data_input['Buffer'].value)&
                                           (inventario['AreaDrenagem']>=data_input['AreaDrenagem_min'].value)&
                                           (inventario['AreaDrenagem']<=data_input['AreaDrenagem_max'].value),
                                           ['Codigo', 'AreaDrenagem']].reset_index(drop=True).copy()

    elif select_filter == 'Estado':
        inventario_filter = inventario.loc[(inventario['nmEstado'].str.contains(data_input['Estado'].value))&
                                           (inventario['AreaDrenagem']>=data_input['AreaDrenagem_min'].value)&
                                           (inventario['AreaDrenagem']<=data_input['AreaDrenagem_max'].value),
                                           ['Codigo', 'AreaDrenagem']].reset_index(drop=True).copy()

    return inventario_filter

@xw.func
def download_HidrowebStation(estado, min_areaDrenagem, max_areaDrenagem, codigo):
    # Try to create a new directory
    cwd = pathlib.Path(__file__).parent.absolute()/f'Hidroweb_Stations_min{min_areaDrenagem}_max{max_areaDrenagem}'
    cwd.mkdir(parents=True, exist_ok=True)

    # Start function to download
    d = Hidroweb_BatchDownload()
    # Returns True if downloads and False if it doesnt
    a = d.download_ANA_stations(station=int(codigo), typeData=3, folder_toDownload=cwd) # Returns True or False
    return a

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
