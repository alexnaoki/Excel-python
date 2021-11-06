import xml.etree.ElementTree as ET
import requests
import pandas as pd
import pathlib
import datetime
import calendar
# Input Data
# station_code_list = [62292200]

# Creation of Folder for the data
# path_folder = pathlib.Path(r"C:\Users\User\Desktop\hidroweb\rioverde")
# try:
#     os.mkdir(pathlib.Path(path_folder))
# except:
#     pass


def download_ANA_stations(list_codes, typeData, folder_toDownload):
    numberOfcodes = len(list_codes)
    count = 0
    path_folder = pathlib.Path(folder_toDownload)
    # floatProgress_loadingDownload.bar_style = 'info'
    dfs_download = []

    for station in list_codes:
        params = {'codEstacao': station, 'dataInicio': '', 'dataFim': '', 'tipoDados': '{}'.format(typeData), 'nivelConsistencia': ''}
        response = requests.get('http://telemetriaws1.ana.gov.br/ServiceANA.asmx/HidroSerieHistorica', params)

        tree = ET.ElementTree(ET.fromstring(response.content))

        root = tree.getroot()

        list_data = []
        list_consistenciaF = []
        list_month_dates = []
        for i in root.iter('SerieHistorica'):
            codigo = i.find("EstacaoCodigo").text
            consistencia = i.find("NivelConsistencia").text
            date = i.find("DataHora").text
            date = datetime.datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
            last_day = calendar.monthrange(date.year, date.month)[1]
            month_dates = [date + datetime.timedelta(days=i) for i in range(last_day)]
            data = []
            list_consistencia = []
            for day in range(last_day):
                if params['tipoDados'] == '3':
                    value = 'Vazao{:02}'.format(day+1)
                    try:
                        data.append(float(i.find(value).text))
                        list_consistencia.append(int(consistencia))
                    except TypeError:
                        data.append(i.find(value).text)
                        list_consistencia.append(int(consistencia))
                    except AttributeError:
                        data.append(None)
                        list_consistencia.append(int(consistencia))
                if params['tipoDados'] == '2':
                    value = 'Chuva{:02}'.format(day+1)
                    try:
                        data.append(float(i.find(value).text))
                        list_consistencia.append(consistencia)
                    except TypeError:
                        data.append(i.find(value).text)
                        list_consistencia.append(consistencia)
                    except AttributeError:
                        data.append(None)
                        list_consistencia.append(consistencia)
            list_data = list_data + data
            list_consistenciaF = list_consistenciaF + list_consistencia
            list_month_dates = list_month_dates + month_dates

        if len(list_data) > 0:
            df = pd.DataFrame({'Date': list_month_dates, 'Consistence_{}_{}'.format(typeData,station): list_consistenciaF, 'Data{}_{}'.format(typeData,station): list_data})

            # if checkbox_downloadIndividual.value == True:
            filename = '{}_{}.csv'.format(typeData, station)
            df.to_csv(path_folder / filename)
            # else:
            #     pass

            count += 1
            # floatProgress_loadingDownload.value = float(count+1)/numberOfcodes
            dfs_download.append(df)
        else:
            count += 1
            # floatProgress_loadingDownload.value = float(count+1)/numberOfcodes

    # try:
    #     dfs_merge_teste0 = reduce(lambda left,right: pd.merge(left, right, on=['Date'], how='outer'), dfs_download)
    #
    #     if checkbox_downloadGrouped.value == True:
    #         dfs_merge_teste0.to_csv(path_folder/'GroupedData_{}.csv'.format(typeData))
    #     else:
    #         pass
    #     selectionMultiple_column.options = list(filter(lambda i: 'Data' in i, dfs_merge_teste0.columns.to_list()))
    # except:
    #     pass
    #
    # floatProgress_loadingDownload.bar_style = 'success'
            # pass

class Hidroweb_BatchDownload:
    def __init__(self):
        pass

    def teste_download(self):
        return 0

    def download_ANA_stations(self, station, typeData, folder_toDownload):
        # numberOfcodes = len(list_codes)
        count = 0
        # path_folder = pathlib.Path(folder_toDownload)
        path_folder = folder_toDownload
        # floatProgress_loadingDownload.bar_style = 'info'
        dfs_download = []

        # for station in list_codes:
        params = {'codEstacao': station, 'dataInicio': '', 'dataFim': '', 'tipoDados': '{}'.format(typeData), 'nivelConsistencia': ''}
        response = requests.get('http://telemetriaws1.ana.gov.br/ServiceANA.asmx/HidroSerieHistorica', params)

        tree = ET.ElementTree(ET.fromstring(response.content))

        root = tree.getroot()

        list_data = []
        list_consistenciaF = []
        list_month_dates = []
        for i in root.iter('SerieHistorica'):
            codigo = i.find("EstacaoCodigo").text
            consistencia = i.find("NivelConsistencia").text
            date = i.find("DataHora").text
            date = datetime.datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
            last_day = calendar.monthrange(date.year, date.month)[1]
            month_dates = [date + datetime.timedelta(days=i) for i in range(last_day)]
            data = []
            list_consistencia = []
            for day in range(last_day):
                if params['tipoDados'] == '3':
                    value = 'Vazao{:02}'.format(day+1)
                    try:
                        data.append(float(i.find(value).text))
                        list_consistencia.append(int(consistencia))
                    except TypeError:
                        data.append(i.find(value).text)
                        list_consistencia.append(int(consistencia))
                    except AttributeError:
                        data.append(None)
                        list_consistencia.append(int(consistencia))
                if params['tipoDados'] == '2':
                    value = 'Chuva{:02}'.format(day+1)
                    try:
                        data.append(float(i.find(value).text))
                        list_consistencia.append(consistencia)
                    except TypeError:
                        data.append(i.find(value).text)
                        list_consistencia.append(consistencia)
                    except AttributeError:
                        data.append(None)
                        list_consistencia.append(consistencia)
            list_data = list_data + data
            list_consistenciaF = list_consistenciaF + list_consistencia
            list_month_dates = list_month_dates + month_dates

        if len(list_data) > 0:
            df = pd.DataFrame({'Date': list_month_dates,
                               'Consistence_{}_{}'.format(typeData,station): list_consistenciaF,
                               'Data{}_{}'.format(typeData,station): list_data})

            # if checkbox_downloadIndividual.value == True:
            filename = '{}_{}.csv'.format(typeData, station)
            df.to_csv(path_folder / filename)
            # else:
            #     pass

            count += 1
            # floatProgress_loadingDownload.value = float(count+1)/numberOfcodes
            # dfs_download.append(df)
            return True
        else:
            count += 1
            return False

                # floatProgress_loadingDownload.value = float(count+1)/numberOfcodes



if __name__ == '__main__':
    list_codes = [20001300, 20008000, 20009000, 21030000, 21100000, 21550000, 24765000, 25070000, 25090000,
 25100000, 25120000, 42430000, 42450100, 60019000, 60321800, 60431100, 60432000, 60432490, 60432500, 60436400,
 60478480, 60495000, 60544000, 60544100, 60620550, 60652000, 60653000, 60664800, 60720000, 60795000, 60811000, 60941000]


    download_ANA_stations(list_codes=list_codes,
                          typeData=3,
                          folder_toDownload=r'C:\Users\User\Desktop\hidroweb\rioverde')
