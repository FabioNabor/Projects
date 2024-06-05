import os
import pandas as pd
from datetime import datetime

class TreatmentDataFrame():
    def __init__(self, diretory, diretory_qrtech):
        try:
            dataatual = datetime.now().strftime('%d-%m-%Y')
            os.makedirs(f'{os.getcwd()}\\Fechamento {dataatual}')
        except:
            pass
        self.diretorysave = f'{os.getcwd()}\\Fechamento {dataatual}'
        self._diretory_qrtech = pd.read_excel(diretory_qrtech, engine='openpyxl')
        self._diretory = diretory
        self.listafatura = pd.read_excel(diretory, skiprows=4, engine='openpyxl')
        self.continuacao = None
        self._context = {
            'Name_Not_Identify': [],
            'Value': [],
            'N_Fatura': [],
        }

    def _splitrow(self, string):
        text = str(string['Complemento/Nr.Docto'])
        if " - " in text:
            vector = text.split(" - ")[1].strip()
            value = self._transformenumber(string['Valor. Movto'])
            faturas = self.listafatura['Complemento/Nr.Docto'].tolist()

            resultado = 'Não Encontrado'

            for indice, linhas in self._diretory_qrtech.iterrows():
                if vector.strip() in linhas['Nome'] and str(linhas['Fatura']).strip() not in faturas:
                    resultado = linhas['Fatura']

            self._context['N_Fatura'].append(resultado)
            self._context['Name_Not_Identify'].append(vector)
            self._context['Value'].append(value)

            return resultado
        return text.strip()

    def _transformenumber(self, number):
        number = str(number).replace("R$ ", "").replace(".","").replace(",",".")
        return float(number)

    def _correcao(self, number):
        return str(number).strip()

    def _detecterror(self, row):
        duplicado = 0
        name = str(row['Name_Not_Identify']).strip()
        fatura = str(row['Name_Not_Identify']).strip()
        for linhas in self._diretory_qrtech['Nome']:
            if name in linhas:
                duplicado += 1
        if duplicado > 2:
            return 'S'
        return 'N'

    def _formatDataFrame(self):
        file = self.listafatura

        file.dropna(how='all', inplace=True)

        file.dropna(axis=1, how='all', inplace=True)
        try: file.drop(columns=["Unnamed: 6"], inplace=True)
        except: pass
        file.drop(columns=["Usuário"], inplace=True)

        file.drop(file.index[-2:],inplace=True)

        file['Result'] = file.apply(self._splitrow, axis=1)
        file['Valor. Movto'] = file['Valor. Movto'].apply(self._transformenumber)

        self.dataframe = pd.DataFrame(self._context)
        self.dataframe['Erro'] = self.dataframe.apply(self._detecterror, axis=1)
        self.dataframe.to_excel(f'{self.diretorysave}\\P_Erros.xlsx', index=False)
        return file

    def _re_avaliation(self):
        dataframe = self._formatDataFrame()
        notencontrados = dataframe[dataframe['Result'] == 'Não Encontrado']
        lista_faturas = self.listafatura['Result'].tolist()
        print(lista_faturas)
        for idx, row in notencontrados.iterrows():
            value = row['Valor. Movto']
            result = self._diretory_qrtech[self._diretory_qrtech['Vlr Recebido'] == value]
            if len(result) >= 1:
                for _, row in result.iterrows():
                    if row['Fatura'] not in lista_faturas and row['A Maior'] != value:
                        dataframe.loc[idx, 'Result'] = row['Fatura']
                        break
        dataframe.to_excel(f'{self.diretorysave}\\ExtratoQR.xlsx', index=False)
        return dataframe

    def pagments_more_than_one(self):
        result = {
            'Cliente': [],
            'NSU': [],
            'Nome': [],
            'Fatura': [],
            'Qt. Pagamento': [],
            'Valor Fatura': [],
            'Total Pago': [],
            'Vlr Liquidado': [],
            'A Maior': [],
        }

        dataframe_tech = self._diretory_qrtech
        dataframe_extrato = self._re_avaliation()

        dataframe_tech['Fatura'] = dataframe_tech['Fatura'].apply(self._correcao)

        for idx, row in dataframe_extrato.iterrows():
            if row['Complemento/Nr.Docto'].strip() == [] or row['Complemento/Nr.Docto'].strip() == '':
                result['NSU'].append(row['NSU'])
                result['Cliente'].append(0)
                result['Nome'].append('Não Identificado')
                result['Fatura'].append(0)
                result['Qt. Pagamento'].append(1)
                result['Valor Fatura'].append(0)
                result['Total Pago'].append(row['Valor. Movto'])
                result['Vlr Liquidado'].append(0)
                result['A Maior'].append(0)
                continue


            tech = dataframe_tech[dataframe_tech['Fatura'] == str(row['Result']).strip()]

            if tech.empty:
                print(row)
                result['NSU'].append(row['NSU'])
                result['Cliente'].append(0)
                result['Nome'].append(row['Natureza. Movto'])
                result['Fatura'].append(row['Result'])
                result['Qt. Pagamento'].append(1)
                result['Valor Fatura'].append(0)
                result['Total Pago'].append(row['Valor. Movto'])
                result['Vlr Liquidado'].append(0)
                result['A Maior'].append(0)
                continue

            tech = tech.iloc[0]  # Pega a primeira linha do resultado

            if row['Result'] not in result['Fatura']:
                result['NSU'].append(row['NSU'])
                result['Cliente'].append(tech['Cliente'])
                result['Nome'].append(tech['Nome'])
                result['Fatura'].append(tech['Fatura'])
                result['Qt. Pagamento'].append(1)
                result['Valor Fatura'].append(tech['Valor da Fatura'])
                result['Total Pago'].append(row['Valor. Movto'])
                result['Vlr Liquidado'].append(tech['Vlr Liquidado'])
                result['A Maior'].append(tech['A Maior'])
            else:
                index = result['Fatura'].index(tech['Fatura'])
                result['Qt. Pagamento'][index] += 1
                result['Total Pago'][index] += row['Valor. Movto']



        dt = pd.DataFrame(result)

        totais = dt[['Qt. Pagamento', 'Valor Fatura', 'Total Pago', 'Vlr Liquidado', 'A Maior']].sum()

        totais['Fatura'] = 'TOTAL'

        dt_totais = dt.append(totais, ignore_index=True)
        dt_totais.to_excel(f'{self.diretorysave}//ExtratoQR_Detalhado.xlsx', index=False)












