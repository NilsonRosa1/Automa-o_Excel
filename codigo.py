import pandas as pd
import numpy as np
import win32com.client as win32
from datetime import datetime
from datetime import datetime, timedelta
import os

class Automacao:
    ''' o Desenvovedor terá que mudar os nomes dos campos utilizados se o nome dos campos forem alterados



    '''
    def __init__(self, df_one,df_two):
        self.df_one= df_one
        self.df_two= df_two
    def fmaior_estoque(self,df_two):  # Função para analisar  o produto com maior estoque
        df_maior = df_two[df_two['Quantidade Total'] == df_two['Quantidade Total'].max()]
        return df_maior[['Produto','Quantidade Total']]

    def fmenor_estoque(self,df_two):  # Função para analisar  o produto com maior estoque
        df_menor = df_two[df_two['Quantidade Total'] == df_two['Quantidade Total'].min()]
        return df_menor[['Produto','Quantidade Total']]

    def listaped(self, df_two):  # Transforma o DataFrame em uma Lista com os Pedidos
        compra = df_two[df_two['Quantidade Total'] <= 0]
        compra = compra['Produto']
        compra = list(compra)
        compra = '{}'.format(compra)
        compra = compra.replace('[', '').replace(']', '').replace("'", '')
        return compra

    def lstaquant(self, df_two):  # Transforma o DataFrame em uma Lista com as Quantidades
        quant = df_two[df_two['Quantidade Total'] <= 0]
        quant = quant['Quantidade Total']
        quant = list((quant))
        quant = '{}'.format(quant)
        quant = quant.replace('[', '').replace(']', '').replace("'", '').replace('-', '')
        return quant

    def start(self, df_two):
        '''  Linha 62: Filtra a quantidade Menor e igual a 0
          : Cria  uma Serie (p), que tranfoma negativos em positivos
          : Deleta o campo x['Quantidade Total']
          : Faz um Inner Join entre os Index, emparelhando x com p
         Envia um Email para os fornecedores, indicando os itens que estão faltando, e tambem a data de envia'''
        def pedido_analise(df_two):
            x = df_two[df_two['Quantidade Total'] <= 0]
            p = x['Quantidade Total'] * -1
            del x['Quantidade Total']
            df_inner = x.merge(p, how='inner', left_index=True, right_index=True)
            return df_inner

        def pedido_email(df_inner):
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'nilsonn.fatec@gmail.com'
            mail.Subject = 'Compras dos Itens '
            mail.HTMLBody = f'''
                      <p>Prezado Fornecedor,</p>
                      <p> O estoque precisa ser reabastecido com esses determinados produtos para o dia {tomorrow()} </p>
                      <p> {df_inner[['Produto', 'Quantidade Total']].to_html()}</p>

                      <p>Att.,</p>
                      <p>Nilson Henrique R. Rosa</p>

                      '''
            mail.Send()
            return print('Email Enviado ')

            df_inner = pedido_analise(df_two)
            pedido_email(df_inner)

    def filtro_porcentagem(self, df_one):
        '''def porcentagem: Analisa  a   diferença emtre a saida de ontem com a de hoje, e tranforma em porcentagem
               def filtro_porcentagem: conta quantas vezes teve a saida no dia de hoje (df_dia) e ontem(df_ontem) '''
        df_dia = df_one[df_one['Data'] == datetime.today().strftime('%Y-%m-%d')]
        df_dia = df_dia['Data'].count()
        df_ontem = df_one[df_one['Data'] == yesterday()]
        df_ontem = df_ontem['Data'].count()
        q = ( df_dia-df_ontem)
        print(df_ontem)
        print(df_dia)
        if q >= 0:
            return print('Teve  {} saidas a menos que ontem'.format(q))
        else:
            return print('Teve  {} saidas a mais que ontem'.format(q*-1))


def presentday(self,presentday):
    return presentday
def yesterday(self,presentday):
    yesterday = presentday - timedelta(1)
    return yesterday
def tomorrow(self,presentday):
    tomorrow = presentday + timedelta(1)
    return tomorrow


tabela= pd.ExcelFile('C:/Users/nilso/Downloads/Estoque.xlsx',)

df_one = pd.read_excel(tabela,'Estoque')
df_two = pd.read_excel(tabela,'Movimento')


automacao= Automacao(df_one,df_one)
automacao_email= automacao.start(df_two)
#automacao_porcentagem= automacao.filtro_porcentagem(df_one) # Só funciona com a base atulizada, pois a uma comparação entre hoje e ontem

maior= automacao.fmaior_estoque(df_two)
menor= automacao.fmenor_estoque(df_two)
