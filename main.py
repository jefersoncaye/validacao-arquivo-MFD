import pandas as pd

arq = open('arquivo MFD.txt', 'rt', encoding='utf-8')
arq2 = open('arquivo MFD.txt', 'rt', encoding='utf-8')
arq2 = arq2.readlines()
lista_linhas = arq.readlines()


pulos = []
for i in range(len(arq2)):
    if 'COO:' in arq2[i] and 'CCF' in arq2[i]:
        pulos.append(i)

lista = []

for indice, pulo in enumerate(pulos):
    dicionario = {}
    count_c = 0
    lista_item_c = []
    lista_item_d = []
    lista_item_d_i = []
    count_d = 0
    try:
        for j in lista_linhas[pulo:pulos[indice+1]]:
            lista_select = []
            if 'COO:' in j and 'CCF' in j:
                cupom = j[j.index('COO:') + 4:-1]
                dicionario['CUPOM'] = cupom
            if 'CUPOM FISCAL CANCELADO' in j:
                dicionario['CUPOM_CANCELADO'] = 'SIM'
                lista_select.append(f'SELECT X_CANCELADO, * FROM TBL_PDV WHERE NR_CUPOM_FISCAL = {cupom};')
            if 'cancelamento de item:' in j:
                lista_item_c.append(j[21:24])
                dicionario['ITEM_CANCELADO'] = 'SIM'
                count_c += 1
                dicionario['QNT_ITEM_CANCELADO'] = count_c
            if 'desconto item' in j:
                lista_item_d.append(f'I: {j[15:19]} {j[42:-1]}')
                lista_item_d_i.append({j[15:19]})
                dicionario['ITEM_DESCONTO'] = 'SIM'
                count_d += 1
                dicionario['QNT_ITEM_DESCONTO'] = count_d
            if 'TOTAL  R$' in j:
                dicionario['TOTAL'] = str(j[j.index('TOTAL  R$') + 10:-1].split()).replace("['", '').replace("']", '')
            if 'TROCO  R$' in j:
                dicionario['TROCO'] = str(j[j.index('TROCO  R$') + 10:-1].split()).replace("['", '').replace("']", '')
            if 'DESCONTO ' in j:
                dicionario['DESCONTO_CUPOM'] = 'SIM'
                dicionario['VL_DESCONTO_CUPOM'] = j[j.index('-'):-1]
                lista_select.append(f'SELECT VL_DESCONTO, VL_DESCONTO_RATEADO, * FROM TBL_PDV WHERE NR_CUPOM_FISCAL = {cupom};')
            if len(lista_item_c) > 0:
                dicionario['ITENS_CANCELADOS'] = lista_item_c
                lista_select.append(f'SELECT PVI.X_CANCELADO, * FROM TBL_PDV P INNER JOIN TBL_PDV_ITENS PVI ON '
                                    f'P.CD_PDV = PVI.CD_PDV WHERE P.NR_CUPOM_FISCAL = {cupom} AND CD_ITEM IN ({lista_item_c});')
            if len(lista_item_d) > 0:
                dicionario['ITENS_DESCONTO'] = lista_item_d
                lista_select.append(
                    f'SELECT PVI.VL_DESCONTO, PVI.VL_DESCONTO_RATEADO, * FROM TBL_PDV P INNER JOIN TBL_PDV_ITENS PVI ON '
                    f'P.CD_PDV = PVI.CD_PDV WHERE P.NR_CUPOM_FISCAL = {cupom} AND CD_ITEM IN ({lista_item_d_i});')
            if len(lista_select) > 0:
                dicionario['SELECTS'] = lista_select
        lista.append(dicionario)
    except:
        break

dataframe = pd.DataFrame(lista)

dataframe = dataframe[['CUPOM', 'TOTAL', 'TROCO', 'CUPOM_CANCELADO', 'ITEM_CANCELADO', 'QNT_ITEM_CANCELADO',
                       'ITENS_CANCELADOS', 'DESCONTO_CUPOM', 'VL_DESCONTO_CUPOM', 'ITEM_DESCONTO',
                       'QNT_ITEM_DESCONTO', 'ITENS_DESCONTO', 'SELECTS']]

dataframe.to_excel(r'caminho\MFD.xlsx', index=False)
