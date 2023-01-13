import pandas as pd 
from docxtpl import DocxTemplate, RichText
doc = DocxTemplate("teste.docx")
# ler planilha 
df = pd.read_excel('planilha.xlsx', dtype=str)








# extrair informações da planilha

total_por_eqpt = df.loc[df['mes'] == 'TOTAL POR EQUIPAMENTO'].to_dict('list')
print(total_por_eqpt)

total_2020 = total_por_eqpt[2020]
total_2021 = total_por_eqpt[2021]

totais = [total_2020, total_2021]

for chave,valor in total_por_eqpt.items():
    

    if chave == 2020:
        context = {'2020': valor}
        doc.render(context)
        doc.save(f'')
    if chave == 2021:


""" x = 0 
for total in totais:
    context = {'num_total': total[0]}
    doc.render(context)
    doc.save(f'doc{x}.docx')
    x = x + 1 """




# transferir as informações para o doc

