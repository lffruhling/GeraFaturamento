from docxtpl import DocxTemplate

## Este trcho tem por objetivo carregar um .docx template, onde substituimos tags pelas informações
## As tags são identificadas entre chaves e atribuídas dentro do context

template = DocxTemplate('C:/Temp/Template.docx')

tipo = "IPCA"
lancamentos = []

def alimentaLancamentos(lista, data, descricao, valor, correcao, corrigido, juros, total):
    lista.append({"data":data, "descricao":descricao, "valor":valor,"correcao":correcao,"corrigido":corrigido,"juros":juros,"total":total})

#teste    
alimentaLancamentos(lancamentos, "01/01/2001","Teste","1500","0","0","0","1500")

context = {
    "nome_associado" : "Andre do Teste",
    "tipo_correcao"  : tipo,
    "numero_titulo"  : "123456-7",
    "forma_calculo"  : "Parcelas Atualizadas Individualmente De 27/09/2013 a 11/04/2023 sem correção Multa de 2,0000 sobre o valor corrigido+juros principais+juros moratórios",
    "forma_juros"    : "Juros ok",
    "lancamentos"    : lancamentos
}

template.render(context)
template.save('C:/Temp/alimentado_' + tipo + '.docx')