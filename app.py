# word
from docx import Document # nos permite criar arquivos word

documento = Document()

# para add um titulo
documento.add_heading('Titulo do documento', 0) # O 0 significa o nível do texto, são varios niveis, ver na documentação



# obs: Toda vez que salvar com o mesmo nome, o arquivo será sobescrevido. Caso queira manter o arquivo atual, mude o nome do save.


# Adicionando um paragrafo
paragrafo = documento.add_paragraph('Um parágrafo simples')
paragrafo.add_run(' e super importante ').bold = True  # deixando esta frase em negrito.
paragrafo.add_run(' do autor ') # sem nenhum tipo de formatação
paragrafo.add_run(' jonathan ').italic = True  # deixando em italico


# Adicionar headings(cabeçalho)
documento.add_heading('Titulo Nível 1', level=1)
documento.add_heading('Titulo Nível 2', level=2)
documento.add_heading('Titulo Nível 3', level=3)
documento.add_heading('Titulo Nível 4', level=4)

# Formatação de estilo
documento.add_paragraph('Formatação "No spacing"', style='No Spacing')


# Salvando o documento criado
documento.save('demo.docx')  # o nome