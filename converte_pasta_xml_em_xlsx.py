from xml.dom import minidom
from os import listdir, path
from openpyxl import Workbook
import time

def find_fields(xml_file_name, elms_attrs):
	xmldoc = minidom.parse(xml_file_name) # Read XML file
	cur_list = []
	for elms in elms_attrs:
		itemlist = xmldoc.getElementsByTagName(elms[1]) # Look for tag
		if elms[3]: # elms[3]==True indicates we want a mere count of this tag
			cur_list.append(len(itemlist))
		else:       # elms[3]==False indicates we want a string with all found tags and attributes
			if (len(itemlist)==0): # Tag not found
				cur_list.append('')
			else: # Tag found
				attrs = [i.getAttribute(cur_attr) for i in itemlist for cur_attr in elms[2]]
				s = '/'
				if elms[4]:
					l = len(elms[2])
					try:
						cur_index = attrs[0:-1:l].index('')
					except ValueError:
						cur_index = -1
					if cur_index!=1:
						cur_index *= l
						cur_list.append(s.join(attrs[cur_index+1:cur_index+l]))
					else:
						cur_list.append('')
				else:
					cur_list.append(s.join(attrs))
	return cur_list

def xmls_2_xlsx(xml_file_folder, elms_attrs, output_file_name):
	wb = Workbook()
	# ws = wb.create_sheet()
	ws = wb.active
	ws.append([e[0] for e in elms_attrs])
	cont = 0
	t0 = time.perf_counter()
	for arq in listdir(xml_file_folder):
		if arq[-4:] == '.xml':
			cont += 1
			# if cont<=100:
			ws.append(find_fields(xml_file_folder+arq, elms_attrs))
			if cont % 100 == 0:
				t1 = time.perf_counter() 
				print("%d: %s (%f s)" % (cont,arq,t1-t0))
				t0 = t1
	newfilename = path.abspath(output_file_name)
	wb.save(newfilename)
	return

folder = "data/2021_04/"
folder = "data/"

xml_fields = [['Nome','DADOS-GERAIS',['NOME-COMPLETO'], False, False],
	['País de nascimento','DADOS-GERAIS',['PAIS-DE-NASCIMENTO'], False, False],
	['Data de atualização', 'CURRICULO-VITAE', ['DATA-ATUALIZACAO'], False, False],
	['Id Lattes', 'CURRICULO-VITAE', ['NUMERO-IDENTIFICADOR'] , False, False],
	['Instituição de atuação', 'ENDERECO-PROFISSIONAL', ['NOME-INSTITUICAO-EMPRESA','NOME-ORGAO','NOME-UNIDADE'], False, False],
	['Vínculo', 'VINCULOS', ['ANO-FIM','OUTRO-ENQUADRAMENTO-FUNCIONAL-INFORMADO','OUTRO-VINCULO-INFORMADO','TIPO-DE-VINCULO'], False, True],
	['País de atuação', 'ENDERECO-PROFISSIONAL', ['PAIS'], False, False],
	['Cidade', 'ENDERECO-PROFISSIONAL', ['CIDADE'], False, False],
	['Estado', 'ENDERECO-PROFISSIONAL', ['UF'], False, False],
	['CEP', 'ENDERECO-PROFISSIONAL', ['CEP'], False, False],
	['Doutorado (instituição/código/conclusão/obtenção)', 'DOUTORADO', ['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], False, False],
	['Mestrado (instituição/código/conclusão/obtenção)', 'MESTRADO', ['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], False, False],
	['Especialização (instituição/código/conclusão/obtenção)', 'ESPECIALIZACAO', ['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], False, False],
	['Graduação (instituição/código/conclusão/obtenção)', 'GRADUACAO', ['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], False, False],
	['Grande área de atuação', 'AREA-DE-ATUACAO', ['NOME-GRANDE-AREA-DO-CONHECIMENTO'], False, False],
	['Área de atuação', 'AREA-DE-ATUACAO', ['NOME-DA-AREA-DO-CONHECIMENTO'], False, False],
	['Sub-área de atuação', 'AREA-DE-ATUACAO', ['NOME-DA-SUB-AREA-DO-CONHECIMENTO'], False, False],
	['Especialidade', 'AREA-DE-ATUACAO', ['NOME-DA-ESPECIALIDADE'], False, False],
	['Trabalhos em eventos', 'TRABALHO-EM-EVENTOS', [''], True, False],
	['Artigos publicados', 'ARTIGO-PUBLICADO', [''], True, False],
	['Livros e capítulos', 'CAPITULO-DE-LIVRO-PUBLICADO', [''], True, False],
	['Participação em projetos', 'PARTICIPACAO-EM-PROJETO', [''], True, False],
	['Patentes', 'PATENTE', [''], True, False],
	['Processos ou técnicas', 'PROCESSOS-OU-TECNICAS', [''], True, False],
	['Trabalho técnico', 'TRABALHO-TECNICO', [''], True, False],
	['Orientações (doutorado)', 'ORIENTACOES-CONCLUIDAS-PARA-DOUTORADO', [''], True, False],
	['Orientações (mestrado)', 'ORIENTACOES-CONCLUIDAS-PARA-MESTRADO', [''], True, False],
	['Orientações (outras)', 'OUTRAS-ORIENTACOES-CONCLUIDAS', [''], True, False],
	]

# País de atuação vs país de graduação - atuação ok, graduação não
# Incluir códigos das instituições (graduação, mestrado e doutorado) - ok
# Conferir origem desses códigos (Google, email ou falar com fuinha)
#    http://di.cnpq.br/di/index.jsp

# Instituição atual onde trabalha é o mais importante. Vínculo atual é importante também - ok
# Incluir contagem de projetos - ok
# Separar brasileiros formados na Itália (graduação etc.)
# EXTRA: Buscar Web of science e Scopus
# https://openpyxl.readthedocs.io/en/stable/usage.html
xmls_2_xlsx(folder, xml_fields, "xml_to_excel.xlsx")