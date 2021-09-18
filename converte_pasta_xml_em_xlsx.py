from xml.dom import minidom
from os import listdir, path
from openpyxl import Workbook
import time

def find_fields(xml_file_name, elms_attrs):
	xmldoc = minidom.parse(xml_file_name) # Read XML file
	cur_list = []
	for elms in elms_attrs:
		itemlist = xmldoc.getElementsByTagName(elms['tag']) # Look for tag
		if elms['count']: # We only want a count of this tag
			cur_list.append(len(itemlist))
		else:
			if (len(itemlist)==0): # Tag not found
				cur_list.append('')
			else: # Tag found
				attrs = [i.getAttribute(cur_attr) for i in itemlist for cur_attr in elms['attr']]
				if elms['get_unq']:
					attrs = list(set(attrs))
				s = '/'
				if elms['fnd_blnk']:
					l = len(elms['attr'])
					try:
						cur_index = attrs[0:-1:l].index('')
					except ValueError:
						cur_index = -1
					if cur_index!=-1:
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
	ws.append([e['fld_name'] for e in elms_attrs])
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
# folder = "data/"

xml_fields = [{'fld_name':'Nome', 'tag':'DADOS-GERAIS', 'attr':['NOME-COMPLETO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'País de nascimento', 'tag':'DADOS-GERAIS', 'attr':['PAIS-DE-NASCIMENTO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Data de atualização', 'tag':'CURRICULO-VITAE', 'attr':['DATA-ATUALIZACAO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Id Lattes', 'tag':'CURRICULO-VITAE', 'attr':['NUMERO-IDENTIFICADOR'] , 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Instituição de atuação', 'tag':'ENDERECO-PROFISSIONAL', 'attr':['NOME-INSTITUICAO-EMPRESA','NOME-ORGAO','NOME-UNIDADE'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Vínculo', 'tag':'VINCULOS', 'attr':['ANO-FIM','OUTRO-ENQUADRAMENTO-FUNCIONAL-INFORMADO','OUTRO-VINCULO-INFORMADO','TIPO-DE-VINCULO'], 'count':False, 'fnd_blnk':True, 'get_unq':False},
	{'fld_name':'País de atuação', 'tag':'ENDERECO-PROFISSIONAL', 'attr':['PAIS'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Cidade', 'tag':'ENDERECO-PROFISSIONAL', 'attr':['CIDADE'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Estado', 'tag':'ENDERECO-PROFISSIONAL', 'attr':['UF'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'CEP', 'tag':'ENDERECO-PROFISSIONAL', 'attr':['CEP'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Doutorado (instituição/código/conclusão/obtenção)', 'tag':'DOUTORADO', 'attr':['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Mestrado (instituição/código/conclusão/obtenção)', 'tag':'MESTRADO', 'attr':['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Especialização (instituição/código/conclusão/obtenção)', 'tag':'ESPECIALIZACAO', 'attr':['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Graduação (instituição/código/conclusão/obtenção)', 'tag':'GRADUACAO', 'attr':['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Grande área de atuação', 'tag':'AREA-DE-ATUACAO', 'attr':['NOME-GRANDE-AREA-DO-CONHECIMENTO'], 'count':False, 'fnd_blnk':False, 'get_unq':True},
	{'fld_name':'Área de atuação', 'tag':'AREA-DE-ATUACAO', 'attr':['NOME-DA-AREA-DO-CONHECIMENTO'], 'count':False, 'fnd_blnk':False, 'get_unq':True},
	{'fld_name':'Sub-área de atuação', 'tag':'AREA-DE-ATUACAO', 'attr':['NOME-DA-SUB-AREA-DO-CONHECIMENTO'], 'count':False, 'fnd_blnk':False, 'get_unq':True},
	{'fld_name':'Especialidade', 'tag':'AREA-DE-ATUACAO', 'attr':['NOME-DA-ESPECIALIDADE'], 'count':False, 'fnd_blnk':False, 'get_unq':True},
	{'fld_name':'Trabalhos em eventos', 'tag':'TRABALHO-EM-EVENTOS', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Artigos publicados', 'tag':'ARTIGO-PUBLICADO', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Livros e capítulos', 'tag':'CAPITULO-DE-LIVRO-PUBLICADO', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Participação em projetos', 'tag':'PARTICIPACAO-EM-PROJETO', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Patentes', 'tag':'PATENTE', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Processos ou técnicas', 'tag':'PROCESSOS-OU-TECNICAS', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Trabalho técnico', 'tag':'TRABALHO-TECNICO', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Orientações (doutorado)', 'tag':'ORIENTACOES-CONCLUIDAS-PARA-DOUTORADO', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Orientações (mestrado)', 'tag':'ORIENTACOES-CONCLUIDAS-PARA-MESTRADO', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Orientações (outras)', 'tag':'OUTRAS-ORIENTACOES-CONCLUIDAS', 'attr':[''], 'count':True, 'fnd_blnk':False, 'get_unq':False},
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