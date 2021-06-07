from xml.dom import minidom
from os import listdir, path
from openpyxl import Workbook
import time

def find_fields(xml_file_name, elms_attrs):
	xmldoc = minidom.parse(xml_file_name)
	cur_list = []
	for elms in elms_attrs:
		itemlist = xmldoc.getElementsByTagName(elms[1])
		if (len(itemlist)==0) or (not itemlist[0].getAttribute(elms[2])):
			if elms[3]:
				cur_list.append("0")
			else:
				cur_list.append('')
		else:
			if elms[3]:
				cur_list.append(str(len(itemlist)))
			else:
				if len(itemlist)==1:
					cur_list.append(itemlist[0].attributes[elms[2]].value)
				else:
					vals = [il.attributes[elms[2]].value for il in itemlist]
					s = ', '
					cur_list.append(s.join(list(set(vals))))
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

folder = "data/2021_04/" #Dados/"
# folder = "Testes/"

xml_fields = [
	['Nome','DADOS-GERAIS','NOME-COMPLETO', False],
	['País de nascimento','DADOS-GERAIS','PAIS-DE-NASCIMENTO', False],
	['Data de atualização', 'CURRICULO-VITAE', 'DATA-ATUALIZACAO', False],
	['Id Lattes', 'CURRICULO-VITAE', 'NUMERO-IDENTIFICADOR', False],
	['País de atuação', 'ENDERECO-PROFISSIONAL', 'PAIS', False],
	['Cidade', 'ENDERECO-PROFISSIONAL', 'CIDADE', False],
	['Estado', 'ENDERECO-PROFISSIONAL', 'UF', False],
	['Instituição', 'ENDERECO-PROFISSIONAL', 'NOME-INSTITUICAO-EMPRESA', False],
	['Órgão', 'ENDERECO-PROFISSIONAL', 'NOME-ORGAO', False],
	['Unidade', 'ENDERECO-PROFISSIONAL', 'NOME-UNIDADE', False],
	['CEP', 'ENDERECO-PROFISSIONAL', 'CEP', False],
	['Doutorado', 'DOUTORADO', 'ANO-DE-CONCLUSAO', False],
	['Doutorado', 'DOUTORADO', 'ANO-DE-OBTENCAO-DO-TITULO', False],
	['Mestrado', 'MESTRADO', 'ANO-DE-CONCLUSAO', False],
	['Mestrado', 'MESTRADO', 'ANO-DE-OBTENCAO-DO-TITULO', False],
	['Especialização', 'ESPECIALIZACAO', 'ANO-DE-CONCLUSAO', False],
	['Graduação', 'GRADUACAO', 'ANO-DE-CONCLUSAO', False],
	['Grande área de atuação', 'AREA-DE-ATUACAO', 'NOME-GRANDE-AREA-DO-CONHECIMENTO', False],
	['Área de atuação', 'AREA-DE-ATUACAO', 'NOME-DA-AREA-DO-CONHECIMENTO', False],
	['Sub-área de atuação', 'AREA-DE-ATUACAO', 'NOME-DA-SUB-AREA-DO-CONHECIMENTO', False],
	['Especialidade', 'AREA-DE-ATUACAO', 'NOME-DA-ESPECIALIDADE', False],
	['Trabalhos em eventos', 'TRABALHO-EM-EVENTOS', '', True],
	['Artigos publicados', 'ARTIGO-PUBLICADO', '', True],
	['Livros e capítulos', 'CAPITULO-DE-LIVRO-PUBLICADO', '', True],
	['Patentes', 'PATENTE', '', True],
	['Processos ou técnicas', 'PROCESSOS-OU-TECNICAS', '', True],
	['Trabalho técnico', 'TRABALHO-TECNICO', '', True],
	['Orientações (doutorado)', 'ORIENTACOES-CONCLUIDAS-PARA-DOUTORADO', '', True],
	['Orientações (mestrado)', 'ORIENTACOES-CONCLUIDAS-PARA-MESTRADO', '', True],
	['Orientações (outras)', 'OUTRAS-ORIENTACOES-CONCLUIDAS', '', True],
	]
xmls_2_xlsx(folder, xml_fields, "xml_to_excel.xlsx")