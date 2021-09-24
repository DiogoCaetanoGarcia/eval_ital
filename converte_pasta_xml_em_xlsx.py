from xml.dom import minidom
from os import listdir, path
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
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

def xmls_2_xlsx(xml_file_folder, elms_attrs, analysis_list, chart_data, output_file_name):
	wb = Workbook()
	ws = wb.create_sheet('Dados')
	del wb['Sheet']
	ws = wb.get_sheet_by_name('Dados')
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
	wb.create_sheet('Italianos')
	ws = wb.get_sheet_by_name('Italianos')
	ws.append(['=ARRAYFORMULA({Dados!A1,Dados!C1:AB1})'])
	ws.append(['=FILTER({Dados!A:A,Dados!C:AB}, Dados!B:B="Itália")'])
	wb.create_sheet('Análises')
	ws = wb.get_sheet_by_name('Análises')
	ws.append([i[0] for i in analysis_list])
	ws.append([i[1] for i in analysis_list])
	for c in chart_data:
		if c[0]=='BarChart':
			chart = BarChart()
			chart.x_axis.title = c[4]
			chart.y_axis.title = c[5]
		if c[0]=='PieChart':
			chart = PieChart()
		if len(c[1]):
			chart.type = c[1]
		chart.title = c[3]
		data = Reference(ws, min_col=c[6][0], max_col=c[6][1], min_row=c[6][2], max_row=c[6][3])
		cats = Reference(ws, min_col=c[7][0], max_col=c[7][1], min_row=c[7][2], max_row=c[7][3])
		chart.add_data(data, titles_from_data=True)
		chart.set_categories(cats)
		cs = wb.create_chartsheet(c[2])
		cs.add_chart(chart)
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
	{'fld_name':'Doutorado', 'tag':'DOUTORADO', 'attr':['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Mestrado', 'tag':'MESTRADO', 'attr':['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Especialização', 'tag':'ESPECIALIZACAO', 'attr':['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
	{'fld_name':'Graduação', 'tag':'GRADUACAO', 'attr':['NOME-INSTITUICAO','CODIGO-INSTITUICAO','ANO-DE-CONCLUSAO','ANO-DE-OBTENCAO-DO-TITULO'], 'count':False, 'fnd_blnk':False, 'get_unq':False},
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
analysis_list = [['Total de italianos', '=COUNTIF(Dados!B:B,"Itália")'],
	['', ''],
	['País de nascimento', '=SORT(UNIQUE(Dados!B2:B1048576))'],
	['Contagem', '=ARRAYFORMULA(COUNTIF(Dados!B2:B1048576,SORT(UNIQUE(Dados!B2:B1048576))))'],
	['', ''],
	['Estado', '={"AC"; "AL"; "AP"; "AM"; "BA"; "CE"; "DF"; "ES"; "GO"; "MA"; "MT"; "MS"; "MG"; "PA"; "PB"; "PR"; "PE"; "PI"; "RJ"; "RN"; "RS"; "RO"; "RR"; "SC"; "SP"; "SE"; "TO"}'],
	['Capital', '={"Rio Branco"; "Maceió"; "Macapá"; "Manaus"; "Salvador"; "Fortaleza"; "Brasília"; "Vitória"; "Goiânia"; "São Luís"; "Cuiabá"; "Campo Grande"; "Belo Horizonte"; "Belém"; "João Pessoa"; "Curitiba"; "Recife"; "Teresina"; "Rio de Janeiro"; "Natal"; "Porto Alegre"; "Porto Velho"; "Boa Vista"; "Florianópolis"; "São Paulo"; "Aracaju"; "Palmas"}'],
	['Contagem', '=ARRAYFORMULA(COUNTIF(Italianos!H2:H1048576,filter(F2:F28,F2:F28<>"")))'],
	['', ''],
	['Produção', '={"Artigos publicados"; "Livros e capítulos"; "Participação em projetos"; "Patentes"; "Processos ou técnicas"; "Trabalho técnico"; "Orientações (doutorado)"; "Orientações (mestrado)"; "Orientações (outras)"}'],
	['Contagem absoluta', '={sum(Italianos!S:S); sum(Italianos!T:T); sum(Italianos!U:U); sum(Italianos!V:V); sum(Italianos!W:W); sum(Italianos!X:X); sum(Italianos!Y:Y); sum(Italianos!Z:Z); sum(Italianos!AA:AA)}'],
	['Contagem relativa', '={K2/$A$2; K3/$A$2; K4/$A$2; K5/$A$2; K6/$A$2; K7/$A$2; K8/$A$2; K9/$A$2; K10/$A$2}'],
	['=J1', '=SORT(J2:K1048576,2,FALSE)'],
	['=K1', ''],
	['=M1', '=SORT({J2:J1048576,L2:L1048576},2,FALSE)'],
	['=L1',''],
	['', ''],
	['=Italianos!N1', '=sort(unique(transpose(split(join("/",unique(Italianos!N2:N1048576)),"/"))))'],
	['Contagem', '=ARRAYFORMULA(COUNTIF(Italianos!$N$2:$N$1048576,ARRAYFORMULA("*" & filter(R2:R1048576,R2:R1048576<>"") & "*")))'],
	['=R1', '=SORT(R2:S1048576,2,FALSE)'],
	['Contagem', ''],
	['', ''],
	['=Italianos!O1', '=sort(unique(transpose(split(join("/",unique(Italianos!O2:O1048576)),"/"))))'],
	['Contagem', '=ARRAYFORMULA(COUNTIF(Italianos!$O$2:$O1048576,ARRAYFORMULA("*" & filter(W2:W1048576,W2:W1048576<>"") & "*")))'],
	['=W1', '=SORT(W2:X1048576,2,FALSE)'],
	['Contagem', ''],
	['',''],
	['Formação', '={Italianos!J1;Italianos!K1;Italianos!L1;Italianos!M1}'],
	['Contagem', '={COUNTA(Italianos!J2:J1048576); COUNTA(Italianos!K2:K1048576); COUNTA(Italianos!L2:L1048576); COUNTA(Italianos!M2:M1048576)}'],
	['', ''],
	['', ''],
	['=Italianos!P1', '=sort(unique(transpose(split(join("/",unique(Italianos!P2:P1048576)),"/"))))'],
	['Contagem', '=ARRAYFORMULA(COUNTIF(Italianos!$P$2:$P1048576,ARRAYFORMULA("*" & filter(AF2:AF1048576,AF2:AF1048576<>"") & "*")))'],
	['=AF1', '=SORT(AF2:AG1048576,2,FALSE)'],
	['Contagem', ''],
	['', ''],
	['=Italianos!Q1', '=sort(unique(transpose(split(join("/",unique(Italianos!Q2:Q1048576)),"/"))))'],
	['Contagem', '=ARRAYFORMULA(COUNTIF(Italianos!$Q$2:$Q1048576,ARRAYFORMULA("*" & filter(AK2:AK1048576,AK2:AK1048576<>"") & "*")))'],
	['=AK1', '=SORT(AK2:AL1048576,2,FALSE)'],
	['Contagem', ''],
	]

chart_data = [['BarChart', 'bar', 'Formação', 'Formação', '', 'Contagem', [29,29,1,5], [28,28,2,5]],
	['BarChart', 'bar', 'Grandes_áreas', 'Grandes áreas de formação', '', 'Contagem', [21,21,1,1000], [20,20,2,1000]],
	['PieChart', 'bar', 'Grandes_áreas_perc', 'Grandes áreas de formação', '', 'Contagem', [21,21,1,1000], [20,20,2,1000]],
	['BarChart', 'bar', 'Áreas', 'Área de atuação', '', 'Contagem', [26,26,1,1000], [25,25,2,1000]],
	['PieChart', 'bar', 'Áreas_perc', 'Área de atuação', '', 'Contagem', [26,26,1,1000], [25,25,2,1000]],
	['BarChart', 'bar', 'Produção', 'Produção', '', 'Contagem', [14,14,1,1000], [13,13,2,1000]],
	['BarChart', 'bar', 'Produção_relativa', 'Produção/pessoa', '', 'Contagem', [16,16,1,1000], [15,15,2,1000]],
	['PieChart', 'bar', 'Italianos_estado', 'Italianos por estado', '', 'Contagem', [8,8,1,30], [6,6,2,30]],
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
xmls_2_xlsx(folder, xml_fields, analysis_list, chart_data, "xml_to_excel.xlsx")