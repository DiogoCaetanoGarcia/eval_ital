# import numpy as np
import pandas as pd
# import matplotlib.pyplot as plt
# import geopandas as gpd
# import pycountry 
import pygal
import json
import plotly as plt
import plotly.express as px

def alpha2code(df, iso_file, sep):
	iso_df = pd.read_csv(iso_file, sep=sep)
	# return [iso_df['alfa-3'][iso_df['País']==country].item() for country in df]
	CODE=[]
	for country in df:
		try:
			code=iso_df['alfa-2'][iso_df['País']==country].item()
			CODE.append(code.lower())
		except:
			CODE.append('')
	return CODE

# def plot_estados(df, json_data_file):
# 	Brazil = json.load(json_data_file)
# 	state_id_map = {}
# 	for feature in Brazil ['features']:
# 		feature['id'] = feature['properties']['name']
# 		state_id_map[feature['properties']['sigla']] = feature['id']
# 	estado_hist = df['Estado'].value_counts()
# 	estado_hist = pd.DataFrame({'Estado':estado_hist.axes[0], 
# 		'Contagem':estado_hist.values})
# 	fig = px.choropleth(
#  		estado_hist, 
#  		locations = "Estado", #define the limits on the map/geography
#  		geojson = Brazil, #shape information
#  		color = "Contagem", #defining the color of the scale through the database
#  		#hover_name = "Estado", #the information in the box
#  		#hover_data =["Produção","Longitude","Latitude"],
#  		title = "Italianos no Brasil", #title of the map
#  		#animation_frame = "ano" #creating the application based on the year
# 	)
# 	fig.update_geos(fitbounds = "locations", visible = False)
# 	fig.show()

df = pd.read_excel("xml_to_excel.xlsx", engine='openpyxl')

worldmap = pygal.maps.world.World()
worldmap.title = 'País de nascimento'
pais_hist = df['País de nascimento'].value_counts()
worldmap.add('País de nascimento', dict(zip(alpha2code(pais_hist.axes[0], 'data/ISO_3166_1.csv', '\t'),
	pais_hist.values)))
worldmap.render_to_file('país_de_nascimento.svg')
# pais_hist = df['País de nascimento'].value_counts()
# pais_hist = pd.DataFrame({'País de nascimento':pais_hist.axes[0], 
# 	'CODE': alpha2code(pais_hist.axes[0], 'data/ISO_3166_1.csv', '\t'),
# 	'Contagem':pais_hist.values})
# print(pais_hist)

df_italia = df[df['País de nascimento']=='Itália']
N_italianos = len(df_italia)
print('%d italianos' % N_italianos)

# plot_estados(df, 'data/brazil-states.geojson')

# estados = {'AC': 'Rio Branco', 'AL': 'Maceió', 'AP': 'Macapá', 'AM': 'Manaus', 'BA': 'Salvador', 'CE': 'Fortaleza', 'DF': 'Brasília', 'ES': 'Vitória', 'GO': 'Goiânia', 'MA': 'São Luís', 'MT': 'Cuiabá', 'MS': 'Campo Grande', 'MG': 'Belo Horizonte', 'PA': 'Belém', 'PB': 'João Pessoa', 'PR': 'Curitiba', 'PE': 'Recife', 'PI': 'Teresina', 'RJ': 'Rio de Janeiro', 'RN': 'Natal', 'RS': 'Porto Alegre', 'RO': 'Porto Velho', 'RR': 'Boa Vista', 'SC': 'Florianópolis', 'SP': 'São Paulo', 'SE': 'Aracaju', 'TO': 'Palmas'}