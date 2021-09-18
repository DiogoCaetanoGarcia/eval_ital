# generate country code  based on country name 
import pycountry 
import pandas as pd
import matplotlib.pyplot as plt
import geopandas
import pycountry 

# def alpha3code(df, iso_file, sep):
# 	iso_df = pd.read_csv(iso_file, sep=sep)
# 	# return [iso_df['alfa-3'][iso_df['País']==country].item() for country in df]
# 	CODE=[]
# 	for country in df:
# 		try:
# 			code=iso_df['alfa-3'][iso_df['País']==country].item()
# 			CODE.append(code)
# 		except:
# 			CODE.append('')
# 	return CODE

def alpha3code(column):
	CODE=[]
	for country in column:
		try:
			code=pycountry.countries.get(name=country).alpha_3
			# .alpha_3 means 3-letter country code 
			# .alpha_2 means 2-letter country code
			CODE.append(code)
		except:
			CODE.append('None')
	return CODE

df = pd.read_csv('data/train.csv',sep=',')
df['CODE']=alpha3code(df.Country_Region)
# df['CODE'] = alpha3code(df['Country_Region'], 'data/ISO_3166_1.csv', '\t')
df.head()
# first let us merge geopandas data with our data
# 'naturalearth_lowres' is geopandas datasets so we can use it directly
world = geopandas.read_file(geopandas.datasets.get_path('naturalearth_lowres'))
# rename the columns so that we can merge with our data
world.columns=['pop_est', 'continent', 'name', 'CODE', 'gdp_md_est', 'geometry']
# then merge with our data 
merge=pd.merge(world,df,on='CODE')
# last thing we need to do is - merge again with our location data which contains each country’s latitude and longitude
location=pd.read_csv('https://raw.githubusercontent.com/melanieshi0120/COVID-19_global_time_series_panel_data/master/data/countries_latitude_longitude.csv')
merge=merge.merge(location,on='name').sort_values(by='Fatalities',ascending=False).reset_index()
# plot confirmed cases world map 
merge.plot(column='Confirmed_Cases', scheme="quantiles",
           figsize=(25, 20),
           legend=True,cmap='coolwarm')
plt.title('2020 Jan-May Confirmed Case Amount in Different Countries',fontsize=25)
# add countries names and numbers 
for i in range(0,10):
    plt.text(float(merge.longitude[i]),float(merge.latitude[i]),"{}\n{}".format(merge.name[i],merge.Confirmed_Cases[i]),size=10)
plt.show()
