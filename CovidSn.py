#NioulBoy
#Load in Our Data
# Section 1 - Loading our Libraries
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
import matplotlib.ticker as ticker
import os


# Section 2 - Loading and Selecting Data\
#rreeqd csv file online
df = pd.read_csv('https://raw.githubusercontent.com/datasets/covid-19/master/data/countries-aggregated.csv', parse_dates=['Date'])
#create variables to store countries
countries = ['Senegal']
countt = ['Senegal']
#date = ['2020-07-16']
#df = df[(df['Country'].isin(countt)) & (df['Date'].isin(date))]
#filter df with country listed on countrries
df = df[df['Country'].isin(countries)]
df_conf = df.sort_values(by=['Date'], ascending=False)[:10]
print(df_conf)


#with pd_conf.
df_conf.plot(x='Date',y= 'Confirmed')
#            x='Date', y='Deaths')
df_conf.plot(x = 'Date', y='Deaths')


plt.show()


#df_Confirmed = df.groupby(['Confirmed']).sum()
#df_Confirmed.plot.pie()
#plt.show()


"""
df_confirmed = df.get('Confirmed')
print(df_confirmed)
df_confirmed(x='Date', y='Confirmed')
plt.show()
"""
