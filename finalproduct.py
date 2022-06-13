from numpy.core.fromnumeric import mean
from openpyxl import load_workbook
import glob
import pandas as pd
import numpy as np

df = pd.DataFrame()
for f in glob.glob("*.csv"):
    data = pd.read_csv(f)
    df = df.append(data, ignore_index=True)

newdf = pd.DataFrame()
anotherdf = pd.DataFrame()
df2 = pd.DataFrame()
df3 = pd.DataFrame()
df6 = pd.DataFrame()
df7 = pd.DataFrame()
finaldf = pd.DataFrame()
stocksdf = pd.DataFrame()

writer = pd.ExcelWriter('Summary.xlsx', engine="openpyxl")
wb = writer.book

newdf['Max Move'] = df['Max Move'].abs()
newdf['Symbol'] = df['Symbol']
answer = newdf.groupby('Symbol')['Max Move'].mean()
newnew = answer.to_frame()
df2['Average'] = newnew['Max Move']

df3['Max Move'] = df['Max Move']
df3['Symbol'] = df['Symbol']
df4, df5 = [x for _, x in df3.groupby(df3['Max Move'] < 0)]

hello = df4.groupby('Symbol')['Max Move'].count()
again = df5.groupby('Symbol')['Max Move'].count()
percentage = again/hello*100

amax_value = df3.groupby('Symbol')['Max Move'].max()
amin_value = df3.groupby('Symbol')['Max Move'].min()
variable = hello.to_frame()
variable2 = again.to_frame()
var = amax_value.to_frame()
var1 = amin_value.to_frame()
apercentage = percentage.to_frame()
df6['Losses'] = variable['Max Move']
df6['Gaines'] = variable2['Max Move']
df6['Largest Value'] = var['Max Move']
df6['Lowest Value'] = var1['Max Move']
df6['Percentage'] = apercentage['Max Move']

anotherdf = pd.concat([df2], ignore_index= True)
anotherdf['Symbol'] = df2.index

anotherdf.drop_duplicates(inplace=True)

df7 = pd.concat([df6], ignore_index= True)
df7['Symbol'] = anotherdf ['Symbol']
df7['Average'] = anotherdf['Average']

finaldf['Symbol'] = df7['Symbol']
finaldf['Largest Move'] = df7['Largest Value']
finaldf['Lowest Move'] = df7['Lowest Value']
finaldf['Gains'] = df7['Gaines']
finaldf['Losses'] = df7['Losses']
finaldf['Percentage'] = df7['Percentage']
finaldf['Average Move'] = df7['Average']

stocksdf['Symbol'] = finaldf['Symbol']
stocksdf['Percentage'] = finaldf['Percentage']
stocksdf['Average Move'] = finaldf['Average Move']
stocksdf['Post Gains'] = finaldf['Gains']
stocksdf['Post Losses'] = finaldf['Losses']
stocksdf = stocksdf.sort_values('Percentage', ascending = False)

stocksdf.to_excel(writer, index= False)
wb.save('Summary.xlsx')

print(stocksdf)
