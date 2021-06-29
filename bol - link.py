import pandas as pd
import numpy as np

bol_data = r"C:\Users\wesl\OneDrive - IRI\Bol.com\Bol_data.xlsx"    ### Laad Bol dataset in, verander naar eigen filepath
link = r"C:\Users\wesl\OneDrive - IRI\Bol.com\Bol_IRI_link.xlsx"    ### Laad Bol IRI link bestand in, verander naar eigen filepath
df1 = pd.read_excel(bol_data)
df2 = pd.read_excel(link)

Unique = df2['Bol.com SubSubGroup'].unique().tolist()
d = {}

for i in Unique:    ###Create dictionary, hier wordt de IRI category uiteindelijk in opgeroepen
    d[i] = {}
    for j in range(len(df2['Bol.com SubSubGroup'])):
        if df2['Bol.com SubSubGroup'][j] == i:
            d[i][df2['Bol.com Chunkname'][j]] = [df2['SMR Category'][j]]
            d[i][df2['Bol.com Chunkname'][j]].append(df2['SMR Type'][j])

def link(BolSubSub,BolChunk):
    try:
        SMR_Cat = d[BolSubSub][BolChunk][0]
    except KeyError:
        SMR_Cat = 'N/A'
    try:
        SMR_Typ = d[BolSubSub][BolChunk][1]
    except KeyError:
        SMR_Typ = 'N/A'
    if SMR_Cat == 'Not yet coded':
        SMR_Cat = 'N/A'
    if SMR_Typ == 'Not yet coded':
        SMR_Typ = 'N/A'
    return pd.Series([SMR_Cat, SMR_Typ], dtype='str')

df1[['SMR Category','SMR Type']] = df1.apply(lambda x: link(x['productSubSubGroup'],x['chunkName']), axis=1)
df1['Geography'] = "Bol.com"
df1 = df1.drop(['productSubSubGroup', 'chunkName', 'chunkId'], axis=1)
df1.rename(columns = {'qtyOrdered': 'Unit sales', 'amtOrderedInclVat': 'Euro sales', 'SMR Category': 'Product_CATEGORY', 'SMR Type': 'Product_TYPE'}, inplace = True)
df1.drop(['Column1'], axis = 1, inplace = True)
df1 = df1[['Geography','Product_TYPE','Euro sales','Product_CATEGORY','Time','Unit sales']]

df3 = pd.read_csv(r"C:\Users\wesl\OneDrive - IRI\Bol.com\EXT_DATADUMP_V9C_C4990RT1_P13_2020.TXT", sep=';',
                  encoding="ISO-8859-1", error_bad_lines=False)
df3['4e kwartaal, 2020'] = df3['4e kwartaal, 2020'].str.replace(',', '.').astype(float)
df3 = df3.groupby(['Geography', 'Measure', 'Product_TYPE']).agg({
    '4e kwartaal, 2020': np.sum,
    'Product_CATEGORY': 'last'}).reset_index()
df3[['Time','Unit sales']] = 'Q4, 2020',1
df3.rename(columns = {'4e kwartaal, 2020' : 'Euro sales'}, inplace = True)
df3.drop(['Measure'],axis = 1, inplace = True)

df3.append(df1)
df3.to_excel(r"C:\Users\wesl\OneDrive - IRI\Bol.com\Dummy_Bol_Appended.xlsx")