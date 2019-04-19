# -*- coding: utf-8 -*-
"""
Created on Thu Jan 24 11:41:52 2019

@author: samir.saci
"""
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import time
import matplotlib as mpl

start = time.time()
pd.set_option('display.max_columns', None), pd.set_option('display.max_colwidth', -1)
pd.options.display.float_format = '{:,}'.format
pictures_file = r'Mois/'     
                           

# Import datasets
def import_csv(Month):
    name = str("Mois\df_mois") + Month + '.csv'
              
    # Importer les parametres
    df_cart = pd.read_excel('Parameters\Cartographie v2.0.xlsx', encoding = 'utf_8_sig',)    
    df_inv = pd.read_excel('Parameters\SH P LOCATION.xlsx', encoding = 'utf_8_sig') 
    xls = pd.ExcelFile('Parameters\Cartographie v2.0.xlsx')
    df_abs = pd.read_excel(xls, 'Abscisses').set_index('Abs').fillna(0)
    df_ord = pd.read_excel(xls, 'Ordonnees').set_index('Ord').fillna(0)
    df_MD = pd.read_excel('Parameters\M_D.xlsx', encoding = 'utf_8_sig')
    
    # Importer le Data Set
    if Month == 'Total':
        name = str("Mois\df_all") + '.csv'
        df = pd.read_csv(name, encoding = 'utf_8_sig')
    else:      
        name = str("Mois\df_mois") + Month + '.csv'
        df = pd.read_csv(name, encoding = 'utf_8_sig')
        # Mapping 
            # 1. Mapping df_inv
        df_inv['ZONE'], df_inv['Alley'], df_inv['Alley Number'], df_inv['Emplacement'] = df_inv['Location'].apply(lambda t: t[0]), df_inv['Location'].apply(lambda t: t[1:3]), df_inv['Location'].apply(lambda t: t[0:3]), df_inv['Location'].apply(lambda t: t[2:])
        df_inv['Cellule'], df_inv['Emplacement'] = df_inv['Location'].apply(lambda t: t[3:5]), df_inv['Location'].apply(lambda t: t[4:])
            # 2. Cartographier 
        df_inv['Coordonnees'] = df_inv['Location'].map(lambda t: abs_ord(t, df_abs, df_ord))
        #   3. Mapping Zone
        df = df.merge(pd.DataFrame(df_inv[['SKU','Location']]), left_on = 'SKU', right_on = 'SKU')
        df['ZONE'], df['Alley'], df['Alley Number'], df['Emplacement'] = df['Location'].apply(lambda t: t[0]), df['Location'].apply(lambda t: t[1:3]), df['Location'].apply(lambda t: t[0:3]), df['Location'].apply(lambda t: t[2:])
        df['Coordonnees'] = df['Location'].map(lambda t: abs_ord(t, df_abs, df_ord))
        df = df.merge(df_MD, left_on = 'SKU', right_on = 'SKU')

    return df, df_cart, df_inv, df_abs, df_ord, df_MD

''' Fonctions d'analyses 
'''
def analysis_top(df, category): #'Brand', 'category', 'Description'
    df_marque = pd.DataFrame(df.groupby(category)['PCS'].sum()/df['PCS'].sum(axis =0))
    df_marque.columns = ['%Pcs']
    #df_marque = df_marque.sort('%Pcs',ascending = False)
    df_marque = df_marque.sort_values('%Pcs',ascending = False)    
    return df_marque
# Cartographie
def abs_ord(Location_Number, df_abs, df_ord):
    if Location_Number == '无拣货库位':
        return 0, 44.5
    else:
        Alley_Number = Location_Number[0:3]
        Cellule = int(Location_Number[3:5])
        x = df_abs.loc[Cellule, Alley_Number]
        y = df_ord.loc[Cellule, Alley_Number]
    return x, y
# Analyses globale par periode (Mois ou Annee)
def figures_analysis(df):                                                           
    monthly_volume, ref, orders = df['PCS'].sum(), df['SKU'].nunique(), df['ORDER NUMBER'].nunique()
    orderlines, linesperorder, qtyperorder = df['ORDER NUMBER'].count(), df['ORDER NUMBER'].count()/df['ORDER NUMBER'].nunique(), df['PCS'].sum()/df['ORDER NUMBER'].nunique()
    brands, categories, pcsperline = df['Brand'].nunique(), df['category'].nunique(), '{0:.2f}'.format((df['PCS'].sum()/df['ORDER NUMBER'].nunique())/(df['ORDER NUMBER'].count()/df['ORDER NUMBER'].nunique()))
    cbm_order, kg_order = df.groupby(['ORDER NUMBER'])['stdcube'].sum().mean(), df.groupby(['ORDER NUMBER'])['stdnetwgt'].sum().mean()
    cities = df['CITY'].nunique()
    
    # Creer la liste
    resultats = [monthly_volume, ref, orders, orderlines, linesperorder, qtyperorder, pcsperline, cbm_order, kg_order, brands, categories, cities]
    return resultats
# Monthly 
def figures_monthly(df):                                                           
    monthly_volume, ref, orders = df['PCS'].sum(), df['SKU'].nunique(), df['ORDER NUMBER'].nunique()
    orderlines, linesperorder, qtyperorder = df['ORDER NUMBER'].count(), df['ORDER NUMBER'].count()/df['ORDER NUMBER'].nunique(), df['PCS'].sum()/df['ORDER NUMBER'].nunique()

    return  monthly_volume, ref, orders, orderlines, linesperorder, qtyperorder


''' Fin de Fonctions d'analyses 
'''  
def global_analysis(df):
    
    ''' I.  General Figures for the whole scope '''
    items = ['Total Volume (Pcs): ', 'Number of References: ', 'Orders: ', 'Orderlines: ', 'Lines per order: ', 
             'Pieces per order: ', 'Pieces per Line:', 'Cbm per order: ', 'Kg per order: ', 'Unique Brands: ', 
             'Unique Categories: ', 'Unique Destination: ']
    items_month = ['Total Volume (Pcs): ', 'Number of References: ', 'Orders: ', 'Orderlines: ', 'Lines per order: ', 
             'Pieces per order: ']
        # 1. Full Scope (fc)
    resultats_fc = figures_analysis(df)
    df_fc = pd.DataFrame(resultats_fc, items)
    df_fc.columns = ['2017/2018 Analysis: ']
        # 2. 2017 et 2018 
#    df17, df18 = df[df['YEAR'] == 2017], df[df['YEAR'] == 2018]
    resultats_17, resultats_18 = figures_analysis(df[df['YEAR'] == 2017]), figures_analysis(df[df['YEAR'] == 2018])
    df_78 = pd.DataFrame(list(zip(resultats_17, resultats_18)), items)
    df_78.columns = ['2017 Analysis: ', '2018 Analysis: ']
        # 3. Monthly Analysis
    mois = []
    for i in [2017, 2018]:
        for j in range(1, 13):
            mois.append(str(j) + '-' + str(i))
    l1, l2, l3, l4, l5, l6 = [], [], [], [], [], []
    for m in mois:
        x1, x2, x3, x4, x5, x6 = figures_monthly(df[df['MONTH-YEAR'] == m])
        l1.append(x1); l2.append(x2); l3.append(x3); l4.append(x4); l5.append(x5), l6.append(x6)
    df_month = pd.DataFrame(list(zip(l1, l2, l3, l4, l5, l6)), mois)
    df_month.columns = items_month
        # Peaks Orders
    df_daily = pd.DataFrame(df.groupby(['DAY', 'MONTH-YEAR'])['ORDER NUMBER'].nunique()).reset_index()
    df_ord = pd.DataFrame(df_daily.groupby(['MONTH-YEAR'])['ORDER NUMBER'].mean())
    df_ord['Max'] = df_daily.groupby(['MONTH-YEAR'])['ORDER NUMBER'].max()
    df_ord.columns = ['Average Orders/Day', 'Max Orders per Day']
        # Peak Order lines
    df_daily = pd.DataFrame(df.groupby(['DAY', 'MONTH-YEAR'])['ORDER NUMBER'].count()).reset_index()
    df_lines = pd.DataFrame(df_daily.groupby(['MONTH-YEAR'])['ORDER NUMBER'].mean())
    df_lines['Max'] = df_daily.groupby(['MONTH-YEAR'])['ORDER NUMBER'].max()
    df_lines.columns = ['Average Lines/Day', 'Max Lines per Day']
    
    # Groupment par 
    df_city = pd.DataFrame(df.groupby(['CITY'])['PCS'].sum())
    df_city['%PCS'] = df_city['PCS']/df_city['PCS'].sum(axis = 0)
    df_city = df_city.sort_values(['PCS'], ascending = False)
    
    # Analysis volumes per cities
    n_cities, c_n, vol_perc= df['CITY'].nunique(), [], []
    for p in range(1, n_cities + 1):
        c_n.append(p)
        vol_perc.append(df_city['%PCS'].head(p).sum())
    df_couv = pd.DataFrame(vol_perc, c_n).reset_index()
    df_couv.columns = ['Cities Number', '%PCS']
    
    # Volumes per SKU
    df_sku = df.groupby(['SKU','MONTH-YEAR'])['PCS'].sum().reset_index()
    df_sku = pd.pivot_table(df_sku, values = 'PCS', index = 'SKU', columns = 'MONTH-YEAR', aggfunc = 'sum').fillna(0)
    df_sku['TOTAL'] = df_sku.sum(axis = 1)
    df_sku = df_sku.sort_values('TOTAL', ascending = False).reset_index()
    df_sku = df_sku.merge(df_MD, left_on = 'SKU', right_on = 'SKU')
    
    return df_fc, df_78, df_month, df_ord, df_lines, df_ord, df_lines, df_city, df_couv, df_sku

''' Resultats 
'''

''' Plot
'''
def plot_analysis(df_fc, df_78, df_month, df_lines, df_couv):
    
    pictures_file = 'Full Scope Analysis'
       
    print("I. Volume per month (Pcs/Month)")
    ax = df_month.plot.bar(y=['Total Volume (Pcs): '], figsize=(10, 6), color = ['black'])
    ax.yaxis.set_major_formatter(mpl.ticker.StrMethodFormatter('{x:,.0f}'))
    plt.title('2017/2018 Monthly Volumes')
    plt.xlabel('(Month)')
    plt.ylabel('(Pcs/Month)')
    plt.savefig(pictures_file + '\Pieces per Month.png')
    plt.show()
    print('\n')

    # Orders received per day
    print("II. Orders/Order lines per month")
    ax = df_month.plot.bar(y=['Orders: ', 'Orderlines: '], figsize=(10, 6), color = ['blue', 'black'])
    ax.yaxis.set_major_formatter(mpl.ticker.StrMethodFormatter('{x:,.0f}'))
    plt.title('2017/2018 Monthly Volumes')
    plt.xlabel('(Month)')
    plt.ylabel('(Orders/Month)')
    plt.savefig(pictures_file + '\Orders per Month.png')
    plt.show()
    print('\n')

    print("II.Average orders per day")
    ax = df_ord.plot.bar(y=['Average Orders/Day', 'Max Orders per Day'], figsize=(10, 6), color = ['blue', 'black'])
    ax.yaxis.set_major_formatter(mpl.ticker.StrMethodFormatter('{x:,.0f}'))
    plt.title('2017/2018 Daily Orders')
    plt.xlabel('(Month)')
    plt.ylabel('(Orders/Day)')
    plt.savefig(pictures_file + '\Orders per Day.png')
    plt.show()
    print('\n')
    
    ax = df_lines.plot.bar(y=['Average Lines/Day', 'Max Lines per Day'], figsize=(10, 6), color = ['blue', 'black'])
    ax.yaxis.set_major_formatter(mpl.ticker.StrMethodFormatter('{x:,.0f}'))
    plt.title('2017/2018 Daily Order lines')
    plt.xlabel('(Month)')
    plt.ylabel('(Lines/Day)')
    plt.savefig(pictures_file + '\Lines per Day.png')
    plt.show()
    print('\n')
    
    print("III. Couverture par cities")
    fig, ax = plt.subplots(figsize=(10, 6))
    Cities_number, Vol_perc = df_couv['Cities Numbers'], df_couv['%PCS']
    plt.plot(Cities_number, Vol_perc, color = 'black')
    ax.yaxis.set_major_formatter(mpl.ticker.StrMethodFormatter('{x:,.2f}'))
    plt.title('Volume Split by Number of Cities')
    plt.xlabel('Number of Cities')
    plt.ylabel('(%Pcs)')
    plt.savefig(pictures_file + 'Volume split by number of Cities.png')
    plt.show()


            
'''
'''
df, df_cart, df_inv, df_abs, df_ord, df_MD = import_csv('Total')
df_fc, df_78, df_month, df_ord, df_lines, df_ord, df_lines, df_city, df_couv, df_sku = global_analysis(df)
plot_analysis(df_fc, df_78, df_month, df_lines, df_couv)


print("End time {0:.2f} s".format(time.time() - start))

# Final Export
filename = 'Full Scope Analysis\Analysis Sheet.xlsx'
writer = pd.ExcelWriter(filename, engine='xlsxwriter')
df_fc.to_excel(writer, sheet_name='Full Scope Analysis', encoding = 'utf_8_sig')
df_78.to_excel(writer, sheet_name='2017-2018', encoding = 'utf_8_sig')                 # Top SKU
df_month.to_excel(writer, sheet_name='Monthly Analysis', encoding = 'utf_8_sig')
df_ord.to_excel(writer, sheet_name='Orders Count', encoding = 'utf_8_sig')
df_lines.to_excel(writer, sheet_name='Lines Count', encoding = 'utf_8_sig')
df_sku.to_excel(writer, sheet_name='SKU Count', encoding = 'utf_8_sig')
writer.save()

# Ajouter une modification
print("Branch 1")