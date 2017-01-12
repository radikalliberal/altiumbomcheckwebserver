import pymssql
import pandas as pd
import numpy as np

def getData():
    conn = pymssql.connect(host='zitleitsql3', database='orderbase_ze')
    cur = conn.cursor()
    
    cur.execute('''SELECT        a.nr, hersteller_name, hersteller_artikelnr, lieferant_art_nummer, ad.name
    FROM            dbo.artikel AS a FULL OUTER JOIN
                     dbo.einkauf_preise AS ep ON ep.nr = a.lfdnr FULL OUTER JOIN
                     dbo.lieferanten AS l ON ep.lieferantennr = l.lfdnr FULL OUTER JOIN
                     dbo.adressen AS ad ON l.adressennr = ad.lfdnr
    WHERE        (a.archiv = 0) AND (ep.art = 1) AND (a.nr LIKE 'z%');''')

    np_a = np.asarray(cur.fetchall())
    spalten = ['Z-Nummer','Hersteller', 'Hersteller-Nummer','Lieferanten-Nummer', 'Lieferant']
    conn.close()
    return pd.DataFrame(np_a,columns=spalten)

def abgleich(row, df):
    
    suppliers = ['Farnell','Digi-Key','Mouser']
    cols = ['Z-Nummer','Lieferant','Hersteller','Hersteller-Nummer','Lieferanten-Nummer']
    
    rex = '(Supplier(\sPart\sNumber)?\s\d{1,3})|(Manufacturer(\sPart\sNumber)?\s\d{1,3})'
    
    solutions = row.filter(regex=rex).copy()
    
    ''' eine Reihe mit Manufacturer X , Manufacturer X+1 in eine Tabelle 
        mit mehreren Reihen umwandeln'''
    
    solutions = (pd.DataFrame(np.reshape(solutions.values,(len(solutions.index)/4,4))
                             ,columns=cols[1:])
                    .dropna())
    
    solutions['Z-Nummer'] = ''
    
    sol2 = pd.DataFrame(columns=solutions.columns)
    
    ''' Alle Solution herraussuchen die eine Z-Nummer hinterlegt haben'''
    
    for i in range(len(solutions.index)):   
        LiefNum = str(solutions.iloc[i]['Lieferanten-Nummer'])
        sol2 = (sol2.append(
                    df.where(df['Lieferanten-Nummer'].str.contains(LiefNum))
                      .dropna()))
    
    ''' Wenn keine Solutions vorhanden sind die in Orderbase vorhanden sind 
        dann suche nach möglicherweise vorhandenen Herstellernummern oder 
        Herstellern'''
    
    if len(sol2.index) < 1 : 
        print('keine Z-Nummer für ' + row['Designator']+ ' vorhanden <br>')
        if pd.notnull(row['Manufacturer']):
            row['Hersteller'] = row['Manufacturer']
        else:
            if pd.notnull(row['Manufacturer 1']):
                row['Hersteller'] = row['Manufacturer 1']
        if pd.notnull(row['ManufacturerNr']):
            row['Hersteller-Nummer'] = row['ManufacturerNr']
        else:
            if pd.notnull(row['Comment']):
                row['Hersteller-Nummer'] = row['Comment']
            else:
                if pd.notnull(row['Manufacturer Part Number 1']):
                    row['Hersteller-Nummer'] = row['Manufacturer Part Number 1']
            
        return row
    #print(sol2.head())
    
    sol2.drop_duplicates()
    for sup in suppliers:
        bed = sol2.Lieferant.str.contains(sup)
        arr = []
        if bed.any():
            length = len(sol2[bed].index)
            if len(sol2[bed].index) > 1:
                for x in range(length):
                    # Wenn es sich um eine größere Verpackungseinheit handelt, nimm diese
                    bed2 = sol2[bed].iloc[x].str.contains('((P)|(RL)|(2-ND)|(REEL))$',regex=True)
                    if bed2.any():
                        for val in cols:
                           row[val] = str(sol2[bed2].get(val).values[0])
                        return row
                    else:
                        for val in cols:
                           row[val] = str(sol2[bed].take([0]).get(val).values[0])
                        return row
            
            for val in cols:
               row[val] = str(sol2[bed].get(val).values[0])
            return row
            break
    print('keine standard Liefderanten für ' + row['Designator'] + ' vorhanden')
    for val in cols:
       row[val] = str(sol2.take([0]).get(val).values[0])
    return row