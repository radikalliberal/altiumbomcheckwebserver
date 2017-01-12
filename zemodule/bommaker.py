import re
import pandas as pd
import numpy as np
import openpyxl
import zemodule.util as util

class bommaker(object):
    
    def __init__(self, bom):
        self.bom = bom

    def to_csv(self, path, quotechar='"', encoding='utf-8', sep=','):
        self.bom.to_csv(path, quotechar=quotechar, encoding=encoding , sep=sep)
    
    def export(self, basepath, info):
        
        wb = openpyxl.load_workbook(basepath + '/Einzelteilstueckliste Vorlage.xltx')
        esl = wb.get_sheet_by_name('Einzelteilstückliste')
        wb.template = False
        entrys = info.split('\n')

        esl['G1'].value = entrys[0]
        esl['G2'].value = entrys[1]
        esl['G3'].value = entrys[2]
        esl['G4'].value = entrys[3]
        esl['G5'].value = entrys[4]

        vorlage = ['A','B','C','D','G','H','J','K']
        bom_ex = ['Designator','Z-Nummer','Layer','Description','Hersteller','Hersteller-Nummer','Lieferant','Lieferanten-Nummer']

        for col in range(8):
            idx = 7
            for pos in self.bom[bom_ex[col]]:
                esl[vorlage[col] + str(idx)].value = pos
                idx += 1

        wb.save(basepath + '/bom.xlsx')

    def checkBom(self):
        # Die letzten Spalten können noch den Hersteller und die herstellernummer enthalten !!!!
        
        einzel = self.bom.copy()
        bad_parts = []
        
        man = einzel.columns.get_loc('Manufacturer')
        mannr = einzel.columns.get_loc('ManufacturerNr')
        com = einzel.columns.get_loc('Comment')
        
        '''
        check ob identische Bauteile vorhanden sind die bereits Einträge für Lieferant/nr oder Hersteller/nr haben
        diese werden dann ergänzt für Bauteile ohne Lieferant/nr und Hersteller/nr
        '''
        for idx_mM, line_mM in enumerate(einzel.values):

            if pd.isnull(line_mM[2]):
                continue

            if pd.isnull(line_mM[3:7]).any() & pd.isnull(line_mM[man:mannr]).all():
                for idx_bom, line_bom in enumerate(einzel.values):
                    if idx_mM == idx_bom: break
                    if line_mM[0] == line_bom[0]: 
                        raise Exception('FEHLER: Es wurde mehrfach der gleiche Bauteilbezeichner genutzt ('+ str(line_mM[1])+')<br>')

                    if  line_mM[2] == line_bom[2] and line_mM[3:].any() != line_bom[3:].any():
                            einzel.values[idx_mM][2:] = einzel.values[idx_bom][2:]
                            break

        self.bad_ids = []
        for idx_mM, line_mM in enumerate(einzel.values):
            if pd.isnull(line_mM[3:7]).all():
                self.bad_ids.append(idx_mM)
        if len(self.bad_ids) > 0:    
            #print(bad_ids)
            bad_parts = (einzel.values[self.bad_ids])[:,0:2:2].flatten()
            print('Achtung: folgende Teile haben keine Lieferanten/nummer und/oder Hersteller/nummer: <br><br>')
            print(', '.join(map(str, bad_parts)))
            print('<br>')
            
        return bommaker(einzel), len(bad_parts)

    def znum(self):

        einzel = self.bom.copy()

        df = util.getData()

        einzel['Z-Nummer'] = ''
        einzel['Lieferant'] = ''
        einzel['Hersteller'] = ''
        einzel['Hersteller-Nummer'] = ''
        einzel['Lieferanten-Nummer'] = ''

        einzel = einzel.apply(util.abgleich,args=(df,),axis=1)    
        return bommaker(einzel)
            
    