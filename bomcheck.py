#!python
import cgi, logging, os, sys, re
import cgitb; cgitb.enable()
import numpy as np
import pandas as pd
import traceback
from zemodule import bommaker, util

try: # Windows needs stdio set for binary mode.
    import msvcrt
    msvcrt.setmode (0, os.O_BINARY) # stdin  = 0
    msvcrt.setmode (1, os.O_BINARY) # stdout = 1
except ImportError:
    pass
BASE_PATH = "c:/Apache24/"
UPLOAD_DIR = "c:/Apache24/tmp"

HTMLHEAD = """<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><head><title>Bom Check</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head><body><font face="Helvetica"><center><h1>Altium Einzelstücklisten Erstellung für DESY-ZE</h1><hr><br><br><br><br><br>"""

def save_uploaded_file (fileitem, upload_dir):
    """This saves a file uploaded by an HTML form.
       The form_field is the name of the file input field from the form.
       For example, the following form_field would be "file_1":
           <input name="file_1" type="file">
       The upload_dir is the directory where the file will be written.
       If no file was uploaded or if the field does not exist then
       this does nothing.
    """

    fout = open(os.path.join(upload_dir, fileitem.filename), 'wb')
    while 1:
        chunk = fileitem.file.read(100000)
        if not chunk: break
        fout.write (chunk)
    fout.close()


def print_html_form():
    print("content-type: text/html\n")
    print(HTMLHEAD)

logger = logging.getLogger(__name__)

logger.setLevel(logging.DEBUG)

# create a file handler
handler = logging.FileHandler(BASE_PATH + '/logs/bomcheck.log')
handler.setLevel(logging.INFO)

# create a logging format
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

# add the handlers to the logger
logger.addHandler(handler)

def parent():
    print_html_form()

    logger.info('new Bom uploaded: <br>'+ load_bom_info())

    logger.info('reading bom.csv')
    try:
        stueckliste = pd.read_csv(UPLOAD_DIR + '/bom.csv', quotechar='"', encoding ='utf-8', dtype='str')  

    except Exception as e:
        logger.error('Datei ist nicht lesbar: {0}'.format(e))  
        raise Exception('Die Datei scheint nicht im geforderten format zu sein, oder es wurde keine Datei bereitgestellt.<br>')  
        return

    logger.info('done reading bom.csv')
    logger.info('initialising BOM')
    einzel = bommaker(stueckliste)
    logger.info('checking Entrys')
    einzel, btOhne = einzel.checkBom()

    emailRegex = r'(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)'

    email = load_bom_info().split('\n')[5]
    validEmail = re.match(emailRegex,email)
    if not validEmail:
        logger.warning('User has not given valid email address')
        raise Exception('Sie haben keine gültige Emailadresse angegeben, es kann keine Email an Sie gesandt werden.<br>')

    print("</body></html>")
    return einzel

def load_bom_info():
    fout = open(os.path.join(UPLOAD_DIR, 'BomMeta.txt'), 'r')
    return fout.read()

if __name__ == '__main__':

    btOhne = 0

    arguments = cgi.FieldStorage()
    try:
        einzel = parent()
        einzel.to_csv(UPLOAD_DIR + '/bom.csv')
    except Exception as e:
        print('Fehler: {0}'.format(e))
        logger.error('Bom couldnt be saved: {0}'.format(e))
    else:
        print('Die Stückliste wird nun mit der DESY-ZE-Datenbank verglichen, dies kann einige Minuten<br>'
          'in Anspruch nehmen. Wenn alle Bauteile eine Herstellernummern und einen Hersteller hinterlegt<br>'
          'haben, schicken sie bitte die Stückliste, im Anhang der Email, zusammen mit den Fertigungsdaten<br>'
          'an desy-ze@desy.de.<br><br>'
          'Sie können die Seite nun schließen, die Email wird ihnen schnellstmöglich zugesandt')
        logger.info('starting new script for Z-Number-Search')
        os.system('start python -W ignore c:/Apache24/cgi-bin/znums.py ' + str(btOhne))
        #print('<br><br><font color="#0B610B">Die Datei ist fertig und wird Ihnen nun per Mail geschickt.')
