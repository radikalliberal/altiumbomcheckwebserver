#!python
import cgi
import cgitb; cgitb.enable()
import os, sys

try: # Windows needs stdio set for binary mode.
    import msvcrt
    msvcrt.setmode (0, os.O_BINARY) # stdin  = 0
    msvcrt.setmode (1, os.O_BINARY) # stdout = 1
except ImportError:
    pass

BASE_PATH = "c:/Apache24/"
UPLOAD_DIR = "c:/Apache24/tmp"

HTML_TEMPLATE = '''<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
   <head>
      <title>BOM Upload</title>
      <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   </head>
   <body>
    <font face="Helvetica">
    <center>
      <h1>Altium Einzelstücklisten Erstellung für DESY-ZE</h1>
      <hr>
      Laden Sie bitte die Stückliste, die mit <a href="../Job_DESY_ZE.OutJob" target="_blank" download>diesem</a> Outputjob (siehe Abbildung) erstellt wird, hoch.
      <br><br>
      <img src="../img/AltiumOutputjob.jpg" alt="Generierung der Stückliste" >
      <br><br>
      <form action="upload.py" method="POST" enctype="multipart/form-data">
          File name: <input name="file_1" type="file"><br><br>
          <input name="submit_file" type="submit" value="Datei senden" style="height:50px; width:148px">
      </form>
      <br><br>
      Nachdem die Datei hochgeladen wurde und mit der Datenbank von DESY-ZE abgeglichen wurde, <br>
      wird ihnen eine Email gesendet mit der Stückliste in der gewünschten Vorlage von DESY-ZE. <br>
      Sollten noch Herstellernummern fehlen tragen Sie diese bitte in Altium nach und exportieren <br>
      Sie erneut. So wird sichergestellt, dass bereits vorhandene Bauteile in der Datenbank von <br>
      DESY-ZE, gefunden werden können.<br><br>
      Es kann einige Minuten dauern bis Sie die Email erreicht.
    </center>
    </font>
   </body>
</html>'''

def get_item(form_field, form):
    if not form_field in form: 
        raise ValueError('Das angeforderte Feld ' + form_field + ' ist leider ohne Wert!!')
        return
    item = form[form_field]
    return item

def get_item_value(form_field, form):
    if not form_field in form: 
        return ''
    return form[form_field].value

def print_redirect(form):

    link = 'http://szepcx16631:8080/cgi-bin/bomcheck.py'

    REDIRECT = '''Content-Type: text/html
            Location: ''' + link + '''

          <html>
          <head>
            <meta http-equiv="refresh" content="0;url=''' + link + '''" />
            <title>You are going to be redirected</title>
          <body>
            <font face="Helvetica">
            <center>
              <h1>Altium Einzelst&uuml;cklisten Erstellung f&uuml;r DESY-ZE</h1>
              <hr>
                <br><br><br><br><br><br>
                Verarbeite Stückliste... 
            </center>
            </font>
          </body>
        </html>'''
    print(REDIRECT)


def print_html_form ():
    """This prints out the html form. Note that the action is set to
      the name of the script which makes this is a self-posting form.
      In other words, this cgi both displays a form and processes it.
    """
    print("content-type: text/html\n")
    print(HTML_TEMPLATE)

def save_bom_info(form):

    fout = open(os.path.join(UPLOAD_DIR, 'BomMeta.txt'), 'w')

    emailingOptions = ''

    if 'emailme' in form:
        emailingOptions = get_item_value('emailme',form) + '\n'
    if 'emailze' in form:
        emailingOptions += get_item_value('emailze',form)

    fout.write(
        get_item_value('desc_1',form) + '\n' + \
        get_item_value('desc_2',form) + '\n' + \
        get_item_value('desc_3',form) + '\n' + \
        get_item_value('desc_4',form) + '\n' + \
        get_item_value('desc_5',form) + '\n' + \
        get_item_value('email',form)
    )
        
    fout.close()

def load_bom_info():
    fout = open(os.path.join(UPLOAD_DIR, 'BomMeta.txt'), 'r')
    return fout.read()


def save_uploaded_file (form_field, upload_dir, form):
    """This saves a file uploaded by an HTML form.
       The form_field is the name of the file input field from the form.
       For example, the following form_field would be "file_1":
           <input name="file_1" type="file">
       The upload_dir is the directory where the file will be written.
       If no file was uploaded or if the field does not exist then
       this does nothing.
    """
    if not form_field in form: 
        save_bom_info(form)
        print_html_form()
        return
    fileitem = form[form_field]
    if not fileitem.file: 
        print_html_form()
        return
    fout = open(os.path.join(upload_dir, 'bom.csv'), 'wb')
    while 1:
        chunk = fileitem.file.read(100000)
        if not chunk: break
        fout.write (chunk)
    fout.close()
    print_redirect(form)



params = cgi.FieldStorage()
info = load_bom_info()
save_uploaded_file('file_1', UPLOAD_DIR, params)

    






