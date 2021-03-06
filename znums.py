import cgi, logging, os, sys
import pandas as pd
from zemodule import bommaker, util
import smtplib, re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def main():

    emailRegex = r'(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)'

    email = load_bom_info().split('\n')[5]
    name = load_bom_info().split('\n')[3]
    sendMe = re.match(emailRegex,email)


    stueckliste = pd.read_csv(UPLOAD_DIR + '/bom.csv', quotechar='"', encoding ='utf-8', dtype='str', sep=',')
    ei = bommaker(stueckliste)
    logger.info('searching Orderbase-Numbers')
    try:
        ei = ei.znum()
        logger.info('creating Excelfile')
        ei.export(UPLOAD_DIR, load_bom_info())
        logger.info('done creating Excelfile')

        btOhneZ = len(ei.bom[ei.bom['Z-Nummer']==''].index)

        if sendMe:
            sendEmail(email,name,btOhneZ)
            logger.info('email sent to ' + email)
        #if sendZe:
        #    sendEmail('jan.scholz@desy.de',name,btOhneZ)
        #    logger.info('email send to jan.scholz@desy.de' )
        
    except Exception as e:
        logging.error(e)
    
def load_bom_info():
    fout = open(os.path.join(UPLOAD_DIR, 'BomMeta.txt'), 'r')
    return fout.read()
 
def sendEmail(toaddr,name,btOhneZ):
    logger.info('Versuche Email an ' + toaddr + ' zu senden')
    fromaddr = "desystuecklistenupload@gmail.com"

    btoz = ''
    herstN = ''
    sendReq = ''
     
    msg = MIMEMultipart()
     
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = "Neue Einzelstückliste für DESY-ZE"

    if  btOhneZ > 0:
        btoz = 'Für ' + str(btOhneZ) + ' Bauteile gibt es noch keine Z-Nummer'
    else:
        btoz = 'Alle Bauteile haben Z-Nummern hinterlegt.'

    try:
        if sys.argv[1] != '0':
            herstN = ('Es sind {0} Bauteile ohne Herstellernummer/Hersteller vorhanden.\n\n' 
                      '{1}\n\nBitte ergänzen Sie alle Hersteller und Herstellernummern '
                      'bevor Sie uns die Stückliste zusenden.').format(sys.argv[1],btoz) 
        else:
            herstN = ('Alle Bauteile haben Hersteller und Herstellernummern hinterlegt\n\n{0}' 
                      '\n\nSenden Sie uns die Stückliste, falls Sie vollständig ist gerne zusammen '
                      'mit den Fertigungsdaten an desy-ze@desy.de.').format(btoz)

    except Exception as e:
        logger.warn('Fehler: '.format(e))
        herst = ('Über den Zustand der Einzelstückliste lässt sich keine Aussage machen, bitte schauen ' 
                 'Sie selbst nach ob alle Herstellernummern und vorzugsweise auch Lieferanten und Liefer'
                 'antennummern vorhanden sind.')

            


    body =  load_bom_info() + '\n\n' + herstN  + '\n\n\nMfg\n\n DESY-ZE'
            

    
     
    msg.attach(MIMEText(body, 'plain'))
     
    filename = "bom.xlsx"
    attachment = open(BASE_PATH + "/tmp/bom.xlsx", "rb")
     
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
     
    msg.attach(part)
     
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, getpw())
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()

def getpw():
    f = open(BASE_PATH + '/cgi-bin/password','r')
    return f.read()


if __name__ == '__main__':
    BASE_PATH = "c:/Apache24/"
    UPLOAD_DIR = "c:/Apache24/tmp"

    logger = logging.getLogger(__name__)

    logger.setLevel(logging.INFO)

    # create a file handler
    handler = logging.FileHandler(BASE_PATH + '/logs/bomcheck.log')
    handler.setLevel(logging.INFO)

    # create a logging format
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)

    # add the handlers to the logger
    logger.addHandler(handler)

    main()