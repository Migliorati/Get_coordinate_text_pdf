#pour parse pdf
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
import pdfminer

#pour export excel
import xlrd
from xlwt import Workbook, Formula

#pour date et time
import time
t = time.localtime()
current_time = time.strftime("%H_%M_%S", t)


#emplacement et definition du fichier d'export
path = r"C:\Users\bastien.migliorati\Documents\@_PYTHON\Pdf_parser\export_" + current_time + ".xls"
# On créer un "classeur"
classeur = Workbook()
# On ajoute une feuille au classeur
feuille = classeur.add_sheet("Fieldwire_syno")

# Open a PDF file.
fp = open('/Users/bastien.migliorati/Documents/@_PYTHON/Pdf_parser/test.pdf', 'rb')

# Create a PDF parser object associated with the file object.
parser = PDFParser(fp)

# Create a PDF document object that stores the document structure.
# Password for initialization as 2nd parameter
document = PDFDocument(parser)
print(str(document))
# Check if the document allows text extraction. If not, abort.
if not document.is_extractable:
    print('Non autorise')
    raise PDFTextExtractionNotAllowed
print('Extraction possible')

# Create a PDF resource manager object that stores shared resources.
rsrcmgr = PDFResourceManager()
print('Creation d un fichier de resources')

# Create a PDF device object.
device = PDFDevice(rsrcmgr)
print('Creation d un fichier de device : ' + str(device))

# BEGIN LAYOUT ANALYSIS
# Set parameters for analysis.
laparams = LAParams()
print('Definition des parametres pour analyse')

# Create a PDF page aggregator object.
device = PDFPageAggregator(rsrcmgr, laparams=laparams)

# Create a PDF interpreter object.
interpreter = PDFPageInterpreter(rsrcmgr, device)

def parse_obj(lt_objs, x1, y1):

    list_exp = []
    compteur = 0
    # loop over the object list
    print('Start loop')
    for obj in lt_objs:

        # if it's a textbox, print text and location
        if isinstance(obj, pdfminer.layout.LTTextBoxHorizontal):

            
            #clean du texte
            t_clean2 = obj.get_text().replace('\n', '|')
            t_clean1 = str(t_clean2).replace('  ', '')
            t_clean = str(t_clean1).replace(' ', '|')
            
            #essaye de recuperer uniquement les boites
            if not 'FO' in str(t_clean):
                if not '/' in str(t_clean):
                    if not 'POSE' in str(t_clean):
                        if not 'LOVE' in str(t_clean):                        
                            if 'NRO' in str(t_clean) or 'BRP' in str(t_clean) or 'BRD' in str(t_clean) or 'BRP' in str(t_clean) or 'BPI' in str(t_clean) or 'BRI' in str(t_clean) or 'BRA' in str(t_clean) or 'BRF' in str(t_clean):
                                compteur = compteur + 1

                                #print la boite detectée avec sa position x y
                                print(current_time + ' : [syno-task-'+ str(compteur) + ']' + "%6d, %6d, %s" % (obj.bbox[0], obj.bbox[1], str(t_clean) + ' | nbcar : ' + str(len(t_clean))))
                                x_percent = str((float(obj.bbox[0])*100)/float(x1))
                                y_percent = str(100 - (float(obj.bbox[1])*100)/float(y1))

                                #on nourris l'export excel
                                feuille.write(compteur, 0, "[syno] " + str(t_clean))
                                feuille.write(compteur, 1, "1")
                                feuille.write(compteur, 2, "Tache_syno")
                                feuille.write(compteur, 3, "bastien.migliorati@sogetrel.fr")
                                feuille.write(compteur, 6, "test")
                                feuille.write(compteur, 7, str(x_percent))
                                feuille.write(compteur, 8, str(y_percent))

        # if it's a container, recurse
        elif isinstance(obj, pdfminer.layout.LTFigure):
            parse_obj(obj._objs)
            print('Container')

    classeur.save(path)
    print('Export enregistre sous : ' + path)

# loop over all pages in the document
for page in PDFPage.create_pages(document):

    print('Loop sur les pages... veuillez patienter quelques minutes')
    # read the page into a layout object
    interpreter.process_page(page)
    layout = device.get_result()
    print('page : ' + str(page))
    print('layout : ' + str(layout))
    print(page.mediabox) # <- the media box that is the page size as list of 4 integers x0 y0 x1 y1

    #recuperer le format de la page
    x0 = page.mediabox[0]
    print('x0 : ' + str(x0))
    y0 = page.mediabox[1]
    print('y0 : ' + str(y0))
    x1 = page.mediabox[3]
    print('x1 : ' + str(x1))
    y1 = page.mediabox[2]
    print('y1 : ' + str(y1))

    
    #il faut aller chercher l'orientation
    print('orientation : ')
    

    # extract text from this object
    parse_obj(layout._objs, x1, y1)
