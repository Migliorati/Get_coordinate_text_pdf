import os
import pkg_resources.py2_warn

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


#pour dialogue
from tkinter import filedialog
from tkinter import *
import tkinter as tk
from tkinter.ttk import *
from tkinter import LabelFrame, Label, Tk
from tkinter.ttk import Notebook
import tkinter.messagebox

#pour export excel
import xlrd
from xlwt import Workbook, Formula

#pour date et time
import time
import subprocess

#definition de la date
t = time.localtime()
current_time = time.strftime("%H_%M_%S", t)
#emplacement et definition du fichier d'export
path = r"C:\Export_Syno_Fieldwire\export_" + current_time + ".xls"
# On créer un "classeur"
classeur = Workbook()
# On ajoute une feuille au classeur
feuille = classeur.add_sheet("Fieldwire_syno")

##################graphique
#la fenetre
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel
root = Tk()
class Fenetre(QWidget):
    def __init__(self):
        QWidget.__init__(self)

        # creation du bouton

        self.label = QLabel("")
        self.label1 = QLabel("Export terminé")
        self.label2 = QLabel("Par MIGLIORATI Bastien")
        self.label3 = QLabel("bastien.migliorati@gmail.com")
        self.label4 = QLabel("https://github.com/Migliorati/Get_coordinate_text_pdf")
        
        
        # creation du gestionnaire de mise en forme
        layout1 = QVBoxLayout()
        layout1.addWidget(self.label)
        layout1.addWidget(self.label1)
        layout1.addWidget(self.label2)
        layout1.addWidget(self.label3)
        layout1.addWidget(self.label4)
        self.setLayout(layout1)
        
        self.setWindowTitle("PDF vers Fieldwire")


    def parse_obj(lt_objs, x1, y1, filename):

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
                                if 'BPESYA' in str(t_clean) or 'NRO' in str(t_clean) or 'BRP' in str(t_clean) or 'BRD' in str(t_clean) or 'BRP' in str(t_clean) or 'BPI' in str(t_clean) or 'BRI' in str(t_clean) or 'BRA' in str(t_clean) or 'BRF' in str(t_clean):
                                    compteur = compteur + 1

                                    #print la boite detectée avec sa position x y
                                    #print(current_time + ' : [syno-task-'+ str(compteur) + ']' + "%6d, %6d, %s" % (obj.bbox[0], obj.bbox[1], str(t_clean) + ' | nbcar : ' + str(len(t_clean))))
                                    
                                    x_percent = str(((float(obj.bbox[0])*100)/float(x1)))
                                    y_percent = str(100-(float(obj.bbox[1])*100)/float(y1))


                                    #on nourris l'export excel
                                    feuille.write(compteur, 0, "[syno_b] " + str(t_clean))
                                    feuille.write(compteur, 1, "1")
                                    feuille.write(compteur, 2, "Tache_syno")
                                    feuille.write(compteur, 3, str(os.getlogin()) + "@sogetrel.fr")
                                    feuille.write(compteur, 6, str(filename))
                                    feuille.write(compteur, 7, str(x_percent))
                                    feuille.write(compteur, 8, str(y_percent))

                if 'POSE' in str(t_clean) or 'CABSYA' in str(t_clean):
                    compteur = compteur + 1

                    #print la boite detectée avec sa position x y
                    #print(current_time + ' : [syno-task-'+ str(compteur) + ']' + "%6d, %6d, %s" % (obj.bbox[0], obj.bbox[1], str(t_clean) + ' | nbcar : ' + str(len(t_clean))))
                    
                    x_percent = str((((float(obj.bbox[0]) + (float(obj.bbox[2])))/2)*100)/float(x1))
                    y_percent = str(100 - (float(obj.bbox[1])*100)/float(y1))

                    #on nourris l'export excel
                    feuille.write(compteur, 0, "[syno_c] " + str(t_clean))
                    feuille.write(compteur, 1, "1")
                    feuille.write(compteur, 2, "Tache_syno")
                    feuille.write(compteur, 3, str(os.getlogin()) + "@sogetrel.fr")
                    feuille.write(compteur, 6, str(filename))
                    feuille.write(compteur, 7, str(x_percent))
                    feuille.write(compteur, 8, str(y_percent))


            # if it's a container, recurse
            elif isinstance(obj, pdfminer.layout.LTFigure):
                parse_obj(obj._objs)
                print('Container')

        classeur.save(path)
        print('Export enregistre sous : ')
        print(path)
        subprocess.run(['explorer', os.path.realpath(path)])





    
    #user choose path
    outputpath = filedialog.askopenfile()
    if not os.path.exists("C:\Export_Syno_Fieldwire"):
        os.makedirs("C:\Export_Syno_Fieldwire")

    #oon verifie si l'utilisateur a bien chosis un pdf
    if not str(outputpath.name) == '' :
        root.withdraw()
        from os.path import basename
        filename = basename(outputpath.name).replace('.pdf', '')


    # Open a PDF file.
    fp = open(str(outputpath.name), 'rb')

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


    # loop over all pages in the document
    for page in PDFPage.create_pages(document):

        print('Lecture et mise en cache de tous les elements...')
        print('Veuillez patienter quelques minutes')
        # read the page into a layout object
        interpreter.process_page(page)
        layout = device.get_result()

        print(str(page))
        print(str(layout))
        #recuperer le format de la page

        inverse = 0
        x1 = page.mediabox[2]
        y1 = page.mediabox[3]

        if y1 > x1:
            x1 = page.mediabox[3]
            y1 = page.mediabox[2]
            inverse = 1


        print('x :' + str(x1))
        print('y :' + str(y1))
        
        # extract text from this object
        parse_obj(layout._objs, x1, y1, filename)


app = QApplication.instance() 
if not app:
    app = QApplication(sys.argv)
    
fen = Fenetre()
fen.show()
