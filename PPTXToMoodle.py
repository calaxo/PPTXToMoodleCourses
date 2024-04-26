# Crée par A CALENDREAU, le 07/02/2023 en Python 3

#déstiné a prendre des fichiers powerpoint,pouvoir les transformer en pdf,envoyer ses fichier sur un drive via une api
#et renvoyer des lien avec les bonnes balise et argument html pour simplifier l'ajout de cours a un moodle


import tkinter as tk

import sys
import os
import comtypes
import comtypes.client
from pyfiglet import Figlet
from ppt2pdf.utils import generateOutputFilename
from comtypes.client import CreateObject, Constants
import shutil


from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

gauth = GoogleAuth()
gauth.LocalWebserverAuth() 
drive = GoogleDrive(gauth)


class fenetre(tk.Tk):

    
    def __init__(self):

        tk.Tk.__init__(self)
        self.title('PPTXToMoodle')   
        self.geometry("410x350")
        self.resizable(width = True, height = True)  
        self.creer_widgets()    
        self.changement_apparence() 
        self.mainloop() 
        self.parent_id=""
        self.parentID=""

    def pdf(self):

        fichier = os.listdir('input/')
        #lsite des noms des fichier du dossier input

        for file in fichier:#pour tout les fichier
            inputFilePath=os.path.dirname(__file__)+'\\input\\'+file
            outputFilePathpdf=os.path.dirname(__file__)+'\\output\\'+file.replace(".pptx",".pdf")
            outputFilePathpptx=os.path.dirname(__file__)+'\\output\\'+file
            #ecriture de chemin absolue pour comtypes
            
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            #%% Set visibility to minimize
            powerpoint.Visible = 1
            #%% Open the powerpoint slides
            slides = powerpoint.Presentations.Open(i/nputFilePath)
            #%% Save as PDF (formatType = 32)
            slides.SaveAs(outputFilePathpdf, 32)
            #%% Close the slide deck
            slides.Close()
            shutil.copy(inputFilePath,outputFilePathpptx )
            #pour copier les pptx du dossier input dans le dossier output ou sont present les pdf juste créé

    def name(self):
        #future fonction pour un renomage en groupe des fichier
        print("renomage")

    def envoyer(self):
        parent = ""
        # dossier parent pour travailler dans google drive


        file_metadata = {#pour les infos du parent
          'title': self.frameWeb.parent.get(),
          'parents': [{'id': self.frameWeb.parentID.get()}],
          'mimeType': 'application/vnd.google-apps.folder'
        }
        

        folder = drive.CreateFile(file_metadata)
        
        folder.Upload()
        self.parent_id=folder['id']

        self.parent_id = folder['id']
        entries = os.listdir('output')
        upload_file_list = entries
        for upload_file in upload_file_list:
                gfile = drive.CreateFile({'parents': [{'id': self.parent_id}],'title': upload_file})
                # Read file and set it as the content of this instance.
                gfile.SetContentFile("output/"+upload_file)
                gfile['copyRequiresWriterPermission'] = True
                gfile.Upload() # Upload the file.
                gfile.InsertPermission({
                            'type': 'anyone',
                            'value': 'anyone',
                            'role': 'reader'})

    def lien(self):
        self.lien_parent_id_temp=self.frameWeb.id.get()
        if self.lien_parent_id_temp=="":
            self.lien_parent_id_temp=self.parent_id

        #si un id précis est indiqué fournis les lien de cet id de dossier sinon reprend le dosier utilisé pour envoyer des fichier
        debut_pdf="<br>\n<div style=\"text-align: center;\"><iframe src=\"https://drive.google.com/file/d/";

        fin_pdf="/preview\" allow=\"autoplay\" width=\"100%\" height=\"500\"></iframe></div>";

        debut_pptx="<br><div style=\"text-align: center;\"><iframe src=\"https://drive.google.com/file/d/"

        fin_pptx = "/preview\" allow=\"autoplay\" width=\"100%\" height=\"500\"></iframe></div>"


        file_list = drive.ListFile({'q': "'{}' in parents and trashed=false".format(self.lien_parent_id_temp)}).GetList();
        id_list = [];

        nom_list_pdf=[];
        nom_list_pptx=[];
        
        lien_list_pdf=[];
        lien_list_pptx=[];
        numero_pdf=1
        numero_pptx=1
        for file in file_list:
        
        #architecture du html différente pour les pdf et les pptx
            
                if (file['title'][-1]=="f"):    #f du .pdf
                    nom_list_pdf.append(str(numero_pdf)+" "+file['title'][: -4]+" PDF");
                    lien_list_pdf.append(debut_pdf+file['id']+fin_pdf);
                    numero_pdf+=1
                    #pour leur donner un nom en prefixe
                    
                if (file['title'][-1]=="x"):    #x du .pptx
                    nom_list_pptx.append(str(numero_pptx)+" "+file['title'][: -5]+" PPTX");
                    lien_list_pptx.append(debut_pptx+file['id']+fin_pptx);
                    
                    numero_pptx+=1

        fichier = open("cours.txt", "w")    #ouverture du fichier

        indice=0
        fichier.write("PDF:\n\n")
        for x in lien_list_pdf: 
            fichier.write(nom_list_pdf[indice])     #pour recuperer la bone valeur de la liste de nom
            fichier.write("\n")
            fichier.write(x)
            fichier.write("\n\n")
            indice+=1
        indice=0
        fichier.write("PPTX:\n\n")
        for x in lien_list_pptx:
            fichier.write(nom_list_pptx[indice])    #pour recuperer la bone valeur de la liste de nom
            fichier.write("\n")
            fichier.write(x)
            fichier.write("\n\n")
            indice+=1

        print("nom et lien ecrit dans le fichier cours.txt")    #log dans la console python
        fichier.close()                     #fermeture du fichier
        

        
    def creer_widgets(self):    #debut partie graphique creation des objet

        self.labelFramelocal = tk.LabelFrame(self, text = 'local')

        self.labelFramelocal.label2      = tk.Label(self.labelFramelocal, text = 'PPTX', font='Helvetica 14 bold')      #titre de la frame

        self.labelFramelocal.name = tk.Button(self.labelFramelocal, command=self.name, text = 'renomage')
        self.labelFramelocal.pdf = tk.Button(self.labelFramelocal, command=self.pdf, text = 'pdf')


        self.frameWeb = tk.LabelFrame(self, text = 'web')
        
        self.frameWeb.label1      = tk.Label(self.frameWeb, text = 'envoyer au drive', font='Helvetica 14 bold')    #titre de la frame


        self.frameWeb.label1infoa      = tk.Label(self.frameWeb, text = 'rentrer l\'ID du parent ou sera creer le dossier', font='Helvetica 8 ')
        self.frameWeb.parentID = tk.Entry(self.frameWeb) 

        
        self.frameWeb.label1infob      = tk.Label(self.frameWeb, text = 'rentrer le nom du dossier a creer', font='Helvetica 8 ')
        self.frameWeb.parent = tk.Entry(self.frameWeb)      #pour entrer la nom du dossier a créer
        self.frameWeb.buttonEnvoyer = tk.Button(self.frameWeb, command=self.envoyer, text = 'envoyer')

        self.frameWeb.label2      = tk.Label(self.frameWeb, text = 'lien', font='Helvetica 14 bold')
        self.frameWeb.label2info      = tk.Label(self.frameWeb, text = 'rentrer l\'ID du parent si diferent par rapport a l\'upload', font='Helvetica 8 ')        
     
        self.frameWeb.id = tk.Entry(self.frameWeb)              #pour entrer l'id du dossier pour lequel on veut les lien de fichier
        self.frameWeb.buttonname = tk.Button(self.frameWeb, command=self.lien, text = 'titre et html')


    def changement_apparence(self): #fin partie graphique placement et parametrage des objets
        colorWhite = "#F2F3F3"      #blanc doux


        self.labelFramelocal["relief"] = "groove"  #pour le contour de la frame
        self.labelFramelocal.pack()     #la premiere frame est juste placcer avec pack()
        #a l'interieur une grid de 2 de large et de 4 de hauteur

        self.labelFramelocal.label2.grid(row = 3, column = 0,columnspan=2, padx = 5, pady = 5) 
        self.labelFramelocal.name.grid(row = 4, column = 0, padx = 5, pady = 5)
        self.labelFramelocal.pdf.grid(row = 4, column = 1, padx = 5, pady = 5)
        
        self.frameWeb["relief"] = "groove"  #pour le contour de la frame


        self.frameWeb.pack()            #comme la premiere, placé avec pack()
        #a l'interieur une grid de 2 de large et de 4 de hauteur
        
        self.frameWeb.label1.grid(row = 0, column = 0,columnspan=2, padx = 5, pady = 5)
        self.frameWeb.label1infoa.grid(row = 1, column = 0,columnspan=2, padx = 5, pady = 5)
        self.frameWeb.parentID.grid(row = 2, column = 0,columnspan=2, padx = 5, pady = 5)
        self.frameWeb.label1infob.grid(row = 3, column = 0,columnspan=2, padx = 5, pady = 5)
        self.frameWeb.parent.grid(row = 4, column = 0, padx = 5, pady = 5)
        self.frameWeb.buttonEnvoyer.grid(row = 4, column = 1, padx = 5, pady = 5) #grid pour les boutons solitaires
        
        self.frameWeb.label2.grid(row = 5, column = 0,columnspan=2, padx = 5, pady = 5)
        self.frameWeb.label2info.grid(row = 6, column = 0,columnspan=2, padx = 5, pady = 5)
        self.frameWeb.id.grid(row = 7, column = 0, padx = 5, pady = 5)
        self.frameWeb.buttonname.grid(row = 7, column = 1, padx = 5, pady = 5)



applicationPrincipale=fenetre()#lancement de la partie graphique
