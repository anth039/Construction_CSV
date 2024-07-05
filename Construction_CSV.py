# coding: utf-8
#Imports des librairies
import tkinter.ttk
import tkinter.messagebox
import tkinter.filedialog
from tkinter import *
import os
import time

#Variable globale
fichierOriginePremiereLigne = ""
fichierOrigineSecondeLigne = ""

#Permet de choisir un fichier d'origine
def onClickFichierOrigine():
    
    #Vidange de la liste d'origine et du champs possèdent le chemin
    listeChoixOrigine.delete(0,'end')
    monFichierOrigine.configure(state="normal")
    monFichierOrigine.delete(0,'end')
    monFichierOrigine.configure(state="disabled")
    
    #Ouverture de la fenêtre demandant de choisir un fichier
    file_path = tkinter.filedialog.askopenfilename(filetypes=[("Fichier avec délimiteurs","*.csv")])
    
    #Si un choix est fait
    if file_path : 
       monFichierOrigine.configure(state="normal")
       monFichierOrigine.insert(0,file_path)
       monFichierOrigine.configure(state="disabled")
       
       #Ouverture/Fermeture du fichier pour récupérer la première ligne
       fichierOrigine = open(file_path, "r")
       global fichierOriginePremiereLigne,fichierOrigineSecondeLigne
       fichierOriginePremiereLigne = fichierOrigine.readline()
       fichierOrigineSecondeLigne = fichierOrigine.readline()
       fichierOrigine.close()
       
       #Ecriture en brut de la première ligne
       listeChoixOrigine.insert(0, fichierOriginePremiereLigne)

       
#Permet de choisir un dossier de destination
def onClickDossierDestination():
    #Du champs possèdent le chemin
    monDossierDestination.configure(state="normal")
    monDossierDestination.delete(0,'end')
    monDossierDestination.configure(state="disabled")
    
    #Ouverture de la fenêtre demandant de choisir un dossier
    folder_path = tkinter.filedialog.askdirectory()
    
    #Si un choix est fait
    if folder_path : 
       monDossierDestination.configure(state="normal")
       monDossierDestination.insert(0,folder_path)
       monDossierDestination.configure(state="disabled")


#Permet d'ajouter des colonnes au fichier de destination
def doubleClickChoixOrigine(event):
    #Récupération de la position du curseur
    choix = listeChoixOrigine.curselection() 

    if listeChoixOrigine.get(choix) not in listeChoixDestination.get(0, END):
        #Ajout dans la colonne destination
        listeChoixDestination.insert(0, listeChoixOrigine.get(choix))
        
        #Enlève dans la colonne d'origine
        listeChoixOrigine.delete(choix)
        
        #Blocage de la position du délimiteurs
        monComboboxDelimiteurs.configure(state="disabled")
    
#Permet d'enlever des colonnes au fichier de destination
def doubleClickChoixDestination(event):
    #Récupération de la position du curseur
    choix = listeChoixDestination.curselection() 

    if listeChoixDestination.get(choix) not in listeChoixOrigine.get(0, END):
        #Ajout dans la colonne destination
        listeChoixOrigine.insert(0, listeChoixDestination.get(choix))
        
        #Enlève dans la colonne d'origine
        listeChoixDestination.delete(choix)
        
        if listeChoixDestination.size() == 0:
            #Déblocage de la position du délimiteurs
            monComboboxDelimiteurs.configure(state="normal")
    
#Permet d'exporter les données
def onClickExport():
    #Vérifications
    if monFichierOrigine.get():
        if monDossierDestination.get():
            if listeChoixDestination.size() > 1 :
                
                #Création du chemin de destination
                monFichierDestination = monDossierDestination.get() + "/Export_" + time.strftime("%d_%m_%Y-%H_%M") + ".csv"
                
                if os.path.isfile(monFichierDestination):
                    #Suppresion de l'ancien fichier exporté
                    os.remove(monFichierDestination)
                
                #Récupération des nombres des colonnes choisis
                mesPositions = []
                for element in listeChoixDestination.get(0, END):
                    i = 0
                    for elementPremiereLigne in fichierOriginePremiereLigne.split(monComboboxDelimiteurs.get()):
                        if element.split(" => ")[0] == elementPremiereLigne:
                            mesPositions.append(i)
                        i+=1
                
                #Ouverture des fichiers
                with open(monFichierOrigine.get(), "r") as fichierOrigine:
                    with open(monFichierDestination, "x") as ficherDestination:
                        for ligne in fichierOrigine:
                            ligneSplit = ligne.split(monComboboxDelimiteurs.get())
                            ligneFinal = []
                            j = len(mesPositions)-1
                            #Tant que tous les champs ne sont pas parcouru
                            while (len(ligneFinal) != listeChoixDestination.size()):
                               i = 0
                               #Parcours de chaque elements d'une ligne
                               for elementLigne in ligneSplit:
                                   if i == mesPositions[j]:
                                       ligneFinal.append(elementLigne.rstrip())
                                       j-=1
                                       break
                                   i+=1    
                            ficherDestination.write(monComboboxDelimiteurs.get().join(ligneFinal)+'\n')
                
                tkinter.messagebox.showinfo("Résultat", message="Le fichier généré est à cette emplacement : " + monFichierDestination )

            else:
                tkinter.messagebox.showerror("Colonne à exporter", message="Merci de choisir des colonnes à exporter")
        else:
            tkinter.messagebox.showerror("Choix du dossier destination", message="Merci de choisir un dossier destination")
    else:
        tkinter.messagebox.showerror("Choix du fichier d'origine", message="Merci de choisir un fichier .CSV")
        
def choixDelimiteur(event):
    
    #Vidange de la liste d'origine
    listeChoixOrigine.delete(0,'end')
    
    #Récupération de la première ligne extraite
    global fichierOriginePremiereLigne,fichierOrigineSecondeLigne
    
    #Récupération du choix
    choix = monComboboxDelimiteurs.get()
    
    #Filtrage
    if choix == "Choix délimiteurs" :
       listeChoixOrigine.insert(0, fichierOriginePremiereLigne)
    else:
       #Si le délimiteurs est contenu dans la première chaine
       if choix in fichierOriginePremiereLigne :
           
          #Traitement seconde ligne
          elementsSecondeLigne = fichierOrigineSecondeLigne.split(choix)
       
          i = 0
          for element in fichierOriginePremiereLigne.split(choix):
             listeChoixOrigine.insert(i, element + " => "+ elementsSecondeLigne[i] )
             i+=1
       else:
          listeChoixOrigine.insert(0, fichierOriginePremiereLigne)

#**********Programme principal
if __name__ == '__main__':

    # Création d'une fenêtre
    fenetre = Tk()

    # Titre de ma fenêtre
    fenetre.title("Construction CSV")

    # Définition d'une taille
    fenetre.geometry("650x350")
    
    #Redimensionnement impossible
    fenetre.resizable(False, False)

    # Création du widget qui va contenir les autres
    monConteneurPrincipal = Frame(fenetre)
    # Ajout dans le conteneur
    monConteneurPrincipal.pack()
    # Gestion du fichier d'origine
    monLabelFichierOrigine = Label(monConteneurPrincipal, text="Fichier d'origine")
    monLabelFichierOrigine.grid(column=0, row=0 )
    monFichierOrigine = Entry(monConteneurPrincipal,state=DISABLED,width=75)
    monFichierOrigine.grid(column=1, row=0 )
    monBoutonParcourirFichierOrigine = Button(monConteneurPrincipal, text="Parcourir", command=onClickFichierOrigine)
    monBoutonParcourirFichierOrigine.grid(column=2, row=0)
    # Gestion du fichier de destination
    monLabelDossierDestination = Label(monConteneurPrincipal, text="Dossier destination")
    monLabelDossierDestination.grid(column=0, row=1 )
    monDossierDestination = Entry(monConteneurPrincipal,state=DISABLED,width=75)
    monDossierDestination.grid(column=1, row=1 )
    monBoutonParcourirDossierDestination = Button(monConteneurPrincipal, text="Parcourir", command=onClickDossierDestination)
    monBoutonParcourirDossierDestination.grid(column=2, row=1,pady=10)
    # Gestion des délimiteurs
    monComboboxDelimiteurs = tkinter.ttk.Combobox(monConteneurPrincipal, values=["Choix délimiteurs",",", ";","|"," "], state="readonly")
    monComboboxDelimiteurs.current(0)
    monComboboxDelimiteurs.bind('<<ComboboxSelected>>', choixDelimiteur)
    monComboboxDelimiteurs.grid(column=0, row=2,columnspan=3,pady=10)
    #Choix colonne
    monConteneurSecondaire = Frame(monConteneurPrincipal)
    
    monLabelListeChoixOrigine = Label(monConteneurSecondaire, text="Colonne d'origine")
    monLabelListeChoixOrigine.grid(column=0, row=0 )
    espace = Label(monConteneurSecondaire)
    espace.grid(column=1,row=0,ipadx=25)
    monLabelListeChoixDestination = Label(monConteneurSecondaire, text="Colonne destination")
    monLabelListeChoixDestination.grid(column=2, row=0 )
    listeChoixOrigine = Listbox(monConteneurSecondaire,width=45)
    listeChoixOrigine.bind('<Double-Button>', doubleClickChoixOrigine)
    listeChoixOrigine.grid(column=0,row=1)
    espace = Label(monConteneurSecondaire)
    espace.grid(column=1,row=1,ipadx=25)
    listeChoixDestination = Listbox(monConteneurSecondaire,width=45)
    listeChoixDestination.bind('<Double-Button>', doubleClickChoixDestination)
    listeChoixDestination.grid(column=2,row=1)
    monConteneurSecondaire.grid(column=0, row=3,columnspan=3,pady=5)
    
    #Extraction
    monBoutonExtraction = Button(monConteneurPrincipal, text="Exportation", command=onClickExport)
    monBoutonExtraction.grid(column=0, row=4,columnspan=3)
    
    #Lancement de la fenêtre
    fenetre.mainloop()
