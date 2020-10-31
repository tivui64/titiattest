#depuis le dossier app, importer app contenu dans le fichier __init__.py
from app import app
from flask import render_template,request
from docxtpl import DocxTemplate
import zlib
import zipfile
import sys
import os
import comtypes.client
import pythoncom
import win32com.client


dico_motifs={"courses":"Déplacements pour effectuer des achats de fournitures nécessaires à l'activité professionnelle, des achats de première nécessité dans des établissements dont les 3 activités demeurent autorisées, le retrait de commande et les livraisons à domicile.",
            "medical":"Consultations, examens et soins ne pouvant être ni assurés à distance ni différés et l’achat de médicaments",
            "sport":"Déplacements brefs, dans la limite d'une heure quotidienne et dans un rayon maximal d'un kilomètre autour du domicile, liés soit à l'activité physique individuelle des personnes, à l'exclusion de toute pratique sportive collective et de toute proximité avec d'autres personnes, soit à la promenade avec les seules personnes regroupées dans un même domicile, soit aux besoins des animaux de compagnie.",
            "justice":"Convocation judiciaire ou administrative et pour se rendre dans un service public"
            }

@app.route('/')
def afficher_home():
    return render_template("public/home_public.html")

@app.route('/formulaire')
def afficher_formulaire():
    return render_template("public/formulaire.html")

@app.route("/creation_attestation", methods=['POST'])
def creation_attestation():
    reponse = request.form
    heure_de_depart = reponse['heure_de_depart']
    heure_de_retour = reponse['heure_de_retour']
    nom = reponse['nom']
    motif = reponse['motif']
    nb_attestations=1
    heure_dep = int(heure_de_depart[11:13])
    min_dep = int(heure_de_depart[14:16])
    if heure_de_retour:
        heure_ret=int(heure_de_retour[0:2])
        min_ret=int(heure_de_retour[3:5])
        difference=((heure_ret*60)+min_ret)-((heure_dep*60)+min_dep)
        nb_attestations=difference // 60
    document=DocxTemplate(r"C:\Users\Asus\Documents\Formations\Python\Attestation\app\static\text\attestation_deplacement_"+nom+".docx")
    file_names=[]
    for i in range(nb_attestations):
        document.save(r"C:\Users\Asus\Documents\Formations\Python\Attestation\app\static\text\Nouvelle_Attestation_"+str(i+1)+".docx")
        document_2=DocxTemplate(r"C:\Users\Asus\Documents\Formations\Python\Attestation\app\static\text\Nouvelle_Attestation_"+str(i+1)+".docx")
        heure_en_minutes=((heure_dep*60)+min_dep)+(i*60)
        heure_h=heure_en_minutes//60
        heure_min=heure_en_minutes%60
        date="le   "+heure_de_depart[8:10]+"/"+heure_de_depart[5:7]+"/"+heure_de_depart[0:4]+"    à "+str(heure_h).zfill(2)+"h"+str(heure_min).zfill(2)
        for paragraph in document_2.paragraphs:
            if '***' in paragraph.text:
                inline = paragraph.runs
                for j in range(len(inline)):
                    if '***' in inline[j].text:
                        if motif=='courses':
                            text=inline[j].text.replace('***',dico_motifs['courses'])
                            inline[j].text=text
                        elif motif=='medical':
                            text=inline[j].text.replace('***',dico_motifs['medical'])
                            inline[j].text=text
                        elif motif=='sport':
                            text=inline[j].text.replace('***',dico_motifs['sport'])
                            inline[j].text=text
                        elif motif=='justice':
                            text=inline[j].text.replace('***',dico_motifs['justice'])
                            inline[j].text=text
            elif '###' in paragraph.text:
                inline = paragraph.runs
                for j in range(len(inline)):
                    text = inline[j].text.replace('###', date)
                    inline[j].text=text
        document_2.save(r"C:\Users\Asus\Documents\Formations\Python\Attestation\app\static\text\Nouvelle_Attestation_"+str(i+1)+".docx")
        filename="Nouvelle_Attestation_"+str(i+1)+".docx"
        file_names.append(filename)
        """Pour convertir en pdf avec Word Office versions récentes"""
        """wdFormatPDF = 17
        in_file = chemin du fichier word
        out_file = chemin du fichier pdf
        pythoncom.CoInitialize()
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()"""
    # Select the compression mode ZIP_DEFLATED for compression
    # or zipfile.ZIP_STORED to just store the file
    compression = zipfile.ZIP_DEFLATED
    # create the zip file first parameter path/name, second mode
    path="C:\\Users\\Asus\\Documents\\Formations\\Python\\Attestation\\app\\static\\text\\"
    zf = zipfile.ZipFile(r"C:\Users\Asus\Documents\Formations\Python\Attestation\app\static\text\Attestations.zip", mode="w")
    for file_name in file_names:
        # Add file to the zip file
        # first parameter file to zip, second filename in zip
        zf.write(path+file_name, file_name, compress_type=compression)
    # Don't forget to close the file!
    zf.close()
    return render_template("public/creation_attestation.html",nb_attestations=nb_attestations)