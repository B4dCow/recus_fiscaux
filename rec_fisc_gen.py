# Import Packages
import pandas as pd
import PyPDF2
import os
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import AnnotationBuilder
import fitz
from config import annee, prenom_nom_pres

# Import des donnees adherants
df_ad = pd.read_excel("data/input/Base adhérents.xlsx",header=1)
df_ad.drop(columns=["Unnamed: 0"],inplace=True)
df_ad['id'] = df_ad['Référence'].str[3:]

#Import des donnees operations
df_op = pd.read_excel("data/input/Base opérations.xlsx",header=1)
df_op.drop(columns=["Unnamed: 0"],inplace=True)

# selectionner les categories concernees
cat = set(df_op['sous-catégorie'])
cat_select = ['Dons ponctuels','Cotisations Adhérents','Dons mensuels']

# filtre sur les operations concernees par le recu fiscal
df_op = df_op.loc[df_op['sous-catégorie'].isin(cat_select)]

# extraire identifiant unique
df_op['id'] =  df_op['Tiers'].str[-5:]

# pour chaque emetteur recuperer la somme des montants 
# et le premier mode de paiement
df_op = df_op.groupby('id').agg({'Montant':"sum",'Mode de paiement': lambda x: x.head(1)})
df_op.reset_index(inplace=True)

# fusion des deux bases
df_fisc = df_op.merge(df_ad, on='id', how='left')

# filtre les associations
df_fisc = df_fisc.loc[~(df_fisc['Type']=='Association')]

# signale ceux qui n'ont pas d'adresse
print(f"{sum(df_fisc['Code Postale'].isnull())} adherants n'ont pas d'adresse renseignee")
df_fisc.loc[df_fisc['Code Postale'].isnull(),['Référence','Nom','Prénom','Mail']].to_excel('data/output/adresse_manquante.xlsx',index=False)

# ne garde que ceux qui ont une adresse 
df_fisc = df_fisc.loc[~df_fisc['Code Postale'].isnull()]

# cree les champs a renseigner
df_fisc['TypeDes'] = df_fisc['Type'] + " " + df_fisc['Designation']
df_fisc['AdresseComplete'] = df_fisc['Adresse'] + " " + df_fisc['Code Postale'].astype(int).astype(str).str.zfill(5) + " " + df_fisc['Localité'] 
df_fisc['MontantStr'] =  df_fisc["Montant"].apply(lambda x: "***** {:10.2f} € *****".format(x).replace('.', ','))
df_fisc.reset_index(drop=True, inplace=True)
coord_info_asso = {
    f"Cumul {annee-1}":(130, 226, 200, 219),
    f"5 janvier {annee}":(400, 235, 520, 228),
    f"{prenom_nom_pres}\nPrésident":(350, 145, 480, 125)
}
coord_info_ad= {
    'TypeDes' : (150, 350, 400, 341),
    'Référence' : (450, 350, 520, 343),
    'AdresseComplete': (130, 322, 520, 313),
    'MontantStr' : (400, 292, 520, 282),
    'Mode de paiement' : (185, 196, 250, 189)

}
def template_writer(template_path="data/input/RECU FISCAL.pdf",info_asso=coord_info_asso):
    # import le template
    reader = PdfReader(template_path)
    page = reader.pages[0]
    writer = PdfWriter()
    writer.add_page(page)
    
    # ajouter les elements fixes
    for key in info_asso.keys():
        annotation = AnnotationBuilder.free_text(
            key,
            rect=info_asso[key],
            font="Arial",
            bold=True,
            italic=False,
            font_size="8pt",
            font_color="000000",
            border_color="ffffff",
            background_color="ffffff",
        )
        writer.add_annotation(page_number=0, annotation=annotation)
    
    with open(f"data/input/template_{annee}.pdf", "wb") as fp:
        writer.write(fp)
    doc = fitz.open(f"data/input/template_{annee}.pdf")
    doc[0].insert_image(fitz.Rect(350, 500, 480, 700),filename="data/input/SignPresident.png")
    doc.save(f"data/input/template_{annee}.pdf",incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)

def pdf_writer(index, info_ad=coord_info_ad):

    # import le template
    reader = PdfReader(f"data/input/template_{annee}.pdf")
    page = reader.pages[0]
    writer = PdfWriter()
    writer.add_page(page)

    # Ajouter les elements variables
    for key in info_ad.keys():
        # Create the annotation and add it
        annotation = AnnotationBuilder.free_text(
            f"{df_fisc.loc[index,key]}",
            rect=info_ad[key],
            font="Arial",
            bold=True,
            italic=False,
            font_size="8pt",
            font_color="000000",
            border_color="ffffff",
            background_color="ffffff",
        )
        writer.add_annotation(page_number=0, annotation=annotation)

    with open(f"data/output/RF_{annee}_{df_fisc.loc[index,"Référence"]}.pdf", "wb") as fp:
        writer.write(fp)
