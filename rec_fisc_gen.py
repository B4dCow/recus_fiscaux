# Import Packages
import pandas as pd
import os
import fitz
from num2words import num2words
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
df_fisc['MontantStr'] =  df_fisc["Montant"].apply(lambda x: "***** {:10.2f} &#8364; *****".format(x).replace('.', ','))

def montant_lettres(num):
    # fonction tranformant le montant en lettres
    euros,centimes = "{:10.2f}".format(num).split(".")
    if centimes == "00":
        return f"{num2words(int(euros), lang='fr')} EUROS".upper()
    else:
        return f"{num2words(int(euros), lang='fr')} EUROS ET \
{num2words(int(centimes), lang='fr')} CENTIIMES".upper()
    
df_fisc['MontantLettres'] =  df_fisc["Montant"].apply(montant_lettres)

df_fisc.reset_index(drop=True, inplace=True)
coord_info_asso = {
    f"Cumul {annee-1}": (150, 564, 350, 584),
    f"5 janvier {annee}": (405, 555, 520, 575),
    f"{prenom_nom_pres}\nPrésident": (350, 645, 480, 670)
}
coord_info_ad= {
    'TypeDes' : (150, 442, 350, 501),
    'Référence' : (450, 442, 520, 501),
    'AdresseComplete': (130, 465, 520, 500),
    'MontantStr' : (400, 505, 520, 520),
    'Mode de paiement' : (185, 594, 270, 620),
    'MontantLettres' : (160, 540, 565, 565)

}

   
def html_format(txt):
    html_txt = """<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exemple</title>
    <style>
        body {
            font-family: Arial, sans-serif;
        }

        .custom-text {
            font-weight: bold;
            font-size: 7pt;
        }
    </style>
</head>
<body>
    <span class="custom-text">"""+txt+"""</span>
</body>"""
    return html_txt

def template_writer(template_path="data/input/RECU FISCAL.pdf",info_asso=coord_info_asso):
    # importe le template
    doc = fitz.open(template_path)
    page = doc[0]
    # ajouter les elements fixes
    for key in info_asso.keys():    
        rc = page.insert_textbox(info_asso[key],key,fontname="helv",fontsize=8)
 
    page.insert_image(fitz.Rect(350, 500, 480, 700),filename="data/input/SignPresident.png")
    doc.save(f"data/input/template_{annee}.pdf")
 
def pdf_writer(index, info_ad=coord_info_ad):

    # import le template
    doc = fitz.open(f"data/input/template_{annee}.pdf")
    page = doc[0]

    # Ajouter les elements variables
    for key in info_ad.keys():
        # Create the annotation and add it
        page.insert_htmlbox(info_ad[key] ,html_format(f"{df_fisc.loc[index,key]}"))
                            #fontname="helv",fontsize=8)
    doc.save(f"data/output/RF_{annee}_{df_fisc.loc[index,'Référence']}.pdf")    