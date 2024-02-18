from rec_fisc_gen import pdf_writer,template_writer,df_fisc

# Ecrit les informations annuelles sur le template
template_writer()

# Genere les recus fiscaux annuels pour chaque adherants
for i in range(df_fisc.shape[0]):
    pdf_writer(i)

# Affiche les informations importantes
print(f"Nombre de recus fiscaux : {df_fisc.shape[0]}")
print(f"Somme des montants : {sum(df_fisc.Montant)}")