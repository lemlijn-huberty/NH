import pandas as pd
import numpy as np
import re

# last changed 31/12/2020
# only > XLS < files

df1 = pd.read_excel("C:/Python/primesSyndicales/Global.xls")
df2 = df1[0:5]
na1 = df2.to_numpy().reshape(1,50)
li1 = na1[0]
df4 = pd.DataFrame(columns=li1)

arr = np.empty((0, 3))
np.warnings.filterwarnings('ignore', category=np.VisibleDeprecationWarning)
for ind in df1.index:
  if not pd.isnull(df1.loc[ind][0]):
    #print(re.findall(r'\d+', df1.loc[ind][0])[0])
    df2 = df1[ind:ind+5]
    agent = df1.loc[ind][0].split('(')
    integers = re.findall(r'\d+', df1.loc[ind][0])
    arr = np.append(arr, np.array([[agent[0], integers[0], integers[1]]]), axis=0)

    #print(agent[0], integers[0], integers[1])
    #df2.iloc[0, 1] = 'test'
    na2 = df2.to_numpy().reshape(1, 50)
    df3 = pd.DataFrame(na2, columns=li1)
    df4 = df4.append(df3)

df5 = pd.DataFrame(arr, columns=['Agent', 'NISS', 'numAgent'])
df5.to_excel("C:/Python/primesSyndicales/Global_agents.xls", index=False)

df4.to_excel("C:/Python/primesSyndicales/Global_retravaillé.xls", index=False)

#df2 = df4[["Unnamed: 0", "Numéro occupation", "Date début occupation", "Date fin occupation", "N° Commission Paritaire", "Nbr jrs sem. régime trav.", "Type du contrat", "Données de l'occupation", "Justification des jours", "Nbr moyen heures sem pers ref", "Statut Travailleur", "Nbr moyen heures sem trav", "Mesure réorganisation tps trv.", "Mesure Promotion Emploi", "Notion pensionné", "IdNavire", "Code NACE", "Type d'apprentissage", "Mode de rémunération", "Numéro de fonction", "Classe du personnel volant", "Paiement Dixièmes | Douzièmes", "Ref occupation", "ClassePersonnel", "N° version occup"]]
##df6 = df4.iloc[[1, 2], [1, 2]]

#df6 = pd.read_excel("C:/Python/primesSyndicales/Global_retravaillé.xls")
#"df7 = df6[["Unnamed: 0", "Numéro occupation", "Date début occupation", "Date fin occupation", "N° Commission Paritaire", "Nbr jrs sem. régime trav.", "Type du contrat", "Données de l'occupation", "Justification des jours", "Nbr moyen heures sem pers ref", "Statut Travailleur", "Nbr moyen heures sem trav", "Mesure réorganisation tps trv.", "Mesure Promotion Emploi", "Notion pensionné", "IdNavire", "Code NACE", "Type d'apprentissage", "Mode de rémunération", "Numéro de fonction", "Classe du personnel volant", "Paiement Dixièmes | Douzièmes", "Ref occupation", "ClassePersonnel", "N° version occup"]]
#"df7.to_excel("C:/Python/primesSyndicales/Global_retravaillé_2.xls", index=False)

# iterating the columns
#for col in df4.columns:
#    print(col)



