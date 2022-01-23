import pandas as pd
import numpy as np

# path_to_excel = "C:/Users/Legion/OneDrive - Universidade de Aveiro/Universidade/4_ano/1_semestre/FA/trabalho2/DADOSCPParadaPinhov1.xlsx"
# df = pd.read_excel(path_to_excel)

# df.at[:,'DATA'] = [d.strftime('%d-%m-%Y') if not pd.isnull(d) else '' for d in df['DATA']]
# df['DATA'] = df['DATA'].astype(str)

# sheets = df.columns.values[1:-1:3][1:26]
# #print(sheets)
# #print(len(sheets))

# data_column = df["DATA"].values

# psi20_df = df.iloc[:, 1:4]
# returns = [np.nan for i in range(len(data_column))]
# for i in range(1, len(returns)):
#     if not (pd.isna(psi20_df.iloc[i, 1]) or pd.isna(psi20_df.iloc[i - 1, 1])):
#         returns[i] = np.log(psi20_df.iloc[i, 1]) - np.log(psi20_df.iloc[i - 1, 1])

# psi20_df["PSI20RETORNOS"] = returns

# writer = pd.ExcelWriter("C:/Users/Legion/OneDrive - Universidade de Aveiro/Universidade/4_ano/1_semestre/FA/trabalho2/result.xlsx", engine='xlsxwriter')
# c = 4

# for s in sheets:
#     df1 = df.iloc[:, c : c + 3].copy()
#     df1.columns = ["INDICE_PRECO","INDICE_VOL","INDICE_CAP"]
#     returns = [np.nan for i in range(len(data_column))]
#     for i in range(1, len(returns)):
#         if not (pd.isna(df1.iloc[i, 0]) or pd.isna(df1.iloc[i - 1, 0])):
#             returns[i] = np.log(df1.iloc[i, 0]) - np.log(df1.iloc[i - 1, 0])
#     c += 3
#     df1["RETORNOS"] = returns
#     final_df = pd.concat([psi20_df, df1], axis = 1)
#     final_df.insert(loc=0, column="DATA", value=data_column)
#     final_df.to_excel(writer, s, index=False)
#     print(s, "done")

# writer.save()

path_to_excel = "C:/Users/Legion/OneDrive - Universidade de Aveiro/Universidade/4_ano/1_semestre/FA/trabalho2/result.xlsx"

sheets = ['BCP','BES','BPI','BRISA','CIMPOR','EDP','JM','PT','SEMAPA',
'SONAE','ZON','SONAECOM','EDPRENOV','GALP','MOTAENGIL','PORTUCEL','REN',
'SONAEIND','ALTRI','COFINA','IMPRESA','MEDIACAP','NOVABASE','PARAREDE' ,
'TEIXEIRADUARTE']

writer = pd.ExcelWriter("C:/Users/Legion/OneDrive - Universidade de Aveiro/Universidade/4_ano/1_semestre/FA/trabalho2/excel_tratado.xlsx", engine='xlsxwriter')
n_dump = 15


for s in sheets:
    df = pd.read_excel(path_to_excel, sheet_name=s)

    try:
        df['EVENTOS_RESULTADOS'] = pd.to_datetime(df['EVENTOS_RESULTADOS']).dt.strftime('%d-%m-%Y')
        all_dates = df["EVENTOS_RESULTADOS"].dropna().values
        all_dates = all_dates.astype(str)

        marked_data = [np.nan for _ in range(len(df["DATA"]))]
        for d in all_dates:
            for i, data in enumerate(df["DATA"].values):
                if d == data:
                    if i - n_dump > 0:
                        helper = [d for _ in range(i - n_dump, i + n_dump + 1)]
                        marked_data[i - n_dump : i + n_dump + 1] = helper
                    else:
                        marked_data[i] = d

        df["DATA_RESULTADOS"] = marked_data
    except Exception as e:
        print(s, "erro resultados")

    #try:
    df['EVENTOS_DIVIDENDOS'] = pd.to_datetime(df['EVENTOS_DIVIDENDOS']).dt.strftime('%d-%m-%Y')
    all_dates = df["EVENTOS_DIVIDENDOS"].dropna().values
    all_dates = all_dates.astype(str)

    marked_data = [np.nan for _ in range(len(df["DATA"]))]
    for d in all_dates:
        for i, data in enumerate(df["DATA"].values):
            if d == data:
                if i - n_dump > 0 and i + n_dump < len(marked_data):
                    helper = [d for _ in range(i - n_dump, i + n_dump + 1)]
                    marked_data[i - n_dump : i + n_dump + 1] = helper
                else:
                    marked_data[i] = d

    df["DATA_DIVIDENDOS"] = marked_data
    #except Exception as e:
    #    print(s, "erro dividendos")
    
    df.to_excel(writer, s, index=False)
    print(s, "done")

writer.save()

