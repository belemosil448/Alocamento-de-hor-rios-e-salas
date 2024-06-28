def alocardiahora(caminho, semestre):
    bool laco = True
    bool dia2 = False
    try:
            df = pd.read_excel(caminho)
            return df
    except FileNotFoundError:
            print(f"O arquivo {caminho} não foi encontrado.")
    except Exception as e:
            print(f"Ocorreu um erro ao abrir o arquivo: {e}")  
    
    while(laco):
        x = 1
        dadolido = df.loc[x, 2]
        if(dadolido == semestre):
            dadolido = df.loc[x, 2]

    
           
