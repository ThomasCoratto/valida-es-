import pandas as pd

arquivo_excel = r'C:\Users\moesios\Desktop\tratamento contorno python\Resultado_final.xlsx'
df = pd.read_excel(arquivo_excel)

df['Validação_FORMS'] = df.apply(lambda row: 'INVÁLIDO' if pd.isna(row['ID_FORMULARIO']) else ('INVÁLIDO' if df[df['ID_DADOS'] == row['ID_DADOS']]['ID_DADOS'].count() != row['NÚMERO DE PASSAGEIROS'] else 'VÁLIDO'), axis=1)

def dif_hora_validation(row):
    if row['Validação_FORMS'] == 'INVÁLIDO':
        return 'INVÁLIDO', (pd.to_datetime(row['HORÁRIO DE INÍCIO DA VIAGEM']) - pd.to_datetime(row['HORA_y'])).seconds // 60
    
    if pd.isna(row['HORA_y']):
        return 'INVÁLIDO', None  
    
    hora_y = pd.to_datetime(row['HORA_y'])
    hora_inicio_viagem = pd.to_datetime(row['HORÁRIO DE INÍCIO DA VIAGEM'])
    
    if hora_y <= hora_inicio_viagem:
        return 'INVÁLIDO', (hora_inicio_viagem - hora_y).seconds // 60  
    
    return 'VÁLIDO', None

def validacao_logradouro(row, coluna_logradouro):
    if pd.isna(row[coluna_logradouro]):
        return 'INVÁLIDO'  
    elif any(char.isdigit() for char in str(row[coluna_logradouro])):
        return 'Contém número'
    else:
        return 'Verificar via Ponto de referência'

def validar_municipio_origem(row):
    if pd.notna(row['MUNICÍPIO DE ORIGEM DA VIAGEM']) and row['MUNICÍPIO DE ORIGEM DA VIAGEM'] != 'Porto Alegre':
        return 'VÁLIDO'
    return validacao_logradouro(row, 'LOGRADOURO_ORIGEM')

def validar_municipio_destino(row):
    if pd.notna(row['MUNICÍPIO DE DESTINO DA VIAGEM']) and row['MUNICÍPIO DE DESTINO DA VIAGEM'] != 'Porto Alegre':
        return 'VÁLIDO'
    return validacao_logradouro(row, 'LOGRADOURO_DESTINO')

def validar_id_menu(row):
    if df[df['ID_MENU'] == row['ID_MENU']]['DATA_y'].nunique() > 1:
        return 'INVÁLIDO'
    return 'VÁLIDO'

df['Validação_LOGRADOURO_ORIGEM'] = df.apply(validar_municipio_origem, axis=1)
df['Validação_LOGRADOURO_DESTINO'] = df.apply(validar_municipio_destino, axis=1)

df['DIF_HORA'], df['DIF_MIN'] = zip(*df.apply(dif_hora_validation, axis=1))

df['Validação_ORIGEM'] = df.apply(lambda row: 'INVÁLIDO' if row['Validação_FORMS'] == 'INVÁLIDO' else ('INVÁLIDO' if pd.isna(row['ESTADO DE ORIGEM DA VIAGEM']) else ('VÁLIDO' if row['ESTADO DE ORIGEM DA VIAGEM'] != 'RIO GRANDE DO SUL (RS)' or row['MUNICÍPIO DE ORIGEM DA VIAGEM'] != 'Porto Alegre' else 'Requer Análise')), axis=1)

df['Validação_DESTINO'] = df.apply(lambda row: 'INVÁLIDO' if row['Validação_ORIGEM'] == 'INVÁLIDO' else ('INVÁLIDO' if pd.isna(row['ESTADO DE DESTINO DA VIAGEM']) else ('VÁLIDO' if row['MUNICÍPIO DE DESTINO DA VIAGEM'] != 'Rio Grande do Sul (RS)' else 'Requer Análise')), axis=1)


condicao_carga_peso_vazia = (df['TIPO DE VEÍCULO'].str.contains('CARGA', case=False)) & (df['QUAL O PESO DA CARGA? (EM TONELADAS)'].isna())
df.loc[condicao_carga_peso_vazia, 'QUAL O PESO DA CARGA? (EM TONELADAS)'] = 0

df.to_excel(arquivo_excel, index=False)

arquivo_resultado_final = r'C:\Users\moesios\Desktop\tratamento contorno python\Resultado_final.xlsx'
