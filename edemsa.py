import pandas as pd
import numpy as np
import re   #Para evaluar por expresiones regulares el mail

titulo = 'Lectura de archivos. Espere.. '
print('')
print(titulo.center(len(titulo)+70, '-')) 

#Archivo a cargar
try:
    archivo = 'Edemsa/ESTUDIOALTAS1201.xlsx'
    df = pd.read_excel(archivo)  
    print('Primer archivo leído correctamente')
except:
    print("Error al leer el primer archivo")

#Archivo a comparar el documento en columna BN
try:
    archivo2 = 'Edemsa/doc1001.xlsx'
    df2 = pd.read_excel(archivo2)
    print('Segundo archivo leído correctamente')
except:
  print("Error al leer el segundo archivo")

#-------------Validador--------------#
#Evalúa si las columnas válidas(linea 21) se encuentran dentro del 
#archivo a evaluar, si no se encuentra, se muestra cuál falta
titulo = 'Validador de columnas'
print('')
print(titulo.center(len(titulo)+70, '-')) 

columnas = df.columns 
columnas_validas = ['NUMEROOPERACION', 'FECHA',
'NOMBRE', 'DOCUMENTO', 'CUIL', 
'DOMICILIO', 'DIASMORA', 'SALDOTOTAL', 'DEUDAVENCIDA', 
'VTOIMPAGO', 'TIPOPRODUCTO', 'TIPOTELEFONO1',
'DESCRIPCIONTELEFONO1', 'DATOTELEFONO1', 'TIPOTELEFONO2',
'DESCRIPCIONTELEFONO2', 'DATOTELEFONO2', 'TIPOTELEFONO3',
'DESCRIPCIONTELEFONO3',
'DATOTELEFONO3', 'TIPOTELEFONO4', 'DESCRIPCIONTELEFONO4', 
'DATOTELEFONO4',
'EMAIL', 'ANEXO1', 'ANEXO2', 'ANEXO3', 'ANEXO4', 
'ANEXO5', 'ANEXO6', 'ANEXO7',
'ANEXO8', 'ANEXO9', 'ANEXO10', 'IDCOMPANIA']

for i in columnas:
    if i in columnas_validas:
        columnas_validas.remove(i)

if not columnas_validas:
    print('Columnas válidas')
else: 
    print(f'Valores faltantes: {columnas_validas}')

#-------------Info general--------------#
titulo = 'Información general'
print('')
print(titulo.center(len(titulo)+70, '-')) 

#Cantidad de registros
registros = df.shape[0]
print(f'Cantidad de registros: {registros}')

#Cantidad de columnas
print(f'Cantidad de columnas: {df.shape[1]}')

#-------------Utilidades--------------#
lista_vacia = [] #Utilizada para varias columnas que deben estar vacías
for i in range (0, registros):
    lista_vacia.append(np.nan)

#-----------A - NUMEROOPERACION-----------#
df_a = df['NUMEROOPERACION'] #En la columna A del archivo

#----------------B - FECHA-----------------#
df_b = df['FECHA']

#----------------C - NOMBRE-----------------#
df_c = df['NOMBRE'] #En la columna C del archivo

#----------------D - TIPODOCUMENTO-----------------#
#Columna completa de valor "DNI"
lista = []
for i in range(0, registros):
    tipo_doc = 'DNI'
    lista.append(tipo_doc)

df_d = pd.DataFrame({'TIPODOCUMENTO' : lista})

#----------------E - DOCUMENTO-----------------#
df_e = df['DOCUMENTO'] #En columna E del archivo

#----------------F - CUIL-----------------#
df_f = df['CUIL']

#----------------G,H,I - CUIL, ESTADOCIVIL, FECHANACIMIENTO, SEXO-----------------#
dicc = {
    'ESTADOCIVIL' : lista_vacia,
    'FECHANACIMIENTO' : lista_vacia,
    'SEXO' : lista_vacia
}

df_gi = pd.DataFrame(dicc)

#----------------J - DOMICIILIO -----------------#
df_j = df['DOMICILIO']

#----------------K,L,M,N-----------------#
dicc = {
    'LOCALIDAD' : lista_vacia,
    'PROVINCIA' : lista_vacia,
    'CODIGOPOSTAL' : lista_vacia,
    'LABORAL' : lista_vacia
}

df_kn = pd.DataFrame(dicc) 

#----------------O - DIASMORA-----------------#
df_o = df['DIASMORA']

#----------------P,Q,R,S,T-----------------#
dicc = {
    'TOTALCUOTAS' : lista_vacia,
    'CANTIDADCUOTASVENCIDAS' : lista_vacia,
    'CANTIDADCUOTASIMPAGAS' : lista_vacia,
    'NUMEROPRIMERACUOTAVENCIDA' : lista_vacia,
    'CAPITALOTORGADO' : lista_vacia
}

df_pt = pd.DataFrame(dicc) 

#----------------U - SALDOTOTAL-----------------#
#Lo declaro DataFrame para luego poder ser renombrado en la columna V
df_u = pd.DataFrame(df['SALDOTOTAL'])

#----------------V - DEUDAVENCIDA-----------------#
df_v = df_u.rename(columns={'SALDOTOTAL':'DEUDAVENCIDA'})

#----------------W - PAGOMINIMO-----------------#
df_w = pd.DataFrame({'PAGOMINIMO': lista_vacia})

#----------------X - VTOIMPAGO-----------------#
df_x = df['VTOIMPAGO'].apply(str)

fecha_final = []
for i in df_x:
    año = i[0:4]
    mes = i[5:7]
    dia = i[8:10]
    formato = f'{dia}/{mes}/{año}'
    fecha_final.append(formato)

df_x = pd.DataFrame({'VTOIMPAGO' : fecha_final})

#----------------Y, Z, AA, AB, AC, AD, AE, AF, AG-----------------#
dicc = {
    'VALORCUOTA' : lista_vacia,
    'VALORCUOTAULTIMOVTO' : lista_vacia,
    'IMPORTEHONORARIOS' : lista_vacia,
    'COMISIONCANALPAGO' : lista_vacia,
    'CALCULAIVAHONORARIOS' : lista_vacia,
    'IMPORTEIVAGASTOS' : lista_vacia,
    'IMPORTEINTERES' : lista_vacia,
    'IMPORTEGESTION' : lista_vacia,
    'IMPORTEGASTOSFIJOS' : lista_vacia
}

df_yag = pd.DataFrame(dicc)

#----------------AH - TIPOPRODUCTO-----------------#
df_ah = df['TIPOPRODUCTO']

#----------------AI, AJ-----------------#
dicc = {
    'ARTICULO' : lista_vacia,
    'SUCURSAL' : lista_vacia
}

df_aiaj = pd.DataFrame(dicc)

#----------------AK, AL - DESCRIPCIONTELEFONO1, TIPOTELEFONO1-----------------#
celular_1 = df['DATOTELEFONO1']

lista = []
for i in celular_1:
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan )
    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc = {
    'DESCRIPCIONTELEFONO1' : lista,
    'TIPOTELEFONO1' : lista
}

df_akal = pd.DataFrame(dicc)

#----------------AM - DATOTELEFONO1-----------------#
#df_am = pd.DataFrame(df['DATOTELEFONO1'])

dt1 = df['DATOTELEFONO1']
lista = []
lista_nulos = []
lista_no_validos = []
for i in dt1: 
    i = str(i)
    i = i.replace('.0', '')
    
    if '-' in i:
        i = i.replace('-', '')
    if '.' in i:
        i = i.replace('.', '')
    if ' ' in i:
        i = i.replace(' ', '')
        
    if len(i) == 10:
        lista.append(int(i))
    elif len(i) == 3:
        lista.append(np.nan)
        lista_nulos.append(i)
    else:
        lista.append(np.nan)
        lista_no_validos.append(i)

titulo = 'Informe Teléfonos'
print('')
print(titulo.center(len(titulo)+70, '-')) 

print(f'Cantidad de valores distinto a 10 caracteres en DATOTELEFONO1: {len(lista_no_validos)}')
df_am = pd.DataFrame({'DATOTELEFONO1' : lista})
df_am = df_am.replace(lista_no_validos, np.nan)

#----------------AN, AO - DESCRIPCIONTELEFONO2, TIPOTELEFONO2-----------------#
celular_2 = df['DATOTELEFONO2'] #En la columna AH del archivo

lista = []
for i in celular_2:
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan )
    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc = {
    'DESCRIPCIONTELEFONO2' : lista,
    'TIPOTELEFONO2' : lista
}

df_anao = pd.DataFrame(dicc)

#----------------AP - DATOTELEFONO2-----------------#
#Supuestamente los teléfonos están normalizados
#df_ap = pd.DataFrame(df['DATOTELEFONO2'])

dt2 = df['DATOTELEFONO2']
lista = []
lista_nulos = []
lista_no_validos = []
for i in dt2: 
    i = str(i)
    i = i.replace('.0', '')
    
    if '-' in i:
        i = i.replace('-', '')
    if '.' in i:
        i = i.replace('.', '')
    if ' ' in i:
        i = i.replace(' ', '')
        
    if len(i) == 10:
        lista.append(int(i))
    elif len(i) == 3:
        lista.append(np.nan)
        lista_nulos.append(i)
    else:
        lista.append(np.nan)
        lista_no_validos.append(i)

print(f'Cantidad de valores distinto a 10 caracteres en DATOTELEFONO2: {len(lista_no_validos)}')
df_ap = pd.DataFrame({'DATOTELEFONO2' : lista})
df_ap = df_ap.replace(lista_no_validos, np.nan)

#----------------AQ, AR - DESCRIPCIONTELEFONO3, TIPOTELEFONO3-----------------#
tel_fijo = df['DATOTELEFONO3'] 

lista = []
for i in tel_fijo:
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc = {
    'DESCRIPCIONTELEFONO3' : lista,
    'TIPOTELEFONO3' : lista
}

df_aqar = pd.DataFrame(dicc)

#----------------AS - DATOTELEFONO3-----------------#
#Supuestamente los teléfonos están normalizados
#df_as = pd.DataFrame(df['DATOTELEFONO3'])

dt3 = df['DATOTELEFONO3']
lista = []
lista_nulos = []
lista_no_validos = []
for i in dt3: 
    i = str(i)
    i = i.replace('.0', '')
    
    if '-' in i:
        i = i.replace('-', '')
    if '.' in i:
        i = i.replace('.', '')
    if ' ' in i:
        i = i.replace(' ', '')
        
    if len(i) == 10:
        lista.append(int(i))
    elif len(i) == 3:
        lista.append(np.nan)
        lista_nulos.append(i)
    else:
        lista.append(np.nan)
        lista_no_validos.append(i)

print(f'Cantidad de valores distinto a 10 caracteres en DATOTELEFONO3: {len(lista_no_validos)}')
df_as = pd.DataFrame({'DATOTELEFONO3' : lista})
df_as = df_as.replace(lista_no_validos, np.nan)

#----------------AT, AU - DESCRIPCIONTELEFONO4, TIPOTELEFONO4-----------------#
tel_fijo_2 = df['DATOTELEFONO4'] 

lista = []
for i in tel_fijo_2:
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc_columnas = {
    'DESCRIPCIONTELEFONO4' : lista,
    'TIPOTELEFONO4' : lista
}

df_atau = pd.DataFrame(dicc_columnas)

#----------------AV - DATOTELEFONO4-----------------#
#df_av = pd.DataFrame(df['DATOTELEFONO4'])

dt4 = df['DATOTELEFONO4']
lista = []
lista_nulos = []
lista_no_validos = []
for i in dt4: 
    i = str(i)
    i = i.replace('.0', '')
    
    if '-' in i:
        i = i.replace('-', '')
    if '.' in i:
        i = i.replace('.', '')
    if ' ' in i:
        i = i.replace(' ', '')
        
    if len(i) == 10:
        lista.append(int(i))
    elif len(i) == 3:
        lista.append(np.nan)
        lista_nulos.append(i)
    else:
        lista.append(np.nan)
        lista_no_validos.append(i)

print(f'Cantidad de valores distinto a 10 caracteres en DATOTELEFONO4: {len(lista_no_validos)}')
df_av = pd.DataFrame({'DATOTELEFONO4' : lista})
df_av = df_av.replace(lista_no_validos, np.nan)

#----------------AW - TELEFONOSADICIONALES-----------------#
df_aw = pd.DataFrame(df['TELEFONOSADICIONALES'])

#----------------AX, AY, AZ, BA, BB-----------------#
dicc = {
    'FECHAULTIMOPAGO' : lista_vacia,
    'CAPITALADEUDADO(SIST.FRANCES)' : lista_vacia,
    'NUMEROCONVENIO' : lista_vacia,
    'IMPORTETERCERVENCIMIENTO' : lista_vacia,
    'DATOSGESTION' : lista_vacia,
}
df_axbb = pd.DataFrame(dicc)

#----------------EXPRESIONES REGULARES EMAILS-----------------#
df_mails = df['EMAIL']

lista = []
lista2 = []
correctos = []
incorrectos = []
nulos = []
for i in df_mails: 
    #A cada valor lo convierto a formato str
    i = str(i)
    i = i.lower()
           
    if '<' in i:
        i = i.replace('<', '')

    if '>' in i:
        i = i.replace('>', '')
        
    if ',' in i:
        i = i.replace(',', '.')

    if '.@' in i:
        i = i.replace('.@', '@')
                
    if 'á' in i:
        i = i.replace('á', 'a')
                
    if 'é' in i:
        i = i.replace('é', 'e')

    if 'í' in i:
        i = i.replace('í', 'i')
            
    if 'ó' in i:
        i = i.replace('ó', 'o')

    if 'ú' in i:
        i = i.replace('ú', 'u')

    if 'ñ' in i:
        i = i.replace('ñ', 'n')

    if 'mail' in i[-4:]:
        i = i.replace('mail', 'mail.com')

    if 'outlook' in i[-7:]:
        i = i.replace('outlook', 'outlook.com')

    if 'hot' in i[-3:]:
        i = i.replace('hot', 'hotmail.com')

    if ' ' in i:
        i = i.replace(' ', '.')
    
    lista.append(i)

for i in lista:
    if (len(i) <= 3) or (i == 'sinmail@sinmail.com') or (i == '1@1.com'):
        i = np.nan
        nulos.append(i)
        lista2.append(i)
    else:           
        expresion_regular = r"(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|\"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*\")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"
        variable = re.match(expresion_regular, i)
        if variable is None: 
            incorrectos.append(i)
            lista2.append(i)
        else:
            i = i.lower()
            correctos.append(i)   
            lista2.append(i)
        

titulo = 'Informe Mails'
print('')
print(titulo.center(len(titulo)+70, '-')) 

#Valores correctos
print(f'Cantidad Correctos: {len(correctos)}')
#print(f'Valores Correctos: {correctos}\n')

#Valores incorrectos
print(f'Cantidad incorrectos: {len(incorrectos)}')
#print(f'Valores Incorrectos: {incorrectos}\n')
        
#Valores únicos de incorrectos
unique_list = list(dict.fromkeys(incorrectos))
print(f'Valores únicos de incorrectos: {unique_list}\n')

#Cantidad de nulos
print(f'Cantidad de nulos: {len(nulos)}\n')


df_bc = pd.DataFrame({'EMAIL' : lista2})

#-----------------Elimino los valores incorrectos----------------#
df_bc = df_bc.replace(incorrectos, np.nan)

#----------------BD - ANEXO1-----------------#
df_bd = df['ANEXO1']

#----------------BE - ANEXO2-----------------#
df_be = df['ANEXO2']

#----------------BF - ANEXO3-----------------#
df_bf = df['ANEXO3']

#----------------BG - ANEXO4-----------------#
df_bg = df['ANEXO4']

#----------------BH - ANEXO5-----------------#
df_bh = df['ANEXO5']

#----------------BI - ANEXO6-----------------#
df_bi = df['ANEXO6']

#----------------BJ - ANEXO7-----------------#
df_bj = df['ANEXO7']

#----------------BK - ANEXO8-----------------#
df_bk = df['ANEXO8']

#----------------BL - ANEXO9-----------------#
df_bl = df['ANEXO9']

#----------------BM - ANEXO10-----------------#
df_bm = df['ANEXO10']

#----------------BN - ANEXO11-----------------#
dni_anteriores = list(df2['DOCUMENTO'])
dni_actuales = list(df['DOCUMENTO'])

titulo = 'Calculando nuevas asignaciones. Espere.. '
print(titulo.center(len(titulo)+70, '-')) 

lista = []
for i in dni_actuales:
    if i in dni_anteriores:
        lista.append(f'OPERACION|ASIGNACION|STOCK')
    else:
        lista.append(f'OPERACION|ASIGNACION|NUEVA')

df_bn = pd.DataFrame({'ANEXO11' : lista})

contador = 0
for i in lista:
    if 'OPERACION|ASIGNACION|NUEVA' in i:
        contador += 1

print(f'Nuevas Asignaciones: {contador}') 

#----------------BO - IDCOMPANIA-----------------#
df_bo = df['IDCOMPANIA']

#--------------CONCATENAR COLUMNAS-------------#
df_concat = pd.concat([df_a, df_b, df_c, df_d,
df_e, df_f, df_gi, df_j, df_kn, df_o, df_pt,
df_u, df_v, df_w, df_x, df_yag, df_ah, df_aiaj,
df_akal, df_am, df_anao, df_ap, df_aqar,
df_as, df_atau, df_av, df_aw, df_axbb, df_bc,
df_bd, df_be, df_bf, df_bg, df_bh, df_bi, df_bj,
df_bk, df_bl, df_bm, df_bn, df_bo], axis=1)

#--------------OBTENER EXCEL-------------#
titulo = 'Creación del archivo. Espere.. '
print('')
print(titulo.center(len(titulo)+70, '-')) 

try:
    df_concat.to_excel('./Resultados/ESTUDIOALTAS.xlsx', index=False)
    print('Archivo creado correctamente')
    print('')
except:
    print("Error al crear archivo")