import pandas as pd
import numpy as np
from datetime import datetime, timedelta 
import re   #Para evaluar por expresiones regulares el mail

titulo = 'Lectura de archivos. Espere.. '
print('')
print(titulo.center(len(titulo)+70, '-')) 

#Archivo a cargar
try:
    archivo = 'Edenor/resultados.txt'
    #df = pd.read_excel(archivo)
    df = pd.read_csv(archivo, sep=';', encoding='latin-1', low_memory=False)
    print('Archivo leído correctamente')
except:
  print("Error al leer el archivo")

#-------------Validador--------------#
#Evalúa si las columnas válidas(linea 19) se encuentran dentro del 
#archivo a evaluar, si no se encuentra, se muestra cuál falta
titulo = 'Validador de columnas'
print('')
print(titulo.center(len(titulo)+70, '-')) 

columnas = df.columns 
columnas_validas = ['ID_CUENTA', 'NOMBRE_TITULAR',
'DIRECCION', 'DIAS_MORA', 'PISO',
'DEPTO_OFICINA', 'BARRIO_REFERENCIA', 'LOCALIDAD',
'SALDO_BALANCE', 'SALDO_EXIGIBLE', 'TARIFA', 
'SITUACION_SUMINISTRO', 'CELULAR_1', 'CELULAR_2',
'FIJO_1', 'FIJO_2', 'CELULAR_3',
'CELULAR_4', 'CELULAR_5', 'CELULAR_6', 'CELULAR_7',
'CELULAR_8', 'CELULAR_9', 'CELULAR_10', 'OTRO_1',
'OTRO_2', 'OTRO_3', 'MAIL_1', 'MAIL_2', 'MAIL_3', 
'NUMERO_DOCUMENTO', 'ESTADO_CLIENTE', 'VTO_ULTIMA_FACTURA',
'ESTADO_SUS', 'CANTIDAD_FACTURAS_VENCIDAS', 'DIAS_ASIGNACION']

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
id_cuenta = df['ID_CUENTA'] #En la columna D del archivo
df_a = pd.DataFrame(id_cuenta)
df_a.rename(columns={'ID_CUENTA':'NUMEROOPERACION'}, inplace=True)

#----------------B - FECHA-----------------#
#Toda la columna posee la fecha de hoy
lista_b = []
for i in range (0, registros):
    lista_b.append(datetime.today().strftime('%d/%m/%Y'))

df_b = pd.DataFrame({'FECHA' : lista_b})

#----------------C - NOMBRE-----------------#
nombre_titular = df['NOMBRE_TITULAR'] #En la columna S del archivo
direccion = df['DIRECCION'] #En la columna U del archivo
dias_mora = df['DIAS_MORA'] #En la columna J del archivo

lista = []
for i in range(0, registros):
    concatenacion = f'{nombre_titular[i]} / CALLE {direccion[i]} / DDM {dias_mora[i]}'
    lista.append(concatenacion)

df_c = pd.DataFrame({'NOMBRE' : lista})

#----------------D - TIPODOCUMENTO-----------------#
#Columna completa de valor "DNI"
lista = []
for i in range(0, registros):
    tipo_doc = 'DNI'
    lista.append(tipo_doc)

df_d = pd.DataFrame({'TIPODOCUMENTO' : lista})

#----------------E - DOCUMENTO-----------------#
#Se coloca el NUMEROOPERACION
df_e = df_a.rename(columns={'NUMEROOPERACION':'DOCUMENTO'})

#----------------F,G,H,I - CUIL, ESTADOCIVIL, FECHANACIMIENTO, SEXO-----------------#
dicc = {
    'CUIL' : lista_vacia,
    'ESTADOCIVIL' : lista_vacia,
    'FECHANACIMIENTO' : lista_vacia,
    'SEXO' : lista_vacia
}

df_fi = pd.DataFrame(dicc) #De f a i

#----------------J - DOMICIILIO -----------------#
direccion = df['DIRECCION'] #En la columna U del archivo
piso = df['PISO'] #En la columna V del archivo
depto = df['DEPTO_OFICINA'] #En la columna W del archivo
barrio = df['BARRIO_REFERENCIA'] #En la columna X del archivo

lista = []
for i in range(0, registros):
    concatenacion = f'CALLE: {direccion[i]} / PISO: {piso[i]} / DPTO: {depto[i]} / BARRIO: {barrio[i]}'
    lista.append(concatenacion)

df_j = pd.DataFrame({'DOMICILIO' : lista})

#----------------K - LOCALIDAD-----------------#
localidad = df['LOCALIDAD'] #En la columna Y del archivo

df_k = pd.DataFrame(localidad)

#----------------L - PROVINCIA-----------------#
#Columna completa de valores 'BUENOS AIRES'
lista = []
for i in range(0, registros):
    prov = 'BUENOS AIRES'
    lista.append(prov)

df_l = pd.DataFrame({'PROVINCIA' : lista})

#----------------M y N - CODIGOPOSTAL, LABORAL-----------------#
dicc = {
    'CODIGOPOSTAL' : lista_vacia,
    'LABORAL' : lista_vacia,
}

df_mn = pd.DataFrame(dicc)

#----------------O - DIASMORA-----------------#
df_o = pd.DataFrame(dias_mora)
df_o = df_o.rename(columns={'DIAS_MORA':'DIASMORA'})

#----------------P,Q,R,S,T-----------------#
dicc_columnas = {
    'TOTALCUOTAS' : lista_vacia,
    'CANTIDADCUOTASVENCIDAS' : lista_vacia,
    'CANTIDADCUOTASIMPAGAS' : lista_vacia,
    'NUMEROPRIMERACUOTAVENCIDA' : lista_vacia,
    'CAPITALOTORGADO' : lista_vacia
}

df_pt = pd.DataFrame(dicc_columnas)

#----------------U - SALDOTOTAL-----------------#
saldo = df['SALDO_BALANCE'] #En la columna F del archivo
df_u = pd.DataFrame(saldo)
df_u.rename(columns={'SALDO_BALANCE':'SALDOTOTAL'}, inplace=True)

#----------------V - DEUDAVENCIDA-----------------#
df_v = df_u.rename(columns={'SALDOTOTAL':'DEUDAVENCIDA'})

#----------------W - PAGOMINIMO-----------------#
col_w = df['SALDO_EXIGIBLE'] #En la columna E del archivo
df_w = pd.DataFrame(col_w)
df_w.rename(columns={'SALDO_EXIGIBLE':'PAGOMINIMO'}, inplace=True)

#----------------X - VTOIMPAGO-----------------#
fechas = lista_b #Variable proveniente de la Sección B
dias = dias_mora #Variable proveniente de la Sección C

lista_fechas = []
for fecha in fechas: 
    fecha = str(fecha)
    fecha = datetime.strptime(fecha, '%d/%m/%Y')
    lista_fechas.append(fecha)

lista_dias = []
for dia in dias: 
    td = timedelta(dia)
    lista_dias.append(td)

#Diferencia entre fecha y n días
lista_3 = []
for i in range(len(lista_fechas)):
    lista_3.append(lista_fechas[i] - lista_dias[i])

#Formato a la fecha final
fecha_final = []
for i in lista_3:
    i = str(i)
    año = i[0:4]
    mes = i[5:7]
    dia = i[8:10]
    formato = f'{dia}/{mes}/{año}'
    fecha_final.append(formato)

#Transformo a DataFrame
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
col_ah = df['TARIFA'] #En la columna T del archivo
df_ah = pd.DataFrame(col_ah)
df_ah.rename(columns={'TARIFA':'TIPOPRODUCTO'}, inplace=True)

#----------------AI - ARTICULO-----------------#
df_ai = pd.DataFrame({'ARTICULO' : lista_vacia})

#----------------AJ - SUCURSAL-----------------#
aj = df['SITUACION_SUMINISTRO'] #En la columna M del archivo
df_aj = pd.DataFrame(aj)
df_aj.rename(columns={'SITUACION_SUMINISTRO':'SUCURSAL'}, inplace=True)

#----------------AK, AL - DESCRIPCIONTELEFONO1, TIPOTELEFONO1-----------------#
celular_1 = df['CELULAR_1'] #En la columna AG del archivo

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
lista = []
for i in celular_1:
    if np.isnan(i) == True:
        lista.append(np.nan)
    else:
        lista.append(i)

df_am = pd.DataFrame({'DATOTELEFONO1' : lista})

#----------------AN, AO - DESCRIPCIONTELEFONO2, TIPOTELEFONO2-----------------#
celular_2 = df['CELULAR_2'] #En la columna AH del archivo

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
lista = []
for i in celular_2:
    if np.isnan(i) == True:
        lista.append(np.nan)
    else:
        lista.append(i)

df_ap = pd.DataFrame({'DATOTELEFONO2' : lista})

#----------------AQ, AR - DESCRIPCIONTELEFONO3, TIPOTELEFONO3-----------------#
tel_fijo = df['FIJO_1'] #En la columna AQ del archivo

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
lista = []
for i in tel_fijo:
    if np.isnan(i) == True:
        lista.append(np.nan)
    else:
        lista.append(i)

df_as = pd.DataFrame({'DATOTELEFONO3' : lista})

#----------------AT, AU - DESCRIPCIONTELEFONO4, TIPOTELEFONO4-----------------#
tel_fijo_2 = df['FIJO_2'] #En la columna AR del archivo

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
lista = []
for i in tel_fijo_2:
    if np.isnan(i) == True:
        lista.append(np.nan)
    else:
        lista.append(i)

df_av = pd.DataFrame({'DATOTELEFONO4' : lista})

#----------------Análisis de columnas teléfonos-----------------#
titulo = 'Informe Teléfonos'
print('')
print(titulo.center(len(titulo)+70, '-')) 

contador = 0
for i in range(0, registros):
    if (np.isnan(celular_1[i]) == True) and (np.isnan(celular_2[i]) == True) and (np.isnan(tel_fijo[i]) == True) and (np.isnan(tel_fijo_2[i]) == True):
        contador += 1
print(f'Cantidad de registros sin ningún teléfono: {contador}')
print(f'Representa un {round(contador/registros*100, 2)}% del total ')

#----------------AW - TELEFONOSADICIONALES-----------------#
cel_3 = df['CELULAR_3'].fillna('') #AI 
cel_4 = df['CELULAR_4'].fillna('') #AJ
cel_5 = df['CELULAR_5'].fillna('') #AK
cel_6 = df['CELULAR_6'].fillna('') #AL
cel_7 = df['CELULAR_7'].fillna('') #AM
cel_8 = df['CELULAR_8'].fillna('') #AN
cel_9 = df['CELULAR_9'].fillna('') #AO
cel_10 = df['CELULAR_10'].fillna('') #AP
otro_1 = df['OTRO_1'].fillna('') #AS
otro_2 = df['OTRO_2'].fillna('') #AT
otro_3 = df['OTRO_3'].fillna('') #AU

lista_aw = []
for i in range(0, registros): 
    cadena = f'CELULAR:{cel_3[i]}|CELULAR:{cel_4[i]}|CELULAR:{cel_5[i]}|CELULAR:{cel_6[i]}|CELULAR:{cel_7[i]}|CELULAR:{cel_8[i]}|CELULAR:{cel_9[i]}|CELULAR:{cel_10[i]}|CELULAR:{otro_1[i]}|CELULAR:{otro_2[i]}|CELULAR:{otro_3[i]}'
    lista_aw.append(cadena)

lista_nueva = []
for i in lista_aw:
    i = i.replace('.0', '')
    lista_nueva.append(i)

df_aw = pd.DataFrame({'TELEFONOSADICIONALES' : lista_nueva})

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
#Evalúa las columnas MAIL_1 a MAIL_3
lista = []
for i in range(1,4):
    col = f'MAIL_{i}'
    lista.append(col)

df_mails = df[lista]

correctos = []
incorrectos = []
nulos = []

for nombre_columna, contenido in df_mails.items(): 
    #A cada valor lo convierto a formato str
    contenido = contenido.apply(str)

    #Recorro cada valor
    for i in contenido: 
        if (i == 'nan' or i == 'sinmail@sinmail.com' or i == '1@1.com'):
            nulos.append(i)
        
        else: 
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
                
            expresion_regular = r"(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|\"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*\")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"
            variable = re.match(expresion_regular, i)

            if variable is None: 
                incorrectos.append(i)

            else:
                correctos.append(i)

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

#-----------------Elimino los valores incorrectos----------------#
df_mails = df_mails.replace(incorrectos, np.nan)
df_mails = df_mails.replace(0, np.nan)

#----------------BC - EMAIL-----------------#
mail_1 = df_mails['MAIL_1'] #En la columna AV del archivo
lista = []

for i in mail_1:
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        lista.append(i)

df_bc = pd.DataFrame({'EMAIL' : lista})

#----------------BD - ANEXO1-----------------#
mail_2 = df_mails['MAIL_2'] #En la columna AW del archivo
lista = []

for i in mail_2: 
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        cadena = f'OPERACION|MAIL 2|{i}'
        lista.append(cadena)

df_bd = pd.DataFrame({'ANEXO1' : lista})

#----------------BE - ANEXO2-----------------#
mail_3 = df_mails['MAIL_3'] #En la columna AX del archivo
lista = []

for i in mail_3: 
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        cadena = f'OPERACION|MAIL 3|{i}'
        lista.append(cadena)

df_be = pd.DataFrame({'ANEXO2' : lista})

#----------------BF - ANEXO3-----------------#
dni = df['NUMERO_DOCUMENTO'] #En la columna P del archivo
lista = []

for i in range(0, registros): 
    if i == np.nan:
        lista.append(np.nan)
    else:
        cadena = f'OPERACION|DNI|{dni[i]}'
        lista.append(cadena)

df_bf = pd.DataFrame({'ANEXO3' : lista})

#----------------BG - ANEXO4-----------------#
estado = df['ESTADO_CLIENTE'] #En la columna Q del archivo
lista = []

for i in range(0, registros): 
    if i == np.nan:
        lista.append(np.nan)
    else:
        cadena = f'OPERACION|ESTADO CLIENTE|{estado[i]}'
        lista.append(cadena)

df_bg = pd.DataFrame({'ANEXO4' : lista})

#----------------BH - ANEXO5-----------------#
vto = df['VTO_ULTIMA_FACTURA'] #En la columna G del archivo
lista = []

for i in vto: 
    i = str(i)
    if len(i) <= 3:
        lista.append(np.nan)
    else:
        fecha = i
        fecha = str(fecha)
        
        año = fecha[0:4]
        mes = fecha[5:7]
        dia = fecha[8:10]

        cadena = f'OPERACION|VTO ULTIMA FC|{dia}_{mes}_{año}'
        lista.append(cadena)

df_bh = pd.DataFrame({'ANEXO5' : lista})

#----------------BI - ANEXO6-----------------#
estado_susp = df['ESTADO_SUS'] #En la columna R del archivo
lista = []

for i in range(0, registros): 
    if i == np.nan: 
        lista.append(np.nan)
    else:
        cadena = f'OPERACION|ESTADO SUSP|{estado_susp[i]}'
        lista.append(cadena)

df_bi = pd.DataFrame({'ANEXO6' : lista})

#----------------BJ - ANEXO7-----------------#
facturas = df['CANTIDAD_FACTURAS_VENCIDAS'] #En la columna L del archivo
lista = []

for i in range(0, registros): 
    if i == np.nan:
        lista.append(np.nan)
    else:
        cadena = f'OPERACION|FACTURAS VENCIDAS|{facturas[i]}'
        lista.append(cadena)

df_bj = pd.DataFrame({'ANEXO7' : lista})

#----------------BK - ANEXO8-----------------#
asignacion = df['DIAS_ASIGNACION'] #En la columna BR del archivo
asignacion = asignacion.fillna(0)

contador = 0
lista = []
for i in asignacion: 
    i = int(i)
    if (i == 0):
        lista.append(f'OPERACION|ASIGNACION|NUEVA')
        contador += 1
    elif (i >= 1 and i <= 60): 
        lista.append(f'OPERACION|ASIGNACION|STOCK')
    elif i > 60:
        lista.append(f'OPERACION|ASIGNACION|NUEVA')
        contador += 1

df_bk = pd.DataFrame({'ANEXO8' : lista})

titulo = f'Cantidad de Nuevas Asignaciones: {contador}'
print(titulo.center(len(titulo)+70, '-')) 
print('')

#----------------BL - IDCOMPANIA-----------------#
lista_52 = []
for i in range(0, registros):
    lista_52.append(52)

df_bl = pd.DataFrame({'IDCOMPANIA' : lista_52})

#--------------CONCATENAR COLUMNAS-------------#
df_concat = pd.concat([df_a, df_b, df_c, df_d, df_e, 
df_fi, df_j, df_k, df_l, df_mn, 
df_o, df_pt, df_u, df_v, df_w,
df_x, df_yag, df_ah, df_ai, df_aj,
df_akal, df_am, df_anao, df_ap, df_aqar,
df_as, df_atau, df_av, df_aw, df_axbb,
df_bc, df_bd, df_be, df_bf, df_bg,
df_bh, df_bi, df_bj, df_bk, df_bl], axis=1)

#--------------OBTENER EXCEL-------------#
titulo = 'Creación del archivo. Espere.. '
print('')
print(titulo.center(len(titulo)+70, '-')) 

try:
    df_concat = df_concat.applymap(lambda x: x.encode('unicode_escape').
                 decode('utf-8') if isinstance(x, str) else x)

    df_concat.to_excel('Resultados/ESTUDIOALTAS.xlsx', index=False)
    print('Archivo creado correctamente')
    print('')
except:
    print("Error al crear el archivo")