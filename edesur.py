import pandas as pd
import numpy as np
from datetime import datetime, timedelta 
import re   #Para evaluar por expresiones regulares el mail

titulo = 'Lectura de archivos. Espere.. '
print('')
print(titulo.center(len(titulo)+70, '-')) 

#Archivo a cargar
try:
    archivo = 'Edesur/Resultado_MasCobr.txt'
    df = pd.read_csv(archivo, sep='|') 
    print('Archivo leído correctamente')   
except:
    print("Error al leer el archivo")

#-------------Validador--------------#
#Evalúa si las columnas válidas(linea 27) se encuentran dentro del 
#archivo a evaluar, si no se encuentra, se muestra cuál falta

titulo = 'Validador de columnas'
print('')
print(titulo.center(len(titulo)+70, '-')) 

columnas = df.columns 
columnas_validas = ['RUE', 'NOM_CLI',
'NRO_PART', 'C_POSTAL', 'SALDO_GESTIONABLE',
'SALDO_TOTAL', 'TARIFA', 'movil',
'fijo', 'NRO_TELEF', 'movil2', 
'email', 'email2', 'TIPO_CLIENTE',
'EST_CLI', 'MEDIDOR', 'NRO_DOC',
'NUEVA_ASIGNACION', 'FEC_VENC1']

for i in columnas:
    if i in columnas_validas:
        columnas_validas.remove(i)

if not columnas_validas:
    print('Columnas válidas')
else: 
    print(f'Columnas faltantes: {columnas_validas}')

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
id_cuenta = df['RUE'] #En columna A del archivo
df_a = pd.DataFrame(id_cuenta)
df_a.rename(columns={'RUE':'NUMEROOPERACION'}, inplace=True)

#----------------B - FECHA-----------------#
#Toda la columna posee la fecha de hoy
lista_b = []
for i in range (0, registros):
    lista_b.append(datetime.today().strftime('%d/%m/%Y'))

df_b = pd.DataFrame({'FECHA' : lista_b})

#----------------C - NOMBRE-----------------#
nombre_titular = df['NOM_CLI'] #En la columna P del archivo
calle = df['CALLE'].astype(str) #En la columna W del archivo
num = df['NRO'].fillna(0).astype(int) #En la columna X del archivo
dias_mora = df['DIAS_ATRAS'] #En la columna M del archivo

lista = []
for i in range(0, registros):
    if (len(calle[i]) <= 3) and (num[i] == 0): 
        concatenacion = f'{nombre_titular[i]} / SIN DIRECCION / DDM: {dias_mora[i]}'
        lista.append(concatenacion)    
    else: 
        concatenacion = f'{nombre_titular[i]} / CALLE: {calle[i]} / NRO: {num[i]} / DDM: {dias_mora[i]}'
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
lista = []
for i in range(0, registros):
    if (len(calle[i]) <= 3) and (num[i] == 0): 
        concatenacion = 'SIN DIRECCION'
        lista.append(concatenacion)    
    else: 
        concatenacion = f'CALLE: {calle[i]} / NRO: {num[i]}'
        lista.append(concatenacion)

df_j = pd.DataFrame({'DOMICILIO' : lista})

#----------------K - LOCALIDAD-----------------#
localidad = df['NRO_PART'] #En la columna AH del archivo

lista = []
for i in localidad:
    i = str(i)
    if len(i) <= 3:
        lista.append('')
    else: 
        lista.append(i)

df_k = pd.DataFrame({'LOCALIDAD': lista})

#----------------L - PROVINCIA-----------------#
#Columna completa de valores 'BUENOS AIRES'
lista = []
for i in range(0, registros):
    prov = 'BUENOS AIRES'
    lista.append(prov)

df_l = pd.DataFrame({'PROVINCIA' : lista})

#----------------M - CODIGOPOSTAL-----------------#
cp = pd.DataFrame(df['C_POSTAL'].fillna(0).astype(int)) #En la columna AE del archivo
df_m = cp.rename(columns={'C_POSTAL':'CODIGOPOSTAL'})

#----------------N - LABORAL-----------------#
df_n = pd.DataFrame({'LABORAL' : lista_vacia})

#----------------O - DIASMORA-----------------#
df_o = pd.DataFrame(dias_mora)
df_o = df_o.rename(columns={'DIAS_ATRAS':'DIASMORA'})

#----------------P,Q,R,S-----------------#
dicc = {
    'TOTALCUOTAS' : lista_vacia,
    'CANTIDADCUOTASVENCIDAS' : lista_vacia,
    'CANTIDADCUOTASIMPAGAS' : lista_vacia,
    'NUMEROPRIMERACUOTAVENCIDA' : lista_vacia,
}

df_ps = pd.DataFrame(dicc)

#----------------T - CAPITALOTORGADO-----------------#
co = df['SALDO_GESTIONABLE'] #En la columna G del archivo
df_t = pd.DataFrame(co)
df_t.rename(columns={'SALDO_GESTIONABLE':'CAPITALOTORGADO'}, inplace=True)

#----------------U - SALDOTOTAL-----------------#
df_u = df['SALDO_TOTAL'] #En la columna J del archivo
df_u = pd.DataFrame(df_u)
df_u.rename(columns={'SALDO_TOTAL':'SALDOTOTAL'}, inplace=True)

#----------------U - SALDOTOTAL-----------------#
lista = df['SALDO_GESTIONABLE'] + df['SALDO_NO_GESTIONABLE'] #En la columna G y H del archivo
df_u = pd.DataFrame({'SALDOTOTAL' : lista})

#----------------V - DEUDAVENCIDA-----------------#
df_v = df_u.rename(columns={'SALDOTOTAL':'DEUDAVENCIDA'})

#----------------W - PAGOMINIMO-----------------#
df_w = df_t.rename(columns={'CAPITALOTORGADO':'PAGOMINIMO'})

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

#----------------Y, Z-----------------#
dicc = {
    'VALORCUOTA' : lista_vacia,
    'VALORCUOTAULTIMOVTO' : lista_vacia,
}

df_yz = pd.DataFrame(dicc)

#----------------AA - TIPOPRODUCTO-----------------#
col_aa = df['TARIFA'] #En la columna E del archivo
df_aa = pd.DataFrame(col_aa)
df_aa.rename(columns={'TARIFA':'TIPOPRODUCTO'}, inplace=True)

#----------------AB,AC - ARTICULO, SUCURSAL-----------------#
dicc = {
    'ARTICULO' : lista_vacia,
    'SUCURSAL' : lista_vacia
}

df_abac = pd.DataFrame(dicc)

#---------AD, AE - DESCRIPCIONTELEFONO1, TIPOTELEFONO1---------#
celular_1 = df['movil'] #En la columna T del archivo

lista = []
for i in celular_1:
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc = {
    'DESCRIPCIONTELEFONO1' : lista,
    'TIPOTELEFONO1' : lista
}

df_adae = pd.DataFrame(dicc)

#----------------AF - DATOTELEFONO1-----------------#
lista = []
for i in celular_1:
    if np.isnan(i) == True:
        lista.append(np.nan)
    else:
        lista.append(i)

df_af = pd.DataFrame({'DATOTELEFONO1' : lista})

#---------AG, AH - DESCRIPCIONTELEFONO2, TIPOTELEFONO2---------#
celular_2 = df['fijo'] #En la columna U del archivo

lista = []
for i in celular_2:
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc = {
    'DESCRIPCIONTELEFONO2' : lista,
    'TIPOTELEFONO2' : lista
}

df_agah = pd.DataFrame(dicc)

#----------------AI - DATOTELEFONO2-----------------#
lista = []
for i in celular_2:
    if np.isnan(i) == True:
        lista.append(np.nan)
    else:
        lista.append(i)

df_ai = pd.DataFrame({'DATOTELEFONO2' : lista})

#---------AJ, AK - DESCRIPCIONTELEFONO3, TIPOTELEFONO3---------#
tel_fijo = df['NRO_TELEF'] #En la columna V del archivo

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

df_ajak = pd.DataFrame(dicc)

#----------------AL - DATOTELEFONO3-----------------#
lista = []
for i in tel_fijo:
    if np.isnan(i) == True:
        lista.append(np.nan)
    else:
        lista.append(i)

df_as = pd.DataFrame({'DATOTELEFONO3' : lista})

#---------AM, AN - DESCRIPCIONTELEFONO4, TIPOTELEFONO4---------#
tel_fijo_2 = df['movil2'] #En la columna AT del archivo

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

df_aman = pd.DataFrame(dicc_columnas)

#----------------AO - DATOTELEFONO4-----------------#
lista = []
for i in tel_fijo_2:
    if np.isnan(i) == True:
        lista.append(np.nan)
    else:
        lista.append(i)

df_ao = pd.DataFrame({'DATOTELEFONO4' : lista})

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

#----------------AP,AQ,AR,AS,AT,AU-----------------#
dicc_columnas = {
    'TELEFONOSADICIONALES' : lista_vacia,
    'FECHAULTIMOPAGO' : lista_vacia,
    'CAPITALADEUDADO(SIST.FRANCES)' : lista_vacia,
    'NUMEROCONVENIO' : lista_vacia,
    'IMPORTETERCERVENCIMIENTO' : lista_vacia,
    'DATOSGESTION' : lista_vacia,
}

df_apau = pd.DataFrame(dicc_columnas)

df_email = df['email']
df_email2 = df['email2']

correctos = []
incorrectos = []
nulos = []
for nombre_columna, i in df_email.items(): 
    #A cada valor lo convierto a formato str
    i = str(i)
    
    if i == 'nan':
        nulos.append(i)
    elif len(i) <= 14:
        incorrectos.append(i)
            #15 para arriba aceptados
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

#Evalúo la segunda columna
correctos2 = []
incorrectos2 = []
nulos2 = []
for nombre_columna, i in df_email2.items(): 
    #A cada valor lo convierto a formato str
    i = str(i)
    
    if i == 'nan':
        nulos2.append(i)
    elif len(i) <= 14:
        incorrectos2.append(i)
            #15 para arriba aceptados
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
            incorrectos2.append(i)

        else:
            correctos2.append(i)

titulo = 'Informe Mails'
print('')
print(titulo.center(len(titulo)+70, '-')) 

#Valores correctos
print(f'Cantidad Correctos: {len(correctos) + len(correctos2)}')
#print(f'Valores Correctos: {correctos}\n')

#Valores incorrectos
print(f'Cantidad incorrectos: {len(incorrectos) + len(incorrectos2)}')
#print(f'Valores Incorrectos: {incorrectos} \n {incorrectos2}\n')
        
#Valores únicos de incorrectos
unique_list = list(dict.fromkeys(incorrectos))
print(f'Valores únicos de incorrectos columna 1: {unique_list}\n')
unique_list = list(dict.fromkeys(incorrectos2))
print(f'Valores únicos de incorrectos columna 2: {unique_list}\n')

#Cantidad de nulos
print(f'Cantidad de nulos: {len(nulos) + len(nulos2)}\n')

#-----------------Elimino los valores incorrectos----------------#
df_email = df_email.replace(incorrectos, np.nan)
df_email = df_email.replace(0, np.nan)

df_email2 = df_email2.replace(incorrectos, np.nan)
df_email2 = df_email2.replace(0, np.nan)

#----------------AV - EMAIL-----------------#
#Si no encuentra un mail en mail_2, trae el que encuentre en mail_2
mail_1 = df_email.astype(str)
mail_2 = df_email2.astype(str)

lista = []
for i in range(0, registros):
    if len(mail_1[i]) > 3:
        lista.append(mail_1[i])
    else: 
        if len(mail_2[i]) > 3:
            lista.append(mail_2[i])
        else:
            lista.append(np.nan)

df_av = pd.DataFrame({'EMAIL' : lista})

#----------------AW - ANEXO1-----------------#
tipo_cliente = df['TIPO_CLIENTE'] #En la columna AM del archivo

lista = []
for i in tipo_cliente: 
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        cadena = f'OPERACION|TIPO DE CLIENTE|{i}'
        lista.append(cadena)

df_aw = pd.DataFrame({'ANEXO1' : lista})

#----------------AX - ANEXO2-----------------#
estado = df['EST_CLI'] #En la columna C del archivo

lista = []
for i in estado: 
    i = str(i)
    if len(i) == 3: 
        lista.append(np.nan)
    else:
        cadena = f'OPERACION|ESTADO DEL SERVICIO|{i}'
        lista.append(cadena)

df_ax = pd.DataFrame({'ANEXO2' : lista})

#----------------AY - ANEXO3-----------------#
lista = []
for i in range(0, registros):
    concatenado = 'OPERACION|CLASIFICACION CLIENTE|'
    lista.append(concatenado)

df_ay = pd.DataFrame({'ANEXO3' : lista})

#----------------AZ - ANEXO4-----------------#
medidor = df['MEDIDOR'].fillna(0) #En la columna BB del archivo
medidor = medidor.astype(str)

lista = []
for i in medidor:
    if i == '0.0':
        lista.append(np.nan)
    else:
        i = i.replace('.0', '')
        cadena = f'OPERACION|NUMERO DE MEDIDOR|{i}'
        lista.append(cadena)

df_az = pd.DataFrame({'ANEXO4' : lista})
df_az = df_az.replace('.0', '')

#----------------BA - ANEXO5-----------------#
dni = df['NRO_DOC'].fillna(0) #En la columna R del archivo
dni = dni.astype(str)

lista = []
for i in dni: 
    if len(i) <= 4:
        lista.append(np.nan)
    else:
        i = i.replace('.0', '')
        cadena = f'CLIENTE|DNI|{i}'
        lista.append(cadena)

df_ba = pd.DataFrame({'ANEXO5' : lista})

#----------------BB - ANEXO6-----------------#
lista = []
for i in range(0, registros):
    tipo_doc = 'OPERACION|CONSUMO INCLUIDO ULTIMA FACTURA|'
    lista.append(tipo_doc)

df_bb = pd.DataFrame({'ANEXO6' : lista})

#----------------BC - ANEXO7-----------------#
asignacion = df['NUEVA_ASIGNACION'] #En la columna BR del archivo
asignacion = asignacion.fillna('')

contador = 0
lista = []
for i in asignacion: 
    if (i == 1):
        lista.append(f'OPERACION|ASIGNACION|NUEVA')
        contador += 1
    elif (i == 0): 
        lista.append(f'OPERACION|ASIGNACION|STOCK')
    else:
        lista.append(np.nan)

df_bc = pd.DataFrame({'ANEXO7' : lista})

titulo = f'Cantidad de Nuevas Asignaciones: {contador}'
print(titulo.center(len(titulo)+70, '-')) 
print('')

#----------------BD - ANEXO8-----------------#
vencimiento = df['FEC_VENC1'] #En la columna O del archivo

lista = []
for i in vencimiento: 
    i = str(i)
    if len(i) <= 3:
        lista.append(np.nan)
    else:
        fecha = i
        fecha = str(fecha)
        
        año = fecha[0:4]
        mes = fecha[5:7]
        dia = fecha[8:10]

        cadena = f'OPERACION|FECHA FACTURA PROXIMO VENCIMIENTO|{dia}_{mes}_{año}'
        lista.append(cadena)

df_bd = pd.DataFrame({'ANEXO8' : lista})

#----------------BE - ANEXO9-----------------#
#A toda la columna el mismo valor
lista = []
for i in range(0, registros):
    tipo_doc = 'OPERACION|FECHA FACTURA VENCIDA|'
    lista.append(tipo_doc)

df_be = pd.DataFrame({'ANEXO9' : lista})

#----------------BF - IDCOMPANIA-----------------#
#A toda la columna el mismo valor
lista_45 = []
for i in range(0, registros):
    lista_45.append(45)

df_bf = pd.DataFrame({'IDCOMPANIA' : lista_45})

#--------------CONCATENAR COLUMNAS-------------#
df_concat = pd.concat([df_a, df_b, df_c, 
df_d, df_e, df_fi, df_j, df_k, df_l, df_m,
df_n, df_o, df_ps, df_t, df_u, df_v, df_w,
df_x, df_yz, df_aa, df_abac, df_adae,
df_af, df_agah, df_ai, df_ajak, df_as, 
df_aman, df_ao, df_apau, df_av, df_aw, 
df_ax, df_ay, df_az, df_ba, df_bb, df_bc,
df_bd, df_be, df_bf], axis=1)

#--------------OBTENER EXCEL-------------#
titulo = 'Creación del archivo. Espere.. '
print('')
print(titulo.center(len(titulo)+70, '-')) 

try:
    df_concat.to_excel('Resultados/ESTUDIOALTAS.xlsx', index=False)
    print('Archivo creado correctamente')
    print('')
except:
    print("Error al crear el archivo")