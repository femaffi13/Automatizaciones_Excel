import pandas as pd
import numpy as np
from datetime import datetime, timedelta #B
import re     #Para evaluar el mail

#Archivo a cargar
archivo = 'Edesur/original_asigna1001.xlsx'
df = pd.read_excel(archivo)

titulo = 'Información general'
print('')
print(titulo.center(len(titulo)+70, '-')) 

#Cantidad de registros
registros = df.shape[0]
print(f'Cantidad de registros: {registros}')

#Cantidad de columnas
print(f'Cantidad de columnas: {df.shape[1]}')


#-------------Utilidades--------------#
lista_vacia = []
for i in range (0, df.shape[0]):
    lista_vacia.append(np.nan)

#-----------A - NUMEROOPERACION-----------#
id_cuenta = df['RUE'] #A
df_a = pd.DataFrame(id_cuenta)
df_a.rename(columns={'RUE':'NUMEROOPERACION'}, inplace=True)

#----------------B - FECHA-----------------#
lista_b = []
for i in range (0, registros):
    lista_b.append(datetime.today().strftime('%d/%m/%Y'))

df_b = pd.DataFrame({'FECHA' : lista_b})

#----------------C - NOMBRE-----------------#
nombre_titular = df['NOM_CLI'] #P
calle = df['CALLE'] #W
num = df['NRO'].fillna(0).astype(int) #X
dias_mora = df['DIAS_ATRAS'] #M

lista = []
for i in range(0, registros):
    concatenacion = f'{nombre_titular[i]} / CALLE: {calle[i]} / NRO: {num[i]} / DDM: {dias_mora[i]}'
    lista.append(concatenacion)

df_c = pd.DataFrame({'NOMBRE' : lista})

#----------------D - TIPODOCUMENTO-----------------#
lista = []
for i in range(0, registros):
    tipo_doc = 'DNI'
    lista.append(tipo_doc)

df_d = pd.DataFrame({'TIPODOCUMENTO' : lista})

#----------------E - DOCUMENTO-----------------#
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
    concatenacion = f'CALLE: {calle[i]} / NRO: {num[i]}'
    lista.append(concatenacion)

df_j = pd.DataFrame({'DOMICILIO' : lista})

#----------------K - LOCALIDAD-----------------#
localidad = df['NRO_PART'] #AH

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
cp = df['C_POSTAL'].fillna(0).astype(int) #AE
df_m = pd.DataFrame(cp)

# lista = []
# for i in cp:
#     i = str(i)
#     if len(i) <= 3:
#         lista.append('')
#     else:
#         lista.append(i[0:4])

# df_m = pd.DataFrame({'CODIGOPOSTAL' : lista})

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
co = df['SALDO_GESTIONABLE'] / 1000000 #G
df_t = pd.DataFrame(co)
df_t.rename(columns={'SALDO_GESTIONABLE':'CAPITALOTORGADO'}, inplace=True)

#----------------U - SALDOTOTAL-----------------#
df_u = df['SALDO_TOTAL'] / 1000000 #J
df_u = pd.DataFrame(df_u)
df_u.rename(columns={'SALDO_TOTAL':'SALDOTOTAL'}, inplace=True)

#----------------V - DEUDAVENCIDA-----------------#
df_v = df_u.rename(columns={'SALDOTOTAL':'DEUDAVENCIDA'})

#----------------W - PAGOMINIMO-----------------#
df_w = df_t.rename(columns={'CAPITALOTORGADO':'PAGOMINIMO'})

#----------------X - VTOIMPAGO-----------------#
fechas = lista_b
dias = dias_mora

lista_fechas = []
for fecha in fechas: 
    fecha = str(fecha)
    fecha_cadena = fecha
    fecha = datetime.strptime(fecha_cadena, '%d/%m/%Y')
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
dicc = {'VTOIMPAGO' : fecha_final}

df_x = pd.DataFrame(dicc)

#----------------Y, Z-----------------#
dicc = {
    'VALORCUOTA' : lista_vacia,
    'VALORCUOTAULTIMOVTO' : lista_vacia,
}

df_yz = pd.DataFrame(dicc)

#----------------AA - TIPOPRODUCTO-----------------#
col_aa = df['TARIFA'] #E
df_aa = pd.DataFrame(col_aa)
df_aa.rename(columns={'TARIFA':'TIPOPRODUCTO'}, inplace=True)

#----------------AB,AC - ARTICULO, SUCURSAL-----------------#
dicc = {
    'ARTICULO' : lista_vacia,
    'SUCURSAL' : lista_vacia
}

df_abac = pd.DataFrame(dicc)

#---------AD, AE - DESCRIPCIONTELEFONO1, TIPOTELEFONO1---------#
celular_1 = df['movil'] #T

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
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        i = i.replace('.0', '')
        lista.append(i)

df_af = pd.DataFrame({'DATOTELEFONO1' : lista})

#---------AG, AH - DESCRIPCIONTELEFONO2, TIPOTELEFONO2---------#
celular_2 = df['fijo'] #U

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
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        i = i.replace('.0', '')
        lista.append(i)

df_ai = pd.DataFrame({'DATOTELEFONO2' : lista})

#---------AJ, AK - DESCRIPCIONTELEFONO3, TIPOTELEFONO3---------#
tel_fijo = df['NRO_TELEF'] #V

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
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        num = i.replace('.0', '')
        lista.append(num)

df_as = pd.DataFrame({'DATOTELEFONO3' : lista})

#---------AM, AN - DESCRIPCIONTELEFONO4, TIPOTELEFONO4---------#
tel_fijo_2 = df['movil2'] #AT

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
    i = str(i)
    if len(i) == 3:
        lista.append(np.nan)
    else:
        num = i.replace('.0', '')
        lista.append(num)

df_ao = pd.DataFrame({'DATOTELEFONO4' : lista})

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

#----------------EXPRESIONES REGULARES EMAILS-----------------#
df_mails = df[['email', 'email2']] #S,AS

correctos = []
incorrectos = []
nulos = []

for nombre_columna, contenido in df_mails.items(): 
    #A cada valor lo convierto a formato str
    contenido = contenido.apply(str)

    #Recorro cada valor
    for x in contenido: 

        if (x == 'nan' or x == 'sinmail@sinmail.com' or x == '1@1.com'):
            nulos.append(x)
        
        else: 
           expresion_regular = r"(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|\"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*\")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"
           variable = re.match(expresion_regular, x)

           if variable is None: 
                incorrectos.append(x)

           else:
                correctos.append(x)

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

#----------------AV - EMAIL-----------------#
#Si no encuentra un mail en mail_2, trae el que encuentre en mail_2
mail_1 = df_mails['email'].astype(str)
mail_2 = df_mails['email2'].astype(str)

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
tipo_cliente = df['TIPO_CLIENTE'] #AM

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
estado = df['EST_CLI'] #C

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
medidor = df['MEDIDOR'].fillna(0) #BB
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
dni = df['NRO_DOC'].fillna(0) #R
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

#----------------BB - ANEXO5-----------------#
lista = []
for i in range(0, registros):
    tipo_doc = 'OPERACION|CONSUMO INCLUIDO ULTIMA FACTURA|'
    lista.append(tipo_doc)

df_bb = pd.DataFrame({'ANEXO5' : lista})

#----------------BC - ANEXO6-----------------#
asignacion = df['NUEVA_ASIGNACION'] #BR
asignacion = asignacion.fillna('')

lista = []
for i in asignacion: 
    if (i == 1):
        lista.append(f'OPERACION|ASIGNACION|NUEVO')
    elif (i == 0): 
        lista.append(f'OPERACION|ASIGNACION|STOCK')
    else:
        lista.append(np.nan)

df_bc = pd.DataFrame({'ANEXO6' : lista})

#----------------BD - ANEXO7-----------------#
vencimiento = df['FEC_VENC1'] #O

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

df_bd = pd.DataFrame({'ANEXO7' : lista})

#----------------BE - ANEXO8-----------------#
lista = []
for i in range(0, registros):
    tipo_doc = 'OPERACION|FECHA FACTURA VENCIDA|'
    lista.append(tipo_doc)

df_be = pd.DataFrame({'ANEXO8' : lista})

#----------------BF - IDCOMPANIA-----------------#
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
df_concat.to_excel('Resultados/ESTUDIOALTAS1001.xlsx', index=False)
print('Archivo creado correctamente')
