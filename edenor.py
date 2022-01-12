import pandas as pd
import numpy as np
from datetime import datetime, timedelta #B
import re     #Para evaluar el mail

#Excel Asigna - 23/12/21
archivo = 'asigna_0511.xlsx'
df = pd.read_excel(archivo)

#Cantidad de valores
registros = df.shape[0]

#-------------Utilidades--------------#
lista_vacia = []

for i in range (0, registros):
    i = np.nan 
    lista_vacia.append(i)

#-----------A - NUMEROOPERACION-----------#
id_cuenta = df['ID_CUENTA'] #D
df_a = pd.DataFrame(id_cuenta)
df_a.rename(columns={'ID_CUENTA':'NUMEROOPERACION'}, inplace=True)

#----------------B - FECHA-----------------#
lista_b = []
for i in range (0, registros):
    lista_b.append(datetime.today().strftime('%d/%m/%Y'))

dic_b = {
    'FECHA' : lista_b
}

df_b = pd.DataFrame(dic_b)

#----------------C - NOMBRE-----------------#
nombre_titular = df['NOMBRE_TITULAR'] #S
direccion = df['DIRECCION'] #U
dias_mora = df['DIAS_MORA'] #J

concatenacion = f'{nombre_titular[0]} / CALLE {direccion[0]} / DDM {dias_mora[0]}'

lista_c = []

for i in range(0, registros):
    concatenacion = f'{nombre_titular[i]} / CALLE {direccion[i]} / DDM {dias_mora[i]}'
    lista_c.append(concatenacion)

dicc_c = {
    'NOMBRE' : lista_c
}

df_c = pd.DataFrame(dicc_c)

#----------------D - TIPODOCUMENTO-----------------#
lista_d = []

for i in range(0, registros):
    tipo_doc = 'DNI'
    lista_d.append(tipo_doc)

dicc_d = {
    'TIPODOCUMENTO' : lista_d
}

df_d = pd.DataFrame(dicc_d)

#----------------E - DOCUMENTO-----------------#
df_e = df_a.rename(columns={'NUMEROOPERACION':'DOCUMENTO'})

#----------------F,G,H,I - CUIL, ESTADOCIVIL, FECHANACIMIENTO, SEXO-----------------#
dicc_fi = {
    'CUIL' : lista_vacia,
    'ESTADOCIVIL' : lista_vacia,
    'FECHANACIMIENTO' : lista_vacia,
    'SEXO' : lista_vacia
}

df_fi = pd.DataFrame(dicc_fi) #De f a i

#----------------J - DOMICIILIO -----------------#
direccion = df['DIRECCION'] #U
piso = df['PISO'] #V
depto = df['DEPTO_OFICINA'] #W
barrio = df['BARRIO_REFERENCIA'] #X

concatenacion = f'CALLE: {direccion[0]} / PISO {piso[0]} / DPTO: {depto[0]} / BARRIO: {barrio[0]}'

lista_dom = []

for i in range(0, registros):
    concatenacion = f'CALLE: {direccion[i]} / PISO: {piso[i]} / DPTO: {depto[i]} / BARRIO: {barrio[i]}'
    lista_dom.append(concatenacion)

dicc_j = {
    'DOMICILIO' : lista_dom
}

df_j = pd.DataFrame(dicc_j)

#----------------K - LOCALIDAD-----------------#
localidad = df['LOCALIDAD'] #Y

df_k = pd.DataFrame(localidad)

#----------------L - PROVINCIA-----------------#
lista_l = []

for i in range(0, registros):
    prov = 'BUENOS AIRES'
    lista_l.append(prov)

dicc_prov = {
    'PROVINCIA' : lista_l
}

df_l = pd.DataFrame(dicc_prov)

#----------------M y N - CODIGOPOSTAL, LABORAL-----------------#
dicc_columnas = {
    'CODIGOPOSTAL' : lista_vacia,
    'LABORAL' : lista_vacia,
}

df_mn = pd.DataFrame(dicc_columnas)

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
saldo = df['SALDO_BALANCE'] #F
df_u = pd.DataFrame(saldo)
df_u.rename(columns={'SALDO_BALANCE':'SALDOTOTAL'}, inplace=True)

#----------------V - DEUDAVENCIDA-----------------#
df_v = df_u.rename(columns={'SALDOTOTAL':'DEUDAVENCIDA'})

#----------------W - PAGOMINIMO-----------------#
col_w = df['SALDO_EXIGIBLE'] #E
df_w = pd.DataFrame(col_w)
df_w.rename(columns={'SALDO_EXIGIBLE':'PAGOMINIMO'}, inplace=True)

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

#----------------Y, Z, AA, AB, AC, AD, AE, AF, AG-----------------#
dicc_columnas = {
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

df_yag = pd.DataFrame(dicc_columnas)

#----------------AH - TIPOPRODUCTO-----------------#
col_ah = df['TARIFA'] #T
df_ah = pd.DataFrame(col_ah)
df_ah.rename(columns={'TARIFA':'TIPOPRODUCTO'}, inplace=True)

#----------------AI - ARTICULO-----------------#
dicc_columnas = {
    'ARTICULO' : lista_vacia,
}

df_ai = pd.DataFrame(dicc_columnas)

#----------------AJ - SUCURSAL-----------------#
aj = df['SITUACION_SUMINISTRO'] #M
df_aj = pd.DataFrame(aj)
df_aj.rename(columns={'SITUACION_SUMINISTRO':'SUCURSAL'}, inplace=True)

#----------------AK, AL - DESCRIPCIONTELEFONO1, TIPOTELEFONO1-----------------#
celular_1 = df['CELULAR_1'] #AG

lista = []

for i in celular_1:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista.append(i)

    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc_columnas = {
    'DESCRIPCIONTELEFONO1' : lista,
    'TIPOTELEFONO1' : lista
}

df_akal = pd.DataFrame(dicc_columnas)

#----------------AM - DATOTELEFONO1-----------------#
lista_am = []

for i in celular_1:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista_am.append(i)

    else:
        i = i.replace('.0', '')
        lista_am.append(i)

dicc_columnas = {
    'DATOTELEFONO1' : lista_am,
}

df_am = pd.DataFrame(dicc_columnas)

#----------------AN, AO - DESCRIPCIONTELEFONO2, TIPOTELEFONO2-----------------#
celular_2 = df['CELULAR_2'] #AH

lista = []

for i in celular_2:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista.append(i)

    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc_columnas = {
    'DESCRIPCIONTELEFONO2' : lista,
    'TIPOTELEFONO2' : lista
}

df_anao = pd.DataFrame(dicc_columnas)

#----------------AP - DATOTELEFONO2-----------------#
lista_ap = []

for i in celular_2:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista_ap.append(i)

    else:
        i = i.replace('.0', '')
        lista_ap.append(i)

dicc_columnas = {
    'DATOTELEFONO2' : lista_ap,
}

df_ap = pd.DataFrame(dicc_columnas)

#----------------AQ, AR - DESCRIPCIONTELEFONO3, TIPOTELEFONO3-----------------#
tel_fijo = df['FIJO_1'] #AQ

lista = []

for i in tel_fijo:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista.append(i)

    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc_columnas = {
    'DESCRIPCIONTELEFONO3' : lista,
    'TIPOTELEFONO3' : lista
}

df_aqar = pd.DataFrame(dicc_columnas)

#----------------AS - DATOTELEFONO3-----------------#
lista_as = []

for i in tel_fijo:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista_as.append(i)

    else:
        num = i.replace('.0', '')
        lista_as.append(num)

dicc_columnas = {
    'DATOTELEFONO3' : lista_as,
}

df_as = pd.DataFrame(dicc_columnas)

#----------------AT, AU - DESCRIPCIONTELEFONO4, TIPOTELEFONO4-----------------#
tel_fijo_2 = df['FIJO_2'] #AR

lista = []

for i in tel_fijo_2:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista.append(i)

    else:
        i = i.replace(i, 'CELULAR')
        lista.append(i)

dicc_columnas = {
    'DESCRIPCIONTELEFONO4' : lista,
    'TIPOTELEFONO4' : lista
}

df_atau = pd.DataFrame(dicc_columnas)

#----------------AV - DATOTELEFONO4-----------------#
lista_av = []

for i in tel_fijo_2:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista_av.append(i)

    else:
        num = i.replace('.0', '')
        lista_av.append(num)

dicc_columnas = {
    'DATOTELEFONO4' : lista_av,
}

df_av = pd.DataFrame(dicc_columnas)

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

dicc = {
    'TELEFONOSADICIONALES' : lista_nueva
}

df_aw = pd.DataFrame(dicc)

#----------------AX, AY, AZ, BA, BB-----------------#
dicc_columnas = {
    'FECHAULTIMOPAGO' : lista_vacia,
    'CAPITALADEUDADO(SIST.FRANCES)' : lista_vacia,
    'NUMEROCONVENIO' : lista_vacia,
    'IMPORTETERCERVENCIMIENTO' : lista_vacia,
    'DATOSGESTION' : lista_vacia,
}

df_axbb = pd.DataFrame(dicc_columnas)

#----------------EXPRESIONES REGULARES EMAILS-----------------#
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
    for x in contenido: 

        if x == 'nan':
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

#----------------BC - EMAIL-----------------#
mail_1 = df_mails['MAIL_1'] #AV
lista_bc = []

for i in mail_1:
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista_bc.append(i)

    else:
        lista_bc.append(i)

dicc_bc = {
    'EMAIL' : lista_bc,
}

df_bc = pd.DataFrame(dicc_bc)

#----------------BD - ANEXO1-----------------#
mail_2 = df_mails['MAIL_2'] #AW
lista_bd = []

for i in mail_2: 
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista_bd.append(i)

    else:
        cadena = f'OPERACION|MAIL 2|{i}'
        lista_bd.append(cadena)

dicc_bd = {
    'ANEXO1' : lista_bd
}

df_bd = pd.DataFrame(dicc_bd)

#----------------BE - ANEXO2-----------------#
mail_3 = df_mails['MAIL_3'] #AW
lista_be = []

for i in mail_3: 
    i = str(i)
    if len(i) == 3:
        i = np.nan 
        lista_be.append(i)

    else:
        cadena = f'OPERACION|MAIL 3|{i}'
        lista_be.append(cadena)

dicc_be = {
    'ANEXO2' : lista_be
}

df_be = pd.DataFrame(dicc_be)

#----------------BF - ANEXO3-----------------#
dni = df['NUMERO_DOCUMENTO'] #P
lista_bf = []

for i in range(0, registros): 
    if i == np.nan:
        i = np.nan 
        lista_bf.append(i)

    else:
        cadena = f'OPERACION|DNI|{dni[i]}'
        lista_bf.append(cadena)

dicc_bf = {
    'ANEXO3' : lista_bf
}

df_bf = pd.DataFrame(dicc_bf)

#----------------BG - ANEXO4-----------------#
estado = df['ESTADO_CLIENTE'] #Q
lista_bg = []

for i in range(0, registros): 
    if i == np.nan:
        i = np.nan 
        lista_bg.append(i)

    else:
        cadena = f'OPERACION|ESTADO CLIENTE|{estado[i]}'
        lista_bg.append(cadena)

dicc_bg = {
    'ANEXO4' : lista_bg
}

df_bg = pd.DataFrame(dicc_bg)

#----------------BH - ANEXO5-----------------#
vto = df['VTO_ULTIMA_FACTURA'] #G
lista_bh = []

for i in range(0, registros): 
    if i == np.nan:
        i = np.nan 
        lista_bh.append(i)

    else:
        fecha = vto[i]
        fecha = str(fecha)
        
        año = fecha[0:4]
        mes = fecha[5:7]
        dia = fecha[8:10]

        cadena = f'OPERACION|VTO ULTIMA FC|{dia}_{mes}_{año}'
        lista_bh.append(cadena)

dicc = {
    'ANEXO5' : lista_bh
}

df_bh = pd.DataFrame(dicc)

#----------------BI - ANEXO6-----------------#
estado_susp = df['ESTADO_SUS'] #R
lista_bi = []

for i in range(0, registros): 
    if i == np.nan:
        i = np.nan 
        lista_bi.append(i)

    else:
        cadena = f'OPERACION|ESTADO SUSP|{estado_susp[i]}'
        lista_bi.append(cadena)

dicc = {
    'ANEXO6' : lista_bi
}

df_bi = pd.DataFrame(dicc)

#----------------BJ - ANEXO7-----------------#
facturas = df['CANTIDAD_FACTURAS_VENCIDAS'] #L
lista_bj = []

for i in range(0, registros): 
    if i == np.nan:
        i = np.nan 
        lista_bj.append(i)

    else:
        cadena = f'OPERACION|FACTURAS VENCIDAS|{facturas[i]}'
        lista_bj.append(cadena)

dicc = {
    'ANEXO7' : lista_bj
}

df_bj = pd.DataFrame(dicc)

#----------------BK - ANEXO8-----------------#
asignacion = df['DIAS_ASIGNACION'] #BR
asignacion = asignacion.fillna(0)

lista = []
for i in asignacion: 
    i = int(i)
    if (i == 0):
        lista.append(f'OPERACION|ASIGNACION|NUEVO')
    elif (i >= 1 and i <= 60): 
        lista.append(f'OPERACION|ASIGNACION|STOCK')
    elif i > 60:
        lista.append(f'OPERACION|ASIGNACION|NUEVO')

dicc = {
    'ANEXO8' : lista
}
df_bk = pd.DataFrame(dicc)

#----------------BL - IDCOMPANIA-----------------#
lista_52 = []
for i in range(0, registros):
    lista_52.append(52)

dicc_bl = {
    'IDCOMPANIA' : lista_52,
}

df_bl = pd.DataFrame(dicc_bl)

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
df_concat.to_excel('ESTUDIOALTAS0511.xlsx', index=False)
print('Archivo creado correctamente')