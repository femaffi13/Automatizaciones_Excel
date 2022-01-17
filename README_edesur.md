# Edesur

El código **edesur.py** consiste en automatizar la tarea de análisis y procesamiento de columnas a partir del excel "asigna" originando un nuevo archivo de "altas" para luego ser subido al sistema de LOAN.  
El programa a su vez da un informe general de la cantidad de filas y columnas del archivo a analizar, cantidad y porcentaje de registros sin ningún teléfono y otro informe sobre los mails analizados donde se muestra la cantidad de correctos, incorrectos, valores únicos de los incorrectos y cuantos valores nulos de mails se detecta. 

## Librerías Usadas:  
- **Pandas** &#8594; Tratamiento de los datos
- **Numpy** &#8594; Para tratar valores nulos
- **Datetime** &#8594; Realiza operaciones entre fechas y números (Columna X)
- **Re** &#8594; Evalúa y filtra a través de expresiones regulares los emails erróneos  

## Análisis por columna 
Cada columna va a estar dividida en secciones por comentarios. Por ejemplo la primer columna:
"#---A - NUMEROOPERACION---#"  
Cada comentario inicia con la letra correspondiente a la columna del archivo excel resultante  

Cada Sección finaliza con el DataFrame correspondiente de esa sección, con la nomenclatura "df_a" para la columna A  

**obs:** En las columnas F,G,H,I, que son columnas vacías, se tratan en una misma sección, el nombre del DataFrame se forma a partir de la primer letra y la última. Siendo en este caso: "df_fi"



