# Edemsa

El código **edemsa.py** consiste en automatizar la tarea de análisis y procesamiento de columnas a partir del excel "asigna" originando un nuevo archivo de "altas" para luego ser subido al sistema de LOAN.  

## Librerías Usadas:  
- **Pandas** &#8594; Tratamiento de los datos
- **Numpy** &#8594; Para tratar valores nulos
- **Re** &#8594; Evalúa y filtra a través de expresiones regulares los emails erróneos  

## Análisis por columna 
Cada columna va a estar dividida en secciones por comentarios. Por ejemplo la primer columna:
"#---A - NUMEROOPERACION---#"  
Cada comentario inicia con la letra correspondiente a la columna del archivo excel resultante  

Cada Sección finaliza con el DataFrame correspondiente de esa sección, con la nomenclatura "df_a" para la columna A  

**obs:** En las columnas G,H,I, que son columnas vacías, se tratan en una misma sección, el nombre del DataFrame se forma a partir de la primer letra y la última. Siendo en este caso: "df_gi"

Al final del archivo se concatenan todos estos df_ con su letras correspondientes para luego obtener el archivo excel final que va a ser subido al sistema de LOAN.

## Registro - Información mostrada
Al ejecutar el código, se muestra un informe sobre: 
- Si los nombres de las columnas que trae el excel a analizar son correctas, se muestra "Columnas válidas", si alguna no está incluida o si vino con distinto nombre, muestra el nombre de la columna que falta, en este caso se corrige el nombre o se reclama por la columna en cuestión. 

- La cantidad de filas y columnas del archivo a analizar.

- Cantidad de teléfonos detectados con un largo distinto a 10 (LOAN acepta 10) que luego son eliminados

- Cantidad de mails correctos, incorrectos, valores únicos de los incorrectos y cuantos valores nulos se detecta. 

- Cantidad de nuevas asignaciones 
