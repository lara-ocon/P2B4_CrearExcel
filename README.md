# Practica2_Bloque4 - Adquisición de Datos - IMAT
# Reporte Ejecutivo Excel Pizzería Maven Pizzas
# Hecho por: Lara Ocón Madrid

Esta práctica consiste en crear un archivo excel que consistirá en un reporte ejecutivo para la pizzería Maven Pizzas con sus datos de 2016. Para realizar dicho excel será necesario instalarse las librerías indicadas por el fichero requiremnts.txt. Una vez instaladas dichos paquetes, procederemos a lanzar el programa "crear_excel.py", que nos creará el fichero "reporte_ejecutivo.xlsx". 

El excel creado, queda dividido en 3 hojas:
1) 'Pizzas a la semana':
    Contiene una tabla con todas las pizzas de cada tipo que se han pedido para cada semana del año 2016. A partir de dicha tabla, hemos creado otra que con la función SUM de excel (y dividiendo entre el número de semanas), calcula la media que se pide de cada pizza en todo el año. Gracias a esta ultima tabla, hemos generado un diagrama de sectores con la media de cada pizza. Además, hemos generado dos gráficos de barras: "Evolución pedidos thai_ckn" que muestra la evolución temporal de la cantidad de esta pizza que se pide a lo largo del año (por semana) y "Evolución total pedidos" que muestra el total de pizzas pedidas para cada semana a lo largo del año.

2) 'Ingredientes por semana':
    Contiene una tabla con la cantidad que se ha usado de cada tipo de ingrediente para cada semana. A partir de dicha tabla, hemos generado otra que nos muestra la media de cada ingrediente y la predicción (empleando funciones y operadores de excel como SUM). Una vez tenemos esto, hemos generado un gráfico de barras con la predicción de los ingredientes. Por último, nos hemos centrado en un ingrediente 'Garlic' y hemos generado una tabla con la cantidad de Garlic empleada cada semana y una segunda columna con la resta de la predicción y dicha cantidad usada. Gracias a esta tabla, hemos hecho un lineplot con la diferencia prediccion-realidad y hemos podido comprobar que con nuestra predicción nunca faltará Garlic (ya que la diferencia nunca vale menos que 0).

3) 'Reporte ejecutivo':
    Esta última hoja, he decidido centrarla en las ganancias de la pizzería. Para ello, he insertado dos imagenes que he creado en 'creacion_imagenes.ipynb'. Estas imagenes son: 1) Ganancias por semana: muestra un lineplot con las ganancias para cada semana del año. 2) Ganancias por mes: muestra un lineplot con las ganancias para cada mes del año.

Para llevar a cabo la práctica, hemos aprovechado el trabajo de la práctica 2 bloque 2 (pizzas 2016) exportando tanto los dataframes creados a csv (encontrados dentro de la carpeta ficheros), como guardando las gráficas que hemos ido creando a partir de estos dataframes en png's (encontrados dentro de la carpeta imagenes). Para ver como se han creado tanto las imagenes como los ficheros, todo se encuentra del del notebook "creacion_imagenes.ipynb".