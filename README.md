# "Aramex" - Webscraper para el WSC.

**Aramex** es un web scraper pensado para extraer información del [World Spider Catalog](https://wsc.nmbe.ch/), en especifico, arañas que tengan su distribución en el área de México.

Una vez con los datos obtenidos, permite exportarlos a un archivo CSV o Excel

## Requerimientos

Necesita de las siguiente librerías de python para funcionar:

* [PyQt5](https://pypi.org/project/PyQt5/)
* [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/bs4/doc/#)
* [Requests](https://docs.python-requests.org/en/master/)
* [html2text](https://pypi.org/project/html2text/)
* [xlsxwriter](https://pypi.org/project/XlsxWriter/)

## Uso

Una vez las dependencias estén instaladas, solo hace falta ejecutar el script `Aramex.py`. Este desplegara una venta, donde solo hay que darle en el botón "Buscar arañas en México".

Una vez la barra este llena, se puede guardar el resultado como un archivo en CSV o Excel. 

Cada fila tiene las siguientes columnas:

* Especie
* Autor 
* Familia 
* Genero 
* Sexos descritos 
* Distribución descrita
* [LSID](https://en.wikipedia.org/wiki/LSID)
* Dirección web en el World Spider Catalog.

## Capturas de pantalla

## Windows
![imagen](/capturas/aramex_1.png)

## Linux

![imagen](/capturas/aramex_2.png)
