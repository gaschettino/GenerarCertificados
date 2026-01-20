# Generador de Certificados

Aplicación web desarrollada con Streamlit para generar certificados personalizados en PDF a partir de:

 - Certificado molde en PowerPoint (PPTX)

- Listado de personas en Excel (XLSX)

La aplicación completa automáticamente el Nombre, apellido y opcionalmente el DNI de cada persona, permitiendo configurar fuente, tamaño y color del texto desde la interfaz.


## Selección de archivos

### Requisitos del Template (.pptx)

El archivo PowerPoint debe tener exactamente el valor "Nombre y apellido" donde deba aparecer el identificador de la persona a recibir el certificado.
Opcionalmente se puede agregar el DNI, en ese caso debe contener el texto "Numero de DNI" donde se quiera hacer el reemplazo.

Es importante respetar mayúsculas, minúsculas y espacios.

La fuente, tamaño y color definidos en el PPTX original serán reemplazados por los valores configurados en la app



### Requisitos del archivo Excel (XLSX)


Debe contener obligatoriamente las columnas: Nombre, Apellido y opcionalmente la columna DNI.



## Configuración visual

Desde la interfaz se puede configurar, de forma independiente la fuente (cantidad limitada de opciones) y el color, tanto colores predeterminados como RGB y HEX para: Nombre y Apellido y DNI

Para elegir colores personalizados se puede usar: https://htmlcolorcodes.com/


### Nota

El PDF final utiliza las fuentes disponibles en el servidor donde corre la app.
La vista previa no garantiza una coincidencia exacta al 100%.

## Generación de certificados

Al presionar “Generar certificados” se genera un certificado por cada fila del Excel. Cada certificado se guarda como PDF
Todos los PDFs se comprimen juntos en un archivo ZIP

Este proceso puede tardar unos minutos, dependiendo la cantidad de certificados a realizar.

## Descarga de certificados

Una vez finalizada la generación se habilita la descarga directa desde la app




---- 

Ante cualquier duda o inconveniente, contactar a: gaschettino@garrahan.gov.ar
