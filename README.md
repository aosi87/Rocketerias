Rocketerias
===========

Programa dedicado a la convergencia de diferentes archivos en uno solo.

###Requerimientos

Para realizar la recoleccion de datos de diferentes tipos de archivos es necesario agregar las siguientes librerias al proyecto.

 1. Actualmente se utiliza la version [**Apache POI 3.10-FINAL**](http://poi.apache.org/).
 
 2. Para este proyecto se esta utilizando la version [**Apache PDFBox 1.8.5**](http://pdfbox.apache.org/). 


###Comentarios

  El funcionamiento basico del programa es obtener diferentes datos de diferentes tipos de archivos y unificarlos en otro archivo para su manejo y control por el usuario. Para ello el prorgama requiere de librerias especificas para cada tipo de archivo, es decir, si queremos leer un archivo de EXCEL de Microsoft Office es necesaria la libreria **Apache POI**, lo mismo ocurre si se requiere leer archivos PDF.
  
  Una vez obtenidos los datos el usuario sera capaz de seleccionar el archivo de salida, asi como en que parte del documento de salida seran ingresados los datos obtenidos anteriormente.
