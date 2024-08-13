



# Microsoft Word
  
Modulo para trabajar con archivos de texto mediante Microsoft Word. Crea y edita documentos word, trabaja con tablas, da formato a tus textos y mas.   

*Read this in other languages: [English](README.md), [Português](README.pr.md), [Español](README.es.md)*

## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Overview


1. Nuevo documento  
Crea un nuevo documento word

2. Abrir Documento  
Abre un documento de Word.

3. Leer documento  
Extrae texto de documento Word.

4. Obtener párrafos  
Obtener el listado de los parrafos que componen un documento Word en fomato diccionario {numero: texto}.

5. Obtener rango de texto  
Buscar texto en un documento y obtener su rango de posición.

6. Escribir en documento  
Escribe en un documento Word.

7. Copiar y pegar texto  
Copiar texto entre rangos del documento Word y pegarlo en otro documento.

8. Copiar/pegar sin portapapeles  
Copie y pegue texto entre rangos en un documento de Word y péguelo en otro documento sin usar el portapapeles del SO.

9. Copiar y pegar tabla  
Seleccione una tabla de un documento de Word, cópiela y péguela en el mismo documento o en otro.

10. Copiar texto  
Copiar texto al portapapeles entre rangos del documento Word

11. Pegar texto  
Pegar texto del portapapeles al documento Word

12. Contar caracteres  
Contar caracteres de un párrafo específico

13. Agregar tabla  
Agregar tabla en un documento Word.

14. Agregar datos a tabla  
Este comando permite agregar datos a una tabla. Es necesario que la tabla ya exista en el documento y que los datos propocionados sean del tamaño de la tabla.

15. Agregar imagen a tabla  
Este comando permite agregar una imagen a una tabla. Es necesario que la tabla ya exista en el documento.

16. Leer tablas  
Extrae los datos de las tablas en el documento

17. Editar tabla  
Editar tabla de un documento Word.

18. Actualizar campos vinculados  
Actualizar campos vinculados (ej. hoja de cálculo de Excel)

19. Insertar página  
Inserta una nueva página al documento

20. Agregar imagen  
Agrega una imagen al documento

21. Convertir a PDF  
Convierte documento Word a PDF.

22. Buscar Texto en párrafo  
Busca el párrafo donde se encuentra el texto indicado.

23. Contar párrafos  
Cuenta la cantidad de párrafos del documento. Incluye los campos de tablas.

24. Remplazar texto en párrafo  
Remplaza el texto de un párrafo.

25. Borrar párrafo  
Borra un párrafo del documento. Si se incluyen tablas, debe utilizarse el comando Buscar texto en párrafo para ubicar el que se desea eliminar. Retorna el texto eliminado.

26. Agregar texto a un bookmark  
Agregar texto a un bookmark.

27. Guardar documento  
Guarda el documento Word abierto

28. Cerrar documento  
Cierra el documento que se está ejecutando

29. Escribir en párrafo  
Escribir texto en un párrafo seleccionado. El contenido del párrafo será reemplazado por el texto.  



### Changes
Thu Jul 21 01:32:22 2022  Merge branch qa into branch-nico

----
### OS

- windows

### Dependencies
- [**pywin32**](https://pypi.org/project/pywin32/)
### License
  
![MIT](https://camo.githubusercontent.com/107590fac8cbd65071396bb4d04040f76cde5bde/687474703a2f2f696d672e736869656c64732e696f2f3a6c6963656e73652d6d69742d626c75652e7376673f7374796c653d666c61742d737175617265)  
[MIT](http://opensource.org/licenses/mit-license.ph)