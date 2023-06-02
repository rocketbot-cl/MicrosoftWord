



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

5. Escribir en documento  
Escribe en un documento Word.

6. Copiar y pegar texto  
Copiar texto entre rangos del documento Word y pegarlo en otro documento.

7. Copiar texto  
Copiar texto al portapapeles entre rangos del documento Word

8. Pegar texto  
Pegar texto del portapapeles al documento Word

9. Contar caracteres  
Contar caracteres de un párrafo específico

10. Agregar tabla  
Agregar tabla en un documento Word.

11. Leer tablas  
Extrae los datos de las tablas en el documento

12. Editar tabla  
Editar tabla de un documento Word.

13. Actualizar campos vinculados  
Actualizar campos vinculados (ej. hoja de cálculo de Excel)

14. Insertar página  
Inserta una nueva página al documento

15. Agregar imagen  
Agrega una imagen al documento

16. Convertir a PDF  
Convierte documento Word a PDF.

17. Buscar Texto en párrafo  
Busca el párrafo donde se encuentra el texto indicado.

18. Contar párrafos  
Cuenta la cantidad de párrafos del documento. Incluye los campos de tablas.

19. Remplazar texto en párrafo  
Remplaza el texto de un párrafo.

20. Borrar párrafo  
Borra un párrafo del documento. Si se incluyen tablas, debe utilizarse el comando Buscar texto en párrafo para ubicar el que se desea eliminar. Retorna el texto eliminado.

21. Agregar texto a un bookmark  
Agregar texto a un bookmark.

22. Guardar documento  
Guarda el documento Word abierto

23. Cerrar documento  
Cierra el documento que se está ejecutando  



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