



# Microsoft Word
  
Modulo para trabajar con archivos de texto mediante Microsoft Word. Crea y edita documentos word, trabaja con tablas, da formato a tus textos y mas.   

*Read this in other languages: [English](Manual_MicrosoftWord.md), [Português](Manual_MicrosoftWord.pr.md), [Español](Manual_MicrosoftWord.es.md)*
  
![banner](imgs/Banner_MicrosoftWord.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Descripción de los comandos

### Nuevo documento
  
Crea un nuevo documento word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|

### Abrir Documento
  
Abre un documento de Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo|Abre el documento especificado|archivo.docx|
|Abrir sin alertas|Si se marca esta opción, no se mostraran alertas al abrir un archivo.|True|
|Sesión|Sesión del archivo|Word1|

### Leer documento
  
Extrae texto de documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Resultado|Almacena el resultado en una variable|Variable|
|Sesión|Sesión del archivo|Word1|
|Agregar Detalles|Escoje si los datos almacenados se guardarán con detalles como estilo, alineación, etc.|True|

### Obtener párrafos
  
Obtener el listado de los parrafos que componen un documento Word en fomato diccionario {numero: texto}.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Obtener rango|Obtener listado de párrafos con su rango.|True|
|Resultado|Almacena el resultado en una variable|Variable|

### Obtener rango de texto
  
Buscar texto en un documento y obtener su rango de posición.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto a encontrar|Texto a buscar en el documento para obtener el rango en que se ubica.|Hola|
|Sesión|Sesión del archivo|Word1|
|Resultado|Almacena el resultado en una variable|Variable|

### Escribir en documento
  
Escribe en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Escriba texto|Texto que se escribirá en el documento|Lorem ipsum |
|Número de párrafo - Opcional|Número de párrafo de referencia para insertar el texto|1|
|Metodo de inserción - Opcional|Método a utilizar para insertar el nuevo texto||
|Tipo de texto|Selector del tipo de texto que tendrá el texto escrito.|Subtitle|
|Nivel|Nivel que tendrá el texto escrito.|1-9|
|Tamaño de fuente|Tamaño de fuente que tendrá el texto escrito.|12|
|Nombre de fuente|Nombre de la fuente que tendrá el texto escrito.|Arial|
|Alineación|Alineación que tendrá el texto escrito.|Left|
|Color de texto|Color que tendrá el texto escrito|Black|
|Negrita|Seleccionar si el texto irá en negrita.|True|
|Cursiva|Seleccionar si el texto irá en cursiva.|True|
|Subrayar|Seleccionar si el texto irá subrayado.|False|

### Copiar y pegar texto
  
Copiar texto entre rangos del documento Word y pegarlo en otro documento.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Inicio del rango|Posición del rango desde donde comienza a copiar el comando.|0|
|Fin del rango|Posición del rango hasta donde copia el comando.|40|
|Método de pegado|Método de pegado del contenido copiado||
|Sesión del archivo a copiar|Sesión del archivo|Word1|
|Archivo|Elige el documento donde se pega el contenido copiado.|archivo.docx|

### Copiar/pegar sin portapapeles
  
Copie y pegue texto entre rangos en un documento de Word y péguelo en otro documento sin usar el portapapeles del SO.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Inicio del rango|Posición del rango desde donde comienza a copiar el comando.|0|
|Fin del rango|Posición del rango hasta donde copia el comando.|40|
|Rango donde pegar|Posición del rango desde donde pegar.|0|
|Sesión del archivo a copiar|Sesión del archivo|Word1|
|Archivo|Elige el documento donde se pega el contenido copiado.|archivo.docx|

### Copiar y pegar tabla
  
Seleccione una tabla de un documento de Word, cópiela y péguela en el mismo documento o en otro.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Tabla a copiar|Número de tabla a copiar|1|
|Rango|Posición del rango donde pegar.|0|
|Método de pegado|Método de pegado del contenido copiado||
|Sesión|Sesión del archivo|Word1|
|Archivo|Elige el documento donde se pega el contenido copiado.|archivo.docx|

### Copiar texto
  
Copiar texto al portapapeles entre rangos del documento Word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Inicio del rango|Posición del rango desde donde comienza a copiar el comando.|0|
|Fin del rango|Posición del rango hasta donde copia el comando.|40|
|Sesión|Sesión del archivo|Word1|

### Pegar texto
  
Pegar texto del portapapeles al documento Word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|

### Contar caracteres
  
Contar caracteres de un párrafo específico
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Párrafo|Párrafo a contar caracteres|1|
|Resultado|Almacena el resultado en una variable|Variable|

### Agregar tabla
  
Agregar tabla en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de filas|Numero de filas que tendrá la tabla|3 |
|Numero de columnas|Numero de columnas que tendrá la tabla|4 |
|Estilo de tabla|Estilo de tabla predeterminada de Microsoft Word|Colorful Grid|
|Sesión|Sesión del archivo|Word1|
|Estilos del borde|Estilo de los bordes de la tabla. Tipo de línea y tamaño.|Line type: Single wavy / Line size: 1 1/2 points|

### Agregar datos a tabla
  
Este comando permite agregar datos a una tabla. Es necesario que la tabla ya exista en el documento y que los datos propocionados sean del tamaño de la tabla.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Número de tabla|Número de tabla donde se agregarán los datos.|1|
|Datos de la tabla|Datos de la tabla. Debe ser un array de arrays que contengan la información de cada fila.|[ ["Name", "Age", "Gender"], ["John Doe", "32", "Male"], ["Jane Doe", "30", "Female"]]|

### Leer tablas
  
Extrae los datos de las tablas en el documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Tabla a leer|Número de tabla de la cual se leerá el contenido|1|
|Sesión|Sesión del archivo|Word1|
|Resultado|Almacena el resultado en una variable|Variable|

### Editar tabla
  
Editar tabla de un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de tabla|Número de tabla que será editada|1|
|Sesión|Sesión del archivo|Word1|
|Ingrese el numero de fila a eliminar|Opcional. El numero de fila ingresado determina qué fila será eliminada de la tabla.| |
|Ingrese el numero de columna a eliminar|Opcional. El numero de columna ingresado determina qué columna será eliminada de la tabla| |
|Insertar fila|Si se selecciona, agrega una fila al final de la tabla|True|
|Insertar columna|Si se selecciona, agrega una columna al final de la tabla|False|
|Ancho de columna|Ancho en puntos que tendrá cada columna de la tabla|140|
|Alto de fila|Alto en puntos que tendrá cada fila de la tabla|25|

### Actualizar campos vinculados
  
Actualizar campos vinculados (ej. hoja de cálculo de Excel)
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de campo|Número de campo que será actualizado|1|
|Sesión|Sesión del archivo|Word1|

### Insertar página
  
Inserta una nueva página al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|

### Agregar imagen
  
Agrega una imagen al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Ruta de la imagen|Ruta de imagen que sera agregada debajo del ultimo parrafo|imagen.jpg|

### Convertir a PDF
  
Convierte documento Word a PDF.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Guardar archivo|Ruta del archivo donde se creará el PDF|archivo.pdf|

### Buscar Texto en párrafo
  
Busca el párrafo donde se encuentra el texto indicado.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Texto a Buscar|Texto que sera usado para localizar el parrafo|Hola mundo|
|Nombre de la variable|Almacena el resultado en una variable|Variable|

### Contar párrafos
  
Cuenta la cantidad de párrafos del documento. Incluye los campos de tablas.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Nombre de la variable|Almacena el número de párrafos en una variable|Variable|

### Remplazar texto en párrafo
  
Remplaza el texto de un párrafo.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Texto a Buscar|Texto que será buscado en los párrafos listados.|Hola mundo|
|Texto a Remplazar|Texto que sera reemplazado|Hola mundo|
|Lista de párrafo|Parrafos donde buscara el texto especificado|Separados por comas ',' ejemplo: 1,2|

### Borrar párrafo
  
Borra un párrafo del documento. Si se incluyen tablas, debe utilizarse el comando Buscar texto en párrafo para ubicar el que se desea eliminar. Retorna el texto eliminado.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Número de párrafo|Numero de parrafo que será eliminado|1|
|Nombre de la variable donde se guardará el párrafo eliminado|Variable donde se guardará el texto que incluía el párrafo eliminado|Variable|

### Agregar texto a un bookmark
  
Agregar texto a un bookmark.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Texto a agregar|Texto que será agregado al marcador elegido.|Hola mundo|
|Nombre del Marcador|Nombre del marcador donde se agregará el texto.|Marcador 1|

### Guardar documento
  
Guarda el documento Word abierto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Guardar archivo|Guarda el archivo en la ruta especificada|archivo.docx|

### Cerrar documento
  
Cierra el documento que se está ejecutando
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|

### Escribir en párrafo
  
Escribir texto en un párrafo seleccionado. El contenido del párrafo será reemplazado por el texto.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Sesión|Sesión del archivo|Word1|
|Párrafo|Párrafo donde se escribirá el texto|5|
|Escriba texto|Texto que se escribirá en el documento|Lorem ipsum |
