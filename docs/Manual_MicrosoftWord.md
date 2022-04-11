



# Microsoft Word
  
Modulo para trabajar con Microsoft Word  
  
![banner](imgs/Banner_MicrosoftWord.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de rocketbot.  



## Descripción de los comandos

### Nuevo documento
  
Crea un nuevo documento word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Abrir Documento
  
Abre un documento de Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo|Abre el documento especificado|archivo.docx|

### Leer documento
  
Extrae texto de documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Resultado|Almacena el resultado en una variable|Variable|
|Agregar Detalles|Agrega informacion del estilo del parrafo||

### Leer tabla
  
Extrae los datos de una tabla
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Resultado|Almacena el resultado en una variable|Variable|

### Escribir en documento
  
Escribe en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Escriba texto|Texto que se escribira en el documento|Lorem ipsum |
|Tipo de texto|Selector del tipo de texto||
|Nivel||1-9|
|Tamaño de fuente||12|
|Alineación|||
|Negrita|||
|Cursiva|||
|Subrayar|||

### Insertar página
  
Inserta una nueva página al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Agregar imagen
  
Agrega una imagen al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta de la imagen|Ruta de imagen que sera agregada debajo del ultimo parrafo|imagen.jpg|

### Convertir a PDF
  
Convierte documento Word a PDF.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo Word|Seleccione el archivo Word a convertir.|archivo.docx|
|Guardar archivo|Especifique la ruta donde guardar el resultado.|archivo.pdf|

### Buscar Texto en párrafo
  
Busca el párrafo donde se encuentra el texto indicado.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto a Buscar|Texto que sera usado para localizar el parrafo|Hola mundo|
|Nombre de la variable|Almacena el resultado en una variable|Varible|

### Contar párrafos
  
Cuenta la cantidad de párrafos del documento.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la variable|Almacena el resultado en una variable|Varible|

### Remplazar texto en párrafo
  
Remplaza el texto de un párrafo.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto a Buscar|Texto que sera reemplazado|Hola mundo|
|Texto a Remplazar|Texto que reemplazara el anterior|Hola mundo|
|Lista de párrafos|Parrafos donde buscara el texto especificado|Separados por comas ',' ejemplo: 1,2|
|Nombre de la variable|Almacena el resultado en una variable|Varible|

### Agregar texto a un bookmark
  
Remplaza el texto de un párrafo.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto a agregar||Hola mundo|
|Nombre del Marcador||Marcador 1|

### Copiar de Excel a Word
  
Copia el rango especificado de un archivo Excel en un Documento Word abierto previamente.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta archivo Excel|Ruta del archivo excel desde el que se copiaran los datos|archivo.xlsx|
|Hoja|Hoja del archivo Excel donde se encuentran los datos|Hoja1|
|Rango de valores a copiar|Rango donde se encuentran los valores que seran pegados en el archivo Word|A1:B2|
|Número de parrafo|Numero del parrafo donde se pegara el contenido del rango copiado|3|

### Alineación de párrafo
  
Alinear párrafo o rango de párrafos a la izquierda, derecha, centro o justificado.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Primer valor del rango|Numero de parrafo inicial en el rango|1|
|Segundo valor del rango|Numero de parrafo final en el rango|4|
|Alineación|||

### Guardar documento
  
Guarda el documento Word abierto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar archivo|Guarda el archivo en la ruta especificada|archivo.docx|

### Cerrar documento
  
Cierra el documento que se está ejecutando
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
