# Microsoft Word
  
Modulo para trabajar con Microsoft Word  

*Read this in other languages: [English](Manual_MicrosoftWord.md), [Portugues](Manual_MicrosoftWord.pr.md), [Español](Manual_MicrosoftWord.es.md).*
  
![banner](/docs/imgs/Banner_MicrosoftWord.png)
## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  



## Description of the commands

### New Document
  
Create a new word document
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|

### Open Document
  
Open a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|File|Open the specified document|file.docx|
|Session|File session|Word1|

### Read Document
  
Extract text from a Word document
|Parameters|Description|example|
| --- | --- | --- |
|Result|Store the result in a variable|Variable|
|Session|File session|Word1|
|Add Details|Choose if the stored data will be saved with details like style, alignment, etc.|True|

### Copy and paste text
  
Copy and paste text between ranges in a Word document and paste it in another document
|Parameters|Description|example|
| --- | --- | --- |
|Start of range|Position of the range from where the command starts to copy.|0|
|End of range|Position of the range to which the command copies.|40|
|Session of the archive to copy|File session|Word1|
|File|Choose the document where the copied content is pasted.|file.docx|

### Copy text
  
Copy text to clipboard between ranges in a Word document
|Parameters|Description|example|
| --- | --- | --- |
|Start of range|Position of the range from where the command starts to copy.|0|
|End of range|Position of the range to which the command copies.|40|
|Session|File session|Word1|

### Paste text
  
Paste text from clipboard in a Word document
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|

### Count characters
  
Count characters in a specific paragraph
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Paragraph|Paragraph to count characters|1|
|Result|Store the result in a variable|Variable|

### Add table
  
Add table in a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Number of rows|Number of rows that the table will have|3 |
|Number of columns|Number of columns that the table will have|4 |
|Table style|Microsoft Word default table style|Colorful Grid|
|Session|File session|Word1|
|Border styles|Table border style. Line type and size.|Line type: Single wavy / Line size: 1 1/2 points|

### Read Tables
  
Extract data from the Tables in the document
|Parameters|Description|example|
| --- | --- | --- |
|Table to read|Table number from which the content will be read|1|
|Session|File session|Word1|
|Result|Store the result in a variable|Variable|

### Edit table
  
Edit table from a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Table number|Table number to be edited|1|
|Session|File session|Word1|
|Enter the row number to delete|Optional. The row number entered determines which row will be removed from the table.| |
|Enter the column number to delete|Optional. The column number entered determines which column will be removed from the table.| |
|Insert row|If selected, adds a row to the end of the table|True|
|Insert column|If selected, adds a column to the end of the table|False|
|Column Width|Width in points that each column of the table will have|140|
|Row height|Height in points that each row of the table will have|25|

### Save document
  
Extract text from file.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Save file|Save the file to the specified path|file.docx|

### Write in Document
  
Write in a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Write text|Text to be written on the document|Lorem ipsum |
|Text type|Text type selector that will have the written text.|Subtitle|
|Level|Level that the written text will have.|1-9|
|Font size|Font size that the written text will have.|12|
|Align|Align that the written text will have.|Left|
|Bold|Select whether the text will be bold.|True|
|Italic|Select whether the text will be italic.|True|
|Underline|Select whether the text will be underlined.|False|

### Close Document
  
Close the document that is running
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|

### Add Page
  
Add a new page to the document
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|

### Add Picture
  
Add an image to the document.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Image path|Image path that will be added below the last paragraph|image.jpg|

### Convert to PDF
  
Convert Word document to PDF.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Save file|Path of the file where the PDF will be created|file.pdf|

### Locate Text in Paragraph
  
Locate in which paragraph there is an indicated text.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Text to Search|Text that will be used to locate the paragraph|Hello Word|
|variable name|Store the result in a variable|Variable|

### Count Paragraphs
  
Count the number of paragraphs in the document. Includes table fields.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|variable name|Store the number of paragraphs in a variable|Variable|

### Replace text in paragraph
  
Replace the text of a paragraph.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Text to Search|Text to be searched for in the listed paragraphs.|Hello Word|
|Text to replace|Text to be replaced|Hello Word|
|Paragraph numbers|Paragraphs where the specified text will be searched|Comma separated ',' example: 1,2|

### Delete paragraph
  
Delete paragraph from the document.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Paragraph number|Paragraph number to be deleted|1|
|Variable name where the deleted paragraph will be saved|Variable where the text that included the deleted paragraph will be saved|Variable|

### Add text at the end of bookmark
  
Add text at the end of bookmark.
|Parameters|Description|example|
| --- | --- | --- |
|Session|File session|Word1|
|Text to add|Text that will be added to the chosen bookmark.|Hello Word|
|Bookmark Name|Name of the bookmark where the text will be added.|Bookmark 1|
