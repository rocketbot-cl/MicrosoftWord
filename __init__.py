# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
import os
import sys
import re
import time

base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'MicrosoftWord' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)

# Import local libraries
import win32com.client

global mod_microsoft_word
SESSION_DEFAULT = "default"
try:
    if not mod_microsoft_word: #type:ignore
        mod_microsoft_word = {SESSION_DEFAULT: {}}
except NameError:
    mod_microsoft_word = {SESSION_DEFAULT: {}}


module = GetParams("module")
session = GetParams("session")
if not session:
    session = SESSION_DEFAULT
    

try:
    ms_word = mod_microsoft_word[session].get("app", None)
    word_document = mod_microsoft_word[session].get("doc", None)
except:
    pass

def alignments(WdParagraphAlignment):
    return ["Left", "Center", "Rigth", "Justify"][WdParagraphAlignment]


WdBuiltinStyle = {
    "paragraph": -1,
    "heading1": -2,
    "heading2": -3,
    "heading3": -4,
    "heading4": -5,
    "heading5": -6,
    "heading6": -7,
    "heading7": -8,
    "heading8": -9,
    "heading9": -10,
    "caption": -35,
    "bullet1": -49,
    "number1": -50,
    "bullet2": -55,
    "bullet3": -56,
    "bullet4": -57,
    "bullet5": -58,
    "number2": -59,
    "number3": -60,
    "number4": -61,
    "number5": -62,
    "title": -63,
    "subtitle": -75,
    "colorful_grid": -172,
    "colorful_list": -171,
    "colorful_shading": -170,
    "dark_list": -169,
    "light_grid": -161,
    "light_grid_accent_1": -175,
    "light_list": -160,
    "light_list_accent_1": -174,
    "light_shading": -159,
    "light_shading_accent_1": -173,
    "medium_grid_1": -166,
    "medium_grid_2": -167,
    "medium_grid_3": -168,
    "medium_list_1": -164,
    "medium_list_accent_1": -178,
    "medium_list_2": -165,
    "medium_shading_1": -162,
    "medium_shading_1_accent_1": -176,
    "medium_shading_2": -163,
    "medium_shading_2_accent_1": -177,
    "quote": -181,
    "intense_quote": -182,
    "book": -265
}

WdLineWidth = {
    "25_point": 2,
    "50_point": 4,
    "75_point": 6,
    "100_point": 8,
    "150_points": 12,
    "225_points": 18,
    "300_points": 24,
    "450_points": 36,
    "600_points": 48
}

WdBorderType = {
    "border_top": -1,
    "border_left": -2,
    "border_bottom": -3,
    "border_right": -4,
    "border_horizontal": -5,
    "border_vertical": -6
}

WdLineStyle = {
    "dash_dot": 5,
    "dash_dot_dot": 6,
    "dash_dot_stroked": 20,
    "dash_large_gap": 4,
    "dash_small_gap": 3,
    "dot": 2,
    "double": 7,
    "double_wavy": 19,
    "emboss_3d": 21,
    "engrave_3d": 22,
    "inset": 24,
    "none": 0,
    "outset": 23,
    "single": 1,
    "single_wavy": 18,
    "thick_thin_large_gap": 16,
    "thick_thin_med_gap": 13,
    "thick_thin_small_gap": 10,
    "thin_thick_large_gap": 15,
    "thin_thick_med_gap": 12,
    "thin_thick_small_gap": 9,
    "thin_thick_thin_large_gap": 17,
    "thin_thick_thin_med_gap": 14,
    "thin_thick_thin_small_gap": 11,
    "triple": 8
}

WdColorIndex = {
    "black": 1,
    "blue": 2,
    "brightGreen": 4,
    "darkBlue": 9,
    "darkRed": 13,
    "darkYellow": 14,
    "gray25": 16,
    "gray50": 15,
    "green": 11,
    "pink": 5,
    "red": 6,
    "teal": 10,
    "turquoise": 3,
    "white": 8,
    "yellow": 7
}



if module == "new":
    try:
        ms_word = win32com.client.DispatchEx("Word.Application")
        word_document = ms_word.Documents.Add()
        ms_word.Visible = True
        mod_microsoft_word[session] = {
            "app": ms_word, 
            "doc": word_document
        }

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "open":
    path = GetParams("path")
    no_alerts = GetParams("no_alerts")
    try:
        
        if no_alerts and eval(no_alerts):
            alerts = False
        else:
            alerts = True
        
        path = path.replace("/", os.sep)
        if os.path.exists(path):
            try:
                ms_word = win32com.client.DispatchEx("Word.Application")
                ms_word.DisplayAlerts = alerts
                word_document = ms_word.Documents.Open(path)
                ms_word.Visible = True
                mod_microsoft_word[session] = {
                    "app": ms_word, 
                    "doc": word_document
                }
            except:
                os.startfile(path)
                time.sleep(3)
                ms_word = win32com.client.GetObject(None, "Word.Application")
                ms_word.DisplayAlerts = alerts
                time.sleep(1)
                word_document = ms_word.Application.Documents.Open(path)
                mod_microsoft_word[session] = {
                    "app": ms_word.Application, 
                    "doc": word_document
                }
        else:
            raise Exception ('File does not exist...')
        
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "read":
    # Rocketbot params
    result = GetParams("result")
    details = GetParams("details")

    try:
        text = []
        word_document = mod_microsoft_word[session]["doc"]
        paragraphs = word_document.Paragraphs
        for paragraph in paragraphs:
            range_ = paragraph.Range
            font = range_.Font
            if details == "True":
                text.append({
                    "text": range_.Text,
                    "style": str(paragraph.Style),
                    "alignment": alignments(paragraph.Alignment),
                    "font": {
                        "name": font.Name,
                        "size": int(font.Size),
                        "bold": bool(font.Bold),
                        "italic": bool(font.Italic),
                        "underline": bool(font.Underline)
                    }
                })
            else:
                text.append("" + range_.Text)

        if result:
            SetVar(result, text)
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "addTable":
        
    rows = GetParams("rows")
    cols = GetParams("cols")
    tableStyle = GetParams("tableStyle")
    iframe = eval(GetParams("iframe"))
    lineType = iframe.get('lineType')
    lineSize = iframe.get('lineSize')
    
        
    try:

        word_document.Paragraphs.Add()
        paragraph = word_document.Paragraphs.Last
        range_ = paragraph.Range
        
        word_document.Tables.Add(Range=range_, NumRows=rows, NumColumns=cols)
        
        lastTable = word_document.Tables[str(word_document.Tables.count)]
        
        
        if tableStyle in WdBuiltinStyle:
            lastTable.Style = WdBuiltinStyle[tableStyle]
            
        try:
            for border in WdBorderType:
                lastTable.Borders(WdBorderType[border]).LineStyle = WdLineStyle[lineType]
                lastTable.Borders(WdBorderType[border]).LineWidth = WdLineWidth[lineSize]
        except Exception as e:
            print("\x1B[" + "31;40mError\x1B[" + "0m")
            print('El estilo de linea elegido no acepta ese tamano de linea. Por favor, escoja un tamano valido')
            PrintException()
            raise e
            



    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
    
  
if module == "copyPasteText":
    
    startRange = GetParams("startRange")
    endRange = GetParams("endRange")
    path = GetParams("path")
    path = path.replace("/", os.sep)

    try:

        mod_microsoft_word[session] = {
            "app": ms_word, 
            "doc": word_document
        }
        
        word_document = mod_microsoft_word[session]["doc"]
        
        range_ = word_document.Range(Start=startRange, End=endRange)
        range_.Copy()
        
        

        word_documentPaste = ms_word.Documents.Open(path)
        ms_word.Visible = True

        # word_documentPaste.Paragraphs.Last.Range.PasteAndFormat(Type=16)
        word_documentPaste.Paragraphs.Last.Range.PasteSpecial()
        word_documentPaste.Paragraphs.Add()
        
        word_documentPaste.Save()
        word_documentPaste.Close(SaveChanges=0)
        

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "copyPasteText_2":
    
    startRangeSource = GetParams("startRangeSource")
    endRangeSource = GetParams("endRangeSource")
    startRangeTarget = GetParams("startRangeTarget")
    path = GetParams("path")
    path = path.replace("/", os.sep)

    try:
        target = ms_word.Documents.Open(path)

        source_range = word_document.Range(Start=startRangeSource, End=endRangeSource)
        target_range = target.Range(Start=startRangeTarget)
        target_range.FormattedText = source_range.FormattedText

        target.Save()
        target.Close(SaveChanges=0)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e


if module == "copyPasteTable":
    try:
        result = GetParams("result")
        tableToCopy = GetParams("tableToCopy")
        startRange = GetParams("startRange")
        path = GetParams("path")
        
        startRange = int(startRange) if startRange else 0

        table = word_document.tables(tableToCopy).Range
        table.Copy()
        
        if path:
            path = path.replace("/", os.sep)
            word_documentPaste = ms_word.Documents.Open(path)
            ms_word.Visible = True
            paste = word_documentPaste.Range(Start=startRange).PasteAndFormat(Type=0)
            
            paste_table = word_documentPaste.Range(Start=startRange).Tables[0]
            paste_table.Rows.Alignment = 1

            word_documentPaste.Save()
            word_documentPaste.Close(SaveChanges=0)
        else:
            word_document.Range(Start=startRange).PasteAndFormat(Type=0)
        
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "copyText":
    
    startRange = GetParams("startRange")
    endRange = GetParams("endRange")
    
    try:
        mod_microsoft_word[session] = {
            "app": ms_word, 
            "doc": word_document
        }

        word_document = mod_microsoft_word[session]["doc"]
        
        range_ = word_document.Range(Start=startRange, End=endRange)
        range_.Copy()
    
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "pasteText":
    
    try:
        mod_microsoft_word[session] = {
            "app": ms_word, 
            "doc": word_document
        }
    
        word_document = mod_microsoft_word[session]["doc"]
        word_document.Paragraphs.Last.Range.PasteAndFormat(Type=16)
        word_document.Paragraphs.Add()

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e



if module == "countCharacters":
    
    numParagraph = GetParams("numParagraph")
    result = GetParams("result")
    
    try:
        word_document = mod_microsoft_word[session].get("doc", None)
        count = word_document.Paragraphs.count
    
        paragraph =  word_document.Paragraphs(numParagraph)
        range_ = paragraph.Range
        characters = range_.Characters.Count - 1
        
        SetVar(result, characters)
        
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
    
    
    
if module == "readTable":
    try:

        result = GetParams("result")
        tableToRead = GetParams("tableToRead")
        
        tablesDoc = []

        if tableToRead:
            table_ = []
            table = word_document.tables(tableToRead)
            for row in table.rows:
                array_row = []
                for cell in row.cells:
                    information = cell.range.text

                    text1 = re.sub(r"[\x07\xa0]", "", information.replace("\r", "\n").replace("\x0b", "\n"))

                    if text1.startswith("\n") and len(text1) >= 1:
                        text1 = text1[1:]

                    if text1.endswith("\n") and len(text1) >= 1:
                        text1 = text1[:-1]

                    array_row.append(text1)
                table_.append(array_row)
            if result:
                SetVar(result, table_)
        else:
            for table in word_document.tables:
                table_ = []
                for row in table.rows:
                    array_row = []
                    for cell in row.cells:
                        information = cell.range.text
                        
                        text1 = re.sub(r"[\x07\xa0]", "", information.replace("\r", "\n").replace("\x0b", "\n"))

                        if text1.startswith("\n") and len(text1) >= 1:
                            text1 = text1[1:]

                        if text1.endswith("\n") and len(text1) >= 1:
                            text1 = text1[:-1]

                        
                        array_row.append(text1)
                    table_.append(array_row)
                tablesDoc.append(table_)
            if result:
                SetVar(result, tablesDoc)
            
        
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
    
    
if module == "editTable":
    
    numTable = GetParams("numTable")
    session = GetParams("session")
    deleteRow = GetParams("deleteRow")
    deleteColumn = GetParams("deleteColumn")
    addRow = GetParams("addRow")
    addColumn = GetParams("addColumn")
    columnWidth = GetParams("columnWidth")
    rowHeight = GetParams("rowHeight")
    
    try:
        
        table_ = word_document.Tables[numTable]
    
        if deleteRow:
            table_.Rows(deleteRow).Delete()
        if deleteColumn:
            table_.Columns(deleteColumn).Delete()
        if columnWidth:
            table_.Columns.SetWidth(columnWidth, 0)
        if rowHeight:
            table_.Rows.SetHeight(rowHeight, 2)
        if addRow == "True":
            table_.Rows.Add()
        if addColumn == "True":
            table_.Columns.Add()
    
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
    
if module == "updateExcelChart":
    
    numTable = GetParams("numTable")
    session = GetParams("session")    
    
    try:
        
        if numTable:
            table_ = word_document.Fields(numTable).Update()
        else:
            for table in word_document.Fields:
                table.Update()

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
    
    
if module == "deleteParagraph":
    
    try:
        numParagraph = GetParams("numParagraph")
        var = GetParams("variable")
        word_document = mod_microsoft_word[session].get("doc", None)
        
        count = word_document.Paragraphs.count
        
        if int(numParagraph) < count:
            paragraph =  word_document.Paragraphs(numParagraph)
            range_ = paragraph.Range
            SetVar(var, range_)
            range_.Delete()
        else:
            range_ = word_document.Paragraphs.Last.Range
            SetVar(var, range_)
            range_.Delete()
        
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
     

if module == "addTextBookmark":

    import copy

    bookmark_searched = GetParams("bookmark")
    text = GetParams("text")
    #clean = GetParams("Clean")
    #print(clean)

    try:
        """tmp_doc = Document()
        # Generate content in tmp_doc document
        tmp_doc.add_paragraph(text)
        # Reference the tmp_doc XML content
        tmp_doc_body = tmp_doc._element.body

        ele = document._element[0]
        bookmarks_list = ele.findall('.//' + qn('w:bookmarkStart'))
        for bookmark in bookmarks_list:
            # print(bookmark)
            name = bookmark.get(qn('w:name'))
            if name == bookmark_searched:
                par = bookmark.getparent()

                if clean:
                    next_element = bookmark.getnext()
                    if not isinstance(next_element, docx.oxml.CT_R):
                        next_element = next_element.getnext()
                    t = next_element.findall('.//' + qn('w:t'))
                    if len(t) == 1:
                        t[0].text = text
                elif isinstance(par, docx.oxml.CT_P):
                    bookmark_par_parent = par.getparent()
                    index = bookmark_par_parent.index(par)
                    for child in tmp_doc_body:
                        bookmark_par_parent.insert(index, child)
                        index = index + 1

                break
            else:
                name = None

        if not name:
            raise Exception("Bookmark not found")"""
        if word_document.Bookmarks.Exists(bookmark_searched):
            range = word_document.Bookmarks(bookmark_searched).Range
            range.Text = text
            #word_document.Bookmarks.Add(bookmark_searched)
        else:
            pass
    except Exception as e:
        PrintException()
        raise e

if module == "save":

    path = GetParams("path")
    try:
        if path:
            word_document.SaveAs2(path)
        else:
            word_document.Save()
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "to_pdf":
    to = GetParams("to")
    wdFormatPDF = 17
    try:

        word_document.ExportAsFixedFormat(OutputFileName=to, ExportFormat=wdFormatPDF, IncludeDocProps=True)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "getParagraphs":
    result = GetParams("result")
    more = GetParams("more")
    
    paragraphs = {}
    try:
        for i in range(word_document.Paragraphs.count):
            p_range = word_document.Paragraphs(i+1).Range
            if more and eval(more):
                paragraphs["Paragraph "+str(i+1)] = {'text': p_range.Text.replace("\r", ""), 'range': [p_range.Start, p_range.End-1]}
            else:
                paragraphs["Paragraph "+str(i+1)] = p_range.Text.replace("\r", "")
            
        SetVar(result, paragraphs)
    except Exception as e:
        SetVar(result, False)
        PrintException()
        raise e
    
if module == "findRange":
    text = GetParams("text")
    result = GetParams("result")

    try:
        found_range = word_document.Content
        find = found_range.Find
        find.Text = text
        found = find.Execute()
        if found:
            text_range = found_range.Duplicate
            SetVar(result, [text_range.Start, text_range.End])
        else:
            SetVar(result, False)
    except Exception as e:
        SetVar(result, False)
        PrintException()
        raise e
    
if module == "write":

    text = GetParams("text")
    type_ = GetParams("type")
    level = GetParams("level")
    align = GetParams("align")
    color = GetParams("color")
    size = GetParams("size")
    bold = GetParams("bold")
    italic = GetParams("italic")
    underline = GetParams("underline")
    paragraph_num = GetParams("numParagraph")
    insert = GetParams("insert")
    try:     
        if not text:
            text = ''
        text = text.replace("\\n", "\n")
        
        if not type_:
            type_ = "paragraph"
            
        count = word_document.Paragraphs.count
        
        if paragraph_num:
            paragraph_num = int(paragraph_num)
            paragraph_num = count if paragraph_num > count else paragraph_num
            insert = "after" if not insert else insert

            paragraph = word_document.Paragraphs(paragraph_num)
            range_ = paragraph.Range
            if insert == "before":
                range_.InsertBefore(text + "\n")
            if insert == "after":
                range_.InsertAfter(text + "\n")
            if insert == "replace":
                range_.Text = text + "\n"
            font = paragraph.Range.Font
        else:
            word_document.Paragraphs.Add()
            paragraph = word_document.Paragraphs.Last
            range_ = paragraph.Range
            text = text.replace("\\n", "\n")
            range_.Text = text
            font = range_.Font
        
        size = float(size) if size else 12

        if color == None:
            color = "black"

        style = type_ + level if level is not None else ""


        if style in WdBuiltinStyle:
            paragraph.Style = WdBuiltinStyle[style]
        elif (type_ == "number" or type_ == "bullet") and int(level) > 5:
            level = 5
            style = type_ + str(level)
            paragraph.Style = WdBuiltinStyle[style]
        else:
            style = type_
            paragraph.Style = WdBuiltinStyle[style]
            
        font.ColorIndex = WdColorIndex[color]
        paragraph.Alignment = int(align) if align else 0
        font.Size = size
        
        if bold == "True":
            boldInt = -1
        else:
            boldInt = 0
        font.Bold = boldInt
        if italic == "True":
            italicInt = -1
        else:
            italicInt = 0
        font.Italic = italicInt
        if underline == "True":
            underlineInt = -1
        else:
            underlineInt = 0
        font.Underline = underlineInt
            
    except Exception as e:
        PrintException()
        raise e

if module == "close":

    try:
        ms_word = mod_microsoft_word[session]["app"]
        word_document = mod_microsoft_word[session]["doc"]
        word_document.Close(SaveChanges=0)
        ms_word.Quit(SaveChanges=0)
        del mod_microsoft_word[session]

    except Exception as e:
        PrintException()
        raise e

if module == "new_page":
    try:
        word_document.Paragraphs.Add()
        paragraph = word_document.Paragraphs.Last
        paragraph.Range.InsertBreak()
    except Exception as e:
        PrintException()
        raise e

if module == "add_pic":
    img_path = GetParams("img_path")

    try:
        # Only work with \
        img_path = img_path.replace("/", os.sep)

        count = word_document.Paragraphs.Count  # Count number paragraphs
        if count > 1:
            word_document.Paragraphs.Add()

        paragraph = word_document.Paragraphs.Last
        img = paragraph.Range.InlineShapes.AddPicture(FileName=img_path, LinkToFile=False, SaveWithDocument=True)
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e
    
if module == "count_paragraphs":
    number = GetParams("variable")
    try:
        count = word_document.Paragraphs.count
        SetVar(number, count)

    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "search_replace_text":
    text_search = GetParams("text_search")
    text_replace = GetParams("text_replace")
    numParagraphs = GetParams("numParagraphs")
    if text_search == text_replace:
        pass
    else:

        if numParagraphs:
            paragraphList = [int(s) for s in numParagraphs.split(',')]
            for i in paragraphList:
                paragraph = word_document.Paragraphs(i)
                range_ = paragraph.Range
                if text_search in range_.Text:
                    range_.Text = range_.Text.replace(text_search, text_replace)
        else:
            paragraphs = word_document.Paragraphs
            #fullRange = word_document.content
            for paragraph in paragraphs:
                range_ = paragraph.Range

                find_obj = range_.Find
                find_obj.Text = text_search
                find_obj.Replacement.Text = text_replace
                #        Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, 
                #                Forward,  Wrap, Format, ReplaceWith, Replace)
                find_obj.Execute(text_search, False, False, False, False, False, 
                                 True, 1, False, text_replace, 2)

if module == "search_text":
    try:
        text_search = GetParams("text_search")
        whichParagraph = GetParams("variable")
        paragraphList = []
        count = 1
        for paragraph in word_document.Paragraphs:
            range_ = paragraph.Range
            range_.Find.Text = text_search
            if range_.Find.Execute(Forward=True, MatchWholeWord=True):
                paragraphList.append(count)
            count += 1
        SetVar(whichParagraph, paragraphList)
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

#deprecated, kept for compatibility
if module == "write_paragraph":
    paragraph = GetParams("paragraph")
    text = GetParams("text")

    try:
        paragraph = int(paragraph)
        word_document.Paragraphs(paragraph).Range.Text = text

    except Exception as e:
        PrintException()
        raise e

if module == "addDataTable": 
    numTable = int(GetParams("numTable")) - 1
    data = GetParams("data")
    
    try:
        table = word_document.Tables[numTable]
        
        if data:
            data = eval(data)
            for i in range(len(data)):
                for j in range(len(data[0])):
                    table.Cell(i+1, j+1).Range.Text = data[i][j]
                    
        else:
            raise Exception("No data provided")
    
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e