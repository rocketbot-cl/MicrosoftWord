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
import xlwings as xl
import time
base_path = tmp_global_obj["basepath"]
cur_path = base_path + 'modules' + os.sep + 'MicrosoftWord' + os.sep + 'libs' + os.sep
sys.path.append(cur_path)

# Import local libraries
import win32com.client

module = GetParams("module")
global word_document
global ms_word


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
    "quote": -181,
    "intense_quote": -182,
    "book": -265
}

if module == "new":
    try:
        ms_word = win32com.client.DispatchEx("Word.Application")
        word_document = ms_word.Documents.Add()
        ms_word.Visible = True
        # print("test")
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "open":
    path = GetParams("path")

    try:
        path = path.replace("/", os.sep)
        ms_word = win32com.client.DispatchEx("Word.Application")
        word_document = ms_word.Documents.Open(path)
        ms_word.Visible = True
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "read":
    # Rocketbot params
    result = GetParams("result")
    details = GetParams("details")

    try:
        text = []
        paragraphs = word_document.Paragraphs
        for paragraph in paragraphs:
            range_ = paragraph.Range
            font = range_.Font
            if details:
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
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "readTable":
    try:

        result = GetParams("result")
        tablesDoc = []
        for table in word_document.tables:
            table_ = []
            for row in table.rows:
                array_row = []
                for cell in row.cells:
                    """if len(array_row) > 0:
                        if array_row[-1] != cell.text:
                            array_row.append(cell.text)
                    else:
                        print("test")"""
                    information = cell.range.text
                    information.replace("\r", "").replace("\x07", "")
                    array_row.append(information)
                table_.append(array_row)
            tablesDoc.append(table_)
        if result:
            SetVar(result, tablesDoc)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
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
            word_document.SaveAs2()
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "to_pdf":
    path = GetParams("from")
    to = GetParams("to")
    wdFormatPDF = 17
    try:
        if path:
            ms_word = win32com.client.DispatchEx("Word.Application")
            word_document = ms_word.Documents.Open(path)
        word_document.ExportAsFixedFormat(OutputFileName=to, ExportFormat=wdFormatPDF, IncludeDocProps=True)
        if path:
            word_document.Close()
            ms_word.Quit()
    except Exception as e:
        print("\x1B[" + "31;40mError\x1B[" + "0m")
        PrintException()
        raise e

if module == "write":

    text = GetParams("text")
    type_ = GetParams("type")
    level = GetParams("level")
    align = GetParams("align")
    size = GetParams("size")
    bold = GetParams("bold")
    italic = GetParams("italic")
    underline = GetParams("underline")

    try:
        word_document.Paragraphs.Add()
        paragraph = word_document.Paragraphs.Last
        range_ = paragraph.Range
        range_.Text = text
        font = paragraph.Range.Font

        size = float(size) if size else 12

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

        paragraph.Alignment = int(align) if align else 0
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
    except Exception as e:
        PrintException()
        raise e

if module == "close":

    try:
        word_document.Close()
        ms_word.Quit()
        word_document = None
        ms_word = None
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
        # print(img)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e
if module == "count_paragraphs":
    number = GetParams("variable")
    try:
        count = word_document.Paragraphs.count
        SetVar(number, count)

    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
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
                # print(range_.Find.Text)
                range_.Find.Text = text_search
                range_.Find.Replacement.Text = text_replace
                range_.Find.Execute(Replace=2,Forward=True,MatchWholeWord=True)
                #print(range_.Find.Execute(FindText=text_search, ReplaceWith="text_replace", Replace=2))
                #if text_search in range_.Text:
                    #range_.Text = range_.Text.replace(text_search,text_replace)

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
        # print(paragraphList)
    except Exception as e:
        print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
        PrintException()
        raise e

if module == "copy_excel":
    try:
        path_excel = GetParams("path_excel")
        path_word = GetParams("path_word")
        hoja_excel = GetParams("hoja_excel")
        rango = GetParams("rango")
        numParagraphs = GetParams("numParagraphs")
        password = ""
        path_excel = path_excel.replace("/", os.sep)
    
        if path_excel:
            app = xw.App(add_book=False, visible=False)
            app.api.DisplayAlerts = False
            wb = app.api.Workbooks.Open(path_excel, False, None, None, password, password, IgnoreReadOnlyRecommended=True,
                                        CorruptLoad=2)
        else:
            excel = GetGlobals("excel")
            if excel.actual_id in excel.file_:
                wb = excel.file_[excel.actual_id]['workbook'].api

        sheet = wb.Worksheets(hoja_excel)
        sheet.Range(rango).Select()
        sheet.Range(rango).Copy()
        time.sleep(10)
        if not numParagraphs:
            numParagraphs = word_document.Paragraphs.Count
        myRange = word_document.Paragraphs(numParagraphs).Range
        myRange.Collapse(0)
        myRange.PasteExcelTable(False, False, False)
        try:
            app.quit()
        except:
            pass
    except Exception as e:
       print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
       PrintException()
       app.quit()
       raise e

if module == "adjust_paragraphs":
    type_ = GetParams("type")
    align = GetParams("align")
    size = GetParams("size")
    bold = GetParams("bold")
    italic = GetParams("italic")
    underline = GetParams("underline")
    num_paragraph1 = int(GetParams("num_paragraph1"))
    num_paragraph2 = int(GetParams("num_paragraph2"))+1
    level = GetParams("level")
    rango = range(num_paragraph1,num_paragraph2)
    try:
        for p in rango:
            paragraph = word_document.Paragraphs(p)
            paragraph.Alignment = align
            font = paragraph.Range.Font
            size = float(size) if size else 12
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
            style = type_ + level if level is not None else ""
            if style in WdBuiltinStyle:
                paragraph.Style = WdBuiltinStyle[style]
            elif (type_ == "number" or type_ == "bullet") and int(level) > 5:
                style = type_ + str(level)
                paragraph.Style = WdBuiltinStyle[style]
            
    except Exception as e:
       print("\x1B[" + "31;40mError\u2193\x1B[" + "0m")
       PrintException()
       raise e
    
    
   