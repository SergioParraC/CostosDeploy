Sub copy_to_word()
    'Desactivando las animaciones de Office
    Application.EnableEvents = False

    'Declarando un diccionario con los tipo de elementos calculados
    Dim diccElemento As Object
    Set diccElemento = CreateObject("Scripting.Dictionary")
    
    'Introduciendo datos de los tipos de elementos
    ThisWorkbook.Worksheets("EDS").Select
    For col = 6 To 18
        diccElemento.Add Cells(1, col).Value, Cells(2, col).Value
    Next col
    
    'Modificando documento
    Set wordApp = CreateObject("Word.Application")
    
    'Creando un documento nuevo
    Set mydoc = wordApp.Documents.Add()
    
    'Haciendo visible la app
    wordApp.Visible = True
    Set objSelection = wordApp.Selection
    
    'Modificando margenes
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2.5)
        .BottomMargin = CentimetersToPoints(2.5)
        .LeftMargin = CentimetersToPoints(2.5)
        .RightMargin = CentimetersToPoints(2.5)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(21.59)
        .PageHeight = CentimetersToPoints(27.94)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
    End With
    
    
    'final de EDS
    fin = 500
    'Final de hojas capitulos
    fin2 = 500
    'Recorre la estructura de trabajo
    For i = 1 To fin
        capitulo = ThisWorkbook.Worksheets("EDS").Cells(i, 1).Value
        'Si hay algo en la celda revisa los items
        If Not IsEmpty(capitulo) Then
            'Recupera el nombre del capitulo y lo cambia a primera en mayusculas
            textCapitulo = ThisWorkbook.Worksheets("EDS").Cells(i, 4).Value
            textCapitulo = capitulo & " " & UCase(Left(textCapitulo, 1)) & LCase(Right(textCapitulo, Len(textCapitulo) - 1)) & vbCrLf
            'Copia el titulo del capitulo al documento, le pone el estilo de titulo y mueve el cursor
            objSelection.InsertBreak Type:=wdPageBreak
            mydoc.Content.InsertAfter Text:=textCapitulo
            objSelection.Style = wordApp.ActiveDocument.Styles("Título 1")
            
            'Recorre dentro de los items
            For j = i + 1 To i + 15
                codItem = ThisWorkbook.Worksheets("EDS").Cells(j, 2).Value
                'Toma el valor donde deberia estar el item
                If Not IsEmpty(codItem) Then
                    'Si no esta vacia, reclama los datos restantes para copiar al documento
                    numItem = ThisWorkbook.Worksheets("EDS").Cells(j, 3).Value
                    textItem = ThisWorkbook.Worksheets("EDS").Cells(j, 4).Value
                    'Pega en el documento el capitulo al que pertenece
                    textItem = numItem & " " & textItem
                    mydoc.Content.InsertAfter Text:=textItem
                    objSelection.MoveDown Unit:=wdLine, Count:=4
                    objSelection.Style = wordApp.ActiveDocument.Styles("Título 2")
                    mydoc.Content.InsertAfter Text:=vbCrLf
                    objSelection.MoveDown Unit:=wdLine, Count:=2
                    mydoc.Content.InsertAfter Text:="Cantidades de obra" & vbCrLf
                    objSelection.Style = wordApp.ActiveDocument.Styles("Normal")
                    objSelection.MoveDown Unit:=wdLine, Count:=1
                    'Se va hasta el capitulo donde esta el item
                    ThisWorkbook.Worksheets(capitulo).Select
                    
                    
                    'Busca el item dentro de la hoja del capitulo
                    For k = 5 To fin2
                        'Adquiere el valor de cada celda en la columna E
                        item2 = ThisWorkbook.Worksheets(capitulo).Cells(k, 5).Value
                        'Si encuentra el item
                        If codItem = item2 Then
                            'Captura el tipo de objeto que tiene la cantidad
                            objeto = ThisWorkbook.Worksheets(capitulo).Cells(k, 4).Value
                            'Busca el fin de la tabla
                            For m = k To k + 100
                                If ThisWorkbook.Worksheets(capitulo).Cells(m, 6).Value = "Total" Then
                                    finalTabla = m
                                    m = k + 100
                                End If
                            Next m
                            objSelection.MoveRight Unit:=wdCharacter, Count:=3
                            'Copia la tabla
                            ThisWorkbook.Worksheets(capitulo).Range(Cells(k, 2), Cells(finalTabla, 9)).Copy
                            'Pega en el documento el titulo del objeto y la tabla
                            texto = numItem & "." & Left(objeto, 1) & " " & diccElemento(objeto)
                            mydoc.Content.InsertAfter Text:=texto
                            objSelection.Style = wordApp.ActiveDocument.Styles("Título 3")
                            mydoc.Content.InsertAfter Text:=vbCrLf
                            objSelection.MoveDown Unit:=wdLine, Count:=1
                            mydoc.Paragraphs.Add.Range.PasteSpecial Link:=False, DataType:=wdPasteEnhancedMetafile, _
                                Placement:=wdInLine, DisplayAsIcon:=False
                            mydoc.Content.InsertAfter Text:=vbCrLf
                        End If
                    Next k
                    
                    'Pasando el APU
                    mydoc.Content.InsertAfter Text:="Análisis de precios unitarios" & vbCrLf
                    objSelection.MoveDown Unit:=wdLine, Count:=2
                    ThisWorkbook.Worksheets("APU").Select
                    'Buscando limites de tabla
                    For n = 1 To 5620
                        valor = Val(ThisWorkbook.Worksheets("APU").Cells(n, 1).Value)
                        If valor = numItem Then
                            inicioTabla = n
                            n = 5620
                        End If
                    Next n
                    For o = inicioTabla To inicioTabla + 200
                        If ThisWorkbook.Worksheets("APU").Cells(o, 4).Value = "Total" Then
                            finalTabla = o
                            o = inicioTabla + 200
                        End If
                    Next o
                    
                    'Copiando tabla
                    ThisWorkbook.Worksheets("APU").Range(Cells(inicioTabla, 1), Cells(finalTabla, 8)).Copy
                    'Pegando en el documento
                    mydoc.Paragraphs.Add.Range.PasteSpecial Link:=False, DataType:=wdPasteEnhancedMetafile, _
                        Placement:=wdInLine, DisplayAsIcon:=False
                    mydoc.Content.InsertAfter Text:=vbCrLf
                    
                    'Verificando si se acabo el capitulo, si es asi, termina el bucle
                    If IsEmpty(ThisWorkbook.Worksheets("EDS").Cells(j + 1, 2).Value) Then
                        j = i + 15
                        objSelection.MoveDown Unit:=wdLine, Count:=1
                    End If
                End If
            Next j
        End If
    Next i
End Sub