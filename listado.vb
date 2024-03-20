Public Sub WDP(blockRef As AcadBlockReference)

Dim properties As Variant
Dim acDinamicProperty As AcadDynamicBlockReferenceProperty
properties = blockRef.GetDynamicBlockProperties

For j = LBound(properties) To UBound(properties)
    Set acDynamicProperty = properties(j)
    If acDynamicProperty.PropertyName = "Visibility" Then
        MsgBox acDynamicProperty.Value
    End If
Next

End Sub

Public Sub listado()

    Dim element As AcadObject
    Dim block As AcadBlockReference
    Dim att As AcadAttributeReference
    Dim varatt As Variant
    Dim n As Integer
    Dim r As Integer
    Dim j As Integer
    Dim code As String
    Dim delegacion As String
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim wsob As Object
    Dim wsof As Object
    Dim ws2 As Object
    Dim estado1 As String, estado2 As String, estado3 As String
    Dim nombreArchivo As String, numobra As String, del As String, x2x As String, cod As String
    Dim ruta As String, siglas As String, listado As String, to1 As String, pais As String, rellenar_inicio As String
    ' Create a new Excel application

    nombreArchivo = ThisDrawing.Name

    ' Obtener los primeros 11 caracteres
    primeros11Caracteres = Left(nombreArchivo, 11)
    
    If Right(primeros11Caracteres, 1) = "T" Then
        pais = "francia"
        numobra = primeros11Caracteres
        del = Right(numobra, 3)
        x2x = Left(numobra, 2)
        cod = Left(numobra, 8)
        cod = Right(cod, 6)
        ruta = "C:\Users\" & Environ$("Username") & "\Incye\France - Projets\" & numobra & "\02 Plans\2_Nouveaux\2_PDF\"
        listado = ruta & numobra & "_ListeDesPlans.xlsm"
        to1 = "C:\Users\" & Environ$("Username") & "\Incye\France - Projets\" & numobra & "\01 Info\" & numobra & ".xlsm"
        If Dir(listado) <> "" Then
            ' El archivo ya existe en la ruta especificada
            ' no rellenar pestaña de inicio
            rellenar_inicio = "no"
        Else
            rellenar_inicio = "si"
            'generar el archivo basándonos en el francés
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            objFSO.CopyFile "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\13_Formacion\23XXXX_MACROS\Listado de planos - fr.xlsm", listado

        End If
        
    Else
        numobra = Left(primeros11Caracteres, 10)
        If Right(numobra, 1) = "R" Then
            pais = "francia"
            numobra = numobra
            del = Right(numobra, 2)
            x2x = Left(numobra, 2)
            cod = Left(numobra, 8)
            cod = Right(cod, 6)
            ruta = "C:\Users\" & Environ$("Username") & "\Incye\France - Projets\" & numobra & "\02 Plans\2_Nouveaux\2_PDF\"
            listado = ruta & numobra & "_ListeDesPlans.xlsm"
            to1 = "C:\Users\" & Environ$("Username") & "\Incye\France - Projets\" & numobra & "\01 Info\" & numobra & ".xlsm"
            If Dir(listado) <> "" Then
                ' El archivo ya existe en la ruta especificada
                ' no rellenar pestaña de inicio
                rellenar_inicio = "no"
            Else
                rellenar_inicio = "si"
                'generar el archivo basándonos en el francés
                Set objFSO = CreateObject("Scripting.FileSystemObject")
                objFSO.CopyFile "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\13_Formacion\23XXXX_MACROS\Listado de planos - fr.xlsm", listado
            End If
            
        Else
            pais = "españa"
            numobra = Left(numobra, 9)
            del = Right(numobra, 1)
            If del = "T" Then
                siglas = "BCN"
            ElseIf del = "N" Then
                siglas = "BLB"
            ElseIf del = "F" Then
                siglas = "CRN"
            ElseIf del = "V" Then
                siglas = "LEV"
            ElseIf del = "M" Then
                siglas = "MAD"
            ElseIf del = "X" Then
                siglas = "SEV"
            End If
            x2x = Left(numobra, 2)
            cod = Left(numobra, 8)
            cod = Right(cod, 6)
            ruta = "C:\Users\" & Environ$("Username") & "\Incye\Proyectos - Documentos\" & siglas & "\" & numobra & "\02 Planos\2_Nuevos\2_PDF\"
            listado = ruta & numobra & "_ListadoDePlanos.xlsm"
            to1 = "C:\Users\" & Environ$("Username") & "\Incye\Proyectos - Documentos\" & siglas & "\" & numobra & "\01 Info\" & numobra & ".xlsm"
            If Dir(listado) <> "" Then
                ' El archivo ya existe en la ruta especificada
                ' no rellenar pestaña de inicio
                rellenar_inicio = "no"
            Else
                rellenar_inicio = "si"
                'generar el archivo basándonos en el español
                Set objFSO = CreateObject("Scripting.FileSystemObject")
                objFSO.CopyFile "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\13_Formacion\23XXXX_MACROS\Listado_de_planos.xlsm", listado
            End If
        End If
    End If
    
    If pais = "españa" Then
        Call esp(listado, to1, ruta, rellenar_inicio)
    ElseIf pais = "francia" Then
        Call fr(listado, to1, ruta, rellenar_inicio)
    End If

End Sub

Public Sub esp(listado As String, to1 As String, ruta As String, rellenar_inicio As String)

    Dim element As AcadObject
    Dim block As AcadBlockReference
    Dim att As AcadAttributeReference
    Dim varatt As Variant
    Dim n As Integer
    Dim r As Integer
    Dim j As Integer
    Dim code As String
    Dim delegacion As String
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim wt As Object
    Dim wsob As Object
    Dim wsof As Object, t01 As Object
    Dim ws2 As Object
    Dim estado1 As String, estado2 As String, estado3 As String
    Dim nombreArchivo As String, numobra As String, del As String, x2x As String, cod As String
    Dim siglas As String
    ' Create a new Excel application

    nombreArchivo = ThisDrawing.Name

    On Error GoTo Terminar
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False ' Set to False if you don't want Excel to be visible
    

    Set wb = excelApp.Workbooks.Open(listado)
    Set wsinicio = wb.Worksheets("Inicio")
    Set ws = wb.Worksheets("Listado Planos Obra")
    Set wt = wb.Worksheets("Listado Planos Oferta")
    
' vamos a obtener las filas sobra las que tenemos que comenzar a escribir en cada pestaña
    
    Dim primeraFilaVacia_oferta As Long
    Dim lastRow_oferta As Long
    
    lastRow_oferta = wt.Cells(wt.Rows.Count, "A").End(xlUp).row
    

' de aquí extraemos la primera fila vacía de la pestaña de oferta
    For r = 10 To lastRow_oferta
        If Trim(wt.Cells(r, 1).Value) = "" Then ' Verificar si la celda en la columna A está vacía
            primeraFilaVacia_oferta = r
            Exit For
        End If
    Next r
    
    
    Dim layoutNames As String
    layoutNames = InputBox("Planos a añadir al listado: (separados sólo mediante comas)")
    
    'Split into array
    Dim layoutArray() As String
    layoutArray = Split(layoutNames, ",")
    
    'Get number of layouts
    Dim numLayouts As Long
    numLayouts = UBound(layoutArray) + 1
    
    'Loop through layouts
    Dim i As Long
    For i = 1 To numLayouts
    
        'Get layout name
        Dim layoutName As String
        layoutName = layoutArray(i - 1)
    
        'Get layout
        Dim layout As AcadLayout
        Set layout = ThisDrawing.Layouts(layoutName)
        identificador = Left(layoutName, 1)
        If identificador = "0" Then
            Set ws = wb.Worksheets("Listado Planos Oferta")
        ElseIf identificador = "1" Then
            Set ws = wb.Worksheets("Listado Planos Obra")
        End If
        
        ' de aquí extraemos la primera fila vacía
        Dim primeraFilaVacia As Long
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        
        For q = 10 To lastRow
            If Trim(ws.Cells(q, 1).Value) = "" Then ' Verificar si la celda en la columna A está vacía
                primeraFilaVacia = q
                Exit For
            End If
        Next q
        
    
        If Not layout Is Nothing Then
    
            'Switch to layout
            ThisDrawing.ActiveLayout = layout
    
            'Get attributes
            n = ThisDrawing.PaperSpace.Count
    
            For r = 0 To n - 1
                Set element = ThisDrawing.PaperSpace.Item(r)
                
                If element.ObjectName = "AcDbBlockReference" Then
                    Set block = element
                    If block.effectiveName = "description2" Then
                        varatt = block.GetAttributes
                        For j = LBound(varatt) To UBound(varatt)
                            Set att = varatt(j)
                            If att.TagString = "NPLANO" Then
                                nplano = att.TextString
                                ws.Cells(primeraFilaVacia, 2).Value = att.TextString
                            ElseIf att.TagString = "TEC" Then
                                tecnico = att.TextString
                                ws.Cells(primeraFilaVacia, 10).Value = att.TextString
                                Dim tec As String
                                If tecnico = "JMM" Then
                                    tec = "José M. Maldonado"
                                ElseIf tecnico = "DLM" Then
                                    tec = "David Lara."
                                ElseIf tecnico = "ESG" Then
                                    tec = "Ezequiel Sánchez."
                                ElseIf tecnico = "ARP" Then
                                    tec = "Andrés Rodríguez."
                                ElseIf tecnico = "AAM" Then
                                    tec = "Alberto Aldama."
                                ElseIf tecnico = "ASC" Then
                                    tec = "Adelaida Sáez."
                                ElseIf tecnico = "AAB" Then
                                    tec = "Alejandro Ángel Builes."
                                ElseIf tecnico = "JJM" Then
                                    tec = "Juan José Morón."
                                ElseIf tecnico = "MGA" Then
                                    tec = "Manuel González."
                                ElseIf tecnico = "RMC" Then
                                    tec = "Rafael Mansilla."
                                ElseIf tecnico = "ELF" Then
                                    tec = "Esteban López Fernández."
                                ElseIf tecnico = "FB" Then
                                    tec = "Filippo Brusca."
                                End If
                                wsinicio.Cells(2, 3).Value = tec
                            ElseIf att.TagString = "DATE" Then
                                fecha = att.TextString
                                ws.Cells(primeraFilaVacia, 5).Value = att.TextString
                                ws.Cells(primeraFilaVacia, 9).Value = att.TextString
                                ws.Cells(primeraFilaVacia, 14).Value = att.TextString
                            ElseIf att.TagString = "RV" Then
                                revisor = att.TextString
                                ws.Cells(primeraFilaVacia, 11).Value = att.TextString
                            ElseIf att.TagString = "DATERV" Then
                                fecharv = att.TextString
                                ws.Cells(primeraFilaVacia, 12).Value = att.TextString
                            ElseIf att.TagString = "EQUIPMENT" Then
                                equipamiento = att.TextString
                                ws.Cells(primeraFilaVacia, 6).Value = att.TextString
                            ElseIf att.TagString = "STRUCTURE" Then
                                estructura = att.TextString
                            ElseIf att.TagString = "PROJECT" Then
                                proyecto = att.TextString
                                wsinicio.Cells(6, 3).Value = estructura & " " & proyecto
                            ElseIf att.TagString = "CUSTOMER" Then
                                wsinicio.Cells(8, 3).Value = att.TextString
                            ElseIf att.TagString = "2X" Then
                                cod1 = att.TextString
                            ElseIf att.TagString = "COD" Then
                                cod2 = att.TextString
                            ElseIf att.TagString = "DEL" Then
                                cod3 = att.TextString
                                ' veamos la delegación para construir el path
                                If cod3 = "M" Then
                                    delegacion = "MAD"
                                ElseIf cod3 = "T" Then
                                    delegacion = "BCN"
                                ElseIf cod3 = "N" Then
                                    delegacion = "BLB"
                                ElseIf cod3 = "F" Then
                                    delegacion = "CRN"
                                ElseIf cod3 = "E" Then
                                    delegacion = "EXP"
                                ElseIf cod3 = "V" Then
                                    delegacion = "LEV"
                                ElseIf cod3 = "P" Then
                                    delegacion = "PT"
                                ElseIf cod3 = "X" Then
                                    delegacion = "SEV"
                                End If
                            ElseIf att.TagString = "REVISION" Then
                                revision = att.TextString
                                ws.Cells(primeraFilaVacia, 3).Value = att.TextString
                            End If
                        Next
                        wsinicio.Cells(1, 16).Value = delegacion
                        wsinicio.Cells(4, 3).Value = cod1 & cod2 & cod3
                        wsinicio.Cells(22, 3).Value = cod1 & cod2 & cod3 & "_LP_" & fecha
                        wsinicio.Cells(1, 15).Value = "" & Environ$("Username") & ""
                        ws.Cells(primeraFilaVacia, 1).Value = cod1 & cod2 & cod3
                        code = cod1 & cod2 & cod3
                    End If
                    If element.effectiveName = "status_sp" And element.IsDynamicBlock Then
                        Dim properties As Variant
                        Dim acDinamicProperty As AcadDynamicBlockReferenceProperty
                        properties = element.GetDynamicBlockProperties
                        For t = LBound(properties) To UBound(properties)
                            Set acDynamicProperty = properties(t)
                            If acDynamicProperty.PropertyName = "Visibility" Then
                                If acDynamicProperty.Value = "Plano de Trabajo" Then
                                    estado1 = "TRABAJO: Comprobado para montaje"         'estatus
                                    estado2 = "-"      ' marca de agua
                                    estado3 = "Plano de Trabajo"
                                ElseIf acDynamicProperty.Value = "Plano de Aprobación" Then
                                    estado1 = "PROPUESTA: No comprobado para montaje"    'estatus
                                    estado2 = "Para aprobación de cliente"     ' marca de agua
                                    estado3 = "Plano de Aprobación"
                                ElseIf acDynamicProperty.Value = "Plano Preliminar" Then
                                    estado1 = "PROPUESTA: No comprobado para montaje"    'estatus
                                    estado2 = "Preliminar"           ' marca de agua
                                    estado3 = "Plano Preliminar"
                                ElseIf acDynamicProperty.Value = "Plano de Conceptos" Then
                                    estado1 = "PROPUESTA: No comprobado para montaje"    'estatus
                                    estado2 = "Conceptual"           ' marca de agua
                                    estado3 = "Plano de Conceptos"
                                End If
                            End If
                        Next
                    ws.Cells(primeraFilaVacia, 7).Value = estado1 'estatus
                    ws.Cells(primeraFilaVacia, 8).Value = estado2 ' marca de agua
                    wsinicio.Cells(18, 3).Value = estado3
                    End If

                End If
            Next
 
        Else
            MsgBox "Asegúrate de introducir bien el nombre de las pestañas"
            GoTo Terminar
        End If
        Dim fileName As String
        fileName = Left(ThisDrawing.Name, InStrRev(ThisDrawing.Name, ")"))
        
        'ws.Cells(1, 13).Value = "Nombre cajetín"
        ws.Cells(primeraFilaVacia, 13).Value = nombreArchivo
    Next i
    
    
    
    
    Set ws = wb.Worksheets("Listado Planos Obra")
    Dim x As Long
    
    Dim uniqueValues As Object
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    
    ' Identificar números de plano únicos
    For x = 11 To 40
        If Not uniqueValues.Exists(ws.Cells(x, 2).Value) Then
            uniqueValues.Add ws.Cells(x, 2).Value, x
            ws.Cells(x, 4).Value = "OK"
        End If
    Next x
    
    ' Encontrar la versión más alta para cada número de plano repetido
    For Each key In uniqueValues.Keys
        Dim maxValue As String
        maxValue = ""
        Dim maxIndex As Long
        maxIndex = 0
        
        For x = 11 To 40
            If ws.Cells(x, 2).Value = key Then
                If ws.Cells(x, 3).Value > maxValue Then
                    maxValue = ws.Cells(x, 3).Value
                    maxIndex = x
                End If
            End If
        Next x
    
        ' Marcar como "OK" la versión más alta y como "ANULADO" las demás
        For x = 11 To 40
            If ws.Cells(x, 2).Value = key Then
                    If x = maxIndex Then
                        ws.Cells(x, 4).Value = "OK"
                    Else
                        ws.Cells(x, 4).Value = "ANULADO"
                    End If
                End If
           
        Next x
        ' Vaciamos todos los "ANULADO" que mete el código por defecto
        For x = 11 To 40
            If IsEmpty(ws.Cells(x, 2)) Then
                ws.Cells(x, 4).Value = " "
            End If
        Next x
    Next

    Set wsf = wb.Worksheets("Listado Planos Oferta")
    Dim y As Long
    
    Dim uniqueValues2 As Object
    Set uniqueValues2 = CreateObject("Scripting.Dictionary")
    
    ' Identificar números de plano únicos
    For y = 11 To 40
        If Not uniqueValues2.Exists(wsf.Cells(y, 2).Value) Then
            uniqueValues2.Add wsf.Cells(y, 2).Value, y
            wsf.Cells(y, 4).Value = "OK"
        End If
    Next y
    
    ' Encontrar la versión más alta para cada número de plano repetido
    For Each key In uniqueValues2.Keys
        Dim maxValue2 As String
        maxValue2 = ""
        Dim maxIndex2 As Long
        maxIndex2 = 0
        
        For y = 11 To 40
            If wsf.Cells(y, 2).Value = key Then
                If wsf.Cells(y, 3).Value > maxValue2 Then
                    maxValue2 = wsf.Cells(y, 3).Value
                    maxIndex2 = y
                End If
            End If
        Next y
    
        ' Marcar como "OK" la versión más alta y como "ANULADO" las demás
        For y = 11 To 40
            If wsf.Cells(y, 2).Value = key Then
                If y = maxIndex2 Then
                    wsf.Cells(y, 4).Value = "OK"
                Else
                    wsf.Cells(y, 4).Value = "ANULADO"
                End If
            End If
        Next y
        ' Vaciamos todos los "ANULADO" que mete el código por defecto
        For y = 11 To 40
            If IsEmpty(wsf.Cells(y, 2)) Then
                wsf.Cells(y, 4).Value = " "
            End If
        Next y
    Next

    If rellenar_inicio = "si" Then
        If Dir(to1) <> "" Then
            Set t01 = excelApp.Workbooks.Open(to1)
            Set hoja = t01.Worksheets("DATOS")
            delegado = hoja.Cells(7, 3).Value
            
            If delegado = "AVS" Then
                contactodel = "antonio.vazquez@incye.com"
            ElseIf delegado = "CENT" Then
                contactodel = "madrid@incye.com"
            ElseIf delegado = "BCN" Then
                contactodel = "barcelona@incye.com"
            ElseIf delegado = "BLB" Then
                contactodel = "bilbao@incye.com"
            ElseIf delegado = "ASP" Or delegado = "CRN" Then
                contactodel = "ana.seoane@incye.com"
            ElseIf delegado = "SPG" Then
                contactodel = "galicia@incye.com"
            ElseIf delegado = "EXP" Then
                contactodel = "antonio.vazquez@incye.com"
            ElseIf delegado = "Xavier Marty" Or delegado = "FR" Then
                contactodel = "xavier.marty@incye.com"
            ElseIf delegado = "LEV" Then
                contactodel = "valencia@incye.com"
            ElseIf delegado = "AND" Then
                contactodel = "malaga@incye.com"
            End If
            
            nomcontactocliente = hoja.Cells(17, 3).Value
            numconcactocliente = hoja.Cells(23, 3).Value
            
            wsinicio.Cells(10, 3).Value = nomcontactocliente
            wsinicio.Cells(12, 3).Value = numcontactocliente
            wsinicio.Cells(20, 3).Value = ruta
            wsinicio.Cells(24, 3).Value = listado
            wsinicio.Cells(14, 3).Value = contactodel
            
            t01.Close (False)
        Else
            MsgBox "Comprueba que el T01 tenga el nombre adecuado y vuelve a lanzarlo."
            GoTo Terminar
        End If
    Else
    End If

    
    'Dim codigo As String
    'n = Left(ThisDrawing.Name, InStrRev(ThisDrawing.Name, "(") - 1)
    'identificador = Left(nplano, 1)
    wb.Save
    MsgBox "Planos inlcuidos en el Listado correctamente."
Terminar:
    
    wb.Close (False)
    excelApp.Quit
    Set excelApp = Nothing
    

End Sub



Public Sub fr(listado As String, to1 As String, ruta As String, rellenar_inicio As String)

    Dim element As AcadObject
    Dim block As AcadBlockReference
    Dim att As AcadAttributeReference
    Dim varatt As Variant
    Dim n As Integer
    Dim r As Integer
    Dim j As Integer
    Dim code As String
    Dim delegacion As String
    Dim excelApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim wt As Object
    Dim wsob As Object
    Dim wsof As Object, t01 As Object
    Dim ws2 As Object
    Dim estado1 As String, estado2 As String, estado3 As String
    Dim nombreArchivo As String, numobra As String, del As String, x2x As String, cod As String
    Dim siglas As String
    ' Create a new Excel application

    nombreArchivo = ThisDrawing.Name

    On Error GoTo Terminar
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False ' Set to False if you don't want Excel to be visible
    

    Set wb = excelApp.Workbooks.Open(listado)
    Set wsinicio = wb.Worksheets("Inicio")
    Set ws = wb.Worksheets("Listado Planos Obra")
    Set wt = wb.Worksheets("Listado Planos Oferta")
    
' vamos a obtener las filas sobra las que tenemos que comenzar a escribir en cada pestaña
    
    Dim primeraFilaVacia_oferta As Long
    Dim lastRow_oferta As Long
    
    lastRow_oferta = wt.Cells(wt.Rows.Count, "A").End(xlUp).row
    

' de aquí extraemos la primera fila vacía de la pestaña de oferta
    For r = 10 To lastRow_oferta
        If Trim(wt.Cells(r, 1).Value) = "" Then ' Verificar si la celda en la columna A está vacía
            primeraFilaVacia_oferta = r
            Exit For
        End If
    Next r
    
    
    Dim layoutNames As String
    layoutNames = InputBox("Planos a añadir al listado: (separados sólo mediante comas)")
    
    'Split into array
    Dim layoutArray() As String
    layoutArray = Split(layoutNames, ",")
    
    'Get number of layouts
    Dim numLayouts As Long
    numLayouts = UBound(layoutArray) + 1
    
    'Loop through layouts
    Dim i As Long
    For i = 1 To numLayouts
    
        'Get layout name
        Dim layoutName As String
        layoutName = layoutArray(i - 1)
    
        'Get layout
        Dim layout As AcadLayout
        Set layout = ThisDrawing.Layouts(layoutName)
        identificador = Left(layoutName, 1)
        If identificador = "0" Then
            Set ws = wb.Worksheets("Listado Planos Oferta")
        ElseIf identificador = "1" Then
            Set ws = wb.Worksheets("Listado Planos Obra")
        End If
        
        ' de aquí extraemos la primera fila vacía
        Dim primeraFilaVacia As Long
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        
        For q = 10 To lastRow
            If Trim(ws.Cells(q, 1).Value) = "" Then ' Verificar si la celda en la columna A está vacía
                primeraFilaVacia = q
                Exit For
            End If
        Next q
        
    
        If Not layout Is Nothing Then
    
            'Switch to layout
            ThisDrawing.ActiveLayout = layout
    
            'Get attributes
            n = ThisDrawing.PaperSpace.Count
    
            For r = 0 To n - 1
                Set element = ThisDrawing.PaperSpace.Item(r)
                
                If element.ObjectName = "AcDbBlockReference" Then
                    Set block = element
                    If block.effectiveName = "description2" Then
                        varatt = block.GetAttributes
                        For j = LBound(varatt) To UBound(varatt)
                            Set att = varatt(j)
                            If att.TagString = "NPLANO" Then
                                nplano = att.TextString
                                ws.Cells(primeraFilaVacia, 2).Value = att.TextString
                            ElseIf att.TagString = "TEC" Then
                                tecnico = att.TextString
                                ws.Cells(primeraFilaVacia, 10).Value = att.TextString
                                Dim tec As String
                                If tecnico = "JMM" Then
                                    tec = "José M. Maldonado"
                                ElseIf tecnico = "DLM" Then
                                    tec = "David Lara."
                                ElseIf tecnico = "ESG" Then
                                    tec = "Ezequiel Sánchez."
                                ElseIf tecnico = "ARP" Then
                                    tec = "Andrés Rodríguez."
                                ElseIf tecnico = "AAM" Then
                                    tec = "Alberto Aldama."
                                ElseIf tecnico = "ASC" Then
                                    tec = "Adelaida Sáez."
                                ElseIf tecnico = "AAB" Then
                                    tec = "Alejandro Ángel Builes."
                                ElseIf tecnico = "JJM" Then
                                    tec = "Juan José Morón."
                                ElseIf tecnico = "MGA" Then
                                    tec = "Manuel González."
                                ElseIf tecnico = "RMC" Then
                                    tec = "Rafael Mansilla."
                                ElseIf tecnico = "ELF" Then
                                    tec = "Esteban López Fernández."
                                ElseIf tecnico = "FB" Then
                                    tec = "Filippo Brusca."
                                End If
                                wsinicio.Cells(2, 3).Value = tec
                            ElseIf att.TagString = "DATE" Then
                                fecha = att.TextString
                                ws.Cells(primeraFilaVacia, 5).Value = att.TextString
                                ws.Cells(primeraFilaVacia, 9).Value = att.TextString
                                ws.Cells(primeraFilaVacia, 14).Value = att.TextString
                            ElseIf att.TagString = "RV" Then
                                revisor = att.TextString
                                ws.Cells(primeraFilaVacia, 11).Value = att.TextString
                            ElseIf att.TagString = "DATERV" Then
                                fecharv = att.TextString
                                ws.Cells(primeraFilaVacia, 12).Value = att.TextString
                            ElseIf att.TagString = "EQUIPMENT" Then
                                equipamiento = att.TextString
                                ws.Cells(primeraFilaVacia, 6).Value = att.TextString
                            ElseIf att.TagString = "STRUCTURE" Then
                                estructura = att.TextString
                            ElseIf att.TagString = "PROJECT" Then
                                proyecto = att.TextString
                                wsinicio.Cells(6, 3).Value = estructura & " " & proyecto
                            ElseIf att.TagString = "CUSTOMER" Then
                                wsinicio.Cells(8, 3).Value = att.TextString
                            ElseIf att.TagString = "2X" Then
                                cod1 = att.TextString
                            ElseIf att.TagString = "COD" Then
                                cod2 = att.TextString
                            ElseIf att.TagString = "DEL" Then
                                cod3 = att.TextString
                                ' veamos la delegación para construir el path
                                If cod3 = "M" Then
                                    delegacion = "MAD"
                                ElseIf cod3 = "T" Then
                                    delegacion = "BCN"
                                ElseIf cod3 = "N" Then
                                    delegacion = "BLB"
                                ElseIf cod3 = "F" Then
                                    delegacion = "CRN"
                                ElseIf cod3 = "E" Then
                                    delegacion = "EXP"
                                ElseIf cod3 = "V" Then
                                    delegacion = "LEV"
                                ElseIf cod3 = "P" Then
                                    delegacion = "PT"
                                ElseIf cod3 = "X" Then
                                    delegacion = "SEV"
                                End If
                            ElseIf att.TagString = "REVISION" Then
                                revision = att.TextString
                                ws.Cells(primeraFilaVacia, 3).Value = att.TextString
                            End If
                        Next
                        wsinicio.Cells(1, 16).Value = delegacion
                        wsinicio.Cells(4, 3).Value = cod1 & cod2 & cod3
                        wsinicio.Cells(22, 3).Value = cod1 & cod2 & cod3 & "_LP_" & fecha
                        wsinicio.Cells(1, 15).Value = "" & Environ$("Username") & ""
                        ws.Cells(primeraFilaVacia, 1).Value = cod1 & cod2 & cod3
                        code = cod1 & cod2 & cod3
                    End If
                    If element.effectiveName = "status_sp-fr" And element.IsDynamicBlock Then
                        Dim properties As Variant
                        Dim acDinamicProperty As AcadDynamicBlockReferenceProperty
                        properties = element.GetDynamicBlockProperties
                        For t = LBound(properties) To UBound(properties)
                            Set acDynamicProperty = properties(t)
                            If acDynamicProperty.PropertyName = "Visibility" Then
                                If acDynamicProperty.Value = "Plano de Trabajo" Then
                                    estado1 = "Exécution: valable pour le montage"         'estatus
                                    estado2 = "-"      ' marca de agua
                                    estado3 = "Plan d'exécution"
                                ElseIf acDynamicProperty.Value = "Plano de Aprobación" Then
                                    estado1 = "PROPOSITION: non valable pour le montage"    'estatus
                                    estado2 = "Plan d'Aprobation"     ' marca de agua
                                    estado3 = "Plan d'Aprobation"
                                ElseIf acDynamicProperty.Value = "Plano Preliminar" Then
                                    estado1 = "PROPOSITION: non valable pour le montage"    'estatus
                                    estado2 = "Preliminaire"           ' marca de agua
                                    estado3 = "Plano Preliminar"
                                ElseIf acDynamicProperty.Value = "Plano de Conceptos" Then
                                    estado1 = "PROPOSITION: non valable pour le montage"    'estatus
                                    estado2 = "Conceptuel"           ' marca de agua
                                    estado3 = "Plano de Conceptos"
                                End If
                            End If
                        Next
                    ws.Cells(primeraFilaVacia, 7).Value = estado1 'estatus
                    ws.Cells(primeraFilaVacia, 8).Value = estado2 ' marca de agua
                    wsinicio.Cells(18, 3).Value = estado3
                    End If

                End If
            Next
 
        Else
            MsgBox "Asegúrate de introducir bien el nombre de las pestañas"
            GoTo Terminar
        End If
        Dim fileName As String
        fileName = Left(ThisDrawing.Name, InStrRev(ThisDrawing.Name, ")"))
        
        'ws.Cells(1, 13).Value = "Nombre cajetín"
        ws.Cells(primeraFilaVacia, 13).Value = nombreArchivo
    Next i
    
    
    
    
    Set ws = wb.Worksheets("Listado Planos Obra")
    Dim x As Long
    
    Dim uniqueValues As Object
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    
    ' Identificar números de plano únicos
    For x = 11 To 40
        If Not uniqueValues.Exists(ws.Cells(x, 2).Value) Then
            uniqueValues.Add ws.Cells(x, 2).Value, x
            ws.Cells(x, 4).Value = "OK"
        End If
    Next x
    
    ' Encontrar la versión más alta para cada número de plano repetido
    For Each key In uniqueValues.Keys
        Dim maxValue As String
        maxValue = ""
        Dim maxIndex As Long
        maxIndex = 0
        
        For x = 11 To 40
            If ws.Cells(x, 2).Value = key Then
                If ws.Cells(x, 3).Value > maxValue Then
                    maxValue = ws.Cells(x, 3).Value
                    maxIndex = x
                End If
            End If
        Next x
    
        ' Marcar como "OK" la versión más alta y como "ANULADO" las demás
        For x = 11 To 40
            If ws.Cells(x, 2).Value = key Then
                    If x = maxIndex Then
                        ws.Cells(x, 4).Value = "OK"
                    Else
                        ws.Cells(x, 4).Value = "ANNULÉ"
                    End If
                End If
           
        Next x
        ' Vaciamos todos los "ANULADO" que mete el código por defecto
        For x = 11 To 40
            If IsEmpty(ws.Cells(x, 2)) Then
                ws.Cells(x, 4).Value = " "
            End If
        Next x
    Next

    Set wsf = wb.Worksheets("Listado Planos Oferta")
    Dim y As Long
    
    Dim uniqueValues2 As Object
    Set uniqueValues2 = CreateObject("Scripting.Dictionary")
    
    ' Identificar números de plano únicos
    For y = 11 To 40
        If Not uniqueValues2.Exists(wsf.Cells(y, 2).Value) Then
            uniqueValues2.Add wsf.Cells(y, 2).Value, y
            wsf.Cells(y, 4).Value = "OK"
        End If
    Next y
    
    ' Encontrar la versión más alta para cada número de plano repetido
    For Each key In uniqueValues2.Keys
        Dim maxValue2 As String
        maxValue2 = ""
        Dim maxIndex2 As Long
        maxIndex2 = 0
        
        For y = 11 To 40
            If wsf.Cells(y, 2).Value = key Then
                If wsf.Cells(y, 3).Value > maxValue2 Then
                    maxValue2 = wsf.Cells(y, 3).Value
                    maxIndex2 = y
                End If
            End If
        Next y
    
        ' Marcar como "OK" la versión más alta y como "ANULADO" las demás
        For y = 11 To 40
            If wsf.Cells(y, 2).Value = key Then
                If y = maxIndex2 Then
                    wsf.Cells(y, 4).Value = "OK"
                Else
                    wsf.Cells(y, 4).Value = "ANNULÉ"
                End If
            End If
        Next y
        ' Vaciamos todos los "ANULADO" que mete el código por defecto
        For y = 11 To 40
            If IsEmpty(wsf.Cells(y, 2)) Then
                wsf.Cells(y, 4).Value = " "
            End If
        Next y
    Next

    If rellenar_inicio = "si" Then
        If Dir(to1) <> "" Then
            Set t01 = excelApp.Workbooks.Open(to1)
            Set hoja = t01.Worksheets("DONÉES")
            delegado = hoja.Cells(7, 3).Value
            
            If delegado = "AVS" Then
                contactodel = "antonio.vazquez@incye.com"
            ElseIf delegado = "CENT" Then
                contactodel = "madrid@incye.com"
            ElseIf delegado = "BCN" Then
                contactodel = "barcelona@incye.com"
            ElseIf delegado = "BLB" Then
                contactodel = "bilbao@incye.com"
            ElseIf delegado = "ASP" Or delegado = "CRN" Then
                contactodel = "ana.seoane@incye.com"
            ElseIf delegado = "SPG" Then
                contactodel = "galicia@incye.com"
            ElseIf delegado = "EXP" Then
                contactodel = "antonio.vazquez@incye.com"
            ElseIf delegado = "Xavier Marty" Or delegado = "XM" Then
                contactodel = "xavier.marty@incye.com"
            ElseIf delegado = "Xavier Raynal" Or delegado = "XR" Then
                contactodel = "xavier.raynal@incye.com"
            ElseIf delegado = "AND" Then
                contactodel = "malaga@incye.com"
            End If
            
            nomcontactocliente = hoja.Cells(17, 3).Value
            numconcactocliente = hoja.Cells(23, 3).Value
            wsinicio.Cells(10, 3).Value = nomcontactocliente
            wsinicio.Cells(12, 3).Value = numcontactocliente
            wsinicio.Cells(14, 3).Value = contactodel
            wsinicio.Cells(20, 3).Value = ruta
            wsinicio.Cells(24, 3).Value = listado
            
            t01.Close (False)
        Else
            MsgBox "Comprueba que el T01 tenga el nombre adecuado y vuelve a lanzarlo."
            GoTo Terminar
        End If
    Else
    End If


    
    'Dim codigo As String
    'n = Left(ThisDrawing.Name, InStrRev(ThisDrawing.Name, "(") - 1)
    'identificador = Left(nplano, 1)
    wb.Save
    MsgBox "Planos inlcuidos en el Listado correctamente."
    
Terminar:
    wb.Close (False)
    excelApp.Quit
    Set excelApp = Nothing
    

End Sub




















