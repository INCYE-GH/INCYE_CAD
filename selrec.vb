Sub selrec()
    Dim obj As AcadEntity
    Dim blk As AcadBlockReference
    Dim blockName As String
    Dim blockDict As Object
    Set blockDict = CreateObject("Scripting.Dictionary")
    Dim noContableDict As Object
    Set noContableDict = CreateObject("Scripting.Dictionary")

    ' Verificar si hay un dibujo activo
    If ThisDrawing Is Nothing Then
        MsgBox "No hay dibujo activo. Abre un dibujo y vuelve a intentarlo."
        Exit Sub
    End If
    
    On Error GoTo terminar
    
    ' Crear una nueva aplicación de Excel
    Dim excelApp As Object
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False

    Dim excelWorkbook As Object
    Set excelWorkbook = excelApp.Workbooks.Add

    Dim excelWorksheet As Object
    Set excelWorksheet = excelWorkbook.Sheets(1)
    excelWorksheet.Name = "Recuento Miscelánea"

    ' Solicitar al usuario que ingrese los nombres de las columnas
    excelWorksheet.Cells(1, 1).Value = "BLOQUE"
    excelWorksheet.Cells(1, 3).Value = "Número"
    excelWorksheet.Cells(1, 2).Value = "Capa"

    ' Verificar si estamos en el espacio modelo
    If ThisDrawing.ActiveSpace = acModelSpace Then
        Dim ss As AcadSelectionSet
        On Error Resume Next
        Set ss = ThisDrawing.SelectionSets("SS1")
        On Error GoTo 0

        If ss Is Nothing Then
            Set ss = ThisDrawing.SelectionSets.Add("SS1")
        Else
            ss.clear
        End If

        ' Solicitar al usuario que seleccione los bloques en la pantalla
        ss.SelectOnScreen

        ' Contar los bloques seleccionados y obtener la capa asociada
        Call CountBlocks2(ss, blockDict, noContableDict, excelWorkbook, excelWorksheet)

        ' Eliminar el conjunto de selección
        ss.Delete

        ' Activar la aplicación de Excel
        excelApp.Visible = True
        excelApp.WindowState = -4137 ' xlNormal
    Else
        MsgBox "Por favor, cambia al espacio modelo y vuelve a intentarlo."
    End If
terminar:
End Sub

Private Sub CountBlocks2(ByVal ss As AcadSelectionSet, ByRef blockDict As Object, ByRef noContableDict As Object, ByVal excelWorkbook As Object, ByVal excelWorksheet As Object)
    Dim obj As AcadEntity
    Dim blk As AcadBlockReference
    Dim blkdef As acadBlock
    Dim blockName As String
    Dim innername As String
    Dim innerlayer As String
    Dim entity As AcadEntity
    Dim tempArray As Variant
    Dim newWorksheet As Object
    Dim miscWS As Object

    For Each obj In ss
        If TypeOf obj Is AcadBlockReference Then
            Set blk = obj
            Set blkdef = ThisDrawing.Blocks.Item(blk.Name)
            blockName = blk.effectiveName
           
            
            If blk.Layer = "NoContable" Then
                If Not noContableDict.Exists(blockName) Then
                    Set newWorksheet = excelWorkbook.Sheets.Add
                    newWorksheet.Name = blockName
                    newWorksheet.Cells(1, 1).Value = "BLOQUE"
                    newWorksheet.Cells(1, 2).Value = "Capa"
                    newWorksheet.Cells(1, 3).Value = "Número"
                    noContableDict.Add blockName, CreateObject("Scripting.Dictionary")
                End If
                For Each entity In blkdef
                    If TypeOf entity Is AcadBlockReference Then
                        Dim innerBlkRef As AcadBlockReference
                        Set innerBlkRef = entity
                        innername = innerBlkRef.effectiveName
                        innerlayer = innerBlkRef.Layer
                        If Not noContableDict(blockName).Exists(innername) Then
                            noContableDict(blockName).Add innername, Array(1, innerlayer)
                        Else
                            tempArray = noContableDict(blockName)(innername)
                            tempArray(0) = tempArray(0) + 1
                            noContableDict(blockName)(innername) = tempArray
                        End If
                    End If
                Next entity
                Dim i As Integer
                i = 2
                For Each key In noContableDict(blockName).Keys
                    newWorksheet.Cells(i, 1).Value = key
                    newWorksheet.Cells(i, 2).Value = noContableDict(blockName)(key)(1)
                    newWorksheet.Cells(i, 3).Value = noContableDict(blockName)(key)(0)
                    i = i + 1
                Next key
            Else
                If Not blockDict.Exists(blockName) Then
                    blockDict.Add blockName, Array(1, blk.Layer) ' Guardar el recuento y la capa asociada
                Else
                    tempArray = blockDict(blockName)
                    tempArray(0) = tempArray(0) + 1
                    blockDict(blockName) = tempArray
                End If
                
                
                
                WriteToExcel2 excelWorksheet, blockDict
            End If
        End If
    Next obj
End Sub



Private Sub WriteToExcel2(ByVal excelWorksheet As Object, ByVal blockDict As Object)
    Dim i As Integer
    i = 2

    For Each key In blockDict.Keys
        excelWorksheet.Cells(i, 1).Value = key
        excelWorksheet.Cells(i, 3).Value = blockDict(key)(0) ' Obtener el recuento
        excelWorksheet.Cells(i, 2).Value = blockDict(key)(1) ' Obtener la capa
        i = i + 1
    Next key
End Sub















Sub antiguocomandorecuento()
    Dim obj As AcadEntity
    Dim blk As AcadBlockReference
    Dim blockName As String
    Dim blockDict As Object
    Set blockDict = CreateObject("Scripting.Dictionary")

    ' Verificar si hay un dibujo activo
    If ThisDrawing Is Nothing Then
        MsgBox "No hay dibujo activo. Abre un dibujo y vuelve a intentarlo."
        Exit Sub
    End If
    
    On Error GoTo terminar
    
    ' Crear una nueva aplicación de Excel
    Dim excelApp As Object
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False

    Dim excelWorkbook As Object
    Set excelWorkbook = excelApp.Workbooks.Add

    Dim excelWorksheet As Object
    Set excelWorksheet = excelWorkbook.Sheets(1)

    ' Solicitar al usuario que ingrese los nombres de las columnas
    excelWorksheet.Cells(2, 1).Value = "BLOQUE"
    excelWorksheet.Cells(2, 3).Value = "Número"
    excelWorksheet.Cells(2, 2).Value = "Capa"

    ' Verificar si estamos en el espacio modelo
    If ThisDrawing.ActiveSpace = acModelSpace Then
        Dim ss As AcadSelectionSet
        On Error Resume Next
        Set ss = ThisDrawing.SelectionSets("SS1")
        On Error GoTo 0

        If ss Is Nothing Then
            Set ss = ThisDrawing.SelectionSets.Add("SS1")
        Else
            ss.clear
        End If

        ' Solicitar al usuario que seleccione los bloques en la pantalla
        ss.SelectOnScreen

        ' Contar los bloques seleccionados y obtener la capa asociada
        Call CountBlocks(ss, blockDict)

        ' Escribir los resultados en el libro de Excel
        WriteToExcel excelWorksheet, blockDict

        ' Eliminar el conjunto de selección
        ss.Delete

        ' Activar la aplicación de Excel
        excelApp.Visible = True
        excelApp.WindowState = -4137 ' xlNormal
    Else
        MsgBox "Por favor, cambia al espacio modelo y vuelve a intentarlo."
    End If
terminar:
End Sub

Private Sub CountBlocks(ByVal ss As AcadSelectionSet, ByRef blockDict As Object)
    Dim obj As AcadEntity
    Dim blk As AcadBlockReference
    Dim blkdef As acadBlock
    Dim blockName As String
    Dim innername As String
    Dim innerlayer As String
    Dim entity As AcadEntity
    Dim tempArray As Variant

    For Each obj In ss
        If TypeOf obj Is AcadBlockReference Then
            Set blk = obj
            Set blkdef = ThisDrawing.Blocks.Item(blk.Name)
            blockName = blk.effectiveName
            
            If blk.Layer = "NoContable" Then
                ' Verificar si el bloque contiene otros bloques anidados
                For Each entity In blkdef
                    If TypeOf entity Is AcadBlockReference Then
                        Dim innerBlkRef As AcadBlockReference
                        Set innerBlkRef = entity
                        innername = innerBlkRef.effectiveName
                        innerlayer = innerBlkRef.Layer
                        If Not blockDict.Exists(innername) Then
                            blockDict.Add innername, Array(1, innerBlkRef.Layer) ' Guardar el recuento y la capa asociada
                        Else
                            tempArray = blockDict(innername)
                            tempArray(0) = tempArray(0) + 1
                            blockDict(innername) = tempArray
                        End If
                    End If
                Next entity
            Else
            
                If Not blockDict.Exists(blockName) Then
                    blockDict.Add blockName, Array(1, blk.Layer) ' Guardar el recuento y la capa asociada
                Else
                    
                    tempArray = blockDict(blockName)
                    tempArray(0) = tempArray(0) + 1
                    blockDict(blockName) = tempArray
                End If
                
            End If

        End If
    Next obj
End Sub

Private Sub WriteToExcel(ByVal excelWorksheet As Object, ByVal blockDict As Object)
    Dim i As Integer
    i = 3

    For Each key In blockDict.Keys
        excelWorksheet.Cells(i, 1).Value = key
        excelWorksheet.Cells(i, 3).Value = blockDict(key)(0) ' Obtener el recuento
        excelWorksheet.Cells(i, 2).Value = blockDict(key)(1) ' Obtener la capa
        i = i + 1
    Next key
End Sub





























Sub selrecantiguo()
    Dim obj As AcadEntity
    Dim blk As AcadBlockReference
    Dim blockName As String
    Dim blockDict As Object
    Set blockDict = CreateObject("Scripting.Dictionary")

    ' Verificar si hay un dibujo activo
    If ThisDrawing Is Nothing Then
        MsgBox "No hay dibujo activo. Abre un dibujo y vuelve a intentarlo."
        Exit Sub
    End If

    ' Crear una nueva aplicación de Excel
    Dim excelApp As Object
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    Dim excelWorkbook As Object
    Set excelWorkbook = excelApp.Workbooks.Add
    Dim excelWorksheet As Object
    Set excelWorksheet = excelWorkbook.Sheets(1)

    ' Solicitar al usuario que ingrese los nombres de las columnas
    Dim columna1 As String
    columna1 = InputBox("Ingrese el nombre para el recuento:")
    excelWorksheet.Cells(1, 1).Value = columna1
    excelWorksheet.Cells(2, 1).Value = "BLOQUE"
    excelWorksheet.Cells(2, 3).Value = "Número"
    excelWorksheet.Cells(2, 2).Value = "Capa"


    ' Verificar si estamos en el espacio modelo
    If ThisDrawing.ActiveSpace = acModelSpace Then
        Dim ss As AcadSelectionSet
        On Error Resume Next
        Set ss = ThisDrawing.SelectionSets("SS1")
        On Error GoTo 0

        If ss Is Nothing Then
            Set ss = ThisDrawing.SelectionSets.Add("SS1")
        Else
            ss.clear
        End If

        ' Solicitar al usuario que seleccione los bloques en la pantalla
        ss.SelectOnScreen

        ' Contar los bloques seleccionados y obtener la capa asociada
        For Each obj In ss
            If TypeOf obj Is AcadBlockReference Then
                Set blk = obj
                blockName = blk.effectiveName
                If Not blockDict.Exists(blockName) Then
                    blockDict.Add blockName, Array(1, blk.Layer) ' Guardar el recuento y la capa asociada
                Else
                    Dim tempArray As Variant
                    tempArray = blockDict(blockName)
                    tempArray(0) = tempArray(0) + 1
                    blockDict(blockName) = tempArray
                End If
            End If
        Next obj

        ' Escribir los resultados en el libro de Excel
        Dim i As Integer
        i = 3
        For Each key In blockDict.Keys
            excelWorksheet.Cells(i, 1).Value = key
            excelWorksheet.Cells(i, 3).Value = blockDict(key)(0) ' Obtener el recuento
            excelWorksheet.Cells(i, 2).Value = blockDict(key)(1) ' Obtener la capa
            i = i + 1
        Next key

        ' Eliminar el conjunto de selección
        ss.Delete

        ' Activar la aplicación de Excel
        excelApp.Visible = True
        excelApp.WindowState = -4137 ' xlNormal

    Else
        MsgBox "Por favor, cambia al espacio modelo y vuelve a intentarlo."
    End If
End Sub



