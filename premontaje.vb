Public consecutivo As Integer

Sub premo()
    Dim obj As GcadEntity
    Dim blk As GcadBlockReference
    Dim blockName As String
    
    Dim PA(0 To 2) As Double, PB(0 To 2) As Double
    Dim insertionPoint2(0 To 2) As Double
    Dim insertionT(0 To 2) As Double
    Dim insertionP(0 To 2) As Double
    
    Dim orientation As Double
    Dim orientationa As Double
    
    Dim textHeight As Double
    Dim textHeight2 As Double
    Dim textToInsert As String
    Dim textObj As GcadText
    Dim newText As GcadObject
    Dim newText2 As GcadObject
    
    Dim rutaps As String, rutapl As String, rutap6 As String
    Dim ps As String
    
    Dim ventanaacti As Integer
    
    Dim newLeader As Object
    Dim nombresItems As Variant
    Dim nombresItems2 As Variant
    Dim i As Integer
    Dim anglePerpendicular As Double
    Dim leaderEndPoint(0 To 2) As Double
    Dim textEndPoint(0 To 2) As Double
    Dim newLine As Object
    Dim rotationAngle As Double
    Dim k As String
    
    Dim textoSinUltimasSeisLetras As String
    Dim distancia As Double
    Dim obj2 As GcadEntity
    Dim blk2 As GcadBlockReference
    
    Dim blockRef As GcadBlockReference
    
    Dim pi As Double
    pi = 4 * Atn(1)
    
    Set gcadDoc = GetObject(, "Gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    
    On Error GoTo terminar
    
    rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
    rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
    rutap6 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_6\"
    rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"
    rutacajon = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Cajon hidraulico\"
    rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
    rutacu = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"

    ' Verificar si hay un dibujo activo
    If ThisDrawing Is Nothing Then
        MsgBox "No hay dibujo activo. Abre un dibujo y vuelve a intentarlo."
        Exit Sub
    End If
    
    repite = 1
    Do While repite = 1
    
    On Error GoTo terminar
    
    ' Definir la distancia de offset para el siguiente bloque
    offsetDistance = 50
    
    'Posicion texto
    Dim offsetXTexto As Double, offsetYTexto As Double
    offsetXTexto = 30
    offsetYTexto = 2.5
    
    textHeight = 5
       
    ' Verificar si estamos en el espacio modelo
    If ThisDrawing.ActiveSpace = acModelSpace Then
        Dim ss As GcadSelectionSet

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
        
        ' Seleccionar punto de inserción
        Dim puntoInicio As Variant, puntoFin As Variant
        
        puntoInicio = ThisDrawing.Utility.GetPoint(, "Selecciona el punto de Insercion: ")
        puntoFin = ThisDrawing.Utility.GetPoint(, "Selecciona el punto Fin del Eje: ")
            
        PA(0) = puntoInicio(0): PA(1) = puntoInicio(1): PA(2) = puntoInicio(2)
        PB(0) = puntoFin(0): PB(1) = puntoFin(1): PB(2) = puntoFin(2)
        
        ' Establecer la escala deseada en el espacio de papel
        
        'calculo distancia total
            xa = PA(0) - PB(0)
            ya = PA(1) - PB(1)
            
            distancia = Val(Sqr((xa ^ 2 + ya ^ 2)))
            
            If distancia < 10000 Then
                ' Acción para distancias menores a 10
            escalaPapel = 0.03
            textHeight2 = 50
            ElseIf distancia >= 10000 And distancia <= 15000 Then
                ' Acción para distancias entre 10 y 15
            escalaPapel = 0.02
            textHeight2 = 75
            Else
                ' Acción para distancias mayores a 15
            escalaPapel = 0.015
            textHeight2 = 100
            End If
        
Inicio:
        'nombre de las hojas de las pestañas
        layoutName = "50"
        
        Dim num As Integer
        num = 1

        ' Intenta encontrar un nombre único para la pestaña
        Do While LayoutExists(layoutName & CStr(num))
            num = num + 1
        Loop
        
        num = num - 1
        
        existingLayerName = layoutName & CStr(num)
        
        'si no existe la primera 501 crear nueva
        If existingLayerName = "500" Then
        
        Call CrearNuevaPestaña2
        num = num + 1
        
        End If
        
        'seccion asegurar estar en la pestaña activa correcta
        ventanaacti = layoutName & num
        
        Call SeleccionarPestana(ventanaacti)
        
        'contar cuantos elemento ya existen en esa pestaña layout
        Call ContarElementos(num)
        
        Dim nconsecutivo As String
        nconsecutivo = consecutivo
                               
        ' Verificar si es el primer bloque
        If nconsecutivo = "1" Then
            ' Es el primer bloque, insertarlo en el punto especificado
            insertionP(0) = 50#: insertionP(1) = 250#: insertionP(2) = 0#
                    
        ElseIf nconsecutivo = "2" Or nconsecutivo = "3" Or nconsecutivo = "4" Or nconsecutivo = "5" Then
        
            insertionP(0) = 50#: insertionP(1) = 250#: insertionP(2) = 0#
        
            ' Calcular la posición del siguiente bloque en la dirección hacia abajo
            insertionP(0) = insertionP(0)
            insertionP(1) = insertionP(1) - (offsetDistance * (nconsecutivo - 1))
            insertionP(2) = insertionP(2)
        
        'si ya esta el maximo de elementos en el layout crear un nuevo layout
        ElseIf nconsecutivo = "6" Then
        
        Call CrearNuevaPestaña2
        
        GoTo Inicio
                
        End If
                
        'Posicion texto
        insertionT(0) = insertionP(0) - offsetXTexto
        insertionT(1) = insertionP(1) - offsetYTexto
        insertionT(2) = insertionP(2)
                
        k = "ELEMENTO50" & num & nconsecutivo
        
        Dim blockNamedelet As String
    Dim block As Object
    
    ' Especifica el nombre del bloque que deseas eliminar
    blockNamedelet = k
    
    ' Verifica si el bloque existe en la colección de bloques
    If BloqueExiste(blockNamedelet) Then
        ' Elimina el bloque de la colección de bloques
        ThisDrawing.Blocks.Item(blockNamedelet).Delete
        MsgBox "Bloque con el mismo nombre eliminado con éxito.", vbInformation
    End If
        
        ' Crear un nuevo bloque
        Set b = ThisDrawing.Blocks.Add(puntoInicio, k)
        
        ' Obtener el ángulo de las líneas
        orientation = gcadUtil.AngleFromXAxis(PA, PB)
        
        'Giro de direccion 90 grados para hallar perpendicular
        orientationa = orientation - ((pi) / 2)
        
        ' Contar los bloques seleccionados y obtener la capa asociada
        For Each obj In ss
            If TypeOf obj Is GcadBlockReference Then
                Set blk = obj
                blockName = blk.effectiveName
                contador = 0
                                                        
                ' Obtener el punto de inserción del bloque seleccionado
                insertionPoint2(0) = blk.insertionPoint(0)
                insertionPoint2(1) = blk.insertionPoint(1)
                insertionPoint2(2) = blk.insertionPoint(2)

                ' Definir el array con los nombres de los items para acotar
                nombresItems = Array("PS_NUDO_PLANTA", "PS_NUDO_ALZADO", "GS_FUSIBLE_PLANTA", "GS_FUSIBLE_ALZADO", "MSHOR270PLA", "MSHOR270ALZ", "MSHOR180PLA", "MSHOR180ALZ", "MSHOR90PLAFUSIBLE", "MSHOR90PLA", "MSHOR90ALZFUSIBLE", "MSHOR90ALZ", "PS_GATO_PLANTA", "PS_GATO_ALZADO", "PS_PLACA35MM_ALZADO", "PS_PLACA35MM_PLANTA", "PS_PLACA50MM_ALZADO", "PS_PLACA50MM_PLANTA", "P6_CAJON_PLANTA", "P6_CONO", "P6_MACHO_PLANTA", "MGHUSILLOGATO", "CAJONH_ALZADO", "CAJONH_PLANTA", "MSHORJACKPLATE")

                ' Definir el array con los nombres de los items que no apareceran
                nombresItems2 = Array("PLACAMP_C_DALZADO", "PLACAMP_C_IALZADO", "PLACAMP_C_PLANTA", "PLACAMP_C_SECCION", "PLACAMP_DALZADO", "PLACAMP_IALZADO", "PLACAMP_PLANTA", "PLACAMP_SECCION", "GS_GIRO_ALZADO", "GS_GIRO_PLANTA", "GS_PLACAANCLAJE_DALZADO", "GS_PLACAANCLAJE_IALZADO", "GS_PLACAANCLAJE_PLANTA", "GS_PLACACOMPACTA_DALZADO", "GS_PLACACOMPACTA_IALZADO", "GS_PLACACOMPACTA_PLANTA", "GS_PLACACOMPACTA_SECCION", "GS_GIRO80PLANTA", "GS_GIRO80ALZADO", "ANGGIRO", "MG_ANGULOGIROPLA", "MG_ANGULOGIROALZ", "PL_GCODAL_C_PLA", "PL_GCODAL_C_ALZ", "PL_GCODAL_PLA", "PL_GCODAL_ALZ", "MG_CUNAAZ_DALZADO", "MG_CUNAAZ_IALZADO", "MG_CUNAAZ_PLANTA", "MG_CUNANAR_DALZADO", "MG_CUNANAR_IALZADO", "MG_CUNANAR_PLANTA", "HGHUSILLOGATO", "GS_BULON80_ALZADO", "GS_BULON80_PLANTA", "GS_BULON120MM_ALZADO", "GS_BULON120MM_PLANTA")

                     'Verificar si el nombre está en el array
                    If UBound(Filter(nombresItems2, UCase(blockName))) > -1 Then
                    
                    GoTo proximo
                    
                    ElseIf UCase(Mid(blockName, 2, 3)) = "VAR" Then
                        ' Inicializar el contador para este conjunto de bloques VAR
                                             
                        ' Bucle interno para contar bloques VAR en el mismo punto de inserción
                        For Each obj2 In ss
                            If TypeOf obj2 Is GcadBlockReference Then
                                Set blk2 = obj2
                                blockName2 = blk2.effectiveName
                                
                                Dim insertionPoint3(0 To 2) As Double
                                
                                ' Obtener el punto de inserción del bloque seleccionado
                                insertionPoint3(0) = blk2.insertionPoint(0)
                                insertionPoint3(1) = blk2.insertionPoint(1)
                                insertionPoint3(2) = blk2.insertionPoint(2)

                                ' Comprobar si el nombre del bloque y el punto de inserción coinciden
                                If UCase(Mid(blockName2, 2, 3)) = "VAR" And insertionPoint2(0) = insertionPoint3(0) And insertionPoint2(1) = insertionPoint3(1) Then
                                    contador = contador + 1
                                End If
                            End If
                        Next obj2
                        
                        ' Agregar un MLeader en el punto de inserción del conjunto de bloques
                        
                        'Quitar el 1 del nombre
                            textoSinUltimasSeisLetras = contador & Right(blockName, Len(blockName) - 1)
                                                                                
                            ' Calcular las coordenadas del punto final del MLeader
                            leaderEndPoint(0) = insertionPoint2(0) + 400 * Cos(orientationa)
                            leaderEndPoint(1) = insertionPoint2(1) + 400 * Sin(orientationa)
                            leaderEndPoint(2) = insertionPoint2(2)
                            
                            textEndPoint(0) = leaderEndPoint(0) - 100 * Cos(orientation)
                            textEndPoint(1) = leaderEndPoint(1) - 100 * Sin(orientation)
                            textEndPoint(2) = leaderEndPoint(2)
                            
                            textEndPoint(0) = textEndPoint(0) + 100 * Cos(orientationa)
                            textEndPoint(1) = textEndPoint(1) + 100 * Sin(orientationa)
                            textEndPoint(2) = textEndPoint(2)

                            ' Calcular el ángulo de rotación en radianes (90 grados)
                            rotationAngle = pi * 60 / 180
                                                        
                            Set newText2 = ThisDrawing.Blocks.Item(k).AddMText(textEndPoint, textHeight2, textoSinUltimasSeisLetras)
                                newText2.Rotate textEndPoint, orientation
                                newText2.Rotate textEndPoint, rotationAngle
                                'newText2.StyleName = "MODELO"
                                'newText2.TextStyleName = "SIMPLEX"
                                newText2.Layer = "Dimension"
                                newText2.AttachmentPoint = acAttachmentPointTopRight
                                newText2.height = textHeight2
                                newText2.color = 250
                            
                            ' Crear una nueva línea con los puntos de inicio y final
                              Set newLine = ThisDrawing.Blocks.Item(k).AddLine(insertionPoint2, leaderEndPoint)
                              newLine.Layer = "Dimension"
                               
                    End If

                    ' Verificar si el nombre está en el array
                    If UBound(Filter(nombresItems, UCase(blockName))) > -1 Then
                    
                        ' Verificar si las últimas cinco letras del nombre del bloque son "planta"
                        If UCase(Right(blockName, 6)) = "PLANTA" Then
                        
                            ' Obtener el texto de blockName sin las últimas cinco letras
                            textoSinUltimasSeisLetras = Left(blockName, Len(blockName) - 6)
                        
                            blockName = textoSinUltimasSeisLetras & "alzado.dwg"
                                                                            
                                If UCase(Left(blockName, 2)) = "PS" Or UCase(Left(blockName, 3)) = "ZPS" Then
                                
                                ps = rutaps & blockName
                                
                                    If UCase(textoSinUltimasSeisLetras) = "PS_PLACA35MM_" Then
                                        textoSinUltimasSeisLetras = "Chapón 35"
                                    ElseIf UCase(textoSinUltimasSeisLetras) = "PS_PLACA50MM_" Then
                                        textoSinUltimasSeisLetras = "Chapón 50"
                                    ElseIf UCase(textoSinUltimasSeisLetras) = "PS_GATO_" Then
                                        textoSinUltimasSeisLetras = "Gato PS"
                                    ElseIf UCase(textoSinUltimasSeisLetras) = "PS_NUDO_" Then
                                        textoSinUltimasSeisLetras = "Nudo PS"
                                    End If
                                                               
                                ElseIf UCase(Left(blockName, 2)) = "GS" Or UCase(Left(blockName, 3)) = "ZGS" Then
                                
                                ps = rutags & blockName
                                    
                                    If UCase(textoSinUltimasSeisLetras) = "GS_FUSIBLE_" Then
                                        textoSinUltimasSeisLetras = "Fusible GS"
                                    End If
                                
                                ElseIf UCase(Left(blockName, 2)) = "CA" Then
                                
                                ps = rutacajon & blockName
                                
                                    If UCase(textoSinUltimasSeisLetras) = "CAJONH_" Then
                                        textoSinUltimasSeisLetras = "Cajón Pipeshor"
                                    End If
                                                            
                                End If
                                
                                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(insertionPoint2, ps, 1#, 1#, 1#, blk.Rotation)
                                blockRef.Layer = "0"
                                
                                If textoSinUltimasSeisLetras = "Cajón Pipeshor" Or textoSinUltimasSeisLetras = "Nudo PS" Then
                                        'mover texto al medio del cajon
                                        insertionPoint2(0) = insertionPoint2(0) - (300 * Cos(orientation))
                                        insertionPoint2(1) = insertionPoint2(1) - (300 * Sin(orientation))
                                        insertionPoint2(2) = insertionPoint2(2)
                                        
                                ElseIf textoSinUltimasSeisLetras = "Fusible GS" Then
                                        'mover texto al medio del fusible
                                        insertionPoint2(0) = insertionPoint2(0) + (90 * Cos(orientation))
                                        insertionPoint2(1) = insertionPoint2(1) + (90 * Sin(orientation))
                                        insertionPoint2(2) = insertionPoint2(2)
                                End If
                                
                                Call Textsup(orientation, orientationa, insertionPoint2, textoSinUltimasSeisLetras, k, textHeight2)
                        
                        ElseIf UCase(Right(blockName, 3)) = "PLA" Then
                        
                            ' Obtener el texto de blockName sin las últimas cinco letras
                            textoSinUltimasSeisLetras = Left(blockName, Len(blockName) - 3)
                        
                            blockName = textoSinUltimasSeisLetras & "ALZ.dwg"
                            
                                If UCase(Left(blockName, 2)) = "MS" Then
                                
                                ps = rutamp & blockName
                                
                                    If UCase(textoSinUltimasSeisLetras) = "MSHOR90" Then
                                        textoSinUltimasSeisLetras = "MP 90"
                                    ElseIf UCase(textoSinUltimasSeisLetras) = "MSHOR180" Then
                                        textoSinUltimasSeisLetras = "MP 180"
                                    ElseIf UCase(textoSinUltimasSeisLetras) = "MSHOR270" Then
                                        textoSinUltimasSeisLetras = "MP 270"
                                    End If
                                                                                            
                                End If
                                
                                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(insertionPoint2, ps, 1#, 1#, 1#, blk.Rotation)
                                blockRef.Layer = "0"
                                
                                Call Textsup(orientation, orientationa, insertionPoint2, textoSinUltimasSeisLetras, k, textHeight2)
                                                    
                        Else
                            ' Agregar el bloque seleccionado al nuevo bloque con escala
                            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(insertionPoint2, blk.effectiveName, 1#, 1#, 1#, blk.Rotation)
                            blockRef.Layer = "0"
                            
                                    If UCase(blockName) = "PS_PLACA35MM_ALZADO" Then
                                        textoSinUltimasSeisLetras = "Chapón 35"
                                    ElseIf UCase(blockName) = "PS_PLACA50MM_ALZADO" Then
                                        textoSinUltimasSeisLetras = "Chapón 50"
                                    ElseIf UCase(blockName) = "PS_GATO_ALZADO" Then
                                        textoSinUltimasSeisLetras = "Gato PS"
                                    ElseIf UCase(blockName) = "CAJONH_ALZADO" Then
                                        textoSinUltimasSeisLetras = "Cajón Pipeshor"
                                        
                                         'mover texto al medio del cajon
                                        insertionPoint2(0) = insertionPoint2(0) - (300 * Cos(orientation))
                                        insertionPoint2(1) = insertionPoint2(1) - (300 * Sin(orientation))
                                        insertionPoint2(2) = insertionPoint2(2)
                                        
                                    ElseIf UCase(blockName) = "GS_FUSIBLE_ALZADO" Then
                                        textoSinUltimasSeisLetras = "Fusible Gs"
                                        
                                        'mover texto al medio del fusible
                                        insertionPoint2(0) = insertionPoint2(0) + (90 * Cos(orientation))
                                        insertionPoint2(1) = insertionPoint2(1) + (90 * Sin(orientation))
                                        insertionPoint2(2) = insertionPoint2(2)
                                        
                                    ElseIf UCase(blockName) = "MGHUSILLOGATO" Then
                                        textoSinUltimasSeisLetras = "Gato MP"
                                    ElseIf UCase(blockName) = "MSHOR90ALZ" Then
                                        textoSinUltimasSeisLetras = "MP 90"
                                    ElseIf UCase(blockName) = "MSHOR90ALZFUSIBLE" Then
                                        textoSinUltimasSeisLetras = "Fusible MP"
                                    ElseIf UCase(blockName) = "MSHOR180ALZ" Then
                                        textoSinUltimasSeisLetras = "MP 180"
                                    ElseIf UCase(blockName) = "MSHOR90PLAFUSIBLE" Then
                                        textoSinUltimasSeisLetras = "Fusible MP"
                                    ElseIf UCase(blockName) = "MSHOR270ALZ" Then
                                        textoSinUltimasSeisLetras = "MP 270"
                                    ElseIf UCase(blockName) = "MSHORJACKPLATE" Then
                                        textoSinUltimasSeisLetras = "JackPlate"
                                    ElseIf UCase(blockName) = "P6_CONO" Then
                                        textoSinUltimasSeisLetras = "Cono P6"
                                    ElseIf UCase(blockName) = "PS_NUDO_ALZADO" Then
                                        textoSinUltimasSeisLetras = "Nudo PS"
                                         'mover texto al medio del nudo
                                        insertionPoint2(0) = insertionPoint2(0) - (300 * Cos(orientation))
                                        insertionPoint2(1) = insertionPoint2(1) - (300 * Sin(orientation))
                                        insertionPoint2(2) = insertionPoint2(2)
                                    End If
                                    
                                Call Textsup(orientation, orientationa, insertionPoint2, textoSinUltimasSeisLetras, k, textHeight2)
                            
                        End If
                    
                    'verificar si el bloque es tornillo para acotarlo
                    ElseIf UCase(Mid(blockName, 2, 1)) = "M" Or UCase(Mid(blockName, 3, 1)) = "M" Or UCase(Mid(blockName, 4, 1)) = "M" Then
                    
                        If UCase(Left(blockName, 1)) = "Z" Then
                        GoTo basegato
                        ElseIf UCase(Left(blockName, 1)) = "0" Then
                        ' Obtener el texto de blockName sin la primera letras
                        textoSinUltimasSeisLetras = Right(blockName, Len(blockName) - 1)
                        Else
                        textoSinUltimasSeisLetras = blockName
                        End If

                    ' Agregar el bloque seleccionado al nuevo bloque con escala
                    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(insertionPoint2, blk.effectiveName, 1#, 1#, 1#, blk.Rotation)
                    blockRef.Layer = "0"
                    
                    ' Calcular las coordenadas del punto final del MLeader
                            leaderEndPoint(0) = insertionPoint2(0) + 400 * Cos(orientationa)
                            leaderEndPoint(1) = insertionPoint2(1) + 400 * Sin(orientationa)
                            leaderEndPoint(2) = insertionPoint2(2)
                            
                            textEndPoint(0) = leaderEndPoint(0) - 100 * Cos(orientation)
                            textEndPoint(1) = leaderEndPoint(1) - 100 * Sin(orientation)
                            textEndPoint(2) = leaderEndPoint(2)
                            
                            textEndPoint(0) = textEndPoint(0) + 100 * Cos(orientationa)
                            textEndPoint(1) = textEndPoint(1) + 100 * Sin(orientationa)
                            textEndPoint(2) = textEndPoint(2)

                            ' Calcular el ángulo de rotación en radianes (90 grados)
                            rotationAngle = pi * 60 / 180
                                                        
                            Set newText2 = ThisDrawing.Blocks.Item(k).AddMText(textEndPoint, textHeight2, textoSinUltimasSeisLetras)
                                newText2.Rotate textEndPoint, orientation
                                newText2.Rotate textEndPoint, rotationAngle
                                'newText2.StyleName = "MODELO"
                                'newText2.TextStyleName = "SIMPLEX"
                                newText2.Layer = "Dimension"
                                newText2.AttachmentPoint = acAttachmentPointTopRight
                                newText2.height = textHeight2
                                newText2.color = 250
                            
                            ' Crear una nueva línea con los puntos de inicio y final
                              Set newLine = ThisDrawing.Blocks.Item(k).AddLine(insertionPoint2, leaderEndPoint)
                              newLine.Layer = "Dimension"
                    
                    ' Verificar si las últimas cinco letras del nombre del bloque son "planta"
                    ElseIf UCase(Right(blockName, 6)) = "PLANTA" Then
                    
                        ' Obtener el texto de blockName sin las últimas cinco letras
                        textoSinUltimasSeisLetras = Left(blockName, Len(blockName) - 6)
                    
                        blockName = textoSinUltimasSeisLetras & "alzado.dwg"
                            
                            If UCase(Left(blockName, 2)) = "PL" Or UCase(Left(blockName, 3)) = "ZPL" Then
                            
                            ps = rutapl & blockName
                            
                            ElseIf UCase(Left(blockName, 2)) = "PS" Or UCase(Left(blockName, 3)) = "ZPS" Then
                            
                            ps = rutaps & blockName
                            
                            ElseIf UCase(Left(blockName, 2)) = "P6" Then
                            
                            ps = rutap6 & blockName
                            
                            ElseIf UCase(Left(blockName, 2)) = "GS" Or UCase(Left(blockName, 3)) = "ZGS" Then
                            
                            ps = rutags & blockName
                            
                            ElseIf UCase(Left(blockName, 2)) = "CA" Or UCase(Left(blockName, 2)) = "MO" Then
                            
                            ps = rutacajon & blockName
                                                        
                            ElseIf UCase(Left(blockName, 2)) = "MG" Or UCase(Left(blockName, 4)) = "PL_G" Or UCase(Left(blockName, 4)) = "PLAC" Or UCase(Left(blockName, 4)) = "SOPO" Then
                            
                            ps = rutacu & blockName
                                                        
                            End If
                            
                            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(insertionPoint2, ps, 1#, 1#, 1#, blk.Rotation)
                            blockRef.Layer = "0"
                    
                    ElseIf UCase(Right(blockName, 3)) = "PLA" Then
                    
                        ' Obtener el texto de blockName sin las últimas cinco letras
                        textoSinUltimasSeisLetras = Left(blockName, Len(blockName) - 3)
                    
                        blockName = textoSinUltimasSeisLetras & "ALZ.dwg"
                        
                            If UCase(Left(blockName, 2)) = "MS" Then
                            
                            ps = rutamp & blockName
                                                        
                            ElseIf UCase(Left(blockName, 2)) = "MG" Or UCase(Left(blockName, 4)) = "PL_G" Then
                            
                            ps = rutacu & blockName
                                                        
                            End If
                            
                            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(insertionPoint2, ps, 1#, 1#, 1#, blk.Rotation)
                            blockRef.Layer = "0"
                                                
                    Else
basegato:
                        ' Agregar el bloque seleccionado al nuevo bloque con escala
                        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(insertionPoint2, blk.effectiveName, 1#, 1#, 1#, blk.Rotation)
                        blockRef.Layer = "0"
                        
                    End If
                                 
            ElseIf TypeOf obj Is GcadMText Or TypeOf obj Is GcadText Then
                Set textEntity = obj
                textToInsert = textEntity.TextString
                
                ' Configurar capa y color para el nuevo texto
                Set textObj = ThisDrawing.PaperSpace.AddText(textToInsert, insertionT, textHeight)
                textObj.Layer = "0"
                textObj.color = acRed
                
            End If
            
proximo:
            
        Next obj
        
        ' Insertar el nuevo bloque en el espacio de papel con escala menor
        Set blockRef = ThisDrawing.PaperSpace.InsertBlock(insertionP, b.Name, escalaPapel, escalaPapel, escalaPapel, -orientation)
        blockRef.Layer = "0"
        
        ' Eliminar el conjunto de selección
        ss.Delete
        
        ' Regenerar el espacio modelo
        ThisDrawing.Regen acAllViewports

    Else
       ' MsgBox "Por favor, cambia al espacio modelo y vuelve a intentarlo."
        
        GoTo terminar
        
    End If
    
clear:

' Liberar objetos y recursos
    Set ss = Nothing
    Set b = Nothing
    Set blockRef = Nothing
    Set textEntity = Nothing
    Set textObj = Nothing
    
    Loop
    
terminar:

' Liberar objetos y recursos
    Set ss = Nothing
    Set b = Nothing
    Set blockRef = Nothing
    Set textEntity = Nothing
    Set textObj = Nothing

End Sub

Sub ContarElementos(num As Integer)
    Dim n As Integer
    Dim i As Integer
    Dim myVal As GcadBlockReference
    Dim blockNameOnPaper As String
    Dim currentConsecutivo As Integer
    Dim existingConsecutivo As Integer

    n = ThisDrawing.PaperSpace.Count
    existingConsecutivo = 0

    For i = 1 To n
        On Error Resume Next
        Set myVal = ThisDrawing.PaperSpace.Item(i)
        On Error GoTo 0

        If Not myVal Is Nothing Then
            If TypeOf myVal Is GcadBlockReference Then
                blockNameOnPaper = myVal.effectiveName

                If Left(blockNameOnPaper, Len("ELEMENTO50" & num)) = "ELEMENTO50" & num Then
                    currentConsecutivo = Val(Mid(blockNameOnPaper, Len("ELEMENTO50" & num) + 1))

                    If currentConsecutivo >= existingConsecutivo Then
                        existingConsecutivo = currentConsecutivo
                    End If
                End If
            End If
        End If
    Next i

    consecutivo = existingConsecutivo + 1

   ' MsgBox "Número consecutivo más alto encontrado: " & existingConsecutivo
    
    'MsgBox "Nuevo número consecutivo a utilizar: " & consecutivo
End Sub

Function LayoutExists(layoutName As String) As Boolean
    ' Verifica si la pestaña con el nombre especificado ya existe
    Dim layout As GcadLayout
    For Each layout In ThisDrawing.Layouts
        If layout.Name = layoutName Then
            LayoutExists = True
            Exit Function
        End If
    Next layout
    LayoutExists = False
End Function

Sub CrearNuevaPestaña2()

    Dim layoutName As String
    Dim newLayoutName As String
    Dim newLayout As GcadLayout
    Dim bloqueRuta As String
    Dim gcadDoc As GcadDocument
    Dim GcadPaper As GcadPaperSpace
    Dim P(0 To 2) As Double
    Dim Xs As Double
    Dim Ys As Double
    Dim Zs As Double
    Dim blockRef As Object
    Dim plotconfig As GcadPlotConfiguration
    Dim nombreArchivo As String
    Dim primeros11Caracteres As String
    Dim numobra As String
    Dim x2x As String, cod As String, del As String
    Xs = 1
    Ys = 1
    Zs = 1

    ' Obtener el nombre del archivo DWG actual
    nombreArchivo = ThisDrawing.Name

    ' Obtener los primeros 11 caracteres
    primeros11Caracteres = Left(nombreArchivo, 11)
    
    If Right(primeros11Caracteres, 1) = "T" Then
        numobra = primeros11Caracteres
        del = Right(numobra, 3)
        x2x = Left(numobra, 2)
        cod = Left(numobra, 8)
        cod = Right(cod, 6)
    Else
        numobra = Left(primeros11Caracteres, 10)
        If Right(numobra, 1) = "R" Then
            numobra = numobra
            del = Right(numobra, 2)
            x2x = Left(numobra, 2)
            cod = Left(numobra, 8)
            cod = Right(cod, 6)
        Else
            numobra = Left(numobra, 9)
            del = Right(numobra, 2)
            x2x = Left(numobra, 2)
            cod = Left(numobra, 8)
            cod = Right(cod, 6)
        End If
    End If

    ' Especifica el nombre base para la pestaña
    layoutName = "50"
    bloqueRuta = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\bloque_cajetin.dwg"
    P(0) = -9.88: P(1) = 0.78: P(2) = 0

    ' Inicializa el número de pestaña
    Dim num As Integer
    num = 1

    ' Intenta encontrar un nombre único para la pestaña
    Do While LayoutExists(layoutName & CStr(num))
        num = num + 1
    Loop

    ' Construye el nuevo nombre de la pestaña
    newLayoutName = layoutName & CStr(num)

    ' Crea la nueva pestaña
    Set newLayout = ThisDrawing.Layouts.Add(newLayoutName)
    
    ThisDrawing.ActiveLayout = newLayout
    newLayout.CanonicalMediaName = "ISO_A3_(297.00_x_420.00_MM)"

    ' Obtener el documento activo y el espacio papel
    Set gcadDoc = GetObject(, "Gcad.Application").ActiveDocument
    Set GcadPaper = gcadDoc.PaperSpace
    Set blockRef = ThisDrawing.PaperSpace.InsertBlock(P, bloqueRuta, Xs, Ys, Zs, 0)
    'blockRef.Explode

    ThisDrawing.Regen True
    
    n = ThisDrawing.PaperSpace.Count

    For r = 0 To n - 1
        Set element = ThisDrawing.PaperSpace.Item(r)
        
        If element.ObjectName = "AcDbBlockReference" Then
            Set block = element
            If block.effectiveName = "description3" Then
                varatt = block.GetAttributes
                For j = LBound(varatt) To UBound(varatt)
                    Set att = varatt(j)
                    If att.TagString = "EQUIPMENT" Then
                        att.TextString = "PLANO DE PRE-MONTAJE"
                    ElseIf att.TagString = "STRUCTURE" Then
                        att.TextString = layoutName & CStr(num)
                    ElseIf att.TagString = "2X" Then
                        att.TextString = x2x
                    ElseIf att.TagString = "COD" Then
                        att.TextString = cod
                    ElseIf att.TagString = "DEL" Then
                        att.TextString = del
                    End If
                Next
            End If
        End If
    Next
        ' Mensaje de confirmación
    MsgBox "Se ha creado una nueva pestaña llamada " & newLayoutName

End Sub

Sub SeleccionarPestana(ventanaacti As Integer)
    Dim layoutName As String
    Dim targetLayoutName As String
        
    ' Número de la pestaña objetivo
    Dim targetLayoutNumber As Integer
    targetLayoutNumber = ventanaacti
    
    ' Construye el nombre de la pestaña objetivo
    targetLayoutName = CStr(targetLayoutNumber)
    
    ' Intenta encontrar la pestaña por nombre
    Dim layout As GcadLayout
    For Each layout In ThisDrawing.Layouts
        If layout.Name = targetLayoutName Then
            ' Establece la pestaña activa
            ThisDrawing.ActiveLayout = layout
            Exit Sub
        End If
    Next layout

End Sub

Function BloqueExiste(blockNamedelet As String) As Boolean
    ' Función para verificar si un bloque existe en la colección
    Dim blk As Object
    BloqueExiste = False
    
    ' Itera a través de la colección de bloques
    For Each blk In ThisDrawing.Blocks
        If UCase(blk.Name) = UCase(blockNamedelet) Then
            ' Encuentra el bloque con el nombre especificado
            BloqueExiste = True
            Exit Function
        End If
    Next blk
End Function

Sub Textsup(orientation As Double, orientationa As Double, insertionPoint2() As Double, textoSinUltimasSeisLetras As String, k As String, textHeight2 As Double)

    Dim leaderEndPoint(0 To 2) As Double
    Dim rotationAngle As Double
    Dim newText As GcadObject
    Dim newLine As Object
    
    Dim pi As Double
    pi = 4 * Atn(1)

    ' Calcular las coordenadas del punto final del MLeader
    leaderEndPoint(0) = insertionPoint2(0) - 400 * Cos(orientationa)
    leaderEndPoint(1) = insertionPoint2(1) - 400 * Sin(orientationa)
    leaderEndPoint(2) = insertionPoint2(2)

    ' Calcular el ángulo de rotación en radianes (90 grados)
    rotationAngle = pi * 60 / 180
                                                            
    Set newText = ThisDrawing.Blocks.Item(k).AddText(textoSinUltimasSeisLetras, leaderEndPoint, textHeight2)
    newText.Rotate leaderEndPoint, orientation
    newText.Rotate leaderEndPoint, rotationAngle
    'newText.StyleName = "MODELO"
    'newText.TextStyleName = "SIMPLEX"
    newText.Layer = "Dimension"
    newText.color = 250
                                
    ' Crear una nueva línea con los puntos de inicio y final
    Set newLine = ThisDrawing.Blocks.Item(k).AddLine(insertionPoint2, leaderEndPoint)
    newLine.Layer = "Dimension"
    
End Sub




