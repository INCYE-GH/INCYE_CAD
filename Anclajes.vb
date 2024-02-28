Option Explicit

Sub an()
    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, blockRef As Object
    Dim rutator As String, capa As String, referencia As String, metrica As String, bloque As String, Respuesta As String, rutabul19 As String, rutabul23 As String
    Dim punto1 As Variant
    Dim Gcapa As Object
    Dim Ncapa As String, kwordList As String
    Dim repite As Double, ANG As Double, Xs As Double, Ys As Double, Zs As Double, P1(0 To 2) As Double
    Dim numero As Variant
    Dim i As Integer
    
    Dim acObject As GcadObject
    Dim acBlock As GcadBlockReference
    Dim numElements As Integer
    
    Set gcadDoc = GetObject(, "gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility

    Ncapa = "Nonplot"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 2

    On Error GoTo terminar

    rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"

    kwordList = "Varillas Tornillos Bulon"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    bloque = ThisDrawing.Utility.GetKeyword(vbLf & "Tornillo: [Varillas/Tornillos/Bulon]")

    If bloque = "Varillas" Then
        kwordList = "A-VarM12X100 B-VarM16X166 C-VarM16X200 D-VarM16X250 E-VarM16X500 F-VarM16X200{12.9} G-VarM20X200 H-VarM20X250 I-VarM20X500 J-VarM20X250{12.9} K-VarM24X200 L-VarM24X250 M-VarM30X250 N-VarM30X250{12.9} O-VarM30X330{12.9}"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        referencia = ThisDrawing.Utility.GetKeyword(vbLf & "Varillas: [A-VarM12X100/B-VarM16X166/C-VarM16X200/D-VarM16X250/E-VarM16X500/F-VarM16X200{12.9}/G-VarM20X200/H-VarM20X250/I-VarM20X500/J-VarM20X250{12.9}/K-VarM24X200/L-VarM24X250/M-VarM30X250/N-VarM30X250{12.9}/O-VarM30X330{12.9}]")

        bloque = rutator & "1" & Right(referencia, Len(referencia) - 2) & ".dwg"
        numero = "a"
        Do While IsNumeric(numero) = False
            numero = InputBox("Introduce el número de unidades", "NÚMERO DE UNIDADES", "1", 750, 750, "DEMO.HLP", 10)
            If IsNumeric(numero) = False Then Respuesta = MsgBox("Debes introducir un número entero, inténtalo de nuevo", , "ERROR", "DEMO.HLP", 750)
        Loop

        repite = 1
        Do While repite = 1
            punto1 = gcadUtil.GetPoint(, "1º Punto: ")
            P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)

            Xs = 1
            Ys = 1
            Zs = 1
            i = 0

            Do While i < numero
                If Dir(bloque) <> "" Then  ' Verificar si el archivo del bloque existe
                    Set blockRef = gcadModel.InsertBlock(P1, bloque, Xs, Ys, Zs, 0)
                    blockRef.Layer = "Nonplot"
                Else
                    MsgBox "El archivo del bloque no existe: " & bloque, vbExclamation, "Error"
                    Exit Do
                End If
                i = i + 1
            Loop
        Loop

	ElseIf bloque = "Bulon" Then
		kwordList = "A-M19X180 B-M23X160"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        referencia = ThisDrawing.Utility.GetKeyword(vbLf & "Metrica: [A-M19X180/B-M23X160]")
		
		repite = 1
		Do While repite = 1
            punto1 = gcadUtil.GetPoint(, "1º Punto: ")
            P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)

            Xs = 1
            Ys = 1
            Zs = 1
			
			rutabul19 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\1M19_BULOND19.dwg"
			rutabul23 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\1M23_BULOND23.dwg"
		
		If Referencia = "A-M19X180" Then
		Set blockRef = gcadModel.InsertBlock(P1, rutabul19, Xs, Ys, Zs, 0) 
            blockRef.Layer = "Nonplot"
			
		ElseIf Referencia = "B-M23X160" Then
		Set blockRef = gcadModel.InsertBlock(P1, rutabul23, Xs, Ys, Zs, 0) 
            blockRef.Layer = "Nonplot"
			
		End If
		loop

    ElseIf bloque = "Tornillos" Then
        kwordList = "A-M16 B-M20 C-M24 D-M30"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        referencia = ThisDrawing.Utility.GetKeyword(vbLf & "Metrica: [A-M16/B-M20/C-M24/D-M30]")

            If referencia = "A-M16" Then
            kwordList = "A-M16X40 B-M16X40A C-M16X60 D-M16X90 E-M16X110 F-M16X150"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            metrica = ThisDrawing.Utility.GetKeyword(vbLf & "Referencia: [A-M16X40/B-M16X40A/C-M16X60/D-M16X90/E-M16X110/F-M16X150]")
            
            ElseIf referencia = "B-M20" Then
            kwordList = "A-M20X40 B-M20X50 C-M20X60 D-M20X60A E-M20X80A F-M20X80{12.9} G-M20X80{12.9}RP H-M20X90  I-M20X110 J-M20X130 K-M20X150 L-M20X160{12.9}A"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            metrica = ThisDrawing.Utility.GetKeyword(vbLf & "Referencia: [A-M20X40/B-M20X50/C-M20X60/D-M20X60A/E-M20X80A/F-M20X80{12.9}/G-M20X80{12.9}RP/H-M20X90/I-M20X110/J-M20X130/K-M20X150/L-M20X160{12.9}A]")
            
            ElseIf referencia = "C-M24" Then
            kwordList = "A-M24x60 B-M24x110"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            metrica = ThisDrawing.Utility.GetKeyword(vbLf & "Referencia: [A-M24x60/B-M24x110]")
            
            ElseIf referencia = "D-M30" Then
            kwordList = "A-M30X100 B-M30X100{10.9} C-M30X150{10.9}"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            metrica = ThisDrawing.Utility.GetKeyword(vbLf & "Referencia: [A-M30X100/B-M30X100{10.9}/C-M30X150{10.9}]")
            
            End If
            
        numero = "a"
        Do While IsNumeric(numero) = False
            numero = InputBox("Introduce el número de unidades", "NÚMERO DE UNIDADES", "1", 750, 750, "DEMO.HLP", 10)
            If IsNumeric(numero) = False Then Respuesta = MsgBox("Debes introducir un número entero, inténtalo de nuevo", , "ERROR", "DEMO.HLP", 750)
        Loop

        repite = 1
        Do While repite = 1
            punto1 = gcadUtil.GetPoint(, "1º Punto: ")
            P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)

            Xs = 1
            Ys = 1
            Zs = 1

            bloque = rutator & numero & Right(metrica, Len(metrica) - 1) & ".dwg"

            If Dir(bloque) <> "" Then  ' Verificar si el archivo del tornillo existe
                Set blockRef = gcadModel.InsertBlock(P1, bloque, Xs, Ys, Zs, 0)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Else
                'MsgBox "El archivo del Tornillo " & numero & Right(metrica, Len(metrica) - 1) & " no existe " & vbCrLf & "Se creara un nuevo archivo", vbExclamation, "Error"
                
                ' Ruta del archivo que deseas abrir
                Dim rutaArchivo As String
                rutaArchivo = rutator & "1" & Right(metrica, Len(metrica) - 1) & ".dwg"

                ' Ruta y nombre del nuevo archivo
                Dim rutaNuevoArchivo As String
                rutaNuevoArchivo = rutator & numero & Right(metrica, Len(metrica) - 1) & ".dwg"

                ' Abrir el archivo existente
                Documents.Open rutaArchivo
                
                Dim nombreBloqueExistente As String
                nombreBloqueExistente = "1" & Right(metrica, Len(metrica) - 2)

                ' Nuevo nombre para el bloque existente
                Dim nuevoNombre As String
                nuevoNombre = numero & Right(metrica, Len(metrica) - 2)

                ' Modificar el nombre del bloque existente
                Dim blk As gcadBlock
                Set blk = ThisDrawing.Blocks.Item(nombreBloqueExistente)
        
                ' Cambiar el nombre del bloque existente
                blk.Name = nuevoNombre
                
                ' Cambiar nombre o atributo a bloque
                numElements = ThisDrawing.ModelSpace.Count

                For i = 0 To numElements - 1
                    Set acObject = ThisDrawing.ModelSpace.Item(i)
                    If acObject.ObjectName = "AcDbBlockReference" Then
                        Set acBlock = acObject
                        If acBlock.effectiveName = nuevoNombre Then
                            Call ModPropAtri(acBlock, CStr(numero & Right(metrica, Len(metrica) - 1)))
                        End If
                    End If
                Next i

               ' MsgBox "Nombre del bloque modificado a: " & nuevoNombre
                
                ' Guardar con un nuevo nombre
                ActiveDocument.SaveAs rutaNuevoArchivo

                ' Cerrar el archivo original sin guardar cambios
                ActiveDocument.Close False

                'MsgBox "Proceso completado."
                
                Set blockRef = gcadModel.InsertBlock(P1, bloque, Xs, Ys, Zs, 0)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
                
            End If
        Loop
    End If

terminar:
End Sub

Sub ModPropAtri(acBlock As GcadBlockReference, strText As String)
    Dim NameTag As String
    Dim valueTag As String
    Dim arrayAttributes As Variant
    Dim acAttribute As GcadAttributeReference
    Dim i As Integer

    arrayAttributes = acBlock.GetAttributes

    For i = LBound(arrayAttributes) To UBound(arrayAttributes)
        Set acAttribute = arrayAttributes(i)
        NameTag = acAttribute.TagString
        If NameTag = "TORNILLO" Then
            acAttribute.TextString = strText
        End If
    Next i
End Sub



