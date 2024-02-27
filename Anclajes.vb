Option Explicit

Sub an()
    Dim gcadDoc As Object, gcadUtil As Object, gcadModel As Object, blockRef As Object
    Dim rutator As String, capa As String, referencia As String, bloque As String, Respuesta As String
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

    kwordList = "Varillas Tornillos"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    bloque = ThisDrawing.Utility.GetKeyword(vbLf & "Tornillo: [Varillas/Tornillos]")

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

    ElseIf bloque = "Tornillos" Then
        kwordList = "A-M20X40 B-M20X60 C-M20X80{12.9} D-M20X80{12.9}RP E-M20X90 F-M20X130 G-M30X100{10.9}"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        referencia = ThisDrawing.Utility.GetKeyword(vbLf & "Referencia: [A-M20X40/B-M20X60/C-M20X80{12.9}/D-M20X80{12.9}RP/E-M20X90/F-M20X130/G-M30X100{10.9}]")

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

            bloque = rutator & numero & Right(referencia, Len(referencia) - 1) & ".dwg"

            If Dir(bloque) <> "" Then  ' Verificar si el archivo del tornillo existe
                Set blockRef = gcadModel.InsertBlock(P1, bloque, Xs, Ys, Zs, 0)
                blockRef.Layer = "Nonplot"
                blockRef.Update
                blockRef.Explode
                blockRef.Delete
            Else
                MsgBox "El archivo del Tornillo " & numero & Right(referencia, Len(referencia) - 1) & " no existe " & vbCrLf & "Se creara un nuevo archivo", vbExclamation, "Error"
                
                ' Ruta del archivo que deseas abrir
                Dim rutaArchivo As String
                rutaArchivo = rutator & "1" & Right(referencia, Len(referencia) - 1) & ".dwg"

                ' Ruta y nombre del nuevo archivo
                Dim rutaNuevoArchivo As String
                rutaNuevoArchivo = rutator & numero & Right(referencia, Len(referencia) - 1) & ".dwg"

                ' Abrir el archivo existente
                Documents.Open rutaArchivo
                
                Dim nombreBloqueExistente As String
                nombreBloqueExistente = "1" & Right(referencia, Len(referencia) - 2)

                ' Nuevo nombre para el bloque existente
                Dim nuevoNombre As String
                nuevoNombre = numero & Right(referencia, Len(referencia) - 2)

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
                            Call ModPropAtri(acBlock, CStr(numero & Right(referencia, Len(referencia) - 1)))
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

