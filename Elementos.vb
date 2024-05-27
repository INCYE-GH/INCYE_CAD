Option Explicit

Sub el()

Dim GcadDoc As Object, GcadUtil As Object, GcadModel As Object, Eje1 As Object, blockRef As Object
Dim rutags As String, rutamp As String, rutator As String, rutaperf As String, rutampacc As String, rutacuña As String, capa As String, rutass As String, rutaps As String, rutapl As String, tipo As String, naranja As String, plalz As String, rutap6 As String
Dim plalzslim As String, plalzmp As String, slim As String, mp As String, ps As String, gs As String
Dim Gcapa As Object
Dim Ncapa As String, cuña As String, lado As String, disposicion As String, kwordList As String, M20x90_4 As String
Dim repite As Double, ANG As Double, x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, P1(0 To 2) As Double, P2(0 To 2) As Double, Punto_inial(0 To 2) As Double
Dim punto1 As Variant, punto2 As Variant, PI As Variant
Dim perf As String, jr As String, perfil As String, P As String
Dim lon As String, alma450 As String, alma As String, junta As String, plalzp As String


Set GcadDoc = GetObject(, "Gcad.Application").ActiveDocument
Set GcadModel = GcadDoc.ModelSpace
Set GcadUtil = GcadDoc.Utility

Ncapa = "Mega"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Granshor"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150
Ncapa = "Pipeshor4S"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor6"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5
Ncapa = "Slims"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30


On Error GoTo terminar

rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"
rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
rutampacc = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"
rutass = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\"
rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutap6 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_6\"
rutaperf = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Perfiles\"
repite = 1



kwordList = "Slim NSlim MP PL PS P6 GS Perfil Tornilleria Cuñas"
tipo = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
tipo = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Slim/NSlim/MP/PL/PS/P6/GS/Perfil/Tornilleria/Cuñas]")
If tipo = "" Or tipo = "Slim" Then
    tipo = "SSlimsG"
    naranja = ""
ElseIf tipo = "NSlim" Then
    tipo = "SSlimsG"
    naranja = "N"
ElseIf tipo = "MP" Then
    tipo = "MSHOR"
ElseIf tipo = "PL" Then
    tipo = "Pshor_4L"
ElseIf tipo = "PS" Then
    tipo = "Pshor_4S"
ElseIf tipo = "P6" Then
    tipo = "Pshor_6"
ElseIf tipo = "GS" Then
    tipo = "Gshor"
ElseIf tipo = "Perfil" Then
    tipo = "Perfil"
ElseIf tipo = "Tornilleria" Then
    Call Anclajes.an
ElseIf tipo = "Cuñas" Then
    Call Cunas.cu
Else
GoTo terminar
End If

If tipo = "Perfil" Then
    If perf = "Bisagra" Then
    Else
        kwordList = "Alzado PlantaInferior PlantaSuperior"
        plalz = ""
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        plalz = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Alzado/PlantaInferior/PlantaSuperior]")
        If plalz = "Alzado" Or plalz = "" Then
            plalzp = "AL"
        ElseIf plalz = "PlantaSuperior" Then
            plalzp = "PLSUP"
        ElseIf plalz = "PlantaInferior" Then
            plalzp = "PLINF"
        End If
    End If
ElseIf tipo = "Tornilleria" Then
    GoTo terminar

Else
    kwordList = "Planta Alzado"
    plalz = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    plalz = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")
    If plalz = "" Or plalz = "Planta" Then
        plalz = "planta"
        plalzmp = "PLA"
        plalzslim = "PL"
    ElseIf plalz = "Alzado" Then
        plalz = "alzado"
        plalzmp = "ALZ"
        plalzslim = ""
    Else
        GoTo terminar
    End If
End If

Do While repite = 1

If tipo = "SSlimsG" Then
    
    kwordList = "3600 2700 1800 0900 0720 0540 0360 0360OS 0180 0090"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    slim = ThisDrawing.Utility.GetKeyword(vbLf & "Viga: [3600/2700/1800/0900/0720/0540/0360/0360OS/0180/0090]")
    If slim = "" Then slim = "3600"
End If
    
If tipo = "MSHOR" Then
    kwordList = "5400 2700 1800 900 450 270 180 90 fusible"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    mp = ThisDrawing.Utility.GetKeyword(vbLf & "Viga: [5400/2700/1800/900/450/270/180/90/fusible]")
    If mp = "" Then mp = "5400"
End If
    
If tipo = "Pshor_4L" Then
    kwordList = "6000 4500 3000 1500 750 280"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    ps = ThisDrawing.Utility.GetKeyword(vbLf & "Viga: [6000/4500/3000/1500/750/280]")
    If ps = "" Then ps = "6000"
End If

If tipo = "Pshor_4S" Then
    kwordList = "6000 4500 3000 1500 750 560"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    ps = ThisDrawing.Utility.GetKeyword(vbLf & "Viga: [6000/4500/3000/1500/750/560]")
    If ps = "" Then ps = "6000"
End If

If tipo = "Pshor_6" Then
    kwordList = "4500 3000 Cono"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    ps = ThisDrawing.Utility.GetKeyword(vbLf & "Viga: [4500/3000/Cono]")
    If ps = "" Then ps = "4500"
End If

If tipo = "Gshor" Then
    kwordList = "6000 4500 3000 1500 750 GS2-6000"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    gs = ThisDrawing.Utility.GetKeyword(vbLf & "Viga: [6000/4500/3000/1500/750/GS2-6000]")
    If gs = "" Then gs = "6000"
End If

If tipo = "Perfil" Then
    kwordList = "600 450 300 Bisagra"
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    perf = ThisDrawing.Utility.GetKeyword(vbLf & "Viga: [300/450/600/Bisagra]")
    If perf = "300" Or perf = "" Then
        kwordList = "Sí No"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        jr = ThisDrawing.Utility.GetKeyword(vbLf & "Juntas reforzadas?: [Sí/No]")
        If jr = "Sí" Or jr = "" Then
            kwordList = "3070 4570 6070"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [3070/4570/6070]")
        ElseIf jr = "No" Then
            kwordList = "900 1500 3000 4500 6000"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [900/1500/3000/4500/6000]")
        End If
    ElseIf perf = "450" Then
        kwordList = "Triple Simple"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        alma450 = ThisDrawing.Utility.GetKeyword(vbLf & "Alma?: [Triple/Simple]")
        If alma450 = "Triple" Or alma450 = "" Then
            alma = ""
        ElseIf alma450 = "Simple" Then
            alma = "SA"
        End If
        lon = ""
        kwordList = "3000 4500 6000"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [3000/4500/6000]")
        If lon = "" Then
            lon = "3000"
        End If
    ElseIf perf = "600" Then
        kwordList = "Sí No"
        ThisDrawing.Utility.InitializeUserInput 0, kwordList
        jr = ThisDrawing.Utility.GetKeyword(vbLf & "Juntas reforzadas?: [Sí/No]")
        If jr = "Sí" Or jr = "" Then
            junta = "JR"
        ElseIf jr = "No" Then
            junta = ""
        End If
    End If
End If




'Geometría:
punto1 = GcadUtil.GetPoint(, "1º Punto: ")
punto2 = GcadUtil.GetPoint(punto1, "2º Punto: ")
P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)

Set Eje1 = GcadModel.AddLine(P1, P2)
ANG = GcadUtil.AngleFromXAxis(P1, P2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
'Distancia = Val(Sqr((X ^ 2 + Y ^ 2)))



If tipo = "SSlimsG" Then
      
    If slim = "0360OS" Then
        slim = rutass & "SS" & plalzslim & "0360Offsecc.dwg"
        Set blockRef = GcadModel.InsertBlock(P1, slim, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
    Else
        slim = rutass & "SS" & plalzslim & slim & naranja & ".dwg"
        Set blockRef = GcadModel.InsertBlock(P1, slim, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
    End If
    
ElseIf tipo = "MSHOR" Then

    If mp = "fusible" Then
        mp = rutamp & "Mshor90" & plalzmp & "fusible.dwg"
        Set blockRef = GcadModel.InsertBlock(P1, mp, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
    Else
        mp = rutamp & "Mshor" & mp & plalzmp & ".dwg"
        Set blockRef = GcadModel.InsertBlock(P1, mp, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
    End If

ElseIf tipo = "Pshor_4L" Then

    ps = rutapl & "PL_" & ps & "_" & plalz & ".dwg"
    Set blockRef = GcadModel.InsertBlock(P1, ps, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4L"

ElseIf tipo = "Pshor_4S" Then

    If ps = "560" Then
    
    ps = rutaps & "PS_" & ps & ".dwg"
    Set blockRef = GcadModel.InsertBlock(P1, ps, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    
    Else
    
    ps = rutaps & "PS_" & ps & "_" & plalz & ".dwg"
    Set blockRef = GcadModel.InsertBlock(P1, ps, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    
    End If

ElseIf tipo = "Pshor_6" Then

    If ps = "Cono" Then
        ps = rutap6 & "P6_cono.dwg"
        Set blockRef = GcadModel.InsertBlock(P1, ps, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Pipeshor6"
    Else
        ps = rutap6 & "P6_" & ps & "_" & plalz & ".dwg"
        Set blockRef = GcadModel.InsertBlock(P1, ps, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Pipeshor6"
    End If

ElseIf tipo = "Gshor" Then
    
    If gs = "GS2-6000" Then
        gs = rutags & "GS_2_6000_" & plalz & ".dwg"
    Else
        gs = rutags & "GS_" & gs & "_" & plalz & ".dwg"
    End If
    
    
    Set blockRef = GcadModel.InsertBlock(P1, gs, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"
    
ElseIf tipo = "Perfil" Then
    If perf = "300" Then
        If jr = "Sí" Or jr = "" Then
            P = rutaperf & "Incye_300JR_" & lon & "_" & plalzp & ".dwg"
            Set blockRef = GcadModel.InsertBlock(P1, P, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Perfiles"
        ElseIf jr = "No" Then
            P = rutaperf & "Incye_300_" & lon & "_" & plalzp & ".dwg"
            Set blockRef = GcadModel.InsertBlock(P1, P, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Perfiles"
        End If
    ElseIf perf = "450" Then
        P = rutaperf & "Incye_450" & alma & "JR_" & lon & "_" & plalzp & ".dwg"
        Set blockRef = GcadModel.InsertBlock(P1, P, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Perfiles"
    ElseIf perf = "600" Then
        P = rutaperf & "Incye_600" & junta & "_4000_" & plalzp & ".dwg"
        Set blockRef = GcadModel.InsertBlock(P1, P, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Perfiles"
    ElseIf perf = "Bisagra" Then
            P = rutaperf & "BIS_DIN.dwg"
            Set blockRef = GcadModel.InsertBlock(P1, P, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Perfiles"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Else
            
    End If
End If


    
    
    
Eje1.Erase
Loop
terminar:
End Sub


Sub InsertDynamicBlock_2()

    ' Declara variables para el documento y el bloque
    Dim gcadApp As Object
    Dim GcadDoc As Object
    Dim gcadBlock As String
    
    ' Intenta obtener la instancia activa de gcad
    On Error Resume Next
    Set gcadApp = GetObject(, "Gcad.Application")
    On Error GoTo 0
    
    ' Si no hay instancia activa, crea una nueva
    If gcadApp Is Nothing Then
        Set gcadApp = CreateObject("Gcad.Application")
    End If
    
    ' Verifica si hay documentos abiertos
    If gcadApp.Documents.Count = 0 Then
        MsgBox "No hay documentos abiertos en gcad.", vbExclamation
        Exit Sub
    End If
    
    ' Establece el documento activo
    Set GcadDoc = gcadApp.ActiveDocument
    
    ' Especifica la ruta del archivo del bloque
    Dim blockName As String
    blockName = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Perfiles\BIS_DIN.dwg"
    
    ' Intenta abrir el bloque
    On Error Resume Next
    gcadBlock = blockName
    On Error GoTo 0
    
    ' Si no se puede abrir el bloque, muestra un mensaje y sale
    'If gcadBlock Is Nothing Then
        'MsgBox "No se pudo encontrar el bloque especificado.", vbExclamation
        'Exit Sub
    'End If
    
    ' Especifica el punto de inserción del bloque
    Dim insertPoint(0 To 2) As Double
    insertPoint(0) = 0
    insertPoint(1) = 0
    insertPoint(2) = 0
    
    ' Inserta el bloque en el modelo del espacio
    Dim newBlockRef As Object
    Set newBlockRef = GcadDoc.ModelSpace.InsertBlock(insertPoint, blockName, 1#, 1#, 1#, 0#)
    
    ' Actualiza los atributos del bloque insertado
    Dim att As Object
    For Each att In newBlockRef.GetAttributes
        ' Setea los valores de los atributos deseados aquí
        If att.TagString = "Ángulo1" Then
            att.TextString = "220"
        ElseIf att.TagString = "Ángulo2" Then
            att.TextString = "0"
        ElseIf att.TagString = "Estado de simetría1" Then
            att.TextString = "No volteado"
        End If
    Next att
    
    ' Regenera el dibujo para que se muestren los cambios
    GcadDoc.Regen acAllViewports

End Sub


Sub InsertDynamicBlock()

    ' Declara variables para el documento y el bloque
    Dim gcadApp As Object
    Dim GcadDoc As Object
    Dim gcadBlock As Object
    
    ' Intenta obtener la instancia activa de gcad
    On Error Resume Next
    Set gcadApp = GetObject(, "Gcad.Application")
    On Error GoTo 0
    
    ' Si no hay instancia activa, crea una nueva
    If gcadApp Is Nothing Then
        Set gcadApp = CreateObject("Gcad.Application")
    End If
    
    ' Verifica si hay documentos abiertos
    If gcadApp.Documents.Count = 0 Then
        MsgBox "No hay documentos abiertos en gcad.", vbExclamation
        Exit Sub
    End If
    
    ' Establece el documento activo
    Set GcadDoc = gcadApp.ActiveDocument
    
    ' Especifica la ruta del archivo del bloque
    Dim blockName As String
    blockName = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Perfiles\BIS_DIN.dwg"
    
    ' Intenta abrir el bloque
    ' Especifica el punto de inserción del bloque
    Dim insertPoint(0 To 2) As Double
    insertPoint(0) = 0
    insertPoint(1) = 0
    insertPoint(2) = 0
    
    ' Inserta el bloque en el modelo del espacio
    Dim newBlockRef As Object
    Set newBlockRef = GcadDoc.ModelSpace.InsertBlock(insertPoint, blockName, 1#, 1#, 1#, 0#)
    
    ' Actualiza los atributos del bloque insertado
    Dim att As GcadDynamicBlockReferenceProperty
    For Each att In newBlockRef.GetDynamicBlockProperties
        ' Setea los valores de los atributos deseados aquí
        If att.TagString = "Ángulo1" Then
            att.TextString = "220"
        ElseIf att.TagString = "Ángulo2" Then
            att.TextString = "0"
        ElseIf att.TagString = "Estado de simetría1" Then
            att.TextString = "No volteado"
        End If
    Next att
    
    ' Regenera el dibujo para que se muestren los cambios
    GcadDoc.Regen acAllViewports

End Sub


