Option Explicit

Sub ps()
Dim ruta As String, rutaps As String, rutapl As String, rutags As String
Dim ruta2 As String
Dim AcadDoc As Object
Dim AcadUtil As Object
Dim AcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim M20x90 As String
Dim M20x150 As String
Dim M20x160 As String
Dim M20x90_16 As String
Dim GS_Bulon120mm As String
Dim GS_Giro As String
Dim GS_Fusible As String
Dim PS_280 As String
Dim PS_750 As String, PS_560 As String
Dim PS_1500 As String
Dim PS_3000 As String
Dim PS_4500 As String
Dim PS_6000 As String
Dim PS_Husillo As String
Dim PS_Placa50mm As String
Dim zPS_Gato_Cono As String
Dim zPS_Gato_Tope As String
Dim PS_Gato As String
Dim lgiro As Double
Dim lfusible As Double
Dim l280 As Double
Dim l750 As Double, l560 As Double
Dim l1500 As Double
Dim l3000 As Double
Dim l4500 As Double
Dim l6000 As Double
Dim l50 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lalt1 As Double
Dim lalt2 As Double
Dim lgatomin As Double
Dim n6000 As Integer
Dim n4500 As Integer
Dim n3000 As Integer
Dim n1500 As Integer
Dim n750 As Integer, n560 As Integer
Dim n280 As Integer
Dim nfusible As Integer
Dim blockRef As Object
Dim repite As Double
Dim Punto_inial(0 To 2) As Double
Dim Punto_final(0 To 2) As Double
Dim Punto_inial2(0 To 2) As Double
Dim Punto_final2(0 To 2) As Double
Dim PI As Variant
Dim Eje1 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim ANG As Double, Direcc As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim dato1 As String
Dim dato2 As String
Dim dato3 As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object
Dim k As String, b As acadBlock, entity As Object

Dim longitud As String, orientacion As String


Set AcadDoc = GetObject(, "Autocad.Application").ActiveDocument
Set AcadModel = ThisDrawing.ModelSpace
Set AcadUtil = AcadDoc.Utility

Ncapa = "Pipeshor4S"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5
Ncapa = "Granshor"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150

'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgiro = 205
lfusible = 187.5
l280 = 280
l560 = 560
l750 = 750
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
l50 = 50
l_tope = 325
l_conogato = 170
lgatomin = 620
lfija = (2 * lgiro) + lfusible + (2 * l50) + lgatomin

On Error GoTo terminar

kwordList = "S L"
dato2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato2 = "Pshor_4" & ThisDrawing.Utility.GetKeyword(vbLf & "Introduce PS4S o PS4L: [S/L]")

If dato2 = "Pshor_4L" Then
capa = "Pipeshor4L"
dato3 = "PL"
ElseIf dato2 = "Pshor_4S" Or dato2 = "Pshor_4" Then
dato2 = "Pshor_4S"
capa = "Pipeshor4S"
dato3 = "PS"
Else
GoTo terminar
End If

ruta = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\" & dato2 & "\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"

kwordList = "Planta Alzado"
dato1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")
If dato1 = "" Or dato1 = "Planta" Then
dato1 = "Planta"
ElseIf dato1 = "Alzado" Then
Else
GoTo terminar
End If

Do While repite = 1
'Geometría:
punto1 = AcadUtil.GetPoint(, "1º Punto: ")
punto2 = AcadUtil.GetPoint(punto1, "2º Punto: ")
P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)


k = InputBox("Ingrese nombre: ")

If k = "" Then
nop:
    MsgBox "Introduzca un nombre, por favor"
    k = InputBox("Ingrese nombre: ")
    If k = "" Then
        GoTo nop
    End If
End If
    
If BloqueExiste(k) Then
    Dim Respuesta As String
    kwordList = "Sobreescribir Renombrar"
    Respuesta = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    Respuesta = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Sobreescribir/Renombrar]")
    
    If Respuesta = "Sobreescribir" Or Respuesta = "" Then
    
        For Each entity In ThisDrawing.ModelSpace
            If TypeOf entity Is AcadBlockReference Then
                If entity.effectiveName = k Then
                    entity.Delete
                End If
            End If
        Next entity
                
            
        
        ThisDrawing.Blocks.Item(k).Delete
    ElseIf Respuesta = "Renombrar" Then
        GoTo nop
    End If
End If

Set b = ThisDrawing.Blocks.Add(punto1, k)

Set Eje1 = ThisDrawing.Blocks.Item(k).AddLine(P1, P2)
Eje1.Layer = "Nonplot"
ANG = AcadUtil.AngleFromXAxis(P1, P2)
Direcc = AcadUtil.AngleFromXAxis(P1, P2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

longitud = CStr(Distancia)
orientacion = CStr(ANG)

If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If


'Introducir el bulón de 120 mm en los extremos siempre, ángulo de giro, fusible fijo y chapas de 50mm:
GS_Bulon120mm = rutags & "GS_Bulon120mm_" & dato1 & ".dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, GS_Bulon120mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P2, GS_Bulon120mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
GS_Giro = rutags & "GS_Giro_" & dato1 & ".dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, GS_Giro, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P2, GS_Giro, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Granshor"


Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
GS_Fusible = rutags & "GS_Fusible_" & dato1 & ".dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
M20x90 = ruta2 & "4M20X90.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Pipeshor4S"
M20x150 = ruta2 & "4M20X150.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x150, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Pipeshor4S"
M20x160 = ruta2 & "4M20X160.dwg"
Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x160, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"


lpuntal = Distancia - lfija
n6000 = Fix(lpuntal / l6000)
lpuntal = lpuntal - n6000 * l6000
n4500 = Fix(lpuntal / l4500)
lpuntal = lpuntal - n4500 * l4500
n3000 = Fix(lpuntal / l3000)
lpuntal = lpuntal - n3000 * l3000
n1500 = Fix(lpuntal / l1500)
lpuntal = lpuntal - n1500 * l1500
n750 = Fix(lpuntal / l750)
lpuntal = lpuntal - n750 * l750

Select Case lpuntal

    Case 0 To 230
    nfusible = 1
    n280 = 0
    n560 = 0
    Case 230 To 280
    nfusible = 2
    n280 = 0
    n560 = 0
    Case 280 To 510
    nfusible = 1
    n280 = 1
    n560 = 0
    Case 510 To 560
    nfusible = 2
    n280 = 1
    n560 = 0
    Case 560 To 750
    nfusible = 1
        If dato2 = "Pshor_4L" Then
        n280 = 2
        n560 = 0
        ElseIf dato2 = "Pshor_4S" Then
        n280 = 0
        n560 = 1
        End If
    Case Else
    MsgBox "Longitud no controlada " & lpuntal & "mm, fuera de rango, revisar código"
    GoTo terminar
        
End Select



M20x90_16 = ruta2 & "16M20X90.dwg"

If n280 > 0 Then
    i = 0
    Do While i < n280
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_280 = rutapl & "PL_280_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_280, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Pipeshor4L"
        Punto_final(0) = Punto_inial(0) + l280 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l280 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n560 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_560 = ruta & "PS_560.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n1500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_1500 = ruta & dato3 & "_1500_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_1500, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n3000 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_3000 = ruta & dato3 & "_3000_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_3000, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n4500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_4500 = ruta & dato3 & "_4500_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_4500, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n6000 > 0 Then
    i = 0
    Do While i < n6000
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_6000 = ruta & dato3 & "_6000_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_6000, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If


If n750 > 0 Then
    i = 0
    Do While i < n750
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        PS_750 = ruta & dato3 & "_750_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_750, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
zPS_Gato_Cono = rutaps & "zPS_Gato_Cono_" & dato1 & ".dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, zPS_Gato_Cono, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Pipeshor4S"
'M20x90_16 = ruta & "16M20x90.dwg"
'Set BlockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
'BlockRef.Layer = "TORNILLERIA"
Punto_final(0) = Punto_inial(0) + l_conogato * Cos(ANG): Punto_final(1) = Punto_inial(1) + l_conogato * Sin(ANG): Punto_final(2) = Punto_inial(2)

Punto_inial2(0) = P2(0) - lgiro * Cos(ANG): Punto_inial2(1) = P2(1) - lgiro * Sin(ANG): Punto_inial2(2) = P2(2)
Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
    If nfusible = 2 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Granshor"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Nonplot"
        Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
    End If
Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
zPS_Gato_Tope = rutaps & "zPS_Gato_Tope_" & dato1 & ".dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, zPS_Gato_Tope, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Pipeshor4S"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, M20x90, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Nonplot"
Punto_final2(0) = Punto_inial2(0) - l_tope * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - l_tope * Sin(ANG): Punto_final(2) = Punto_inial2(2)


Punto_inial(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_inial(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_inial(2) = (Punto_final(2) + Punto_final2(2)) / 2

PS_Gato = rutaps & "PS_Gato_" & dato1 & ".dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Gato, Xs, Ys, Zs, ANG)
blockRef.Layer = "Pipeshor4S"
        
If dato3 = "PL" Then
    If nfusible = 1 Then
        If n280 = 1 Then
            lalt1 = Distancia - l280
            lalt2 = Distancia + 470
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        ElseIf n280 = 2 Then
            lalt1 = Distancia - l280
            lalt2 = Distancia + 190
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt2 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt1 & "."
        End If
    ElseIf nfusible = 2 Then
        If n280 = 1 Then
            lalt1 = Distancia - lfusible - l280 + 150
            lalt2 = Distancia - l280 - lfusible + l750
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt2 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt1 & "."
        ElseIf n280 = 2 Then
            lalt1 = Distancia - lfusible - l280 + 150
            lalt2 = Distancia - 560 - lfusible + l750
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt2 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt1 & "."
        End If
    End If
ElseIf dato3 = "PS" Then
    If n280 = 1 Then
        If n560 = 1 And n750 = 1 Then
            lalt1 = Distancia - l280 + 190
            lalt2 = Distancia + l280
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        ElseIf n560 = 0 And n750 = 1 Then
            lalt1 = Distancia - l280
            lalt2 = Distancia + l280
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        ElseIf n560 = 1 And n750 = 0 Then
            lalt1 = Distancia - 90
            lalt2 = Distancia + 280
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        Else
            lalt1 = Distancia - 280
            lalt2 = Distancia + 280
            MsgBox "Es posible reducir el uso de tubos pequeños modificando la longitud del puntal." & vbCrLf & vbCrLf & "Longitud actual: " & Distancia & "" & vbCrLf & vbCrLf & "Opción más óptima: " & lalt1 & "." & vbCrLf & vbCrLf & "Opción alternativa: " & lalt2 & "."
        End If
    End If
End If
        
Dim orientacionatt As AcadAttribute
Set orientacionatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", punto1, "Orientacion", orientacion)
        
Dim longitudatt As AcadAttribute
Set longitudatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", punto1, "Longitud", longitud)
           
        
Set blockRef = ThisDrawing.ModelSpace.InsertBlock(punto1, k, Xs, Ys, Zs, 0)
blockRef.Layer = "NoContable"
        
Eje1.Layer = "Nonplot"
Loop

'Set blockRef = ThisDrawing.ModelSpace.InsertBlock(punto1, b, Xs, Ys, Zs, ANG)
'blockRef.Layer = "Pipeshor4L"



ThisDrawing.Regen acAllViewports

terminar:
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



