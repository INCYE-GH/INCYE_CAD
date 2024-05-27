
Option Explicit

Sub ssc()
Dim rutass As String
Dim rutat As String, rutags As String
Dim GcadDoc As Object
Dim GcadUtil As Object
Dim GcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim M16x40 As String, M24x110 As String, M20x130 As String
Dim VarM16x166 As String
Dim Angulo As String, espadags As String, espadampcorta As String, espadamplarga As String, espadass As String, gato As String, tornillogs1 As String, tornillogs2 As String, husillo As String, Base_drc As String, Base_izq As String
Dim ss_90 As String
Dim ss_180 As String
Dim ss_360 As String
Dim ss_360os As String
Dim ss_540 As String
Dim ss_720 As String
Dim ss_900 As String
Dim ss_1800 As String
Dim ss_2700 As String
Dim ss_3600 As String
Dim langulo As Double
Dim nangulo As Integer
Dim l90 As Double, lespadags As Double, lespadampcorta As Double, lespadamplarga As Double, lespadass As Double, lgato As Double
Dim l180 As Double
Dim l360 As Double
Dim l540 As Double
Dim l720 As Double
Dim l900 As Double
Dim l1800 As Double
Dim l2700 As Double
Dim l3600 As Double
Dim n90 As Integer, nespadags As Integer, nespadampcorta As Integer, nespadamplarga As Integer, nespadass As Integer, n360os As Integer, ngato As Integer
Dim n180 As Integer
Dim n360 As Integer
Dim n540 As Integer
Dim n720 As Integer
Dim n900 As Integer
Dim n1800 As Integer
Dim n2700 As Integer
Dim n3600 As Integer
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
Dim ANG As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim extremo1 As String, extremo2 As String
Dim vigaM As String
Dim vigamenor As String, naranja As String
Dim offsecc As String
Dim lpuntal As Double, laux As Double
Dim plalz As String
Dim kwordList As String
Dim i As Integer
Dim lfija As Double
Dim Ncapaslim As String, Ncapa As String
Dim capaslim As Object, Gcapa As Object

Dim longitud As String, orientacion As String, Pinicio0 As String, Pinicio1 As String, Pinicio2 As String, PreMon As String

Set GcadDoc = GetObject(, "Gcad.Application").ActiveDocument
Set GcadModel = GcadDoc.ModelSpace
Set GcadUtil = GcadDoc.Utility

On Error GoTo terminar
repite = 1

Ncapa = "NoContable"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 40
Ncapaslim = "Slims"
Set capaslim = GcadDoc.Layers.Add(Ncapaslim)
capaslim.color = 30
Ncapa = "Nonplot"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 50

rutass = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\SSlimsG\"
rutat = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"

VarM16x166 = rutat & "2VarM16X166.dwg"
M16x40 = rutat & "4M16X40.dwg"
M24x110 = rutat & "1M24X110.dwg"
M20x130 = rutat & "1M20X130.dwg"

'Valores fijos
PI = 4 * Atn(1)
langulo = 119.5
l90 = 90
l180 = 180
l360 = 360
l540 = 540
l720 = 720
l900 = 900
l1800 = 1800
l2700 = 2700
l3600 = 3600
lespadags = 150
lespadass = 358
lespadamplarga = 358
lespadampcorta = 158
lgato = 420




kwordList = "A-Libre B-EspadaSlimAlzado C-EspadaSlimPlanta D-EspadaMpLargaAlzado E-EspadaMpLargaPlanta F-EspadaMpCortaAlzado G-EspadaMpCortaPlanta H-EspadaGranshorAlzado I-EspadaGranshorPlanta J-AnguloEsquina K-SS360OS L-GatoSlim"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
extremo1 = ThisDrawing.Utility.GetKeyword(vbLf & "Extremo 1: [A-Libre/B-EspadaSlimAlzado/C-EspadaSlimPlanta/D-EspadaMpLargaAlzado/E-EspadaMpLargaPlanta/F-EspadaMpCortaAlzado/G-EspadaMpCortaPlanta/H-EspadaGranshorAlzado/I-EspadaGranshorPlanta/J-AnguloEsquina/K-SS360OS/L-GatoSlim]")

If extremo1 = "" Or extremo1 = "A-Libre" Then
extremo1 = "A-Libre"
nespadags = 0
nespadampcorta = 0
nespadamplarga = 0
nespadass = 0
nangulo = 0
n360os = 0
ngato = 0
ElseIf extremo1 = "B-EspadaSlimAlzado" Or extremo1 = "C-EspadaSlimPlanta" Then
nespadags = 0
nespadampcorta = 0
nespadamplarga = 0
nespadass = 1
nangulo = 0
n360os = 0
ngato = 1
ElseIf extremo1 = "D-EspadaMpLargaAlzado" Or extremo1 = "E-EspadaMpLargaPlanta" Then
nespadags = 0
nespadampcorta = 0
nespadamplarga = 1
nespadass = 0
nangulo = 0
n360os = 0
ngato = 1
ElseIf extremo1 = "F-EspadaMpCortaAlzado" Or extremo1 = "G-EspadaMpCortaPlanta" Then
nespadags = 0
nespadampcorta = 1
nespadamplarga = 0
nespadass = 0
nangulo = 0
n360os = 0
ngato = 1
ElseIf extremo1 = "H-EspadaGranshorAlzado" Or extremo1 = "I-EspadaGranshorPlanta" Then
nespadags = 1
nespadampcorta = 0
nespadamplarga = 0
nespadass = 0
nangulo = 0
n360os = 0
ngato = 1
kwordList = "AM20X40 BM20X60"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
tornillogs1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tornillo en la espada Gshor del extremo 1: [AM20X40/BM20X60]")
    If tornillogs1 = "AM20X40" Or tornillogs1 = "" Then
    tornillogs1 = rutat & "1M20X40.dwg"
    ElseIf tornillogs1 = "BM20X60" Then
    tornillogs1 = rutat & "1M20X60.dwg"
    End If
ElseIf extremo1 = "J-AnguloEsquina" Then
nespadags = 0
nespadampcorta = 0
nespadamplarga = 0
nespadass = 0
nangulo = 1
n360os = 0
ngato = 0
ElseIf extremo1 = "K-SS360OS" Then
nespadags = 0
nespadampcorta = 0
nespadamplarga = 0
nespadass = 0
nangulo = 0
n360os = 1
ngato = 0
ElseIf extremo1 = "L-GatoSlim" Then
nespadags = 0
nespadampcorta = 0
nespadamplarga = 0
nespadass = 0
nangulo = 0
n360os = 0
ngato = 1
Else
End If

kwordList = "A-Libre B-EspadaSlimAlzado C-EspadaSlimPlanta D-EspadaMpLargaAlzado E-EspadaMpLargaPlanta F-EspadaMpCortaAlzado G-EspadaMpCortaPlanta H-EspadaGranshorAlzado I-EspadaGranshorPlanta J-AnguloEsquina K-SS360OS L-GatoSlim"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
extremo2 = ThisDrawing.Utility.GetKeyword(vbLf & "Extremo 2: [A-Libre/B-EspadaSlimAlzado/C-EspadaSlimPlanta/D-EspadaMpLargaAlzado/E-EspadaMpLargaPlanta/F-EspadaMpCortaAlzado/G-EspadaMpCortaPlanta/H-EspadaGranshorAlzado/I-EspadaGranshorPlanta/J-AnguloEsquina/K-SS360OS/L-GatoSlim]")

If extremo2 = "" Or extremo2 = "A-Libre" Then
extremo2 = "A-Libre"
nespadags = nespadags + 0
nespadampcorta = nespadampcorta + 0
nespadamplarga = nespadamplarga + 0
nespadass = nespadass + 0
nangulo = nangulo + 0
n360os = n360os + 0
ngato = ngato + 0
ElseIf extremo2 = "B-EspadaSlimAlzado" Or extremo2 = "C-EspadaSlimPlanta" Then
nespadags = nespadags + 0
nespadampcorta = nespadampcorta + 0
nespadamplarga = nespadamplarga + 0
nespadass = nespadass + 1
nangulo = nangulo + 0
n360os = n360os + 0
ngato = ngato + 1
ElseIf extremo2 = "D-EspadaMpLargaAlzado" Or extremo2 = "E-EspadaMpLargaPlanta" Then
nespadags = nespadags + 0
nespadampcorta = nespadampcorta + 0
nespadamplarga = nespadamplarga + 1
nespadass = nespadass + 0
nangulo = nangulo + 0
n360os = n360os + 0
ngato = ngato + 1
ElseIf extremo2 = "F-EspadaMpCortaAlzado" Or extremo2 = "G-EspadaMpCortaPlanta" Then
nespadags = nespadags + 0
nespadampcorta = nespadampcorta + 1
nespadamplarga = nespadamplarga + 0
nespadass = nespadass + 0
nangulo = nangulo + 0
n360os = n360os + 0
ngato = ngato + 1
ElseIf extremo2 = "H-EspadaGranshorAlzado" Or extremo2 = "I-EspadaGranshorPlanta" Then
nespadags = nespadags + 1
nespadampcorta = nespadampcorta + 0
nespadamplarga = nespadamplarga + 0
nespadass = nespadass + 0
nangulo = nangulo + 0
n360os = n360os + 0
ngato = ngato + 1
kwordList = "AM20X40 BM20X60"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
tornillogs2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tornillo en la espada Gshor del extremo 2: [AM20X40/BM20X60]")
    If tornillogs2 = "AM20X40" Or tornillogs2 = "" Then
    tornillogs2 = rutat & "1M20X40.dwg"
    ElseIf tornillogs2 = "BM20X60" Then
    tornillogs2 = rutat & "1M20X60.dwg"
    End If
ElseIf extremo2 = "J-AnguloEsquina" Then
nespadags = nespadags + 0
nespadampcorta = nespadampcorta + 0
nespadamplarga = nespadamplarga + 0
nespadass = nespadass + 0
nangulo = nangulo + 1
n360os = n360os + 0
ngato = ngato + 0
ElseIf extremo2 = "K-SS360OS" Then
nespadags = nespadags + 0
nespadampcorta = nespadampcorta + 0
nespadamplarga = nespadamplarga + 0
nespadass = nespadass + 0
nangulo = nangulo + 0
n360os = n360os + 1
ngato = ngato + 0
ElseIf extremo2 = "L-GatoSlim" Then
nespadags = nespadags + 0
nespadampcorta = nespadampcorta + 0
nespadamplarga = nespadamplarga + 0
nespadass = nespadass + 0
nangulo = nangulo + 0
n360os = n360os + 0
ngato = ngato + 1
Else
End If

kwordList = "Planta Alzado"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
plalz = ThisDrawing.Utility.GetKeyword(vbLf & "Vista de las Slim en: [Planta/Alzado]")

If plalz = "" Or plalz = "Planta" Then
plalz = "PL"
ElseIf plalz = "Alzado" Then
plalz = ""
Else
GoTo terminar
End If

kwordList = "3600 2700 1800"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vigaM = ThisDrawing.Utility.GetKeyword(vbLf & "Viga de mayor tamaño: [3600/2700/1800]")

kwordList = "90 180 360"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vigamenor = ThisDrawing.Utility.GetKeyword(vbLf & "Viga de menor tamaño: [90/180/360]")

kwordList = "Galvanizada Pintada"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
naranja = ThisDrawing.Utility.GetKeyword(vbLf & "Vigas Slim: [Galvanizada/Pintada]")

If naranja = "" Or naranja = "Galvanizada" Then
naranja = ""
ElseIf naranja = "Pintada" Then
naranja = "N"
Else
GoTo terminar
End If

lfija = nangulo * langulo + n360os * l360 + nespadags * lespadags + nespadampcorta * lespadampcorta + nespadamplarga * lespadamplarga + nespadass * lespadass + ngato * lgato
 


Do While repite = 1



'Geometría:
punto1 = GcadUtil.GetPoint(, "1º Punto: ")
punto2 = GcadUtil.GetPoint(punto1, "2º Punto: ")
P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)

Pinicio0 = CStr(P1(0))
Pinicio1 = CStr(P1(1))
Pinicio2 = CStr(P1(2))

Dim k As String, b As Object, entity As Object
nop:
k = InputBox("Ingrese nombre: ")

If k = "" Then
    k = GenerarNombreAleatorio(30)
End If
    
If BloqueExiste(k) Then
    MsgBox "Ya existe un bloque con el mismo nombre!"
    Dim Respuesta As String
    kwordList = "Sobreescribir Renombrar"
    Respuesta = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    Respuesta = ThisDrawing.Utility.GetKeyword(vbLf & "¿Qué desea hacer?: [Sobreescribir/Renombrar]")
    
    If Respuesta = "Sobreescribir" Or Respuesta = "" Then
    
        For Each entity In ThisDrawing.ModelSpace
            If TypeOf entity Is GcadBlockReference Then
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

Dim check As Double
check = Len(k)

Set b = ThisDrawing.Blocks.Add(punto1, k)

Set Eje1 = ThisDrawing.Blocks.Item(k).AddLine(P1, P2)
Eje1.Layer = "Nonplot"
ANG = GcadUtil.AngleFromXAxis(P1, P2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

Longitud = CStr(Distancia)
orientacion = CStr(ANG)

If Distancia < lfija Then
        MsgBox "Medida " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If

lpuntal = Distancia - lfija
If vigaM = "" Then vigaM = "3600"
If vigaM = 3600 Or vigaM = "" Then
n3600 = Fix(lpuntal / l3600)
lpuntal = lpuntal - n3600 * l3600
n2700 = Fix(lpuntal / l2700)
lpuntal = lpuntal - n2700 * l2700
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
ElseIf vigaM = 2700 Then
n3600 = 0
n2700 = Fix(lpuntal / l2700)
lpuntal = lpuntal - n2700 * l2700
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
ElseIf vigaM = 1800 Then
n3600 = 0
n2700 = 0
n1800 = Fix(lpuntal / l1800)
lpuntal = lpuntal - n1800 * l1800
Else: GoTo terminar
End If

n900 = Fix(lpuntal / l900)
lpuntal = lpuntal - n900 * l900
n720 = Fix(lpuntal / l720)
lpuntal = lpuntal - n720 * l720
n540 = Fix(lpuntal / l540)
lpuntal = lpuntal - n540 * l540

If vigamenor = "" Then vigamenor = "90"

If vigamenor = 360 Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = 0
n90 = 0
ElseIf vigamenor = 180 Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = Fix(lpuntal / l180)
lpuntal = lpuntal - n180 * l180
n90 = 0
ElseIf vigamenor = 90 Or vigamenor = "" Then
n360 = Fix(lpuntal / l360)
lpuntal = lpuntal - n360 * l360
n180 = Fix(lpuntal / l180)
lpuntal = lpuntal - n180 * l180
n90 = Fix(lpuntal / l90)
lpuntal = lpuntal - n90 * l90
Else: GoTo terminar
End If

If ngato = 0 Then
    laux = 0
ElseIf ngato = 1 Then
    If lpuntal > 200 Then
        MsgBox "La abertura del gato " & laux & "mm, es mayor que la máxima admisible de 620mm, el puntal dibujado no llega al segundo extremo solicitado (se dibuja abertura de 620mm)"
        laux = 620
    Else
        laux = lpuntal + lgato
    End If
ElseIf ngato = 2 Then
    laux = lpuntal / 2 + lgato
End If

Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)

If extremo1 = "" Or extremo1 = "A-Libre" Then
'No hacer nada
ElseIf extremo1 = "B-EspadaSlimAlzado" Then
    espadass = rutass & "SSEspada.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, espadass, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
    'Set BlockRef = GcadModel.InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
    'BlockRef.Layer = "Nonplot"
ElseIf extremo1 = "C-EspadaSlimPlanta" Then
    espadass = rutass & "SSPLEspada.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, espadass, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
ElseIf extremo1 = "D-EspadaMpLargaAlzado" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    espadamplarga = rutass & "SSEspadaLarga.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, espadamplarga, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadamplarga * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadamplarga * Sin(ANG): Punto_final(2) = Punto_inial(2)
ElseIf extremo1 = "E-EspadaMpLargaPlanta" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    espadamplarga = rutass & "SSPLEspadaLarga.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, espadamplarga, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadamplarga * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadamplarga * Sin(ANG): Punto_final(2) = Punto_inial(2)
ElseIf extremo1 = "F-EspadaMpCortaAlzado" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    espadampcorta = rutass & "SSEspadaCorta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, espadampcorta, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadampcorta * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadampcorta * Sin(ANG): Punto_final(2) = Punto_inial(2)
ElseIf extremo1 = "G-EspadaMpCortaPlanta" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    espadampcorta = rutass & "SSPLEspadaCorta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, espadampcorta, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadampcorta * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadampcorta * Sin(ANG): Punto_final(2) = Punto_inial(2)
ElseIf extremo1 = "H-EspadaGranshorAlzado" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, tornillogs1, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    espadags = rutags & "GS_espada.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, espadags, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadags * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadags * Sin(ANG): Punto_final(2) = Punto_inial(2)
ElseIf extremo1 = "I-EspadaGranshorPlanta" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, tornillogs1, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    espadags = rutags & "GS_PLespada.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, espadags, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + lespadags * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadags * Sin(ANG): Punto_final(2) = Punto_inial(2)
ElseIf extremo1 = "J-AnguloEsquina" And plalz = "" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, VarM16x166, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Angulo = rutass & "SSAnguloEsquina.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, Angulo, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + langulo * Cos(ANG): Punto_final(1) = Punto_inial(1) + langulo * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo1 = "J-AnguloEsquina" And plalz = "PL" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, VarM16x166, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Angulo = rutass & "SSPLAnguloEsquina.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, Angulo, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + langulo * Cos(ANG): Punto_final(1) = Punto_inial(1) + langulo * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo1 = "K-SS360OS" And plalz = "" Then
    Punto_final(0) = Punto_inial(0) + l360 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l360 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    ss_360os = rutass & "SS0360Offsecc.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, ss_360os, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo1 = "K-SS360OS" And plalz = "PL" Then
    Punto_final(0) = Punto_inial(0) + l360 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l360 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    ss_360os = rutass & "SSPL0360Offsecc.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, ss_360os, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo1 = "L-GatoSlim" And plalz = "" Then
    husillo = rutass & "zSSHusillo.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, husillo, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Base_izq = rutass & "SSGatorefizq.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Base_izq, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo1 = "L-GatoSlim" And plalz = "PL" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, M24x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    husillo = rutass & "zSSHusillo.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, husillo, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Base_izq = rutass & "SSPLGatorefizq.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Base_izq, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
Else
GoTo terminar
End If

If plalz = "" Then
    If extremo1 = "B-EspadaSlimAlzado" Or extremo1 = "C-EspadaSlimPlanta" Or extremo1 = "D-EspadaMpLargaAlzado" Or extremo1 = "E-EspadaMpLargaPlanta" Or extremo1 = "F-EspadaMpCortaAlzado" Or extremo1 = "G-EspadaMpCortaPlanta" Or extremo1 = "H-EspadaGranshorAlzado" Or extremo1 = "I-EspadaGranshorPlanta" Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_final(0) + laux * Cos(ANG): Punto_final(1) = Punto_final(1) + laux * Sin(ANG): Punto_final(2) = Punto_final(2)
        Base_izq = rutass & "SSGatorefizq.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Base_izq, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"

    End If
End If

If plalz = "PL" Then
    If extremo1 = "B-EspadaSlimAlzado" Or extremo1 = "C-EspadaSlimPlanta" Or extremo1 = "D-EspadaMpLargaAlzado" Or extremo1 = "E-EspadaMpLargaPlanta" Or extremo1 = "F-EspadaMpCortaAlzado" Or extremo1 = "G-EspadaMpCortaPlanta" Or extremo1 = "H-EspadaGranshorAlzado" Or extremo1 = "I-EspadaGranshorPlanta" Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_final(0) + laux * Cos(ANG): Punto_final(1) = Punto_final(1) + laux * Sin(ANG): Punto_final(2) = Punto_final(2)
        Base_izq = rutass & "SSPLGatorefizq.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Base_izq, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"

    End If
End If

If n3600 > 0 Then
    i = 0
    Do While i < n3600
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_3600 = rutass & "SS" & plalz & "3600" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_3600, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l3600 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3600 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n2700 > 0 Then
    i = 0
    Do While i < n2700
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_2700 = rutass & "SS" & plalz & "2700" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_2700, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l2700 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l2700 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n1800 > 0 Then
    i = 0
    Do While i < n1800
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_1800 = rutass & "SS" & plalz & "1800" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_1800, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l1800 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1800 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n900 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_900 = rutass & "SS" & plalz & "0900" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_900, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l900 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l900 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n720 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_720 = rutass & "SS" & plalz & "0720" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_720, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l720 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l720 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n540 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_540 = rutass & "SS" & plalz & "0540" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_540, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l540 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l540 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n360 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_360 = rutass & "SS" & plalz & "0360" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_360, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l360 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l360 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n180 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_180 = rutass & "SS" & plalz & "0180" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_180, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n90 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        ss_90 = rutass & "SS" & plalz & "0090" & naranja & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_90, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M16x40, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If plalz = "" Then
    If extremo2 = "B-EspadaSlimAlzado" Or extremo2 = "C-EspadaSlimPlanta" Or extremo2 = "D-EspadaMpLargaAlzado" Or extremo2 = "E-EspadaMpLargaPlanta" Or extremo2 = "F-EspadaMpCortaAlzado" Or extremo2 = "G-EspadaMpCortaPlanta" Or extremo2 = "H-EspadaGranshorAlzado" Or extremo2 = "I-EspadaGranshorPlanta" Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Base_drc = rutass & "SSGatorefdrc.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, Base_drc, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Slims"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"

    End If
End If

If plalz = "PL" Then
    If extremo2 = "B-EspadaSlimAlzado" Or extremo2 = "C-EspadaSlimPlanta" Or extremo2 = "D-EspadaMpLargaAlzado" Or extremo2 = "E-EspadaMpLargaPlanta" Or extremo2 = "F-EspadaMpCortaAlzado" Or extremo2 = "G-EspadaMpCortaPlanta" Or extremo2 = "H-EspadaGranshorAlzado" Or extremo2 = "I-EspadaGranshorPlanta" Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Base_drc = rutass & "SSPLGatorefdrc.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, Base_drc, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Slims"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"

    End If
End If

If extremo2 = "" Or extremo2 = "A-Libre" Then
'No hacer nada
ElseIf extremo2 = "B-EspadaSlimAlzado" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadass = rutass & "SSEspada.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, espadass, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
ElseIf extremo2 = "C-EspadaSlimPlanta" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadass * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadass * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadass = rutass & "SSPLEspada.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, espadass, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
ElseIf extremo2 = "D-EspadaMpLargaAlzado" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadamplarga * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadamplarga * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadamplarga = rutass & "SSEspadaLarga.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, espadamplarga, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo2 = "E-EspadaMpLargaPlanta" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadamplarga * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadamplarga * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadamplarga = rutass & "SSPLEspadaLarga.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, espadamplarga, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo2 = "F-EspadaMpCortaAlzado" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadampcorta * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadampcorta * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadampcorta = rutass & "SSEspadaCorta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, espadampcorta, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo2 = "G-EspadaMpCortaPlanta" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadampcorta * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadampcorta * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadampcorta = rutass & "SSPLEspadaCorta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, espadampcorta, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x130, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo2 = "H-EspadaGranshorAlzado" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadags * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadags * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadags = rutags & "GS_espada.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, espadags, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, tornillogs2, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo2 = "I-EspadaGranshorPlanta" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + lespadags * Cos(ANG): Punto_final(1) = Punto_inial(1) + lespadags * Sin(ANG): Punto_final(2) = Punto_inial(2)
    espadags = rutags & "GS_PLespada.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, espadags, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, tornillogs2, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo2 = "J-AnguloEsquina" And plalz = "" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + langulo * Cos(ANG): Punto_final(1) = Punto_inial(1) + langulo * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Angulo = rutass & "SSAnguloEsquina.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Angulo, Xs, Ys, Zs, ANG - (PI / 2))
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, VarM16x166, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo2 = "J-AnguloEsquina" And plalz = "PL" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    Punto_final(0) = Punto_inial(0) + langulo * Cos(ANG): Punto_final(1) = Punto_inial(1) + langulo * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Angulo = rutass & "SSPLAnguloEsquina.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Angulo, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Slims"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, VarM16x166, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
ElseIf extremo2 = "K-SS360OS" And plalz = "" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    ss_360os = rutass & "SS0360Offsecc.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_360os, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
ElseIf extremo2 = "K-SS360OS" And plalz = "PL" Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    ss_360os = rutass & "SSPL0360Offsecc.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, ss_360os, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Slims"
ElseIf extremo2 = "L-GatoSlim" And plalz = "" Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Base_drc = rutass & "SSGatorefdrc.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, Base_drc, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Slims"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
ElseIf extremo2 = "L-GatoSlim" And plalz = "PL" Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        Base_drc = rutass & "SSPLGatorefdrc.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, Base_drc, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Slims"
        Punto_final(0) = Punto_inial(0) + laux * Cos(ANG): Punto_final(1) = Punto_inial(1) + laux * Sin(ANG): Punto_final(2) = Punto_inial(2)
        husillo = rutass & "zSSHusillo.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, husillo, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Slims"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M24x110, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
Else
GoTo terminar
End If
        
PreMon = ""
Dim NamePre As GcadAttribute
Set NamePre = b.AddAttribute(1, acAttributeModeInvisible, "ey", punto1, "NombrePremontaje", PreMon)

Dim longitudatt As GcadAttribute
Set longitudatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", punto1, "Longitud", longitud)
        
Dim orientacionatt As GcadAttribute
Set orientacionatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", punto1, "Orientacion", orientacion)
           
Dim cooordenadainicio0 As GcadAttribute
Set cooordenadainicio0 = b.AddAttribute(1, acAttributeModeInvisible, "ey", punto1, "Coordenada0", Pinicio0)
        
Dim cooordenadainicio1 As GcadAttribute
Set cooordenadainicio1 = b.AddAttribute(1, acAttributeModeInvisible, "ey", punto1, "Coordenada1", Pinicio1)
        
Dim cooordenadainicio2 As GcadAttribute
Set cooordenadainicio2 = b.AddAttribute(1, acAttributeModeInvisible, "ey", punto1, "Coordenada2", Pinicio2)           
 
Set blockRef = ThisDrawing.ModelSpace.InsertBlock(punto1, k, Xs, Ys, Zs, 0)
If check = 30 Then
    blockRef.Explode
    blockRef.Delete
Else
    blockRef.Layer = "NoContable"
End If

Eje1.Layer = "Nonplot"
Loop
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


Function GenerarNombreAleatorio(Longitud As Integer) As String
    Dim i As Integer
    Dim Nombre As String
    Dim Caracter As String
    Dim Rango As String

    Rango = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    
    Nombre = ""
    For i = 1 To Longitud
        Caracter = Mid(Rango, Int((Len(Rango) * Rnd) + 1), 1)
        Nombre = Nombre & Caracter
    Next i
    
    GenerarNombreAleatorio = Nombre
End Function

