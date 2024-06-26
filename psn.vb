Option Explicit
''''''' Geometría para el cálculo del ajuste de los puntales pipeshor

' falta cambiar el cajón de orientación y añadir la posibilidad de la proyección


Sub mpn(carga As Variant)
Dim ruta As String, rutaps As String, rutapl As String, rutags As String
Dim ruta2 As String
Dim GcadDoc As Object
Dim M20x60_4 As String
Dim M20x50_4 As String
Dim GcadUtil As Object
Dim GcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim line2 As GcadLine
Dim line1 As GcadLine
Dim M20x90 As String
Dim M20x150 As String, M20x110 As String, Var20x250 As String
Dim M20x160 As String
Dim M20x90_16 As String
Dim GS_Bulon120mm As String
Dim GS_Giro As String
Dim GS_Fusible As String
Dim PS_280 As String
Dim PBloque(0 To 2) As Double
Dim PS_750 As String, PS_560 As String
Dim PS_1500 As String
Dim PS_3000 As String
Dim PS_4500 As String
Dim PS_6000 As String
Dim placaanc1 As String
Dim rutaplaca1 As String, rutacajon As String
Dim placaanc2 As String
Dim rutaplaca2 As String
Dim basecajon As String, brazocajon As String
Dim PS_Husillo As String
Dim PS_Placa50mm As String
Dim PS_Placa35mm As String
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
Dim l50 As Double, l35 As Double
Dim l145 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lalt1 As Double
Dim lalt2 As Double
Dim lgatomin As Double
Dim lcajonmax As Double
Dim lcajonmin As Double
Dim n6000 As Integer
Dim n4500 As Integer
Dim OffsetDist As Double
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
Dim Eje2 As Object
Dim Eje3 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim ANG As Double
Dim ANG2 As Double
Dim DirBulon1 As Double
Dim DirBulon2 As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim dato1 As String
Dim plcu As String
Dim dato2 As String
Dim dato3 As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object
Dim puntoA As Variant
Dim puntop1 As Variant
Dim puntoB As Variant
Dim puntop2 As Variant
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim x3 As Double
Dim y3 As Double
Dim x4 As Double
Dim y4 As Double
Dim x5 As Double
Dim y5 As Double
Dim PGato1(0 To 2) As Double
Dim PGato2(0 To 2) As Double
Dim PA(0 To 2) As Double
Dim PAP(0 To 2) As Double
Dim PP1(0 To 2) As Double
Dim PP2(0 To 2) As Double
Dim PB(0 To 2) As Double
Dim PBP(0 To 2) As Double
Dim Esq(0 To 2) As Double
Dim userInput As String
Dim muro1menor  As Double
Dim muromoduladome1  As Double
Dim muromoduladoma1 As Double
Dim mod1c As Double
Dim mod1M As Double
Dim muro2menor As Double
Dim muromoduladome2 As Double
Dim muromoduladoma2 As Double
Dim mod2c As Double
Dim mod2M As Double
Dim objGcadDimAligned As GcadDimAligned
Dim TxtPnt(0 To 2) As Double
Dim TxtPnt2(0 To 2) As Double
Dim TxtPnt3(0 To 2) As Double
Dim Pcerca2(0 To 2) As Double
Dim Plejos2(0 To 2) As Double
Dim Pcerca1(0 To 2) As Double
Dim Plejos1(0 To 2) As Double
Dim Slope1 As Double
Dim Slope2 As Double
Dim D_A0 As Double
Dim D_B0 As Double
Dim D_AB As Double
Dim D_ABP As Double
Dim D_Gato As Double
Dim DirMuro1 As Double
Dim DirMuro2 As Double
Dim DirMuro1Inv As Double
Dim DirMuro2Inv As Double
Dim DirPuntal As Double
Dim rutacu As String
Dim DirPuntal2 As Double


Dim NamePre As GcadAttribute
Dim longitudatt As GcadAttribute
Dim orientacionatt As GcadAttribute
Dim cooordenadainicio0 As GcadAttribute
Dim cooordenadainicio1 As GcadAttribute
Dim cooordenadainicio2 As GcadAttribute


Set GcadDoc = GetObject(, "Gcad.Application").ActiveDocument
Set GcadModel = GcadDoc.ModelSpace
Set GcadUtil = GcadDoc.Utility

Ncapa = "Pipeshor4S"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5
Ncapa = "Granshor"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150
Ncapa = "NoContable"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 40



'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgiro = 205
lfusible = 187.5
l145 = 145
l280 = 280
l560 = 560
l750 = 750
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
l50 = 50
l35 = 35
l_tope = 325
l_conogato = 170
lgatomin = 620
lcajonmin = 835
lcajonmax = 1022


On Error GoTo terminar
Dim rutamp As String

ruta = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\" & dato2 & "\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
rutacu = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"
rutacajon = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Cajon hidraulico\"

Dim plcue1 As String
Dim plcue2 As String
Dim plcu2 As String





kwordList = "Planta Alzado"
dato1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Planta/Alzado]")
If dato1 = "Planta" Or dato1 = "" Then
dato1 = "Planta"
plcue1 = "Ialzado"
plcue2 = "Dalzado"
plcu2 = "PLA"
ElseIf dato1 = "Alzado" Then
plcue1 = "Ialzado"
plcue2 = "Dalzado"
plcu2 = "ALZ"
Else
GoTo terminar
End If


Do While repite = 1


lalt1 = 0
lalt2 = 0



puntoA = GcadUtil.GetPoint(, "punto inserción 1ª placa: ")
puntop1 = GcadUtil.GetPoint(puntoA, "punto direccional del muro 1 (convergente): ")

puntoB = GcadUtil.GetPoint(, "punto inserción 2ª placa: ")
puntop2 = GcadUtil.GetPoint(puntoB, "punto direccional del muro 2 (convergente): ")

'PA es el punto de inserción de la primera placa
PA(0) = puntoA(0): PA(1) = puntoA(1): PA(2) = puntoA(2)
PP1(0) = puntop1(0): PP1(1) = puntop1(1): PP1(2) = puntop1(2)

'PB es el punto de inserción de la segunda placa
PB(0) = puntoB(0): PB(1) = puntoB(1): PB(2) = puntoB(2)
PP2(0) = puntop2(0): PP2(1) = puntop2(1): PP2(2) = puntop2(2)

DirMuro1 = GcadUtil.AngleFromXAxis(PA, PP1)
DirMuro2 = GcadUtil.AngleFromXAxis(PB, PP2)
DirPuntal = GcadUtil.AngleFromXAxis(PA, PB)

'''' caso General
'' seleccionar los 4 puntos en sentido horario
Dim k As String, b As Object, entity As Object
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
On Error GoTo terminar
    Dim Respuesta As String
    MsgBox "Ya existe un bloque con este nombre!"
    kwordList = "Sobreescribir Renombrar"
    Respuesta = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    Respuesta = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Sobreescribir/Renombrar]")
    
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

' conseguir la esquina:
' Calculamos las direcciones de las rectas
Slope2 = Tan(DirMuro2)
If DirMuro1 <> (PI / 2) And DirMuro1 <> (3 * PI / 2) Then
    Slope1 = Tan(DirMuro1)
    If DirMuro1 = DirMuro2 Then
    Else
    ' Calculamos el punto intersección
        Esq(0) = (PB(1) - PA(1) - Slope2 * PB(0) + Slope1 * PA(0)) / (Slope1 - Slope2)
        Esq(1) = PA(1) + Slope1 * (Esq(0) - PA(0))
        Esq(2) = PA(2) ' Assuming the lines are in the same plane

    End If
Else
    ' Caso especial: DirMuro1 es 90 grados
    ' En este caso, la línea es vertical, y la intersección se encuentra en la coordenada X de la línea vertical
    Esq(0) = PA(0)
    Esq(1) = Slope2 * (Esq(0) - PB(0)) + PB(1)
    Esq(2) = PA(2) ' Assuming the lines are in the same plane

End If


If Abs(DirMuro2 - DirMuro1) > PI Then
    If DirMuro2 > DirMuro1 Then
        ' Intercambiar los puntos P1 y P3
        Dim tempP0(0 To 2) As Double
        tempP0(0) = PA(0): tempP0(1) = PA(1): tempP0(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = tempP0(0): PB(1) = tempP0(1): PB(2) = tempP0(2):
        ' Recalcular la dirección del muro 1 y perpendicular al muro
        DirMuro1 = GcadUtil.AngleFromXAxis(PA, Esq)
        ' Recalcular la dirección del muro 2 y perpendicular al muro
        DirMuro2 = GcadUtil.AngleFromXAxis(PB, Esq)
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
    End If
Else
    If DirMuro2 < DirMuro1 Then
        ' Intercambiar los puntos P1 y P3
        Dim tempP(0 To 2) As Double
        tempP(0) = PA(0): tempP(1) = PA(1): tempP(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = tempP(0): PB(1) = tempP(1): PB(2) = tempP(2):
        ' Recalcular la dirección del muro 1 y perpendicular al muro
        DirMuro1 = GcadUtil.AngleFromXAxis(PA, Esq)
        ' Recalcular la dirección del muro 2 y perpendicular al muro
        DirMuro2 = GcadUtil.AngleFromXAxis(PB, Esq)
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
    End If
End If





'If DirMuro2 > DirMuro1 Then
    ' Intercambiar los puntos P1 y P3
    'Dim tempP(0 To 2) As Double
    'tempP(0) = PA(0): tempP(1) = PA(1): tempP(2) = PA(2)
    'PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
    'PB(0) = tempP(0): PB(1) = tempP(1): PB(2) = tempP(2)

    ' Recalcular la dirección del muro 1 y perpendicular al muro
    'DirMuro1 = GcadUtil.AngleFromXAxis(PA, Esq)
    ' Recalcular la dirección del muro 2 y perpendicular al muro
    'DirMuro2 = GcadUtil.AngleFromXAxis(PB, Esq)
    'DirPuntal = GcadUtil.AngleFromXAxis(PA, PB)
'Else
    'DirBulon1 = DirMuro1 - (PI / 2)
    'DirBulon2 = DirMuro2 + (PI / 2)
'End If


DirBulon1 = DirMuro1 - (PI / 2)
DirBulon2 = DirMuro2 + (PI / 2)



'''''Extremo 2 del puntal
Dim AnguloAbsoluto1 As Double
Dim UmbralMin1 As Double
Dim UmbralMax1 As Double

UmbralMin1 = 80 * (PI / 180) ' Convertir 80 grados a radianes
UmbralMax1 = 100 * (PI / 180) ' Convertir 105 grados a radianes

AnguloAbsoluto1 = Abs((DirPuntal + PI) - DirMuro2)

If AnguloAbsoluto1 > UmbralMin1 And AnguloAbsoluto1 < UmbralMax1 Then
    
    Dim decisionGiro1 As String
    kwordList = "GiroMP CompactaMP"
    decisionGiro1 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    decisionGiro1 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Qué prefieres usar?: [GiroMP/CompactaMP]")
    If decisionGiro1 = "" Or decisionGiro1 = "GiroMP" Then
        decisionGiro1 = "GiroMP"
        P1(0) = PB(0) + 90 * Cos(DirBulon2): P1(1) = PB(1) + 90 * Sin(DirBulon2): P1(2) = PB(2)
    Else
        P1(0) = PB(0) + 85 * Cos(DirBulon2): P1(1) = PB(1) + 85 * Sin(DirBulon2): P1(2) = PB(2)
    End If

    placaanc1 = decisionGiro1
Else
    kwordList = "Naranja Azul MP CompactaMP"
    placaanc1 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    placaanc1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de cuña en el primer extremo seleccionado: [Naranja/Azul/MP/CompactaMP]")

    ' aquí vienen los condicionales de las diferentes placas de anclaje, según cuál sea la seleccionada, moveremos para encontrar el P1 en este caso
    ' tenemos todas las medidas para P1
    
    If placaanc1 = "" Or placaanc1 = "Naranja" Then
        P1(0) = PB(0) + 239.2 * Cos(DirBulon2): P1(1) = PB(1) + 239.2 * Sin(DirBulon2): P1(2) = PB(2)
        P1(0) = P1(0) + 179.21 * Cos(DirMuro2): P1(1) = P1(1) + 179.21 * Sin(DirMuro2): P1(2) = P1(2)
    ElseIf placaanc1 = "Azul" Then
        P1(0) = PB(0) + 288.7 * Cos(DirBulon2): P1(1) = PB(1) + 288.7 * Sin(DirBulon2): P1(2) = PB(2)
        P1(0) = P1(0) + 213.7 * Cos(DirMuro2): P1(1) = P1(1) + 213.7 * Sin(DirMuro2): P1(2) = P1(2)
    ElseIf placaanc1 = "MP" Then
        P1(0) = PB(0) + 90 * Cos(DirBulon2): P1(1) = PB(1) + 90 * Sin(DirBulon2): P1(2) = PB(2)
    ElseIf placaanc1 = "CompactaMP" Then
        P1(0) = PB(0) + 85 * Cos(DirBulon2): P1(1) = PB(1) + 85 * Sin(DirBulon2): P1(2) = PB(2)
    End If
        
End If




'''''Extremo 1 del puntal
Dim AnguloAbsoluto2 As Double
Dim UmbralMin2 As Double
Dim UmbralMax2 As Double

UmbralMin2 = 80 * (PI / 180) ' Convertir 80 grados a radianes
UmbralMax2 = 100 * (PI / 180) ' Convertir 105 grados a radianes

AnguloAbsoluto2 = Abs(DirPuntal - DirMuro1)

If AnguloAbsoluto2 > UmbralMin2 And AnguloAbsoluto2 < UmbralMax2 Then
    ' No pasa nada
    
    Dim decisionGiro2 As String
    kwordList = "GiroMP CompactaMP"
    decisionGiro2 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    decisionGiro2 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Qué prefieres usar?: [GiroMP/CompactaMP]")
    If decisionGiro2 = "" Or decisionGiro2 = "GiroMP" Then
        decisionGiro2 = "GiroMP"
        P2(0) = PA(0) + 90 * Cos(DirBulon1): P2(1) = PA(1) + 90 * Sin(DirBulon1): P2(2) = PA(2)
    Else
        P2(0) = PA(0) + 85 * Cos(DirBulon1): P2(1) = PA(1) + 85 * Sin(DirBulon1): P2(2) = PA(2)
    End If
    placaanc2 = decisionGiro2
Else
    kwordList = "Naranja Azul MP CompactaMP"
    placaanc2 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    placaanc2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de cuña en el segundo extremo seleccionado: [Naranja/Azul/MP/CompactaMP]")
    
    ' aquí vienen los condicionales de las diferentes placas de anclaje, según cuál sea la seleccionada, moveremos para encontrar el P1 en este caso
    ' tenemos todas las medidas para P1
    
    If placaanc2 = "" Or placaanc2 = "Naranja" Then
        P2(0) = PA(0) + 239.2 * Cos(DirBulon1): P2(1) = PA(1) + 239.2 * Sin(DirBulon1): P2(2) = PA(2)
        P2(0) = P2(0) + 179.21 * Cos(DirMuro1): P2(1) = P2(1) + 179.21 * Sin(DirMuro1): P2(2) = P2(2)
    ElseIf placaanc2 = "Azul" Then
        P2(0) = PA(0) + 288.7 * Cos(DirBulon1): P2(1) = PA(1) + 288.7 * Sin(DirBulon1): P2(2) = PA(2)
        P2(0) = P2(0) + 213.7 * Cos(DirMuro1): P2(1) = P2(1) + 213.7 * Sin(DirMuro1): P2(2) = P2(2)
    ElseIf placaanc2 = "MP" Then
        P2(0) = PA(0) + 90 * Cos(DirBulon1): P2(1) = PA(1) + 90 * Sin(DirBulon1): P2(2) = PA(2)
    ElseIf placaanc2 = "CompactaMP" Then
        P2(0) = PA(0) + 85 * Cos(DirBulon1): P2(1) = PA(1) + 85 * Sin(DirBulon1): P2(2) = PA(2)
    End If
    
End If




DirPuntal = GcadUtil.AngleFromXAxis(P1, P2)
DirPuntal2 = GcadUtil.AngleFromXAxis(P2, P1)

'''' podemos también añadir las rutas de la placa1 y placa2 como hago aquí abajo pero lo metemos directamente en el condicional que tenemos aquí arriba, para dejar ya cerrada cuál va a ser cada una de las placas
'''' además de añadir las capas. En caso de que sea el ángulo de giro podemos añadirlo también.


If placaanc1 = "Naranja" Or placaanc1 = "" Then
    rutaplaca1 = rutacu & "MG_CunaNar_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "Azul" Then
    rutaplaca1 = rutacu & "MG_CunaAz_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "MP" Then
    rutaplaca1 = rutacu & "PlacaMP_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "CompactaMP" Then
    rutaplaca1 = rutacu & "PlacaMP_C_" & plcue2 & ".dwg"
    capa = "Mega"
ElseIf placaanc1 = "GiroMP" Then
    rutaplaca1 = rutacu & "GiroMP_" & plcue2 & ".dwg"
    capa = "Mega"
End If

If placaanc2 = "Naranja" Or placaanc2 = "" Then
    rutaplaca2 = rutacu & "MG_CunaNar_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "Azul" Then
    rutaplaca2 = rutacu & "MG_CunaAz_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "MP" Then
    rutaplaca2 = rutacu & "PlacaMP_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "CompactaMP" Then
    rutaplaca2 = rutacu & "PlacaMP_C_" & plcue1 & ".dwg"
    capa = "Mega"
ElseIf placaanc2 = "GiroMP" Then
    rutaplaca2 = rutacu & "GiroMP_" & plcue1 & ".dwg"
    capa = "Mega"
End If


Dim agiro As String
agiro = rutacu & "ANGgiro.dwg"

DirMuro1Inv = DirMuro1 - PI
DirMuro2Inv = DirMuro2 - PI

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' Aquí viene la toma de decisiones de si coger un puntal o coger otro::::
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If carga <= 1020 Then
    lfija = 680
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
ElseIf carga > 1020 And carga <= 1200 Then
    lfija = 720
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
ElseIf carga > 1200 Then
    lfija = 760
    If placaanc1 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc1 = "CompactaMP" Then
        lfija = lfija + 5
    End If
    If placaanc2 = "MP" Then
        lfija = lfija + 225
    ElseIf placaanc2 = "CompactaMP" Then
        lfija = lfija + 5
    End If
End If

Xs = 1
Ys = 1
Zs = 1

Dim Longitud As String, orientacion As String, Pinicio0 As String, Pinicio1 As String, Pinicio2 As String, PreMon As String


x = P2(0) - P1(0)
y = P2(1) - P1(1)

Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

    Dim tubo As String
    Dim placa1 As String
    Dim placa2 As String
    Dim vista As String

If carga >= 750 And carga <= 1350 Then
    If Distancia >= 5000 And Distancia <= 12000 Then
        Set b = ThisDrawing.Blocks.Add(PA, k)
        
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca2, Xs, Ys, Zs, DirMuro1)
        blockRef.Layer = "Mega"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca1, Xs, Ys, Zs, DirMuro2)
        blockRef.Layer = "Mega"
        
        ANG = GcadUtil.AngleFromXAxis(P1, P2)
        ANG2 = ANG + (PI / 2)
        
        Pinicio0 = CStr(P1(0))
        Pinicio1 = CStr(P1(1))
        Pinicio2 = CStr(P1(2))
        Longitud = CStr(Distancia)
        orientacion = CStr(ANG)
        
        PAP(0) = PA(0): PAP(1) = PA(1): PAP(2) = PA(2)
        PBP(0) = PB(0): PBP(1) = PB(1): PBP(2) = PB(2)
        x4 = PAP(0) - PBP(0)
        y4 = PAP(1) - PBP(1)
        D_ABP = Val(Sqr((x4 ^ 2 + y4 ^ 2)))
        
        TxtPnt2(0) = PBP(0) + (D_ABP / 2) * Cos(ANG): TxtPnt2(1) = PBP(1) + (D_ABP / 2) * Sin(ANG): TxtPnt2(2) = PBP(2)
        TxtPnt2(0) = TxtPnt2(0) + 860 * Cos(ANG2): TxtPnt2(1) = TxtPnt2(1) + 860 * Sin(ANG2): TxtPnt2(2) = TxtPnt2(2)
        
        TxtPnt(0) = P1(0) + (Distancia / 2) * Cos(ANG): TxtPnt(1) = P1(1) + (Distancia / 2) * Sin(ANG): TxtPnt(2) = P1(2)
        TxtPnt(0) = TxtPnt(0) + 410 * Cos(ANG2): TxtPnt(1) = TxtPnt(1) + 410 * Sin(ANG2): TxtPnt(2) = TxtPnt(2)
        
        Set objGcadDimAligned = GcadModel.AddDimAligned(P1, P2, TxtPnt)
        objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
        objGcadDimAligned.StyleName = "MODELO"
        objGcadDimAligned.TextStyle = "SIMPLEX"
        objGcadDimAligned.VerticalTextPosition = acOutside
        objGcadDimAligned.Update
        objGcadDimAligned.Layer = "Dimension"
        
        Set objGcadDimAligned = GcadModel.AddDimAligned(PBP, PAP, TxtPnt2)
        objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
        objGcadDimAligned.StyleName = "MODELO"
        objGcadDimAligned.TextStyle = "SIMPLEX"
        objGcadDimAligned.VerticalTextPosition = acOutside
        objGcadDimAligned.Update
        objGcadDimAligned.Layer = "Dimension"
        
        tubo = "Pshor_4L"
        ' mandamos al pm la selección de placa 1
        If placaanc1 = "Naranja" Or placaanc1 = "Azul" Then
            placa1 = "Cuña"
        ElseIf placaanc1 = "CompactaMP" Then
            placa1 = "PlacaMPCompacta"
        ElseIf placaanc1 = "MP" Then
            placa1 = "PlacaMP"
        ElseIf placaanc1 = "GiroMP" Then
            placa1 = "Cuña"
        End If
        ' mandamos al pm la selección de placa 2
        If placaanc2 = "Naranja" Or placaanc2 = "Azul" Then
            placa2 = "Cuña"
        ElseIf placaanc2 = "CompactaMP" Then
            placa2 = "PlacaMPCompacta"
        ElseIf placaanc2 = "MP" Then
            placa2 = "PlacaMP"
        ElseIf placaanc2 = "GiroMP" Then
            placa2 = "Cuña"
        End If
        ' mandamos al pm la selección de vista
        If dato1 = "Planta" Or dato1 = "" Then
            vista = "Planta"
        ElseIf dato1 = "Alzado" Then
            vista = "Alzado"
        End If
        
        Call pm(P1, P2, placa1, placa2, vista, tubo, k)
        
        PreMon = ""

        Set NamePre = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "NombrePremontaje", PreMon)

        Set longitudatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Longitud", Longitud)

        Set orientacionatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Orientacion", orientacion)

        Set cooordenadainicio0 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada0", Pinicio0)

        Set cooordenadainicio1 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada1", Pinicio1)

        Set cooordenadainicio2 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada2", Pinicio2)

        
        Set blockRef = ThisDrawing.ModelSpace.InsertBlock(PA, k, Xs, Ys, Zs, 0)
        blockRef.Layer = "NoContable"
        blockRef.Update

        GoTo volver
    ElseIf Distancia > 12000 And Distancia <= 15000 Then
    
        Set b = ThisDrawing.Blocks.Add(PA, k)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca2, Xs, Ys, Zs, DirMuro1)
        blockRef.Layer = "Mega"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca1, Xs, Ys, Zs, DirMuro2)
        blockRef.Layer = "Mega"
        
        ANG = GcadUtil.AngleFromXAxis(P1, P2)
        ANG2 = ANG + (PI / 2)
        
        Pinicio0 = CStr(P1(0))
        Pinicio1 = CStr(P1(1))
        Pinicio2 = CStr(P1(2))
        Longitud = CStr(Distancia)
        orientacion = CStr(ANG)
        
        PAP(0) = PA(0): PAP(1) = PA(1): PAP(2) = PA(2)
        PBP(0) = PB(0): PBP(1) = PB(1): PBP(2) = PB(2)
        x4 = PAP(0) - PBP(0)
        y4 = PAP(1) - PBP(1)
        D_ABP = Val(Sqr((x4 ^ 2 + y4 ^ 2)))
        
        TxtPnt2(0) = PBP(0) + (D_ABP / 2) * Cos(ANG): TxtPnt2(1) = PBP(1) + (D_ABP / 2) * Sin(ANG): TxtPnt2(2) = PBP(2)
        TxtPnt2(0) = TxtPnt2(0) + 860 * Cos(ANG2): TxtPnt2(1) = TxtPnt2(1) + 860 * Sin(ANG2): TxtPnt2(2) = TxtPnt2(2)
        
        TxtPnt(0) = P1(0) + (Distancia / 2) * Cos(ANG): TxtPnt(1) = P1(1) + (Distancia / 2) * Sin(ANG): TxtPnt(2) = P1(2)
        TxtPnt(0) = TxtPnt(0) + 410 * Cos(ANG2): TxtPnt(1) = TxtPnt(1) + 410 * Sin(ANG2): TxtPnt(2) = TxtPnt(2)
        
        Set objGcadDimAligned = GcadModel.AddDimAligned(P1, P2, TxtPnt)
        objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
        objGcadDimAligned.StyleName = "MODELO"
        objGcadDimAligned.TextStyle = "SIMPLEX"
        objGcadDimAligned.VerticalTextPosition = acOutside
        objGcadDimAligned.Update
        objGcadDimAligned.Layer = "Dimension"
        
        Set objGcadDimAligned = GcadModel.AddDimAligned(PBP, PAP, TxtPnt2)
        objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
        objGcadDimAligned.StyleName = "MODELO"
        objGcadDimAligned.TextStyle = "SIMPLEX"
        objGcadDimAligned.VerticalTextPosition = acOutside
        objGcadDimAligned.Update
        objGcadDimAligned.Layer = "Dimension"
        
        tubo = "Pshor_4S"
        ' mandamos al pm la selección de placa 1
        If placaanc1 = "Naranja" Or placaanc1 = "Azul" Then
            placa1 = "Cuña"
        ElseIf placaanc1 = "CompactaMP" Then
            placa1 = "PlacaMPCompacta"
        ElseIf placaanc1 = "MP" Then
            placa1 = "PlacaMP"
        ElseIf placaanc1 = "GiroMP" Then
            placa1 = "Cuña"
        End If
        ' mandamos al pm la selección de placa 2
        If placaanc2 = "Naranja" Or placaanc2 = "Azul" Then
            placa2 = "Cuña"
        ElseIf placaanc2 = "CompactaMP" Then
            placa2 = "PlacaMPCompacta"
        ElseIf placaanc2 = "MP" Then
            placa2 = "PlacaMP"
        ElseIf placaanc2 = "GiroMP" Then
            placa2 = "Cuña"
        End If
        ' mandamos al pm la selección de vista
        If dato1 = "Planta" Or dato1 = "" Then
            vista = "Planta"
        ElseIf dato1 = "Alzado" Then
            vista = "Alzado"
        End If
        Call pm(P1, P2, placa1, placa2, vista, tubo, k)
        
        PreMon = ""

        Set NamePre = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "NombrePremontaje", PreMon)
        

        Set longitudatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Longitud", Longitud)
                

        Set orientacionatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Orientacion", orientacion)
                   

        Set cooordenadainicio0 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada0", Pinicio0)
                

        Set cooordenadainicio1 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada1", Pinicio1)

        Set cooordenadainicio2 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada2", Pinicio2)

        
        Set blockRef = ThisDrawing.ModelSpace.InsertBlock(PA, k, Xs, Ys, Zs, 0)
        blockRef.Layer = "NoContable"
        blockRef.Update
        
        GoTo volver
    ElseIf Distancia > 15000 Then
        MsgBox "Puntal de más de 15 metros, revisar si no es necesario terminación con gato Pipeshor."
        GoTo terminar
    End If
End If




Dim n5400 As Integer
Dim n2700 As Integer
Dim n1800 As Integer
Dim n900 As Integer
Dim n450 As Integer
Dim n270 As Integer
Dim n180 As Integer
Dim n90 As Integer



lpuntal = Distancia - lfija
n5400 = Fix(lpuntal / 5400)
lpuntal = lpuntal - n5400 * 5400
n2700 = Fix(lpuntal / 2700)
lpuntal = lpuntal - n2700 * 2700
n1800 = Fix(lpuntal / 1800)
lpuntal = lpuntal - n1800 * 1800
n900 = Fix(lpuntal / 900)
lpuntal = lpuntal - n900 * 900
n450 = Fix(lpuntal / 450)
lpuntal = lpuntal - n450 * 450
n270 = Fix(lpuntal / 270)
lpuntal = lpuntal - n270 * 270
n180 = Fix(lpuntal / 180)
lpuntal = lpuntal - n180 * 180
n90 = Fix(lpuntal / 90)
lpuntal = lpuntal - n90 * 90


'''''' aquí OPTIMIZACIÓN

Dim D_P20 As Double
Dim D_P10 As Double
Dim Oprima(0 To 2) As Double
Dim mmm1 As Double, mmm2 As Double, modulacionmenorM1 As Double, modulacionmenorM2 As Double, modulacionmayorM1 As Double, modulacionmayorM2 As Double, mmmayor1 As Double, mmmayor2 As Double


If n90 > 0 Then
    lalt1 = Distancia - 90
    lalt2 = Distancia + 90
    
    If DirMuro1 = DirMuro2 Then
    Else
    ' Calculamos el punto intersección entre P1 y P2 con las direcciones del muro
        Oprima(0) = (P1(1) - P2(1) - Slope2 * P1(0) + Slope1 * P2(0)) / (Slope1 - Slope2)
        Oprima(1) = P2(1) + Slope1 * (Oprima(0) - P2(0))
        Oprima(2) = P2(2) ' Assuming the lines are in the same plane
        
        'Distancia entre P2 y la Oprima
        x1 = Oprima(0) - P2(0)
        y1 = Oprima(1) - P2(1)
        D_P20 = Val(Sqr((x1 ^ 2 + y1 ^ 2)))
        
        'Distancia entre P1 y la Oprima
        x2 = Oprima(0) - P1(0)
        y2 = Oprima(1) - P1(1)
        D_P10 = Val(Sqr((x2 ^ 2 + y2 ^ 2)))
        
        
        ' modulamos hacia la esquina con un puntal más pequeño
        mmm1 = (lalt1 * D_P20) / Distancia
        mmm2 = (lalt1 * D_P10) / Distancia
        
        modulacionmenorM1 = D_P20 - mmm1
        modulacionmenorM2 = D_P10 - mmm2
        
        ' modulamos con un puntal más grande
        mmmayor1 = (lalt2 * D_P20) / Distancia
        mmmayor2 = (lalt2 * D_P10) / Distancia
        
        modulacionmayorM1 = mmmayor1 - D_P20
        modulacionmayorM2 = mmmayor2 - D_P10
        
        If ((modulacionmenorM1 + modulacionmayorM1) < 1400) Or ((modulacionmenorM2 + modulacionmayorM2) < 1400) Then
            
            userInput = InputBox("Elija una de las siguientes opciones:" & vbCrLf & vbCrLf & vbCrLf & "1. Dibujar el puntal seleccionado de longitud " & Distancia & "." & vbCrLf & vbCrLf & vbCrLf & "2. Dibujar un puntal MENOR de " & lalt1 & "mm de longitud más cercano a la esquina" & vbCrLf & vbCrLf & vbCrLf & "3. Dibujar un puntal MAYOR de " & lalt2 & "mm de longitud más alejado de la esquina")

            If userInput = "1" Or userInput = "" Then
            ElseIf userInput = "2" Then
                ' modulamos si elegimos coger el puntal más pequeño
                PA(0) = PA(0) + modulacionmenorM1 * Cos(DirMuro1): PA(1) = PA(1) + modulacionmenorM1 * Sin(DirMuro1): PA(2) = PA(2)
                PB(0) = PB(0) + modulacionmenorM2 * Cos(DirMuro2): PB(1) = PB(1) + modulacionmenorM2 * Sin(DirMuro2): PB(2) = PB(2)
                
                P2(0) = P2(0) + modulacionmenorM1 * Cos(DirMuro1): P2(1) = P2(1) + modulacionmenorM1 * Sin(DirMuro1): P2(2) = P2(2)
                P1(0) = P1(0) + modulacionmenorM2 * Cos(DirMuro2): P1(1) = P1(1) + modulacionmenorM2 * Sin(DirMuro2): P1(2) = P1(2)
                
                n90 = n90 - 1
                
            ElseIf userInput = "3" Then
                ' modulamos si elegimos coger el puntal más grande
                PA(0) = PA(0) + modulacionmayorM1 * Cos(DirMuro1Inv): PA(1) = PA(1) + modulacionmayorM1 * Sin(DirMuro1Inv): PA(2) = PA(2)
                PB(0) = PB(0) + modulacionmayorM2 * Cos(DirMuro2Inv): PB(1) = PB(1) + modulacionmayorM2 * Sin(DirMuro2Inv): PB(2) = PB(2)
                
                P2(0) = P2(0) + modulacionmayorM1 * Cos(DirMuro1Inv): P2(1) = P2(1) + modulacionmayorM1 * Sin(DirMuro1Inv): P2(2) = P2(2)
                P1(0) = P1(0) + modulacionmayorM2 * Cos(DirMuro2Inv): P1(1) = P1(1) + modulacionmayorM2 * Sin(DirMuro2Inv): P1(2) = P1(2)
                
                n90 = n90 - 1
                n180 = n180 + 1
            End If
        Else
        End If

    End If
    

End If


' vamos a colocar la cuña/ angulo de giro en el extremo 1
    Set b = ThisDrawing.Blocks.Add(PA, k)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca2, Xs, Ys, Zs, DirMuro1)
    blockRef.Layer = "Mega"
'End If



' vamos a colocar la cuña/angulo de giro en el extremo 2

    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca1, Xs, Ys, Zs, DirMuro2)
    blockRef.Layer = "Mega"



Set Eje1 = ThisDrawing.Blocks.Item(k).AddLine(P1, P2)
Eje1.Layer = "Nonplot"
ANG = GcadUtil.AngleFromXAxis(P1, P2)
ANG2 = ANG + (PI / 2)

Pinicio0 = CStr(P1(0))
Pinicio1 = CStr(P1(1))
Pinicio2 = CStr(P1(2))
Longitud = CStr(Distancia)
orientacion = CStr(ANG)

If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If

'Puntos centrales de las placas

PAP(0) = PA(0): PAP(1) = PA(1): PAP(2) = PA(2)
PBP(0) = PB(0): PBP(1) = PB(1): PBP(2) = PB(2)
x4 = PAP(0) - PBP(0)
y4 = PAP(1) - PBP(1)
D_ABP = Val(Sqr((x4 ^ 2 + y4 ^ 2)))

TxtPnt2(0) = PBP(0) + (D_ABP / 2) * Cos(ANG): TxtPnt2(1) = PBP(1) + (D_ABP / 2) * Sin(ANG): TxtPnt2(2) = PBP(2)
TxtPnt2(0) = TxtPnt2(0) + 860 * Cos(ANG2): TxtPnt2(1) = TxtPnt2(1) + 860 * Sin(ANG2): TxtPnt2(2) = TxtPnt2(2)

TxtPnt(0) = P1(0) + (Distancia / 2) * Cos(ANG): TxtPnt(1) = P1(1) + (Distancia / 2) * Sin(ANG): TxtPnt(2) = P1(2)
TxtPnt(0) = TxtPnt(0) + 410 * Cos(ANG2): TxtPnt(1) = TxtPnt(1) + 410 * Sin(ANG2): TxtPnt(2) = TxtPnt(2)

Set objGcadDimAligned = GcadModel.AddDimAligned(P1, P2, TxtPnt)
objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objGcadDimAligned.StyleName = "MODELO"
objGcadDimAligned.TextStyle = "SIMPLEX"
objGcadDimAligned.VerticalTextPosition = acOutside
objGcadDimAligned.Update
objGcadDimAligned.Layer = "Dimension"

Set objGcadDimAligned = GcadModel.AddDimAligned(PBP, PAP, TxtPnt2)
objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objGcadDimAligned.StyleName = "MODELO"
objGcadDimAligned.TextStyle = "SIMPLEX"
objGcadDimAligned.VerticalTextPosition = acOutside
objGcadDimAligned.Update
objGcadDimAligned.Layer = "Dimension"


Punto_inial(0) = P1(0): Punto_inial(1) = P1(1): Punto_inial(2) = P1(2)
Punto_final(0) = P2(0): Punto_final(1) = P2(1): Punto_final(2) = P2(2)
M20x50_4 = ruta2 & "4M20X50.dwg"
M20x60_4 = ruta2 & "4M20X60.dwg"
'
' meter el ángulo de giro y los jackplates + el gato donde haga falta
' gato en el EXTREMO 1

Dim CuMP As String
Dim CuMPc As String


If placaanc1 = "MP" Then
    CuMP = rutacu & "PL_GCODAL_PLA.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, CuMP, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 315 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 315 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
ElseIf placaanc1 = "CompactaMP" Then
    CuMPc = rutacu & "PL_GCODAL_C_PLA.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, CuMPc, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 95 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 95 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
ElseIf placaanc1 = "GiroMP" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, agiro, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Else
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, agiro, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Mega"
    Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
End If


    



''' aquí va el FUSIBLE
Dim mp_fus As String

Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x50_4, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"

mp_fus = rutamp & "Mshor90" & plcu2 & "fusible.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_fus, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"





Dim mp_90 As String
Dim mp_270 As String
Dim mp_900 As String
Dim mp_2700 As String
Dim mp_5400 As String
Dim mp_1800 As String
Dim mp_450 As String
Dim mp_180 As String

If n90 > 0 Then
        mp_90 = rutamp & "Mshor90" & plcu2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 90 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 90 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n270 > 0 Then
    i = 0
    Do While i < n270
        mp_270 = rutamp & "Mshor270" & plcu2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 270 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 270 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n900 > 0 Then
        mp_900 = rutamp & "Mshor900" & plcu2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_900, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 900 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 900 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n2700 > 0 Then
        mp_2700 = rutamp & "Mshor2700" & plcu2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_2700, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 2700 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 2700 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If

If n5400 > 0 Then
    i = 0
    Do While i < n5400
        mp_5400 = rutamp & "Mshor5400" & plcu2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_5400, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 5400 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 5400 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If n1800 > 0 Then
    i = 0
    Do While i < n1800
        mp_1800 = rutamp & "Mshor1800" & plcu2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_1800, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 1800 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 1800 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If


If n450 > 0 Then
        mp_450 = rutamp & "Mshor450" & plcu2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 450 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 450 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
End If


If n180 > 0 Then
    i = 0
    Do While i < n180
        mp_180 = rutamp & "Mshor180" & plcu2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 180 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 180 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
        i = i + 1
    Loop
End If

If carga > 1020 Then
    '
End If



' primer jack si hace falta

Dim MP_JP As String
Dim M20x110_4 As String
M20x110_4 = ruta2 & "4M20X110.dwg"

If carga > 1020 Then
        MP_JP = rutamp & "MshorJACKPLATE.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, MP_JP, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Punto_inial(0) = Punto_inial(0) + 40 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 40 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x110_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"

End If
PGato2(0) = Punto_inial(0): PGato2(1) = Punto_inial(1): PGato2(2) = Punto_inial(2)

' base azul
Dim base_azul As String
base_azul = rutacu & "zMGBaseGato_azul.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, base_azul, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_inial(0) = Punto_inial(0) + 150 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 150 * Sin(ANG): Punto_inial(2) = Punto_inial(2)


' ahora vamos al P2 a meter lo que haga falta para respetar la apertura del gato y metemos en orden inverso lo demás
' angulo de giro o la terminación de las cuñas que hagan falta
If placaanc2 = "MP" Then
    CuMP = rutacu & "PL_GCODAL_PLA.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, CuMP, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 315 * Cos(ANG): Punto_final(1) = Punto_final(1) - 315 * Sin(ANG): Punto_final(2) = Punto_final(2)
ElseIf placaanc2 = "CompactaMP" Then
    CuMPc = rutacu & "PL_GCODAL_C_PLA.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, CuMPc, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 95 * Cos(ANG): Punto_final(1) = Punto_final(1) - 95 * Sin(ANG): Punto_final(2) = Punto_final(2)
ElseIf placaanc2 = "GiroMP" Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, agiro, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 90 * Cos(ANG): Punto_final(1) = Punto_final(1) - 90 * Sin(ANG): Punto_final(2) = Punto_final(2)
Else
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, agiro, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Punto_final(0) = Punto_final(0) - 90 * Cos(ANG): Punto_final(1) = Punto_final(1) - 90 * Sin(ANG): Punto_final(2) = Punto_final(2)
End If
    

' segundo jack si hace falta
If carga > 1200 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x110_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"

        Punto_final(0) = Punto_final(0) - 40 * Cos(ANG): Punto_final(1) = Punto_final(1) - 40 * Sin(ANG): Punto_final(2) = Punto_final(2)
        MP_JP = rutamp & "MshorJACKPLATE.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, MP_JP, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
End If

PGato1(0) = Punto_final(0): PGato1(1) = Punto_final(1): PGato1(2) = Punto_final(2)

' base naranja del gato
Dim base_naranja As String
base_naranja = rutacu & "zMGBaseGato_naranja.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, base_naranja, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Mega"
If carga < 1200 Then
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x60_4, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
End If


Punto_final(0) = Punto_final(0) - 150 * Cos(ANG): Punto_final(1) = Punto_final(1) - 150 * Sin(ANG): Punto_final(2) = Punto_final(2)



x1 = Punto_final(0) - Punto_inial(0)
y1 = Punto_final(1) - Punto_inial(1)

Distancia = Val(Sqr((x1 ^ 2 + y1 ^ 2)))
Distancia = Distancia / 2


Punto_inial(0) = Punto_inial(0) + Distancia * Cos(ANG): Punto_inial(1) = Punto_inial(1) + Distancia * Sin(ANG): Punto_inial(2) = Punto_inial(2)


' husillo
Dim husillo As String
husillo = rutacu & "MGHusilloGato.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, husillo, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"

x5 = PGato2(0) - PGato1(0)
y5 = PGato2(1) - PGato1(1)
D_Gato = Val(Sqr((x5 ^ 2 + y5 ^ 2)))

TxtPnt3(0) = PGato1(0) + (D_Gato / 2) * Cos(ANG): TxtPnt3(1) = PGato1(1) + (D_Gato / 2) * Sin(ANG): TxtPnt3(2) = PGato1(2)
TxtPnt3(0) = TxtPnt3(0) - 350 * Cos(ANG2): TxtPnt3(1) = TxtPnt3(1) - 350 * Sin(ANG2): TxtPnt3(2) = TxtPnt3(2)

Set objGcadDimAligned = GcadModel.AddDimAligned(PGato1, PGato2, TxtPnt3)
objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objGcadDimAligned.StyleName = "MODELO"
objGcadDimAligned.TextStyle = "SIMPLEX"
objGcadDimAligned.VerticalTextPosition = acOutside
objGcadDimAligned.Update
objGcadDimAligned.Layer = "Dimension"

PreMon = ""

Set NamePre = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "NombrePremontaje", PreMon)


Set longitudatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Longitud", Longitud)
        

Set orientacionatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Orientacion", orientacion)
           

Set cooordenadainicio0 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada0", Pinicio0)
        

Set cooordenadainicio1 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada1", Pinicio1)
        

Set cooordenadainicio2 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada2", Pinicio2)


Set blockRef = ThisDrawing.ModelSpace.InsertBlock(PA, k, Xs, Ys, Zs, 0)
blockRef.Layer = "NoContable"
blockRef.Update

Eje1.Layer = "Nonplot"
volver:
Loop

terminar:

End Sub



Sub psn()
Dim ruta As String, rutaps As String, rutapl As String, rutags As String
Dim ruta2 As String
Dim GcadDoc As Object
Dim GcadUtil As Object
Dim GcadModel As Object
Dim punto1 As Variant
Dim punto2 As Variant
Dim x As Double
Dim y As Double
Dim z As Double
Dim line2 As GcadLine
Dim line1 As GcadLine
Dim M20x90 As String
Dim M20x150 As String, M20x110 As String, Var20x250 As String
Dim M20x100 As String
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
Dim placaanc1 As String
Dim rutaplaca1 As String, rutacajon As String
Dim placaanc2 As String
Dim rutaplaca2 As String
Dim basecajon As String, brazocajon As String
Dim PS_Husillo As String
Dim PS_Placa50mm As String
Dim PS_Placa35mm As String
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
Dim l50 As Double, l35 As Double
Dim l145 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lalt1 As Double
Dim lalt2 As Double
Dim lgatomin As Double
Dim lcajonmax As Double
Dim lcajonmin As Double
Dim n6000 As Integer
Dim n4500 As Integer
Dim OffsetDist As Double
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
Dim Eje2 As Object
Dim Eje3 As Object
Dim Xs As Double
Dim Ys As Double
Dim Zs As Double
Dim ANG As Double
Dim ANG2 As Double
Dim DirBulon1 As Double
Dim DirBulon2 As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim PBloque(0 To 2) As Double
Dim dato1 As String
Dim dato2 As String
Dim dato3 As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object
Dim puntoA As Variant
Dim puntop1 As Variant
Dim puntoB As Variant
Dim puntop2 As Variant
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim x3 As Double
Dim y3 As Double
Dim x4 As Double
Dim y4 As Double
Dim x5 As Double
Dim y5 As Double
Dim PGato1(0 To 2) As Double
Dim PGato2(0 To 2) As Double
Dim PA(0 To 2) As Double
Dim PAP(0 To 2) As Double
Dim PP1(0 To 2) As Double
Dim PP2(0 To 2) As Double
Dim PB(0 To 2) As Double
Dim PBP(0 To 2) As Double
Dim Esq(0 To 2) As Double
Dim userInput As String
Dim muro1menor  As Double
Dim muromoduladome1  As Double
Dim muromoduladoma1 As Double
Dim mod1c As Double
Dim mod1M As Double
Dim muro2menor As Double
Dim muromoduladome2 As Double
Dim muromoduladoma2 As Double
Dim mod2c As Double
Dim mod2M As Double
Dim objGcadDimAligned As GcadDimAligned
Dim TxtPnt(0 To 2) As Double
Dim TxtPnt2(0 To 2) As Double
Dim TxtPnt3(0 To 2) As Double
Dim Pcerca2(0 To 2) As Double
Dim Plejos2(0 To 2) As Double
Dim Pcerca1(0 To 2) As Double
Dim Plejos1(0 To 2) As Double
Dim Slope1 As Double
Dim Slope2 As Double
Dim D_A0 As Double
Dim D_B0 As Double
Dim D_AB As Double
Dim D_ABP As Double
Dim D_Gato As Double
Dim DirMuro1 As Double
Dim DirMuro2 As Double
Dim DirMuro1Inv As Double
Dim DirMuro2Inv As Double
Dim DirPuntal As Double
Dim DirPuntal2 As Double

Set GcadDoc = GetObject(, "Gcad.Application").ActiveDocument
Set GcadModel = GcadDoc.ModelSpace
Set GcadUtil = GcadDoc.Utility

Ncapa = "Pipeshor4S"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5
Ncapa = "Granshor"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 150
Ncapa = "NoContable"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 40
Ncapa = "Nonplot"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 50

'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgiro = 205
lfusible = 187.5
l145 = 145
l280 = 280
l560 = 560
l750 = 750
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
l50 = 50
l35 = 35
l_tope = 325
l_conogato = 170
lgatomin = 620
lcajonmin = 835
lcajonmax = 1022

Dim carga As Double
Do While repite = 1




On Error GoTo terminar




carga = InputBox("Introduce carga soportada por el codal (kN ELU): ", "Carga", 0)

If (carga < 1350) Then
    Call mpn(carga)
    GoTo terminar
ElseIf (carga >= 1350) And (carga < 1500) Then
    lfija = (2 * lgiro) + lfusible + l50 + lgatomin
ElseIf (carga >= 1500) And (carga < 2000) Then
    lfija = (2 * lgiro) + lfusible + l50 + l35 + lgatomin
ElseIf (carga >= 2000) And (carga < 2900) Then
    lfija = (2 * lgiro) + lfusible + (2 * l50) + lgatomin
ElseIf carga >= 2900 Then
    lfija = (2 * lgiro) + lfusible + (5 * l50) + lcajonmin
ElseIf carga = "" Then
    MsgBox "Introduce una carga para continuar"
End If


Dim k As String, b As Object, entity As Object
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
    MsgBox "Ya existe un bloque con este nombre!"
    kwordList = "Sobreescribir Renombrar"
    Respuesta = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    Respuesta = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce: [Sobreescribir/Renombrar]")
    
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

On Error GoTo terminar

kwordList = "S L PS6"
dato2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato2 = ThisDrawing.Utility.GetKeyword(vbLf & "Introduce PS4S o PS4L: [S/L/PS6]")

If dato2 = "PS6" Then
Else
    dato2 = "Pshor_4" & dato2
End If


    If dato2 = "Pshor_4L" Then
    capa = "Pipeshor4L"
    dato3 = "PL"
    ElseIf dato2 = "Pshor_4S" Or dato2 = "Pshor_4" Then
    dato2 = "Pshor_4S"
    capa = "Pipeshor4S"
    dato3 = "PS"
    ElseIf dato2 = "PS6" Then
    Else
    GoTo terminar
    End If


ruta = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\" & dato2 & "\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"
rutacajon = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Cajon hidraulico\"

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


''''''' variables necesarias por meter aquiiiii ?????????????)&/(%/(%/(%/
kwordList = "Granshor Compacta"
placaanc2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
placaanc2 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de placa en el primer extremo seleccionado: [Granshor/Compacta]")


kwordList = "Granshor Compacta"
placaanc1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
placaanc1 = ThisDrawing.Utility.GetKeyword(vbLf & "Tipo de placa en el segundo extremo seleccionado: [Granshor/Compacta]")


If placaanc1 = "" Or placaanc1 = "Granshor" Then
    rutaplaca1 = rutags & "GS_PlacaAnclaje_Ialzado.dwg"
    capa = "Granshor"
    'If dato1 = "Alzado" Then
        'rutaplaca1 = rutags & "GS_PlacaAnclaje_seccion.dwg"
    'End If
ElseIf placaanc1 = "Compacta" Then
    rutaplaca1 = rutags & "GS_Placacompacta_Ialzado.dwg"
    capa = "Granshor"
    'If dato1 = "Alzado" Then
        'rutaplaca1 = rutags & "GS_Placacompacta_seccion.dwg"
    'End If
End If

If placaanc2 = "" Or placaanc2 = "Granshor" Then
    rutaplaca2 = rutags & "GS_PlacaAnclaje_Dalzado.dwg"
    capa = "Granshor"
    'If dato1 = "Alzado" Then
        'rutaplaca2 = rutags & "GS_PlacaAnclaje_seccion.dwg"
    'End If
ElseIf placaanc2 = "Compacta" Then
    rutaplaca2 = rutags & "GS_Placacompacta_Dalzado.dwg"
    capa = "Granshor"
    'If dato1 = "Alzado" Then
        'rutaplaca2 = rutags & "GS_Placacompacta_seccion.dwg"
    'End If
End If



'Do While repite = 1


lalt1 = 0
lalt2 = 0

'''' caso General
'' seleccionar los 4 puntos en sentido horario


puntoA = GcadUtil.GetPoint(, "punto inserción 1ª placa: ")
puntop1 = GcadUtil.GetPoint(puntoA, "punto direccional del muro 1 (convergente): ")

puntoB = GcadUtil.GetPoint(, "punto inserción 2ª placa: ")
puntop2 = GcadUtil.GetPoint(puntoB, "punto direccional del muro 2 (convergente): ")

'PA es el punto de inserción de la primera placa
PA(0) = puntoA(0): PA(1) = puntoA(1): PA(2) = puntoA(2)
PP1(0) = puntop1(0): PP1(1) = puntop1(1): PP1(2) = puntop1(2)

'PB es el punto de inserción de la segunda placa
PB(0) = puntoB(0): PB(1) = puntoB(1): PB(2) = puntoB(2)
PP2(0) = puntop2(0): PP2(1) = puntop2(1): PP2(2) = puntop2(2)

DirMuro1 = GcadUtil.AngleFromXAxis(PA, PP1)
DirMuro2 = GcadUtil.AngleFromXAxis(PB, PP2)
DirPuntal = GcadUtil.AngleFromXAxis(PA, PB)

' conseguir la esquina:
' Calculamos las direcciones de las rectas
Slope2 = Tan(DirMuro2)
If DirMuro1 <> (PI / 2) And DirMuro1 <> (3 * PI / 2) Then
    Slope1 = Tan(DirMuro1)
    If DirMuro1 = DirMuro2 Then
    Else
    ' Calculamos el punto intersección
        Esq(0) = (PB(1) - PA(1) - Slope2 * PB(0) + Slope1 * PA(0)) / (Slope1 - Slope2)
        Esq(1) = PA(1) + Slope1 * (Esq(0) - PA(0))
        Esq(2) = PA(2) ' Assuming the lines are in the same plane
        
        'Distancia entre A y la esquina
        x1 = Esq(0) - PA(0)
        y1 = Esq(1) - PA(1)
        D_A0 = Val(Sqr((x1 ^ 2 + y1 ^ 2)))
        
        'Distancia entre B y la esquina
        x2 = Esq(0) - PB(0)
        y2 = Esq(1) - PB(1)
        D_B0 = Val(Sqr((x2 ^ 2 + y2 ^ 2)))
        
        'Distacia entre A y B
        x3 = PA(0) - PB(0)
        y3 = PA(1) - PB(1)
        D_AB = Val(Sqr((x3 ^ 2 + y3 ^ 2)))
    End If
Else
    ' Caso especial: DirMuro1 es 90 grados
    ' En este caso, la línea es vertical, y la intersección se encuentra en la coordenada X de la línea vertical
    Esq(0) = PA(0)
    Esq(1) = Slope2 * (Esq(0) - PB(0)) + PB(1)
    Esq(2) = PA(2) ' Assuming the lines are in the same plane

    ' Distancia entre A y la esquina
    x1 = Esq(0) - PA(0)
    y1 = Esq(1) - PA(1)
    D_A0 = Sqr((x1 ^ 2 + y1 ^ 2))

    ' Distancia entre B y la esquina
    x2 = Esq(0) - PB(0)
    y2 = Esq(1) - PB(1)
    D_B0 = Sqr((x2 ^ 2 + y2 ^ 2))

    ' Distacia entre A y B
    x3 = PA(0) - PB(0)
    y3 = PA(1) - PB(1)
    D_AB = Sqr((x3 ^ 2 + y3 ^ 2))
End If

''''' NUEVO


If Abs(DirMuro2 - DirMuro1) > PI Then
    If DirMuro2 > DirMuro1 Then
        ' Intercambiar los puntos P1 y P3
        Dim tempP0(0 To 2) As Double
        tempP0(0) = PA(0): tempP0(1) = PA(1): tempP0(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = tempP0(0): PB(1) = tempP0(1): PB(2) = tempP0(2):
        ' Recalcular la dirección del muro 1 y perpendicular al muro
        DirMuro1 = GcadUtil.AngleFromXAxis(PA, Esq)
        ' Recalcular la dirección del muro 2 y perpendicular al muro
        DirMuro2 = GcadUtil.AngleFromXAxis(PB, Esq)
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
    End If
Else
    If DirMuro2 < DirMuro1 Then
        ' Intercambiar los puntos P1 y P3
        Dim tempP(0 To 2) As Double
        tempP(0) = PA(0): tempP(1) = PA(1): tempP(2) = PA(2)
        PA(0) = PB(0): PA(1) = PB(1): PA(2) = PB(2)
        PB(0) = tempP(0): PB(1) = tempP(1): PB(2) = tempP(2):
        ' Recalcular la dirección del muro 1 y perpendicular al muro
        DirMuro1 = GcadUtil.AngleFromXAxis(PA, Esq)
        ' Recalcular la dirección del muro 2 y perpendicular al muro
        DirMuro2 = GcadUtil.AngleFromXAxis(PB, Esq)
        Slope1 = Tan(DirMuro1)
        Slope2 = Tan(DirMuro2)
    End If
End If

'''' FIN DE NUEVO


DirBulon1 = DirMuro1 - (PI / 2)
DirBulon2 = DirMuro2 + (PI / 2)

P2(0) = PA(0) + l145 * Cos(DirBulon1): P2(1) = PA(1) + l145 * Sin(DirBulon1): P2(2) = PA(2)
P1(0) = PB(0) + l145 * Cos(DirBulon2): P1(1) = PB(1) + l145 * Sin(DirBulon2): P1(2) = PB(2)

PBloque(0) = PA(0): PBloque(1) = PA(1): PBloque(2) = PA(2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

DirMuro1Inv = DirMuro1 - PI
DirMuro2Inv = DirMuro2 - PI

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' Aquí viene la toma de decisiones de si coger un puntal o coger otro::::
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



If DirMuro1 = DirMuro2 Then
    Set b = ThisDrawing.Blocks.Add(PBloque, k)

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
    DirPuntal = GcadUtil.AngleFromXAxis(P1, P2)
    DirPuntal2 = GcadUtil.AngleFromXAxis(P2, P1)
    If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
        rutaplaca1 = rutaplaca1
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
        blockRef.Layer = "Granshor"
    Else
        rutaplaca1 = rutaplaca1
        If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
            rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
        ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
            rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
        End If
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
        blockRef.Layer = "Granshor"
    End If
    If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
        rutaplaca2 = rutaplaca2
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
        blockRef.Layer = "Granshor"
    Else
        rutaplaca2 = rutaplaca2
        If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
            rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
        ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
            rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
        End If
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
        blockRef.Layer = "Granshor"
    End If
Else
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



    If dato3 = "PL" Then
        If nfusible = 1 Then
            If n280 = 1 Then
                lalt1 = Distancia - l280
                lalt2 = Distancia + 470
            ElseIf n280 = 2 Then
                lalt1 = Distancia - l280
                lalt2 = Distancia + 190
            End If
        ElseIf nfusible = 2 Then
            If n280 = 1 Then
                lalt1 = Distancia - lfusible - l280 + 150
                lalt2 = Distancia - l280 - lfusible + l750
            ElseIf n280 = 2 Then
                lalt1 = Distancia - lfusible - l280 + 150
                lalt2 = Distancia - 560 - lfusible + l750
            End If
        End If
    ElseIf dato3 = "PS" Then
        If n280 = 1 Then
            If n560 = 1 And n750 = 1 Then
                lalt1 = Distancia - l280 + 190
                lalt2 = Distancia + l280
            ElseIf n560 = 0 And n750 = 1 Then
                lalt1 = Distancia - l280
                lalt2 = Distancia + l280
                
            ElseIf n560 = 1 And n750 = 0 Then
                lalt1 = Distancia - 90
                lalt2 = Distancia + 280
            Else
                lalt1 = Distancia - 280
                lalt2 = Distancia + 280
            End If
        End If
    End If


    '''''''''''''''''''''Cálculo de las posibles modulaciones''''''''''''''''''''''''''''''''''


    '''' MURO 1 '''''''
    muro1menor = (Distancia * D_A0) / D_AB
    muromoduladome1 = (lalt1 * D_A0) / D_AB
    muromoduladoma1 = (lalt2 * D_A0) / D_AB

    mod1c = muro1menor - muromoduladome1
    mod1M = muromoduladoma1 - muro1menor

    Pcerca2(0) = P2(0) + mod1c * Cos(DirMuro1): Pcerca2(1) = P2(1) + mod1c * Sin(DirMuro1): Pcerca2(2) = P2(2)
    Plejos2(0) = P2(0) - mod1M * Cos(DirMuro1): Plejos2(1) = P2(1) - mod1M * Sin(DirMuro1): Plejos2(2) = P2(2)

    '''''' MURO 2 ''''''''''
    muro2menor = (Distancia * D_B0) / D_AB
    muromoduladome2 = (lalt1 * D_B0) / D_AB
    muromoduladoma2 = (lalt2 * D_B0) / D_AB

    mod2c = muro2menor - muromoduladome2
    mod2M = muromoduladoma2 - muro2menor

    Pcerca1(0) = P1(0) + mod2c * Cos(DirMuro2): Pcerca1(1) = P1(1) + mod2c * Sin(DirMuro2): Pcerca1(2) = P1(1)
    Plejos1(0) = P1(0) - mod2M * Cos(DirMuro2): Plejos1(1) = P1(1) - mod2M * Sin(DirMuro2): Plejos1(2) = P1(1)
    

    If ((mod1M + mod1c) < 2000) Or ((mod2c + mod2M) < 2000) Then
        userInput = InputBox("Elija una de las siguientes opciones:" & vbCrLf & vbCrLf & vbCrLf & "1. Dibujar el puntal seleccionado de longitud " & Distancia & "." & vbCrLf & vbCrLf & vbCrLf & "2. Dibujar un puntal MENOR de " & lalt1 & "mm de longitud más cercano a la esquina" & vbCrLf & vbCrLf & vbCrLf & "3. Dibujar un puntal MAYOR de " & lalt2 & "mm de longitud más alejado de la esquina")
        
        If userInput = "1" Or userInput = "" Then
            Set b = ThisDrawing.Blocks.Add(PBloque, k)
            DirPuntal = GcadUtil.AngleFromXAxis(P1, P2)
            DirPuntal2 = GcadUtil.AngleFromXAxis(P2, P1)
            If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
                rutaplaca1 = rutaplaca1
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca1 = rutaplaca1
                If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
                ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
                End If
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
                blockRef.Layer = "Granshor"
            End If
            If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
                rutaplaca2 = rutaplaca2
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca2 = rutaplaca2
                If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
                ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
                End If
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
                blockRef.Layer = "Granshor"
            End If
        ElseIf userInput = "2" Then
            
            P1(0) = Pcerca1(0): P1(1) = Pcerca1(1): P1(2) = Pcerca1(2)
            P2(0) = Pcerca2(0): P2(1) = Pcerca2(1): P2(2) = Pcerca2(2)
            Distancia = lalt1
            PA(0) = PA(0) + mod1c * Cos(DirMuro1): PA(1) = PA(1) + mod1c * Sin(DirMuro1): PA(2) = PA(2)
            PB(0) = PB(0) + mod2c * Cos(DirMuro2): PB(1) = PB(1) + mod2c * Sin(DirMuro2): PB(2) = PB(0)
            DirPuntal = GcadUtil.AngleFromXAxis(P1, P2)
            DirPuntal2 = GcadUtil.AngleFromXAxis(P2, P1)
            PBloque(0) = PA(0): PBloque(1) = PA(1): PBloque(2) = PA(2)
            Set b = ThisDrawing.Blocks.Add(PBloque, k)
            If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
                rutaplaca1 = rutaplaca1
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca1 = rutaplaca1
                If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
                ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
                End If
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
                blockRef.Layer = "Granshor"
            End If
            If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
                rutaplaca2 = rutaplaca2
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca2 = rutaplaca2
                If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
                ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
                End If
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
                blockRef.Layer = "Granshor"
            End If

        ElseIf userInput = "3" Then
            P1(0) = Plejos1(0): P1(1) = Plejos1(1): P1(2) = Plejos1(2)
            P2(0) = Plejos2(0): P2(1) = Plejos2(1): P2(2) = Plejos2(2)
            Distancia = lalt2
            PA(0) = PA(0) - mod1M * Cos(DirMuro1): PA(1) = PA(1) - mod1M * Sin(DirMuro1): PA(2) = PA(2)
            PB(0) = PB(0) - mod2M * Cos(DirMuro2): PB(1) = PB(1) - mod2M * Sin(DirMuro2): PB(2) = PB(0)
            DirPuntal = GcadUtil.AngleFromXAxis(P1, P2)
            DirPuntal2 = GcadUtil.AngleFromXAxis(P2, P1)
            PBloque(0) = PA(0): PBloque(1) = PA(1): PBloque(2) = PA(2)
            Set b = ThisDrawing.Blocks.Add(PBloque, k)
            If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
                rutaplaca1 = rutaplaca1
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca1 = rutaplaca1
                If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
                ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
                    rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
                End If
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
                blockRef.Layer = "Granshor"
            End If
            If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
                rutaplaca2 = rutaplaca2
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
                blockRef.Layer = "Granshor"
            Else
                rutaplaca2 = rutaplaca2
                If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
                ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
                    rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
                End If
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
                blockRef.Layer = "Granshor"
            End If

        End If
    Else
        Set b = ThisDrawing.Blocks.Add(PBloque, k)
        DirPuntal = GcadUtil.AngleFromXAxis(P1, P2)
        DirPuntal2 = GcadUtil.AngleFromXAxis(P2, P1)
        If (Abs(DirMuro1 - DirPuntal2) <= (PI / 2)) Or (Abs(DirMuro1 - DirPuntal2) >= ((3 * PI) / 2)) Then
            rutaplaca1 = rutaplaca1
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1)
            blockRef.Layer = "Granshor"
        Else
            rutaplaca1 = rutaplaca1
            If rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg" Then
                rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg"
            ElseIf rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg" Then
                rutaplaca1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg"
            End If
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PA, rutaplaca1, Xs, Ys, Zs, DirMuro1Inv)
            blockRef.Layer = "Granshor"
        End If
        If (Abs(DirPuntal - DirMuro2) <= (PI / 2)) Or (Abs(DirPuntal - DirMuro2) >= ((3 * PI) / 2)) Then
            rutaplaca2 = rutaplaca2
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2)
            blockRef.Layer = "Granshor"
        Else
            rutaplaca2 = rutaplaca2
            If rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Dalzado.dwg" Then
                rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_PlacaAnclaje_Ialzado.dwg"
            ElseIf rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Dalzado.dwg" Then
                rutaplaca2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\GS_Placacompacta_Ialzado.dwg"
            End If
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(PB, rutaplaca2, Xs, Ys, Zs, DirMuro2Inv)
            blockRef.Layer = "Granshor"
        End If

    End If
End If


''''''''''' Dibujamos el puntal con los antiguos PS
''''''' antiguo PS

Set Eje1 = ThisDrawing.Blocks.Item(k).AddLine(P1, P2)
Eje1.Layer = "Nonplot"
ANG = GcadUtil.AngleFromXAxis(P1, P2)
ANG2 = ANG + (PI / 2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

Dim Longitud As String, orientacion As String, Pinicio0 As String, Pinicio1 As String, Pinicio2 As String, PreMon As String


Pinicio0 = CStr(P1(0))
Pinicio1 = CStr(P1(1))
Pinicio2 = CStr(P1(2))
Longitud = CStr(Distancia)
orientacion = CStr(ANG)

If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If

'Puntos centrales de las placas
PAP(0) = PA(0) - 50 * Cos(DirMuro1): PAP(1) = PA(1) - 50 * Sin(DirMuro1): PAP(2) = PA(2)
PBP(0) = PB(0) - 50 * Cos(DirMuro2): PBP(1) = PB(1) - 50 * Sin(DirMuro2): PBP(2) = PB(2)
x4 = PAP(0) - PBP(0)
y4 = PAP(1) - PBP(1)
D_ABP = Val(Sqr((x4 ^ 2 + y4 ^ 2)))

TxtPnt2(0) = PBP(0) + (D_ABP / 2) * Cos(ANG): TxtPnt2(1) = PBP(1) + (D_ABP / 2) * Sin(ANG): TxtPnt2(2) = PBP(2)
TxtPnt2(0) = TxtPnt2(0) + 860 * Cos(ANG2): TxtPnt2(1) = TxtPnt2(1) + 860 * Sin(ANG2): TxtPnt2(2) = TxtPnt2(2)

TxtPnt(0) = P1(0) + (Distancia / 2) * Cos(ANG): TxtPnt(1) = P1(1) + (Distancia / 2) * Sin(ANG): TxtPnt(2) = P1(2)
TxtPnt(0) = TxtPnt(0) + 410 * Cos(ANG2): TxtPnt(1) = TxtPnt(1) + 410 * Sin(ANG2): TxtPnt(2) = TxtPnt(2)

Set objGcadDimAligned = GcadModel.AddDimAligned(P1, P2, TxtPnt)
objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objGcadDimAligned.StyleName = "MODELO"
objGcadDimAligned.TextStyle = "SIMPLEX"
objGcadDimAligned.VerticalTextPosition = acOutside
objGcadDimAligned.Update
objGcadDimAligned.Layer = "Dimension"


Set objGcadDimAligned = GcadModel.AddDimAligned(PBP, PAP, TxtPnt2)
objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objGcadDimAligned.StyleName = "MODELO"
objGcadDimAligned.TextStyle = "SIMPLEX"
objGcadDimAligned.VerticalTextPosition = acOutside
objGcadDimAligned.Update
objGcadDimAligned.Layer = "Dimension"


If dato2 = "PS6" Then
    Call p6(P1, P2, carga, dato1, k)
    GoTo enchufarp6
End If


'Introducir el bulón de 120 mm en los extremos siempre, ángulo de giro, fusible fijo y chapas de 50mm:
GS_Bulon120mm = rutags & "GS_Bulon120mm_Planta.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, GS_Bulon120mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P2, GS_Bulon120mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
GS_Giro = rutags & "GS_Giro_Planta.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, GS_Giro, Xs, Ys, Zs, ANG)
blockRef.Layer = "Granshor"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P2, GS_Giro, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Granshor"



If (carga < 1350) Then
    MsgBox "Los puntales mixtos actualmente han de lanzarse con el comando PM"
ElseIf (carga >= 1350) And (carga < 1500) Then
    Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
    GS_Fusible = rutags & "GS_Fusible_Planta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"
    M20x90 = ruta2 & "4M20X90.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"

    Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
    M20x110 = ruta2 & "4M20X110.dwg"
    M20x100 = ruta2 & "4M20X100A.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x100, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"

    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"

ElseIf (carga >= 1500) And (carga < 2000) Then
    Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
    GS_Fusible = rutags & "GS_Fusible_Planta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"
    M20x90 = ruta2 & "4M20X90.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"

    Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
    PS_Placa35mm = rutaps & "PS_Placa35mm_" & dato1 & ".dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa35mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    M20x150 = ruta2 & "4M20X150.dwg"
    M20x100 = ruta2 & "4M20X100A.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x100, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
    Punto_inial(0) = Punto_inial(0) + 35 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + 35 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x150, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
ElseIf (carga >= 2000) And (carga < 2900) Then
    Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
    GS_Fusible = rutags & "GS_Fusible_Planta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"
    M20x90 = ruta2 & "4M20X90.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
    Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    M20x150 = ruta2 & "4M20X150.dwg"
    M20x100 = ruta2 & "4M20X100A.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x100, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x150, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
ElseIf carga >= 2900 Then
    M20x90 = ruta2 & "4M20X90.dwg"
    Punto_inial(0) = P2(0) + lgiro * Cos(ANG + PI): Punto_inial(1) = P2(1) + lgiro * Sin(ANG + PI): Punto_inial(2) = P2(2)
    GS_Fusible = rutags & "GS_Fusible_Planta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
    Punto_inial(0) = Punto_inial(0) + lfusible * Cos(ANG + PI): Punto_inial(1) = Punto_inial(1) + lfusible * Sin(ANG + PI): Punto_inial(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Granshor"


    PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato1 & ".dwg"
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG + PI): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG + PI): Punto_inial(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Var20x250 = ruta2 & "1VarM20X250.dwg"
    M20x100 = ruta2 & "4M20X100A.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x100, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG + PI): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG + PI): Punto_inial(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    ' aquí van las varillas
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
End If



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
        n280 = 0
        n560 = 1
        ElseIf dato2 = "Pshor_4S" Then
        n280 = 0
        n560 = 1
        End If
    Case Else
    MsgBox "Longitud no controlada " & lpuntal & "mm, fuera de rango, revisar código"
    GoTo terminar
        
End Select


M20x90_16 = ruta2 & "16M20X90.dwg"
If carga > 2900 Then
    If n280 > 0 Then
        i = 0
        Do While i < n280
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_280 = rutapl & "PL_280_" & dato1 & ".dwg"
            Punto_final(0) = Punto_inial(0) + l280 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + l280 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_280, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4L"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            i = i + 1
        Loop
    End If

    If n560 > 0 Then
        i = 0
        Do While i < n560
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_560 = ruta & "PS_560.dwg"
            Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_560, Xs, Ys, Zs, ANG)
            blockRef.Layer = capa
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            i = i + 1
        Loop
    End If

    If n1500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_1500 = ruta & dato3 & "_1500_" & dato1 & ".dwg"
            Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_1500, Xs, Ys, Zs, ANG)
            blockRef.Layer = capa
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
    End If

    If n3000 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
            PS_3000 = ruta & dato3 & "_3000_" & dato1 & ".dwg"
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = capa
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
    End If

    If n4500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
            PS_4500 = ruta & dato3 & "_4500_" & dato1 & ".dwg"
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = capa
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
    End If

    If n6000 > 0 Then
        i = 0
        Do While i < n6000
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
            PS_6000 = ruta & dato3 & "_6000_" & dato1 & ".dwg"
            Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_6000, Xs, Ys, Zs, ANG)
            blockRef.Layer = capa
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
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_750, Xs, Ys, Zs, ANG)
            blockRef.Layer = capa
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
            i = i + 1
        Loop
    End If


Else
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
        i = 0
        Do While i < n560
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_560 = ruta & "PS_560.dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
            blockRef.Layer = capa
            Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            i = i + 1
        Loop
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

End If



If carga >= 2900 Then
    '
End If
    

If carga < 2900 Then
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    PGato1(0) = Punto_inial(0): PGato1(1) = Punto_inial(1): PGato1(2) = Punto_inial(2)
    Punto_final(0) = Punto_inial(0) + l_conogato * Cos(ANG): Punto_final(1) = Punto_inial(1) + l_conogato * Sin(ANG): Punto_final(2) = Punto_inial(2)
    zPS_Gato_Cono = rutaps & "zPS_Gato_Cono_Planta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, zPS_Gato_Cono, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Pipeshor4S"
    'M20x90_16 = ruta & "16M20x90.dwg"
    'Set BlockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
    'BlockRef.Layer = "TORNILLERIA"



    Punto_inial2(0) = P2(0) - lgiro * Cos(ANG): Punto_inial2(1) = P2(1) - lgiro * Sin(ANG): Punto_inial2(2) = P2(2)
    Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
        If nfusible = 2 Then
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Nonplot"
            '
            '
            '
            Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
        End If
    Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
    zPS_Gato_Tope = rutaps & "zPS_Gato_Tope_Planta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, zPS_Gato_Tope, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Pipeshor4S"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, M20x90, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Nonplot"
    '
        '
        '
    Punto_final2(0) = Punto_inial2(0) - l_tope * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - l_tope * Sin(ANG): Punto_final(2) = Punto_inial2(2)
    
    PGato2(0) = Punto_inial2(0): PGato2(1) = Punto_inial2(1): PGato2(2) = Punto_inial2(2)

    x5 = PGato2(0) - PGato1(0)
    y5 = PGato2(1) - PGato1(1)
    D_Gato = Val(Sqr((x5 ^ 2 + y5 ^ 2)))

    TxtPnt3(0) = PGato1(0) + (D_Gato / 2) * Cos(ANG): TxtPnt3(1) = PGato1(1) + (D_Gato / 2) * Sin(ANG): TxtPnt3(2) = PGato1(2)
    TxtPnt3(0) = TxtPnt3(0) - 350 * Cos(ANG2): TxtPnt3(1) = TxtPnt3(1) - 350 * Sin(ANG2): TxtPnt3(2) = TxtPnt3(2)

    Set objGcadDimAligned = GcadModel.AddDimAligned(PGato1, PGato2, TxtPnt3)
    objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
    objGcadDimAligned.StyleName = "MODELO"
    objGcadDimAligned.TextStyle = "SIMPLEX"
    objGcadDimAligned.VerticalTextPosition = acOutside
    
    objGcadDimAligned.Layer = "Dimension"
    objGcadDimAligned.Update


    Punto_inial(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_inial(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_inial(2) = (Punto_final(2) + Punto_final2(2)) / 2


    PS_Gato = rutaps & "PS_Gato_" & dato1 & ".dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Gato, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
Else
    Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
    PGato1(0) = Punto_inial(0) + 100 * Cos(ANG + PI): PGato1(1) = Punto_inial(1) + 100 * Sin(ANG + PI): PGato1(2) = Punto_inial(2)
    'dos chapones de 50 y 4 varillas 20x250
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG + PI): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG + PI): Punto_inial(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG + PI): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG + PI): Punto_inial(2) = Punto_inial(2)
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    ' base del cajón
    basecajon = rutacajon & "cajonh_" & dato1 & ".dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, basecajon, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Pipeshor4S"
    Punto_final(0) = Punto_inial(0) + 810 * Cos(ANG + PI): Punto_final(1) = Punto_inial(1) + 810 * Sin(ANG + PI): Punto_final(2) = Punto_inial(2)
    
    'nos vamos al final del codal a meter el giro y acercarnos al cajón, haya uno o dos fusibles
    Punto_inial2(0) = P1(0) + lgiro * Cos(ANG): Punto_inial2(1) = P1(1) + lgiro * Sin(ANG): Punto_inial2(2) = P1(2)
    Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
        If nfusible = 2 Then
            Punto_final2(0) = Punto_inial2(0) + lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) + lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Granshor"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Nonplot"
            '
        '
        '
        End If
    Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
    'metemos aquí el cajón hidráulico
    brazocajon = rutacajon & "modcajon_Planta.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, brazocajon, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Pipeshor4S"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, M20x90, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Nonplot"
    '
        '
        '
    
    
    
    PGato2(0) = Punto_inial2(0): PGato2(1) = Punto_inial2(1): PGato2(2) = Punto_inial2(2)

    x5 = PGato2(0) - PGato1(0)
    y5 = PGato2(1) - PGato1(1)
    D_Gato = Val(Sqr((x5 ^ 2 + y5 ^ 2)))

    TxtPnt3(0) = PGato1(0) + (D_Gato / 2) * Cos(ANG): TxtPnt3(1) = PGato1(1) + (D_Gato / 2) * Sin(ANG): TxtPnt3(2) = PGato1(2)
    TxtPnt3(0) = TxtPnt3(0) - 350 * Cos(ANG2): TxtPnt3(1) = TxtPnt3(1) - 350 * Sin(ANG2): TxtPnt3(2) = TxtPnt3(2)

    Set objGcadDimAligned = GcadModel.AddDimAligned(PGato1, PGato2, TxtPnt3)
    objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
    objGcadDimAligned.StyleName = "MODELO"
    objGcadDimAligned.TextStyle = "SIMPLEX"
    objGcadDimAligned.VerticalTextPosition = acOutside
    objGcadDimAligned.Update
    objGcadDimAligned.Layer = "Dimension"
    objGcadDimAligned.Update
End If



enchufarp6:

PreMon = ""
Dim NamePre As GcadAttribute
Set NamePre = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "NombrePremontaje", PreMon)

Dim longitudatt As GcadAttribute
Set longitudatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Longitud", Longitud)
        
Dim orientacionatt As GcadAttribute
Set orientacionatt = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Orientacion", orientacion)
           
Dim cooordenadainicio0 As GcadAttribute
Set cooordenadainicio0 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada0", Pinicio0)
        
Dim cooordenadainicio1 As GcadAttribute
Set cooordenadainicio1 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada1", Pinicio1)
        
Dim cooordenadainicio2 As GcadAttribute
Set cooordenadainicio2 = b.AddAttribute(1, acAttributeModeInvisible, "ey", P1, "Coordenada2", Pinicio2)

Set blockRef = ThisDrawing.ModelSpace.InsertBlock(PBloque, k, Xs, Ys, Zs, 0)
blockRef.Layer = "NoContable"
blockRef.Update

 
 
Eje1.Layer = "Nonplot"
Loop

terminar:

End Sub






''''''''''''''''''''''''''''''''''''' pm pm pm pm pm pm ''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''' pm pm pm pm pm pm ''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''' pm pm pm pm pm pm ''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''' pm pm pm pm pm pm ''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''' pm pm pm pm pm pm ''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''' pm pm pm pm pm pm ''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub pm(punto1 As Variant, punto2 As Variant, placa1 As String, placa2 As String, vista As String, tubo As String, k As String)
Dim ruta1 As String
Dim ruta2 As String
Dim ruta3 As String
Dim ruta4 As String
Dim rutapl1 As String
Dim rutapl2 As String
Dim GcadDoc As Object
Dim GcadUtil As Object
Dim GcadModel As Object
Dim x As Double
Dim y As Double
Dim z As Double
Dim M20x50 As String
Dim M20x90 As String
Dim M20x110 As String
Dim M20x60 As String
Dim M20x90_16 As String
Dim M20x60_4 As String
Dim MP_Giro1 As String
Dim MP_Giro2 As String
Dim MP_Fusible As String
Dim mp_90 As String
Dim mp_180 As String
Dim mp_270 As String
Dim mp_450 As String
Dim mp_900 As String
Dim PS_280 As String
Dim PS_560 As String
Dim PS_750 As String
Dim PS_1500 As String
Dim PS_3000 As String
Dim PS_4500 As String
Dim PS_6000 As String
Dim MP_Husillo As String
Dim zMP_Base As String
Dim PS_Placa50mm As String
Dim MP_Gato As String
Dim MP_Jack As String
Dim lgiro1 As Double
Dim lgiro2 As Double
Dim lfusible As Double
Dim ljack As Double
Dim l280 As Double
Dim l560 As Double
Dim l750 As Double
Dim l1500 As Double
Dim l3000 As Double
Dim l4500 As Double
Dim l6000 As Double
Dim l900 As Double
Dim l450 As Double
Dim l270 As Double
Dim l180 As Double
Dim l90 As Double
Dim l95 As Double
Dim l315 As Double
Dim l50 As Double
Dim l_base As Double
Dim lfija As Double
Dim lpuntal As Double
Dim lgatomin As Double
Dim n6000 As Integer
Dim n4500 As Integer
Dim n3000 As Integer
Dim n1500 As Integer
Dim n750 As Integer
Dim n560 As Integer
Dim n280 As Integer
Dim n900 As Integer
Dim n450 As Integer
Dim n270 As Integer
Dim n180 As Integer
Dim n90 As Integer
Dim njack As Integer
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
Dim ANG2 As Double
Dim x5 As Double
Dim y5 As Double
Dim PGato1(0 To 2) As Double
Dim PGato2(0 To 2) As Double
Dim Distancia As Double
Dim P1(0 To 2) As Double
Dim P2(0 To 2) As Double
Dim dato1 As String
Dim dato2 As String
Dim dato3 As String
Dim dato4 As String
Dim dato5 As String
Dim tipoplaca1 As String
Dim tipoplaca2 As String
Dim objGcadDimAligned As GcadDimAligned
Dim plalz As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object, b As Object
Dim TxtPnt3(0 To 2) As Double

Set GcadDoc = GetObject(, "Gcad.Application").ActiveDocument
Set GcadModel = GcadDoc.ModelSpace
Set GcadUtil = GcadDoc.Utility

On Error GoTo terminar
repite = 1

Ncapa = "NoContable"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 40
Ncapa = "Mega"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Pipeshor4S"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5

'Valores fijos
PI = 4 * Atn(1)
lgiro1 = 90
lgiro2 = 90
lfusible = 90
l280 = 280
l560 = 560
l750 = 750
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
l50 = 50
l900 = 900
l450 = 450
l270 = 270
l180 = 180
l90 = 90
l95 = 95
l315 = 315
lgatomin = 435



' TRAS SELECCION DE PRIMER PUNTO

dato2 = tubo
dato1 = vista
tipoplaca1 = placa1
tipoplaca2 = placa2

kwordList = "Si No"
dato4 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato4 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Introducir Jack Plate?: [Si/No]")

If tipoplaca1 = "" Or tipoplaca1 = "Cuña" Then
lgiro1 = 90
rutapl1 = "MG_AnguloGiro"
ElseIf tipoplaca1 = "PlacaMP" Then
lgiro1 = 315
rutapl1 = "PL_GCODAL_"
ElseIf tipoplaca1 = "PlacaMPCompacta" Then
lgiro1 = 95
rutapl1 = "PL_GCODAL_C_"
Else
GoTo terminar
End If

If tipoplaca2 = "" Or tipoplaca2 = "Cuña" Then
lgiro2 = 90
rutapl2 = "MG_AnguloGiro"
ElseIf tipoplaca2 = "PlacaMP" Then
lgiro2 = 315
rutapl2 = "PL_GCODAL_"
ElseIf tipoplaca2 = "PlacaMPCompacta" Then
lgiro2 = 95
rutapl2 = "PL_GCODAL_C_"
Else
GoTo terminar
End If

If dato4 = "Si" Or dato4 = "" Then
njack = 1
ElseIf dato4 = "No" Then
njack = 0
Else
GoTo terminar
End If
ljack = njack * 40
lfija = (lgiro1 + lgiro2) + lfusible + (2 * l50) + lgatomin + ljack


If dato2 = "Pshor_4L" Then
capa = "Pipeshor4L"
dato5 = "PL"
n560 = 0
ElseIf dato2 = "Pshor_4S" Or dato2 = "Pshor_4" Then
dato2 = "Pshor_4S"
capa = "Pipeshor4S"
dato5 = "PS"
Else
GoTo terminar
End If

ruta1 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\" & dato2 & "\"
ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
ruta3 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\ACCESORIOS\"
ruta4 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"


If dato1 = "" Or dato1 = "Planta" Then
dato1 = "planta"
plalz = "PLA"
ElseIf dato1 = "Alzado" Then
dato1 = "alzado"
plalz = "ALZ"
Else
GoTo terminar
End If

If dato2 = "Pshor_4S" Then
    kwordList = "1500 750 560 280"
    dato3 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    dato3 = ThisDrawing.Utility.GetKeyword(vbLf & "Viga pipeshor de menor longitud en el puntal: [1500/750/560/280]")
ElseIf dato2 = "Pshor_4L" Then
    kwordList = "1500 750 280"
    dato3 = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    dato3 = ThisDrawing.Utility.GetKeyword(vbLf & "Viga pipeshor de menor longitud en el puntal: [1500/750/280]")
End If


P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)

'Set b = ThisDrawing.Blocks.Add(P1, k)

Set Eje1 = ThisDrawing.Blocks.Item(k).AddLine(P1, P2)
Eje1.Layer = "Nonplot"
ANG = GcadUtil.AngleFromXAxis(P1, P2)
ANG2 = ANG + (PI / 2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))


If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."""
        GoTo terminar
End If



lpuntal = Distancia - lfija
n6000 = Fix(lpuntal / l6000)
lpuntal = lpuntal - n6000 * l6000
n4500 = Fix(lpuntal / l4500)
lpuntal = lpuntal - n4500 * l4500
n3000 = Fix(lpuntal / l3000)
lpuntal = lpuntal - n3000 * l3000
n1500 = Fix(lpuntal / l1500)
lpuntal = lpuntal - n1500 * l1500

If dato2 = "Pshor_4S" Then
    If dato3 = "" Or dato3 = "1500" Then
        n750 = 0
        n560 = 0
        n280 = 0
    ElseIf dato3 = 750 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n280 = 0
    ElseIf dato3 = 560 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n560 = Fix(lpuntal / l560)
        lpuntal = lpuntal - n560 * l560
    ElseIf dato3 = 280 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n560 = Fix(lpuntal / l560)
        lpuntal = lpuntal - n560 * l560
        n280 = Fix(lpuntal / l280)
        lpuntal = lpuntal - n280 * l280
    Else
    GoTo terminar
    End If
ElseIf dato2 = "Pshor_4L" Then
    n560 = 0
    If dato3 = "" Or dato3 = "1500" Then
        n750 = 0
        n280 = 0
    ElseIf dato3 = 750 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n280 = 0
    ElseIf dato3 = 280 Then
        n750 = Fix(lpuntal / l750)
        lpuntal = lpuntal - n750 * l750
        n280 = Fix(lpuntal / l280)
        lpuntal = lpuntal - n280 * l280
    Else
    GoTo terminar
    End If
End If

If n280 > 1 Then
    n280 = n280 - 2
    n560 = n560 + 1
End If


n900 = Fix(lpuntal / l900)
lpuntal = lpuntal - n900 * l900
n450 = Fix(lpuntal / l450)
lpuntal = lpuntal - n450 * l450
n270 = Fix(lpuntal / l270)
lpuntal = lpuntal - n270 * l270
n180 = Fix(lpuntal / l180)
lpuntal = lpuntal - n180 * l180
n90 = Fix(lpuntal / l90)
lpuntal = lpuntal - n90 * l90




MP_Giro1 = ruta3 & rutapl1 & "PLA" & ".dwg"
MP_Giro2 = ruta3 & rutapl2 & "PLA" & ".dwg"

Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, MP_Giro1, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
Punto_inial(0) = P1(0) + lgiro1 * Cos(ANG): Punto_inial(1) = P1(1) + lgiro1 * Sin(ANG): Punto_inial(2) = P1(2)
MP_Fusible = ruta2 & "Mshor90" & plalz & "fusible.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, MP_Fusible, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
M20x50 = ruta4 & "4M20X50.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x50, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
'
        '
        '
Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)

If n90 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_90 = ruta2 & "Mshor90" & plalz & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_90, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x50, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        Punto_final(0) = Punto_inial(0) + l90 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l90 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If
M20x60 = ruta4 & "6M20X60.dwg"
If n180 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_180 = ruta2 & "Mshor180" & plalz & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_180, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"

        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        Punto_final(0) = Punto_inial(0) + l180 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l180 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n270 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_270 = ruta2 & "Mshor270" & plalz & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_270, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        Punto_final(0) = Punto_inial(0) + l270 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l270 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n450 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_450 = ruta2 & "Mshor450" & plalz & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_450, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        Punto_final(0) = Punto_inial(0) + l450 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l450 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n900 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
        mp_900 = ruta2 & "Mshor900" & plalz & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, mp_900, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Mega"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x60, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        Punto_final(0) = Punto_inial(0) + l900 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l900 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_final(2) = Punto_final(2)
PS_Placa50mm = ruta1 & "PS_Placa50mm_" & dato1 & ".dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Pipeshor4S"
M20x90 = ruta4 & "4M20X90.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
'
        '
        '
Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
M20x110 = ruta4 & "4M20X110.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x110, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
'
        '
        '
M20x90_16 = ruta4 & "16M20X90.dwg"

If n280 > 0 Then
    i = 0
    Do While i < n280
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_280 = ruta1 & "PL_280_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_280, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Pipeshor4L"
        If i > 0 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        End If
        Punto_final(0) = Punto_inial(0) + l280 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l280 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        i = i + 1
    Loop
End If

If n1500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_1500 = ruta1 & dato5 & "_1500_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_1500, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        End If
        Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n3000 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_3000 = ruta1 & dato5 & "_3000_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_3000, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        End If
        Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If


If n4500 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_4500 = ruta1 & dato5 & "_4500_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_4500, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Or n3000 > 0 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        End If
        Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n6000 > 0 Then
    i = 0
    Do While i < n6000
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_6000 = ruta1 & dato5 & "_6000_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_6000, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Or n3000 > 0 Or n4500 > 0 Or i > 0 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        End If
        Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        i = i + 1
    Loop
End If

If n750 > 0 Then
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_750 = ruta1 & dato5 & "_750_" & dato1 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_750, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Or n3000 > 0 Or n4500 > 0 Or n6000 > 0 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        End If
        Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
End If

If n560 > 0 Then
    i = 0
    Do While i < n560
        Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
        PS_560 = ruta1 & "PS_560.dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
        blockRef.Layer = capa
        If n280 > 0 Or n1500 > 0 Or n3000 > 0 Or n4500 > 0 Or n6000 > 0 Or n750 > 0 Then
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        '
        '
        '
        End If
        Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        i = i + 1
    Loop
End If

Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
blockRef.Layer = "Pipeshor4S"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x110, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
'
        '
        '
Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90, Xs, Ys, Zs, ANG)
blockRef.Layer = "Nonplot"
'
        '
        '
zMP_Base = ruta3 & "zMGBaseGato_azul.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, zMP_Base, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"
PGato2(0) = Punto_inial(0): PGato2(1) = Punto_inial(1): PGato2(2) = Punto_inial(2)
Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)

Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P2, MP_Giro2, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Mega"
Punto_final2(0) = P2(0) - lgiro2 * Cos(ANG): Punto_final2(1) = P2(1) - lgiro2 * Sin(ANG): Punto_final2(2) = P2(2)

If njack = 1 Then
    MP_Jack = ruta2 & "MshorJACKPLATE.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, MP_Jack, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "Mega"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, M20x110, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
    Punto_final2(0) = Punto_final2(0) - ljack * Cos(ANG): Punto_final2(1) = Punto_final2(1) - ljack * Sin(ANG): Punto_final2(2) = Punto_final2(2)
ElseIf njack = 0 Then
    M20x60_4 = ruta4 & "4M20X60.dwg"
    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, M20x60_4, Xs, Ys, Zs, ANG)
    blockRef.Layer = "Nonplot"
    '
        '
        '
End If
zMP_Base = ruta3 & "zMGBaseGato_naranja.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, zMP_Base, Xs, Ys, Zs, ANG + PI)
blockRef.Layer = "Mega"
PGato1(0) = Punto_final2(0): PGato1(1) = Punto_final2(1): PGato1(2) = Punto_final2(2)

Punto_inial(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_inial(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_inial(2) = (Punto_final(2) + Punto_final2(2)) / 2

MP_Husillo = ruta3 & "MGHusilloGato.dwg"
Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, MP_Husillo, Xs, Ys, Zs, ANG)
blockRef.Layer = "Mega"

Dim D_Gato As Double
x5 = PGato2(0) - PGato1(0)
y5 = PGato2(1) - PGato1(1)
D_Gato = Val(Sqr((x5 ^ 2 + y5 ^ 2)))

TxtPnt3(0) = PGato1(0) + (D_Gato / 2) * Cos(ANG): TxtPnt3(1) = PGato1(1) + (D_Gato / 2) * Sin(ANG): TxtPnt3(2) = PGato1(2)
TxtPnt3(0) = TxtPnt3(0) - 350 * Cos(ANG2): TxtPnt3(1) = TxtPnt3(1) - 350 * Sin(ANG2): TxtPnt3(2) = TxtPnt3(2)

Set objGcadDimAligned = GcadModel.AddDimAligned(PGato1, PGato2, TxtPnt3)
objGcadDimAligned.PrimaryUnitsPrecision = acDimPrecisionZero
objGcadDimAligned.StyleName = "MODELO"
objGcadDimAligned.TextStyle = "SIMPLEX"
objGcadDimAligned.VerticalTextPosition = acOutside
objGcadDimAligned.Update
objGcadDimAligned.Layer = "Dimension"


'Set blockRef = ThisDrawing.ModelSpace.InsertBlock(P1, k, Xs, Ys, Zs, 0)
'blockRef.Layer = "NoContable"

Eje1.Layer = "Nonplot"

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













Sub p6(punto1 As Variant, punto2 As Variant, carga As Double, vista As String, k As String)

Dim rutaps As String, rutap6 As String, rutator As String, rutags As String, rutapl As String
Dim GcadDoc As Object
Dim GcadUtil As Object
Dim GcadModel As Object
Dim x As Double
Dim y As Double
Dim z As Double
Dim M20x90_16 As String, M20x150_4 As String, M20x160_4 As String, Revisartornilleria As String, M30x100_16 As String, M20x90_4 As String
Dim GS_Bulon120mm As String
Dim GS_Giro As String
Dim GS_Fusible As String
Dim PS_280 As String, PS_560 As String
Dim PS_750 As String, PS_nudo As String
Dim PS_1500 As String
Dim PS_3000 As String, P6_3000 As String
Dim PS_4500 As String, P6_4500 As String
Dim PS_6000 As String
Dim P6_cono As String
Dim PS_Husillo As String
Dim PS_Placa50mm As String, PS_Placa35mm As String
Dim zPS_Gato_Cono As String
Dim zPS_Gato_Tope As String
Dim PS_Gato As String
Dim lgiro As Double
Dim lfusible As Double
Dim l280 As Double, l560 As Double
Dim l750 As Double
Dim l1500 As Double
Dim l3000 As Double
Dim l4500 As Double
Dim l6000 As Double
Dim l50 As Double, l35 As Double
Dim l_tope As Double
Dim l_conogato As Double
Dim lfija As Double, lfija2 As Double
Dim lpuntal As Double
Dim lgatomin As Double, lcajonmin As Double
Dim lmacho As Double, lcajon As Double
Dim lcono As Double
Dim n6000 As Integer
Dim n4500 As Integer, n4500p6 As Integer
Dim n3000 As Integer, n3000p6 As Integer
Dim n1500 As Integer
Dim n750 As Integer
Dim n280 As Integer, n560 As Integer
Dim nnudo As Integer
Dim nfusible As Integer
Dim n50 As Integer, n35 As Integer, ncono As Integer
Dim nmacho As Integer, ncajon As Integer, ngiro As Integer, n As Integer, lchapones As Integer
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
Dim dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String, dato6 As String
Dim capa As String
Dim condicion As Boolean
Dim kwordList As String
Dim i As Integer
Dim Ncapa As String
Dim Gcapa As Object

Set GcadDoc = GetObject(, "Gcad.Application").ActiveDocument
Set GcadModel = GcadDoc.ModelSpace
Set GcadUtil = GcadDoc.Utility

Ncapa = "NoContable"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 40
Ncapa = "Pipeshor4S"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor6"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 7
Ncapa = "Pipeshor4L"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 5
Ncapa = "Nonplot"
Set Gcapa = GcadDoc.Layers.Add(Ncapa)
Gcapa.color = 50

rutaps = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4S\"
rutapl = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_4L\"
rutags = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Gshor\"
rutap6 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Pshor_6\"
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"

'Valores fijos
PI = 4 * Atn(1)
repite = 1
lgiro = 205
lcono = 375
lfusible = 187.5
l280 = 280
l560 = 560
l750 = 750
l1500 = 1500
l3000 = 3000
l4500 = 4500
l6000 = 6000
l50 = 50
l35 = 35
l_tope = 325
l_conogato = 170
lgatomin = 620
lcajonmin = 850
lmacho = 290

On Error GoTo terminar

kwordList = "No Inicial Final Ambos"
ThisDrawing.Utility.InitializeUserInput 0, kwordList
dato1 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Nudo en algún extremo?: [No/Inicial/Final/Ambos]")

dato2 = vista

'kwordList = "A B C D"
dato4 = ""
'ThisDrawing.Utility.InitializeUserInput 0, kwordList
'dato4 = ThisDrawing.Utility.GetKeyword(vbLf & "A (giro y gato) B (machon y machón) C (giro y machón) D (giro y cajón + machón): [A/B/C/D]")

If dato1 = "" Or dato1 = "No" Then
nnudo = 0
ElseIf dato1 = "Inicial" Or dato1 = "Final" Then
nnudo = 1
ElseIf dato1 = "Ambos" Then
nnudo = 2
Else
GoTo terminar
End If


If dato2 = "" Or dato2 = "Planta" Then
dato2 = "planta"
dato6 = "pla"
ElseIf dato2 = "Alzado" Then
dato2 = "alzado"
dato6 = "alz"
Else
GoTo terminar
End If

If (carga >= 1350) And (carga < 1500) Then
    n35 = 0
    n50 = 1
ElseIf (carga >= 1500) And (carga < 2000) Then
    n35 = 1
    n50 = 1
ElseIf (carga >= 2000) And (carga < 2900) Then
    n35 = 0
    n50 = 2
ElseIf carga >= 2900 Then
    n35 = 0
    n50 = 5
End If

lchapones = (n35 * l35) + (n50 * l50)


Dim M20x110_4 As String, M20x100A As String, Var20x250 As String
M20x110_4 = rutator & "4M20X110.dwg"
M20x100A = rutator & "4M20X100A.dwg"
Var20x250 = rutator & "1VarM20X250.dwg"

M20x90_16 = rutator & "16M20X90.dwg"
M30x100_16 = rutator & "16M30X100.dwg"
M20x150_4 = rutator & "4M20X150.dwg"
M20x160_4 = rutator & "4M20X160.dwg"
M20x90_4 = rutator & "4M20X90.dwg"
Revisartornilleria = rutator & "4MetricasXrevisar.dwg"

If dato4 = "" Or dato4 = "A" Then
        
    Do While repite = 1


        'Geometría:
        'punto1 = GcadUtil.GetPoint(, "1º Punto: ")
        'punto2 = GcadUtil.GetPoint(punto1, "2º Punto: ")
        P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
        P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)
        
'        Dim k As String, b As Object, entity As Object
'        k = InputBox("Ingrese nombre: ")
'
'        If k = "" Then
'nop:
 '           MsgBox "Introduzca un nombre, por favor"
 '           k = InputBox("Ingrese nombre: ")
 '           If k = "" Then
 '               GoTo nop
 '           End If
 '       End If
 '
 '       If BloqueExiste(k) Then
 '           MsgBox "Ya existe un bloque con este nombre!"
 '           Dim Respuesta As String
 '           kwordList = "Sobreescribir Renombrar"
 '           Respuesta = ""
 '           ThisDrawing.Utility.InitializeUserInput 0, kwordList
 '           Respuesta = ThisDrawing.Utility.GetKeyword(vbLf & "¿Qué deseas hacer?: [Sobreescribir/Renombrar]")
'            '
            'If Respuesta = "Sobreescribir" Or Respuesta = "" Then
            '
            ''    For Each entity In ThisDrawing.ModelSpace
           '         If TypeOf entity Is GcadBlockReference Then
          '              If entity.effectiveName = k Then
         '                   entity.Delete
        '                End If
       '             'End If
                'Next entity
                        
                    
                
                'ThisDrawing.Blocks.Item(k).Delete
            'ElseIf Respuesta = "Renombrar" Then
                'GoTo nop
            'End If
        'End If
        
        'Set b = ThisDrawing.Blocks.Add(punto1, k)
        
        'Set Eje1 = ThisDrawing.Blocks.Item(k).AddLine(P1, P2)
        'Eje1.Layer = "Nonplot"
        ANG = GcadUtil.AngleFromXAxis(P1, P2)

        x = P2(0) - P1(0)
        y = P2(1) - P1(1)
        Xs = 1
        Ys = 1
        Zs = 1
        Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

        lfija2 = InputBox("Introduce longitud en mm de tubo PS6: 3000/4500/6000...", "Longitud PS6", 4500)
        If lfija2 Mod 1500 = 0 Then
        n = (lfija2 / 1500) + 0
            If n = 1 Then n = 2
        lfija2 = n * 1500
        ElseIf lfija2 Mod 1500 <> 0 Then
        n = (lfija2 / 1500) + 1
            If n = 1 Then n = 2
        lfija2 = n * 1500
        End If
        
        If n = 0 Then ncono = 0 Else ncono = 2
        
        If carga < 2900 Then
            lfija = (2 * lgiro) + lfusible + lchapones + lgatomin + (ncono * lcono) + nnudo * l750 + lfija2
        Else
            lfija = (2 * lgiro) + lfusible + lchapones + lcajonmin + (ncono * lcono) + nnudo * l750 + lfija2
        End If
        
        If Distancia < lfija Then
        MsgBox "Medida de puntal " & Distancia & "mm, menor que el mínimo necesario de " & lfija & "."
        GoTo terminar
        End If
        
        If n Mod 3 = 0 Then
        n3000p6 = 0
        n4500p6 = n / 3
        ElseIf (n - 2) Mod 3 = 0 Then
        n3000p6 = 1
        n4500p6 = (n - 2) / 3
        ElseIf (n - 4) Mod 3 = 0 Then
        n3000p6 = 2
        n4500p6 = (n - 4) / 3
        End If
        
        lpuntal = Distancia - lfija
        
        
        Dim lpprueba As Double
        lpprueba = lpuntal
        Dim reg560 As String
        Dim reg280 As String
        
        'Dim carga As Double
        
        'carga = InputBox("introduzca carga")
        
        If carga < 2900 Then

            lpuntal = lpuntal / 2
            
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
            n560 = Fix(lpuntal / 560)
            lpuntal = lpuntal - n560 * 560
            n560 = n560 * 2
    ' APERTURA DEL GATO ES DE 280
            Select Case lpuntal
            Case 0 To 115
            ' se puede colocar el P6 en el medio
            nfusible = 1
            n280 = 0
            reg280 = "no"
            reg560 = "no"
            Case 115 To 140
            ' se puede colocar el P6 en el medio
            nfusible = 2
            n280 = 0
            reg280 = "no"
            reg560 = "no"
            Case 140 To 280
            ' ajustar el P6 para una de 280
            nfusible = 2
            n280 = 1
            reg280 = "si"
            reg560 = "no"
            Case 280 To 385
            ' ajustar el P6 para una más de 560
            nfusible = 1
            n560 = n560 + 1
            n280 = 0
            reg280 = "no"
            reg560 = "si"
            Case 385 To 470
            ' ajustar el P6 para una más de 560
            nfusible = 2
            n560 = n560 + 1
            n280 = 0
            reg280 = "no"
            reg560 = "si"
            Case 470 To 510
            ' ajustar el P6 para una más de 560 y una de 280
            nfusible = 1
            n560 = n560 + 1
            n280 = 1
            reg280 = "si"
            reg560 = "si"
            Case 510 To 560
            ' ajustar el P6 para una más de 560 y una de 280
            nfusible = 2
            n560 = n560 + 1
            n280 = 1
            reg280 = "si"
            reg560 = "si"
            Case Else
            MsgBox "Longitud no controlada " & lpuntal & "mm, fuera de rango, revisar código"
            GoTo terminar
            End Select
                
            
        ElseIf carga >= 2900 Then
        
            ' yo lo que haría aquí es sacar el despiece completo
            lpuntal = lpuntal / 2
            
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
            n560 = Fix(lpuntal / 560)
            lpuntal = lpuntal - n560 * 560
            n560 = n560 * 2
    ' REVISAR PORQUE LA APERTURA DEL CAJON ES DE 187!!!!
    'REVISADO
            Select Case lpuntal
            Case 0 To 85
            ' se puede colocar el P6 en el medio
            nfusible = 1
            reg560 = "no"
            reg280 = "no"
            Case 85 To 180
            ' se puede colocar el P6 en el medio
            nfusible = 2
            reg560 = "no"
            reg280 = "no"
            Case 180 To 270
            ' se puede colocar el P6 en el medio
            nfusible = 3
            reg280 = "no"
            Case 270 To 360
            ' ajustar el P6 para una más de 560
            nfusible = 1
            n560 = n560 + 1
            reg280 = "no"
            reg560 = "si"
            Case 360 To 460
            ' ajustar el P6 para una más de 560
            nfusible = 2
            n560 = n560 + 1
            reg280 = "no"
            reg560 = "si"
            Case 460 To 560
            ' ajustar el P6 para una más de 560
            nfusible = 3
            n560 = n560 + 1
            reg280 = "no"
            reg560 = "si"
            Case Else
            MsgBox "Longitud no controlada " & lpuntal & "mm, fuera de rango, revisar código"
            GoTo terminar
            End Select
        
        
        End If
            
        
        
        

        'Introducir el bulón de 120 mm en los extremos siempre, ángulo de giro, fusible fijo:
        GS_Bulon120mm = rutags & "GS_Bulon120mm_" & dato2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, GS_Bulon120mm, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Granshor"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P2, GS_Bulon120mm, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Granshor"
        GS_Giro = rutags & "GS_Giro_" & dato2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P1, GS_Giro, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Granshor"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(P2, GS_Giro, Xs, Ys, Zs, ANG + PI)
        blockRef.Layer = "Granshor"

        Punto_inial(0) = P1(0) + lgiro * Cos(ANG): Punto_inial(1) = P1(1) + lgiro * Sin(ANG): Punto_inial(2) = P1(2)
        GS_Fusible = rutags & "GS_Fusible_" & dato2 & ".dwg"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, GS_Fusible, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Granshor"
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_4, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        
        If nfusible = 3 Then
                Punto_final(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_final(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_final(2) = Punto_inial(2)
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, GS_Fusible, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Granshor"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_4, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                Punto_final(0) = Punto_inial(0): Punto_final(1) = Punto_inial(1): Punto_final(2) = Punto_inial(2)
        End If
            
        
        '
        Punto_final(0) = Punto_inial(0) + lfusible * Cos(ANG): Punto_final(1) = Punto_inial(1) + lfusible * Sin(ANG): Punto_final(2) = Punto_inial(2)

        
        If n50 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato2 & ".dwg"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                n50 = n50 - 1
        End If
        
        Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x100A, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        
        If n50 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato2 & ".dwg"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                n50 = n50 - 1
        End If
        
        If n50 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_Placa50mm = rutaps & "PS_Placa50mm_" & dato2 & ".dwg"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                n50 = n50 - 1
        End If
        
        If n35 > 0 Then
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_Placa35mm = rutaps & "PS_Placa35mm_" & dato2 & ".dwg"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa35mm, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l35 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l35 * Sin(ANG): Punto_final(2) = Punto_inial(2)
        End If
        

        If (carga >= 1350) And (carga < 1500) Then
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x110_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
        ElseIf (carga >= 1500) And (carga < 2000) Then
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x150_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
        ElseIf (carga >= 2000) And (carga < 2900) Then
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x150_4, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
        ElseIf carga >= 2900 Then
            ' varillas no?
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
        End If


        If n280 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_280 = rutapl & "PL_280_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_280, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4L"
            Punto_final(0) = Punto_inial(0) + l280 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l280 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If

        If n560 > 1 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_560 = rutaps & "PS_560.dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If
        
        If n750 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_750 = rutaps & "PS_750_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_750, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If

 
        If n1500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_1500 = rutaps & "PS_1500_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_1500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
        End If   '
            
        If n3000 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_3000 = rutaps & "PS_3000_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If



        If n4500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_4500 = rutaps & "PS_4500_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If
        
        If (n6000 / 2) > 0 Then
            i = 0
            Do While i < (n6000 / 2)
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_6000 = rutaps & "PS_6000_" & dato2 & ".dwg"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_6000, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                '
                '
                '
                i = i + 1
            Loop
        End If
        
        If dato1 = "Ambos" Or dato1 = "Inicial" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            PS_nudo = rutaps & "PS_nudo_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_nudo, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Pipeshor4S"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If
        
        If n > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            P6_cono = rutap6 & "P6_cono.dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, P6_cono, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor6"
            Punto_final(0) = Punto_inial(0) + lcono * Cos(ANG): Punto_final(1) = Punto_inial(1) + lcono * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M30x100_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If

        If n3000p6 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            P6_3000 = rutap6 & "P6_3000_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, P6_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor6"
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M30x100_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If

        If n4500p6 > 0 Then
            i = 0
            Do While i < n4500p6
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            P6_4500 = rutap6 & "P6_4500_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, P6_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor6"
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M30x100_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
            i = i + 1
            Loop
        End If

        If n3000p6 = 2 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            P6_3000 = rutap6 & "P6_3000_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, P6_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor6"
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M30x100_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If

        If n > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            Punto_final(0) = Punto_inial(0) + lcono * Cos(ANG): Punto_final(1) = Punto_inial(1) + lcono * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, P6_cono, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Pipeshor6"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If

        If dato1 = "Ambos" Or dato1 = "Final" Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_nudo = rutaps & "PS_nudo_" & dato2 & ".dwg"
            'Set BlockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, PS_nudo, Xs, Ys, Zs, ANG)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_nudo, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If

        If (n6000 / 2) > 0 Then
            i = 0
            Do While i < (n6000 / 2)
                Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
                PS_6000 = rutaps & "PS_6000_" & dato2 & ".dwg"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_6000, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Pipeshor4S"
                Punto_final(0) = Punto_inial(0) + l6000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l6000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
                blockRef.Layer = "Nonplot"
                '
                '
                '
                i = i + 1
            Loop
        End If
        
        If n4500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_4500 = rutaps & "PS_4500_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_4500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l4500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l4500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If

        If n3000 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_3000 = rutaps & "PS_3000_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_3000, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l3000 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l3000 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If
        
        If n1500 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_1500 = rutaps & "PS_1500_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_1500, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l1500 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l1500 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
        End If   '

        If n750 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_750 = rutaps & "PS_750_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_750, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l750 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l750 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If
        
        If n560 > 0 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_560 = rutaps & "PS_560.dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If
        
        If n560 > 2 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            PS_560 = rutaps & "PS_560.dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_560, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l560 * Cos(ANG): Punto_final(1) = Punto_inial(1) + l560 * Sin(ANG): Punto_final(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, M20x90_16, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            '
            '
            '
        End If
        
        If carga > 2900 Then
            blockRef.Delete
        End If
        ' Aquí vamos a meter el gato si la carga es < 2900 y el cajón pipeshor si es mayor a 2900
        
        If carga < 2900 Then
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            zPS_Gato_Cono = rutaps & "zPS_Gato_Cono_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, zPS_Gato_Cono, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + l_conogato * Cos(ANG): Punto_final(1) = Punto_inial(1) + l_conogato * Sin(ANG): Punto_final(2) = Punto_inial(2)
    
            Punto_inial2(0) = P2(0) - lgiro * Cos(ANG): Punto_inial2(1) = P2(1) - lgiro * Sin(ANG): Punto_inial2(2) = P2(2)
            Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
            If nfusible = 2 Then
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Granshor"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Nonplot"
                '
                '
                '
                Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
            End If
            If nfusible = 3 Then
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Granshor"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Nonplot"
                '
                '
                '
                Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, GS_Fusible, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Granshor"
                Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG + PI)
                blockRef.Layer = "Nonplot"
                '
                '
                '
                Punto_final2(0) = Punto_inial2(0) - lfusible * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - lfusible * Sin(ANG): Punto_final(2) = Punto_inial2(2)
            
            
            End If
            
            
            
            Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
            zPS_Gato_Tope = rutaps & "zPS_Gato_Tope_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, zPS_Gato_Tope, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Pipeshor4S"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Nonplot"
            '
            '
            '
            Punto_final2(0) = Punto_inial2(0) - l_tope * Cos(ANG): Punto_final2(1) = Punto_inial2(1) - l_tope * Sin(ANG): Punto_final(2) = Punto_inial2(2)
    
    
            Punto_inial(0) = (Punto_final(0) + Punto_final2(0)) / 2: Punto_inial(1) = (Punto_final(1) + Punto_final2(1)) / 2: Punto_inial(2) = (Punto_final(2) + Punto_final2(2)) / 2
    
            PS_Gato = rutaps & "PS_Gato_" & dato2 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Gato, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
        Else
            Dim rutacajon As String, basecajon As String, brazocajon As String
            rutacajon = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Cajon hidraulico\"

        ' aquí toca meter el cajón más todos los complementos que harían falta
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            'PGato1(0) = Punto_inial(0) + 100 * Cos(ANG + PI): PGato1(1) = Punto_inial(1) + 100 * Sin(ANG + PI): PGato1(2) = Punto_inial(2)
            'dos chapones de 50 y 4 varillas 20x250
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final, Var20x250, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, PS_Placa50mm, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_inial(0) = Punto_inial(0) + l50 * Cos(ANG): Punto_inial(1) = Punto_inial(1) + l50 * Sin(ANG): Punto_inial(2) = Punto_inial(2)
            
            ' base del cajón
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, M20x90_16, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Nonplot"
            basecajon = rutacajon & "cajonh_" & dato6 & ".dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial, basecajon, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Pipeshor4S"
            Punto_final(0) = Punto_inial(0) + 810 * Cos(ANG): Punto_final(1) = Punto_inial(1) + 810 * Sin(ANG): Punto_final(2) = Punto_inial(2)
                    
            Punto_inial(0) = Punto_final(0): Punto_inial(1) = Punto_final(1): Punto_inial(2) = Punto_final(2)
            
            
            'nos vamos al final del codal a meter el giro y acercarnos al cajón, haya uno o dos fusibles
            Punto_inial2(0) = P2(0) + lgiro * Cos(ANG + PI): Punto_inial2(1) = P2(1) + lgiro * Sin(ANG + PI): Punto_inial2(2) = P2(2)
            Punto_final2(0) = Punto_inial2(0): Punto_final2(1) = Punto_inial2(1): Punto_final2(2) = Punto_inial2(2)
                If nfusible >= 2 Then
                    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG + PI)
                    blockRef.Layer = "Nonplot"
                    Punto_final2(0) = Punto_inial2(0) + lfusible * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + lfusible * Sin(ANG + PI): Punto_final(2) = Punto_inial2(2)
                    Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, GS_Fusible, Xs, Ys, Zs, ANG)
                    blockRef.Layer = "Granshor"
                End If
                'If nfusible = 3 Then
                    'Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, M20x90_4, Xs, Ys, Zs, ANG + PI)
                    'blockRef.Layer = "Nonplot"
                    'Punto_final2(0) = Punto_inial2(0) + lfusible * Cos(ANG + PI): Punto_final2(1) = Punto_inial2(1) + lfusible * Sin(ANG + PI): Punto_final(2) = Punto_inial2(2)
                    'Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, GS_Fusible, Xs, Ys, Zs, ANG)
                    'blockRef.Layer = "Granshor"
                'End If
            Punto_inial2(0) = Punto_final2(0): Punto_inial2(1) = Punto_final2(1): Punto_inial2(2) = Punto_final2(2)
            'metemos aquí el cajón hidráulico
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_final2, M20x90_4, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Nonplot"
            brazocajon = rutacajon & "modcajon_Planta.dwg"
            Set blockRef = ThisDrawing.Blocks.Item(k).InsertBlock(Punto_inial2, brazocajon, Xs, Ys, Zs, ANG + PI)
            blockRef.Layer = "Pipeshor4S"
            
            
        
        End If
    
        'Set blockRef = ThisDrawing.ModelSpace.InsertBlock(punto1, k, Xs, Ys, Zs, 0)
        'blockRef.Layer = "NoContable"
    
    Eje1.Layer = "Nonplot"
    Loop


ElseIf dato4 = "B" Then






ElseIf dato4 = "C" Then





ElseIf dato4 = "D" Then






Else
GoTo terminar
End If










terminar:
End Sub









