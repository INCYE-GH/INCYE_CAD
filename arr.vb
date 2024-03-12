Sub arr()

Dim AcadDoc As Object, AcadUtil As Object, AcadModel As Object, Eje1 As Object, blockRef As Object
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim vista As String
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double
Dim Gcapa As Object
Dim Ncapa As String
Dim punto1 As Variant, punto2 As Variant
Dim PI As Variant, Distancia As Double
Dim PuntoM(0 To 2) As Double, husillo_c As String, cuerpo_c As String


Set AcadDoc = GetObject(, "Autocad.Application").ActiveDocument
Set AcadModel = AcadDoc.ModelSpace
Set AcadUtil = AcadDoc.Utility

kwordList = "RapidTie Tubo"
vista = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vista = ThisDrawing.Utility.GetKeyword(vbLf & "¿Con qué vas a arriostrar?: [RapidTie/Tubo]")

If vista = "RapidTie" Or vista = "" Then
    Call rtie
    GoTo Terminar
ElseIf vista = "Tubo" Then
    Call tubo
    'MsgBox "Aún no disponible... Estamos en ello ;)"
    GoTo Terminar
End If

Terminar:
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------RAPID TIE-------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Sub rtie()

Dim AcadDoc As Object, AcadUtil As Object, AcadModel As Object, Eje1 As Object, blockRef As Object
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim rutatensor As String
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double
Dim Gcapa As Object
Dim Ncapa As String, punto1 As Variant, punto2 As Variant
Dim PI As Variant, Distancia As Double
Dim PuntoM(0 To 2) As Double, husillo_l As String, cuerpo_l As String
Dim ext1 As String, ext2 As String, vista As String, tor_ext1 As String, tor_ext2 As String, nrt As String
Dim margen1 As Double, margen2 As Double
Dim ExtRT1(0 To 2) As Double, ExtRT2(0 To 2) As Double, insRT(0 To 2) As Double
Dim torple1(0 To 2) As Double, torple2(0 To 2) As Double, tormar1up(0 To 2) As Double, tormar1down(0 To 2) As Double, tormar2up(0 To 2) As Double, tormar2down(0 To 2) As Double
Dim rtie As String, plalz As String, rutator As String, tor1 As String, tor2 As String


Set AcadDoc = GetObject(, "Autocad.Application").ActiveDocument
Set AcadModel = AcadDoc.ModelSpace
Set AcadUtil = AcadDoc.Utility

Ncapa = "Mega"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Lolashor"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Slims"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30

rutartie = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\RTie\"
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"

On Error GoTo Terminar

' preguntas -------------------------------------------------------

kwordList = "Pletina AdaptMarsella Cojinete Plato80 Plato110 Plato160 Canal-Arandela Libre"
ext1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
ext1 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Terminación en el primer extremo?: [Pletina/AdaptMarsella/Cojinete/Plato80/Plato110/Plato160/Canal-Arandela/Libre]")

kwordList = "Pletina AdaptMarsella Cojinete Plato80 Plato110 Plato160 Canal-Arandela Libre"
ext2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
ext2 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Terminación en el segundo extremo?: [Pletina/AdaptMarsella/Cojinete/Plato80/Plato110/Plato160/Canal-Arandela/Libre]")

kwordList = "Planta Alzado"
vista = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
vista = ThisDrawing.Utility.GetKeyword(vbLf & "¿Vista de la Rapid Tie?: [Planta/Alzado]")

If vista = "Planta" Or vista = "" Then
    plalz = "pl"
ElseIf vista = "Alzado" Then
    plalz = "alz"
End If


        tor1 = rutator & "1-M20X40.dwg"
        tor2 = rutator & "1-M20X40.dwg"



If ext1 = "Pletina" And ext2 = "Pletina" Then
    
    kwordList = "Una Doble"
    nrt = ""
    ThisDrawing.Utility.InitializeUserInput 0, kwordList
    nrt = ThisDrawing.Utility.GetKeyword(vbLf & "¿Cuántas Rapid Tie?: [Una/Doble]")
   
End If

Dim repite As Integer
repite = 1

'COMIENZA Bucle
Do While repite = 1


punto1 = AcadUtil.GetPoint(, "1º Punto: ")
punto2 = AcadUtil.GetPoint(punto1, "2º Punto: ")

Dim lrt As Double

PI = 4 * Atn(1)

P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)

Set Eje1 = AcadModel.AddLine(P1, P2)
ANG = AcadUtil.AngleFromXAxis(P1, P2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

Dim Dfinal As Double

Dfinal = Distancia

Punto_inicial(0) = P1(0): Punto_inicial(1) = P1(1): Punto_inicial(2) = P1(2)
Punto_final(0) = Punto_inicial(0) + Distancia * Cos(ANG): Punto_final(1) = Punto_inicial(1) + Distancia * Sin(ANG): Punto_final(2) = Punto_inicial(2)

PuntoM(0) = Punto_inicial(0) + (Distancia / 2) * Cos(ANG): PuntoM(1) = Punto_inicial(1) + (Distancia / 2) * Sin(ANG): PuntoM(2) = Punto_inicial(2)



If ext1 = "Pletina" Then
    Dfinal = Dfinal - 330
    ExtRT1(0) = Punto_inicial(0) + 330 * Cos(ANG): ExtRT1(1) = Punto_inicial(1) + 330 * Sin(ANG): ExtRT1(2) = Punto_inicial(2)
ElseIf ext1 = "AdaptMarsella" Then
    Dfinal = Dfinal - 250
    ExtRT1(0) = Punto_inicial(0) + 250 * Cos(ANG): ExtRT1(1) = Punto_inicial(1) + 250 * Sin(ANG): ExtRT1(2) = Punto_inicial(2)
ElseIf ext1 = "Cojinete" Then
    Dfinal = Dfinal + 80
    ExtRT1(0) = Punto_inicial(0) - 80 * Cos(ANG): ExtRT1(1) = Punto_inicial(1) - 80 * Sin(ANG): ExtRT1(2) = Punto_inicial(2)
ElseIf ext1 = "Plato80" Then
    Dfinal = Dfinal + 110
    ExtRT1(0) = Punto_inicial(0) - 110 * Cos(ANG): ExtRT1(1) = Punto_inicial(1) - 110 * Sin(ANG): ExtRT1(2) = Punto_inicial(2)
ElseIf ext1 = "Plato110" Then
    Dfinal = Dfinal + 110
    ExtRT1(0) = Punto_inicial(0) - 110 * Cos(ANG): ExtRT1(1) = Punto_inicial(1) - 110 * Sin(ANG): ExtRT1(2) = Punto_inicial(2)
ElseIf ext1 = "Plato160" Then
    Dfinal = Dfinal + 92
    ExtRT1(0) = Punto_inicial(0) - 92 * Cos(ANG): ExtRT1(1) = Punto_inicial(1) - 92 * Sin(ANG): ExtRT1(2) = Punto_inicial(2)
ElseIf ext1 = "Canal-Arandela" Then
    Dfinal = Dfinal + 80
    ExtRT1(0) = Punto_inicial(0) - 80 * Cos(ANG): ExtRT1(1) = Punto_inicial(1) - 80 * Sin(ANG): ExtRT1(2) = Punto_inicial(2)
ElseIf ext1 = "Libre" Then
    ExtRT1(0) = Punto_inicial(0) + 20 * Cos(ANG): ExtRT1(1) = Punto_inicial(1) + 20 * Sin(ANG): ExtRT1(2) = Punto_inicial(2)
End If

If ext2 = "Pletina" Then
    Dfinal = Dfinal - 330
    ExtRT2(0) = Punto_final(0) - 330 * Cos(ANG): ExtRT2(1) = Punto_final(1) - 330 * Sin(ANG): ExtRT2(2) = Punto_final(2)
ElseIf ext2 = "AdaptMarsella" Then
    Dfinal = Dfinal - 250
    ExtRT2(0) = Punto_final(0) - 250 * Cos(ANG): ExtRT2(1) = Punto_final(1) - 250 * Sin(ANG): ExtRT2(2) = Punto_final(2)
ElseIf ext2 = "Cojinete" Then
    Dfinal = Dfinal + 80
    ExtRT2(0) = Punto_final(0) + 80 * Cos(ANG): ExtRT2(1) = Punto_final(1) + 80 * Sin(ANG): ExtRT2(2) = Punto_final(2)
ElseIf ext2 = "Plato80" Then
    Dfinal = Dfinal + 110
    ExtRT2(0) = Punto_final(0) + 110 * Cos(ANG): ExtRT2(1) = Punto_final(1) + 110 * Sin(ANG): ExtRT2(2) = Punto_final(2)
ElseIf ext2 = "Plato110" Then
    Dfinal = Dfinal + 110
    ExtRT2(0) = Punto_final(0) + 110 * Cos(ANG): ExtRT2(1) = Punto_final(1) + 110 * Sin(ANG): ExtRT2(2) = Punto_final(2)
ElseIf ext2 = "Plato160" Then
    Dfinal = Dfinal + 92
    ExtRT2(0) = Punto_final(0) + 92 * Cos(ANG): ExtRT2(1) = Punto_final(1) + 92 * Sin(ANG): ExtRT2(2) = Punto_final(2)
ElseIf ext2 = "Canal-Arandela" Then
    Dfinal = Dfinal + 80
    ExtRT2(0) = Punto_final(0) + 80 * Cos(ANG): ExtRT2(1) = Punto_final(1) + 80 * Sin(ANG): ExtRT2(2) = Punto_final(2)
ElseIf ext2 = "Libre" Then
    ExtRT2(0) = Punto_final(0) - 20 * Cos(ANG): ExtRT2(1) = Punto_final(1) - 20 * Sin(ANG): ExtRT2(2) = Punto_final(2)
End If





lrt = Dfinal

If lrt < 500 And lrt > 150 Then
    lrt = 500
ElseIf lrt >= 500 And lrt <= 690 Then
    lrt = 750
ElseIf lrt > 690 And lrt <= 940 Then
    lrt = 1000
ElseIf lrt > 940 And lrt <= 1190 Then
    lrt = 1250
ElseIf lrt > 1190 And lrt <= 1450 Then
    lrt = 1500
ElseIf lrt > 1450 And lrt <= 1950 Then
    lrt = 2000
ElseIf lrt > 1950 And lrt <= 2450 Then
    lrt = 2500
ElseIf lrt > 2450 And lrt <= 2950 Then
    lrt = 3000
ElseIf lrt > 2950 And lrt <= 3450 Then
    lrt = 3500
ElseIf lrt > 3450 And lrt <= 3950 Then
    lrt = 4000
ElseIf lrt > 3950 And lrt <= 4450 Then
    lrt = 4500
ElseIf lrt > 4450 And lrt <= 4950 Then
    lrt = 5000
ElseIf lrt > 4950 And lrt <= 5700 Then
    lrt = 5750
Else
    MsgBox "No existe Rapid Tie de estas dimensiones"
End If


insRT(0) = ExtRT1(0) + (Dfinal / 2) * Cos(ANG): insRT(1) = ExtRT1(1) + (Dfinal / 2) * Sin(ANG): insRT(2) = ExtRT1(2)





' insertamos en el medio la rapid tie que sea necesaria
rtie = rutartie & "RTie" & lrt & ".dwg"
Set blockRef = AcadModel.InsertBlock(insRT, rtie, Xs, Ys, Zs, ANG)
blockRef.Layer = "AM"
If nrt = "Doble" Then
    Set blockRef = AcadModel.InsertBlock(insRT, rtie, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
End If


    Dim pletina As String
    pletina = rutartie & "Pletina_" & plalz & ".dwg"
    Dim pletinad As String
    pletinad = rutartie & "Pletinadcha_" & plalz & ".dwg"
    Dim TuercaH As String
    TuercaH = rutartie & "TuercaHexagonal.dwg"
    Dim TuercaA As String
    TuercaA = rutartie & "TuercaAlas.dwg"
    Dim AMarsella As String
    AMarsella = rutartie & "AdaptMarsella_" & plalz & ".dwg"
    Dim cojinete_D As String
    cojinete_D = rutartie & "Cojinete_d_" & plalz & ".dwg"
    Dim cojinete_I As String
    cojinete_I = rutartie & "Cojinete_i_" & plalz & ".dwg"
    Dim plato80i As String
    plato80i = rutartie & "PlatoArandela80izq.dwg"
    Dim plato80d As String
    plato80d = rutartie & "PlatoArandela80dcha.dwg"
    Dim plato110d As String
    plato110d = rutartie & "PlatoArandela110dcha.dwg"
    Dim plato110i As String
    plato110i = rutartie & "PlatoArandela110izq.dwg"
    Dim TuercaArt As String
    TuercaArt = rutartie & "TuercaAlas_rapidtie_20.dwg"
    Dim plato160 As String
    plato160 = rutartie & "PlatoArandela160_" & plalz & ".dwg"
    Dim canalar As String
    canalar = rutartie & "CanalArandela_" & plalz & ".dwg"


' Insertar en el extremo 1 lo que haya sido seleccionado
If ext1 = "Pletina" Then
    If plalz = "pl" Then
        torple1(0) = Punto_inicial(0) - 25.5 * Cos(ANG + (PI / 2)): torple1(1) = Punto_inicial(1) - 25.5 * Sin(ANG + (PI / 2)): torple1(2) = Punto_inicial(2)
        Set blockRef = AcadModel.InsertBlock(torple1, tor1, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    Else
        Set blockRef = AcadModel.InsertBlock(Punto_inicial, tor1, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
    End If
    
    If nrt = "Doble" Then
        If plalz = "pl" Then
            torple1(0) = Punto_inicial(0) - 25.5 * Cos(ANG + (PI / 2)): torple1(1) = Punto_inicial(1) - 25.5 * Sin(ANG + (PI / 2)): torple1(2) = Punto_inicial(2)
            Set blockRef = AcadModel.InsertBlock(torple1, tor1, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        Else
            Set blockRef = AcadModel.InsertBlock(Punto_inicial, tor1, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
        End If
    End If
    
        Set blockRef = AcadModel.InsertBlock(Punto_inicial, pletina, Xs, Ys, Zs, ANG)
        blockRef.Layer = "AM"
    
    
    If nrt = "Doble" Then
        Set blockRef = AcadModel.InsertBlock(Punto_inicial, pletina, Xs, Ys, Zs, ANG)
        blockRef.Layer = "AM"
    End If
    
    Punto_inicial(0) = Punto_inicial(0) + 375 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + 375 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, TuercaH, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"

    If nrt = "Doble" Then
        Set blockRef = AcadModel.InsertBlock(Punto_inicial, TuercaH, Xs, Ys, Zs, ANG)
        blockRef.Layer = "AM"
    End If
    

    
ElseIf ext1 = "AdaptMarsella" Then

    If plalz = "pl" Then
            tormar1up(0) = Punto_inicial(0) + 140 * Cos(ANG + (PI / 2)): tormar1up(1) = Punto_inicial(1) + 140 * Sin(ANG + (PI / 2)): tormar1up(2) = Punto_inicial(2)
            Set blockRef = AcadModel.InsertBlock(tormar1up, tor1, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
            
            tormar1down(0) = Punto_inicial(0) - 140 * Cos(ANG + (PI / 2)): tormar1down(1) = Punto_inicial(1) - 140 * Sin(ANG + (PI / 2)): tormar1down(2) = Punto_inicial(2)
            Set blockRef = AcadModel.InsertBlock(tormar1down, tor1, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
    Else
            Set blockRef = AcadModel.InsertBlock(Punto_inicial, tor1, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            Set blockRef = AcadModel.InsertBlock(Punto_inicial, tor1, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
    End If
    
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, AMarsella, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_inicial(0) = Punto_inicial(0) + 255 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) + 255 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, TuercaH, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"

ElseIf ext1 = "Cojinete" Then

    Set blockRef = AcadModel.InsertBlock(Punto_inicial, cojinete_I, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_inicial(0) = Punto_inicial(0) - 25 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) - 25 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, TuercaH, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"

ElseIf ext1 = "Plato80" Then

    Set blockRef = AcadModel.InsertBlock(Punto_inicial, plato80i, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_inicial(0) = Punto_inicial(0) - 20 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) - 20 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, TuercaA, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "AM"

ElseIf ext1 = "Plato110" Then

    Set blockRef = AcadModel.InsertBlock(Punto_inicial, plato110i, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_inicial(0) = Punto_inicial(0) - 20 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) - 20 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, TuercaA, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "AM"

ElseIf ext1 = "Plato160" Then

    Set blockRef = AcadModel.InsertBlock(Punto_inicial, plato160, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_inicial(0) = Punto_inicial(0) - 48 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) - 48 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, TuercaArt, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"

ElseIf ext1 = "Canal-Arandela" Then

    Set blockRef = AcadModel.InsertBlock(Punto_inicial, canalar, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_inicial(0) = Punto_inicial(0) - 25 * Cos(ANG): Punto_inicial(1) = Punto_inicial(1) - 25 * Sin(ANG): Punto_inicial(2) = Punto_inicial(2)
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, TuercaH, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"

ElseIf ext1 = "Libre" Then
    
End If






If ext2 = "Pletina" Then
    If plalz = "pl" Then
        torple2(0) = Punto_final(0) - 25.5 * Cos(ANG + (PI / 2)): torple2(1) = Punto_final(1) - 25.5 * Sin(ANG + (PI / 2)): torple2(2) = Punto_final(2)
        Set blockRef = AcadModel.InsertBlock(torple2, tor2, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    Else
        Set blockRef = AcadModel.InsertBlock(Punto_final, tor2, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    End If

    If nrt = "Doble" Then
        If plalz = "pl" Then
            Set blockRef = AcadModel.InsertBlock(torple2, tor2, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        Else
            Set blockRef = AcadModel.InsertBlock(Punto_final, tor2, Xs, Ys, Zs, ANG)
            blockRef.Layer = "Nonplot"
            blockRef.Update
            blockRef.Explode
            blockRef.Delete
        End If
    End If

    Punto_final(0) = Punto_final(0) - 470 * Cos(ANG): Punto_final(1) = Punto_final(1) - 470 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = AcadModel.InsertBlock(Punto_final, pletinad, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    
    If nrt = "Doble" Then
        Set blockRef = AcadModel.InsertBlock(Punto_final, pletinad, Xs, Ys, Zs, ANG)
        blockRef.Layer = "AM"
    End If
    
    
    Punto_final(0) = Punto_final(0) + 95 * Cos(ANG): Punto_final(1) = Punto_final(1) + 95 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = AcadModel.InsertBlock(Punto_final, TuercaH, Xs, Ys, Zs, (ANG + PI))
    blockRef.Layer = "AM"
    
    If nrt = "Doble" Then
        Set blockRef = AcadModel.InsertBlock(Punto_final, TuercaH, Xs, Ys, Zs, (ANG + PI))
        blockRef.Layer = "AM"
    End If
    

ElseIf ext2 = "AdaptMarsella" Then

    If plalz = "pl" Then
        tormar2up(0) = Punto_final(0) + 140 * Cos(ANG + (PI / 2)): tormar2up(1) = Punto_final(1) + 140 * Sin(ANG + (PI / 2)): tormar2up(2) = Punto_final(2)
        Set blockRef = AcadModel.InsertBlock(tormar2up, tor1, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
        
        tormar2down(0) = Punto_final(0) - 140 * Cos(ANG + (PI / 2)): tormar2down(1) = Punto_final(1) - 140 * Sin(ANG + (PI / 2)): tormar2down(2) = Punto_final(2)
        Set blockRef = AcadModel.InsertBlock(tormar2down, tor1, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        blockRef.Update
        blockRef.Explode
        blockRef.Delete
    Else
        Set blockRef = AcadModel.InsertBlock(Punto_final, tor1, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
        Set blockRef = AcadModel.InsertBlock(Punto_final, tor1, Xs, Ys, Zs, ANG)
        blockRef.Layer = "Nonplot"
    End If

    Set blockRef = AcadModel.InsertBlock(Punto_final, AMarsella, Xs, Ys, Zs, (ANG + PI))
    blockRef.Layer = "AM"
    Punto_final(0) = Punto_final(0) - 255 * Cos(ANG): Punto_final(1) = Punto_final(1) - 255 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = AcadModel.InsertBlock(Punto_final, TuercaH, Xs, Ys, Zs, (ANG + PI))
    blockRef.Layer = "AM"

ElseIf ext2 = "Cojinete" Then

    Set blockRef = AcadModel.InsertBlock(Punto_final, cojinete_D, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_final(0) = Punto_final(0) + 25 * Cos(ANG): Punto_final(1) = Punto_final(1) + 25 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = AcadModel.InsertBlock(Punto_final, TuercaH, Xs, Ys, Zs, (ANG + PI))
    blockRef.Layer = "AM"

ElseIf ext2 = "Plato80" Then

    Set blockRef = AcadModel.InsertBlock(Punto_final, plato80d, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_final(0) = Punto_final(0) + 20 * Cos(ANG): Punto_final(1) = Punto_final(1) + 20 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = AcadModel.InsertBlock(Punto_final, TuercaA, Xs, Ys, Zs, (ANG))
    blockRef.Layer = "AM"

ElseIf ext2 = "Plato110" Then

    Set blockRef = AcadModel.InsertBlock(Punto_final, plato110d, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
    Punto_final(0) = Punto_final(0) + 20 * Cos(ANG): Punto_final(1) = Punto_final(1) + 20 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = AcadModel.InsertBlock(Punto_final, TuercaA, Xs, Ys, Zs, (ANG))
    blockRef.Layer = "AM"

ElseIf ext2 = "Plato160" Then

    Set blockRef = AcadModel.InsertBlock(Punto_final, plato160, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "AM"
    Punto_final(0) = Punto_final(0) + 48 * Cos(ANG): Punto_final(1) = Punto_final(1) + 48 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = AcadModel.InsertBlock(Punto_final, TuercaArt, Xs, Ys, Zs, (ANG + PI))
    blockRef.Layer = "AM"

ElseIf ext2 = "Canal-Arandela" Then

    Set blockRef = AcadModel.InsertBlock(Punto_final, canalar, Xs, Ys, Zs, ANG + PI)
    blockRef.Layer = "AM"
    Punto_final(0) = Punto_final(0) + 25 * Cos(ANG): Punto_final(1) = Punto_final(1) + 25 * Sin(ANG): Punto_final(2) = Punto_final(2)
    Set blockRef = AcadModel.InsertBlock(Punto_final, TuercaH, Xs, Ys, Zs, (ANG + PI))
    blockRef.Layer = "AM"

ElseIf ext2 = "Libre" Then
    
End If
Loop

Terminar:
End Sub



Sub tubo()


Dim AcadDoc As Object, AcadUtil As Object, AcadModel As Object, Eje1 As Object, blockRef As Object
Dim x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, ANG As Double, lpuntal As Double, lregulacion As Double
Dim rutatubo As String
Dim Punto_inicial(0 To 2) As Double, Punto_final(0 To 2) As Double, Punto_inicial2(0 To 2) As Double, Punto_final2(0 To 2) As Double, Punto_aux1(0 To 2) As Double, Punto_aux2(0 To 2) As Double, P1(0 To 2) As Double, P2(0 To 2) As Double
Dim Gcapa As Object
Dim Ncapa As String, punto1 As Variant, punto2 As Variant
Dim PI As Variant, Distancia As Double
Dim PuntoM(0 To 2) As Double, husillo_l As String, cuerpo_l As String
Dim ext1 As String, ext2 As String, vista As String, tor_ext1 As String, tor_ext2 As String, nrt As String
Dim margen1 As Double, margen2 As Double
Dim ExtRT1(0 To 2) As Double, ExtRT2(0 To 2) As Double, insRT(0 To 2) As Double
Dim torple1(0 To 2) As Double, torple2(0 To 2) As Double, tormar1up(0 To 2) As Double, tormar1down(0 To 2) As Double, tormar2up(0 To 2) As Double, tormar2down(0 To 2) As Double
Dim rtubo As String, plalz As String, rutator As String, tor1 As String, tor2 As String, lt As Double
Dim grapa1 As String, grapa2 As String

Set AcadDoc = GetObject(, "Autocad.Application").ActiveDocument
Set AcadModel = AcadDoc.ModelSpace
Set AcadUtil = AcadDoc.Utility

Ncapa = "Mega"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Lolashor"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30
Ncapa = "Slims"
Set Gcapa = AcadDoc.Layers.Add(Ncapa)
Gcapa.color = 30

rutatubo = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Grapas-Auxiliares\"
rutator = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"

On Error GoTo Terminar

' preguntas -------------------------------------------------------

kwordList = "90 Giratoria MediaGrapa TuboSlimshor GrapaB Perfil-Tubo Libre"
ext1 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
ext1 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Grapa en el primer extremo?: [90/Giratoria/MediaGrapa/TuboSlimshor/GrapaB/Perfil-Tubo/Libre]")

kwordList = "90 Giratoria MediaGrapa TuboSlimshor GrapaB Perfil-Tubo Libre"
ext2 = ""
ThisDrawing.Utility.InitializeUserInput 0, kwordList
ext2 = ThisDrawing.Utility.GetKeyword(vbLf & "¿Grapa en el segundo extremo?: [90/Giratoria/MediaGrapa/TuboSlimshor/GrapaB/Perfil-Tubo/Libre]")


If ext1 = "90" Or ext1 = "" Then
    grapa1 = rutatubo & "Grapa90_alzado.dwg"
ElseIf ext1 = "Giratoria" Then
    grapa1 = rutatubo & "GrapaGiratoria.dwg"
ElseIf ext1 = "MediaGrapa" Then
    grapa1 = rutatubo & "MediaGrapa_planta.dwg"
ElseIf ext1 = "TuboSlimshor" Then
    grapa1 = rutatubo & "GrapaSlimshor_planta.dwg"
ElseIf ext1 = "GrapaB" Then
    grapa1 = rutatubo & "GrapaB_alzado.dwg"
ElseIf ext1 = "Perfil-Tubo" Then
    grapa1 = rutatubo & "Grapa_perfil_tubo_seccion.dwg"
End If


If ext2 = "90" Or ext2 = "" Then
    grapa2 = rutatubo & "Grapa90_alzado.dwg"
ElseIf ext2 = "Giratoria" Then
    grapa2 = rutatubo & "GrapaGiratoria.dwg"
ElseIf ext2 = "MediaGrapa" Then
    grapa2 = rutatubo & "MediaGrapa_planta.dwg"
ElseIf ext2 = "TuboSlimshor" Then
    grapa2 = rutatubo & "GrapaSlimshor_planta.dwg"
ElseIf ext2 = "GrapaB" Then
    grapa2 = rutatubo & "GrapaB_alzado.dwg"
ElseIf ext2 = "Perfil-Tubo" Then
    grapa2 = rutatubo & "Grapa_perfil_tubo_seccion.dwg"
End If


Dim repite As Integer
repite = 1

'COMIENZA Bucle
Do While repite = 1

punto1 = AcadUtil.GetPoint(, "1º Punto: ")
punto2 = AcadUtil.GetPoint(punto1, "2º Punto: ")

Dim lrt As Double



PI = 4 * Atn(1)

P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)

'Set Eje1 = AcadModel.AddLine(P1, P2)
ANG = AcadUtil.AngleFromXAxis(P1, P2)

x = P2(0) - P1(0)
y = P2(1) - P1(1)
Xs = 1
Ys = 1
Zs = 1
Distancia = Val(Sqr((x ^ 2 + y ^ 2)))

lt = Distancia

If lt < 500 And lt > 150 Then
    lt = 500
ElseIf lt >= 500 And lt <= 700 Then
    lt = 750
ElseIf lt > 700 And lt <= 950 Then
    lt = 1000
ElseIf lt > 950 And lt <= 1450 Then
    lt = 1500
ElseIf lt > 1450 And lt <= 1950 Then
    lt = 2000
ElseIf lt > 1950 And lt <= 2450 Then
    lt = 2500
ElseIf lt > 2450 And lt <= 2950 Then
    lt = 3000
ElseIf lt > 2950 And lt <= 3450 Then
    lt = 3500
ElseIf lt > 3450 And lt <= 3950 Then
    lt = 4000
ElseIf lt > 3950 And lt <= 4450 Then
    lt = 4500
ElseIf lt > 4450 And lt <= 4950 Then
    lt = 5000
ElseIf lt > 4950 And lt <= 5450 Then
    lt = 5500
ElseIf lt > 5450 And lt <= 5970 Then
    lt = 6000
Else
    MsgBox "No existe tubo de estas dimensiones"
End If

Punto_inicial(0) = P1(0): Punto_inicial(1) = P1(1): Punto_inicial(2) = P1(2)
Punto_final(0) = Punto_inicial(0) + Distancia * Cos(ANG): Punto_final(1) = Punto_inicial(1) + Distancia * Sin(ANG): Punto_final(2) = Punto_inicial(2)

PuntoM(0) = Punto_inicial(0) + (Distancia / 2) * Cos(ANG): PuntoM(1) = Punto_inicial(1) + (Distancia / 2) * Sin(ANG): PuntoM(2) = Punto_inicial(2)


rtubo = rutatubo & "TuboArr" & lt & ".dwg"
Set blockRef = AcadModel.InsertBlock(PuntoM, rtubo, Xs, Ys, Zs, ANG)
blockRef.Layer = "AM"

If ext1 = "Libre" Then
Else
    Set blockRef = AcadModel.InsertBlock(Punto_inicial, grapa1, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
End If

If ext2 = "Libre" Then
Else
    Set blockRef = AcadModel.InsertBlock(Punto_final, grapa2, Xs, Ys, Zs, ANG)
    blockRef.Layer = "AM"
End If

Loop

Terminar:
End Sub
