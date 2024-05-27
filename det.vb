Option Explicit

Sub det()
Dim GcadDoc As Object, GcadUtil As Object, GcadModel As Object, Eje1 As Object, blockRef As Object, GcadPaper As Object
Dim rutags As String, rutadet As String, rutator As String, rutampacc As String, rutacuña As String, capa As String
Dim Gcapa As Object
Dim Ncapa As String, cuña As String, lado As String, tipo1 As String, tipo2 As String, disposicion As String, kwordList As String, M20x90_4 As String
Dim repite As Double, ANG As Double, x As Double, y As Double, z As Double, Xs As Double, Ys As Double, Zs As Double, P1(0 To 2) As Double, P2(0 To 2) As Double, Punto_inial(0 To 2) As Double
Dim punto1 As Variant, punto2 As Variant, PI As Variant



Set GcadDoc = GetObject(, "Gcad.Application").ActiveDocument
Set GcadModel = GcadDoc.ModelSpace
Set GcadPaper = GcadDoc.PaperSpace
Set GcadUtil = GcadDoc.Utility

rutadet = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\det.dwg"

        punto1 = GcadUtil.GetPoint(, "1º Punto: ")
        punto2 = GcadUtil.GetPoint(punto1, "2º Punto: ")
        P1(0) = punto1(0): P1(1) = punto1(1): P1(2) = punto1(2)
        P2(0) = punto2(0): P2(1) = punto2(1): P2(2) = punto2(2)

        
        ANG = GcadUtil.AngleFromXAxis(P1, P2)

        x = P2(0) - P1(0)
        y = P2(1) - P1(1)
        Xs = 1
        Ys = 1
        Zs = 1
        Set blockRef = GcadPaper.InsertBlock(P1, rutadet, Xs, Ys, Zs, ANG)
        'Set BlockRef = GcadModel.InsertBlock(P1, rutadet, Xs, Ys, Zs, ANG)

End Sub

