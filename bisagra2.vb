 Sub bisagra2()
    Dim doc As Object
    Set doc = ThisDrawing
    
    Dim blockRef As Object
    
    Set gcadDoc = GetObject(, "Gcad.Application").ActiveDocument
    Set gcadModel = gcadDoc.ModelSpace
    Set gcadUtil = gcadDoc.Utility
    
    Ncapa = "Mega"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 30
    Ncapa = "Granshor"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 150
    Ncapa = "Pipeshor4S"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 7
    Ncapa = "Pipeshor6"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 7
    Ncapa = "Pipeshor4L"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 5
    Ncapa = "Slims"
    Set Gcapa = gcadDoc.Layers.Add(Ncapa)
    Gcapa.color = 30
    
    Dim intPoint As Variant
    Dim PA(0 To 2) As Double
    Dim PP1(0 To 2) As Double
    Dim PB(0 To 2) As Double
    Dim PA2(0 To 2) As Double
    Dim PB2(0 To 2) As Double
    Dim Esq(0 To 2) As Double
    Dim Esqb(0 To 2) As Double
    Dim Esqi(0 To 2) As Double
    Dim Esqt(0 To 2) As Double
    Dim DirMuro1 As Double
    Dim DirMuro2 As Double
    Dim DirMuro1a As Double
    Dim DirMuro2a As Double
    Dim Slope1 As Double
    Dim Slope2 As Double
    Dim pl As String
    Dim pr As String
    Dim va As String
    Dim Xs As Double
    Dim Ys As Double
    Dim Zs As Double
    Dim distanciaA As Double
    Dim distanciaB As Double
    Dim xa As Double
    Dim ya As Double
    Dim xb As Double
    Dim yb As Double
    Dim dato1 As String
    Dim dato2 As String
    Dim perf As String, jr As String, jr3 As String, wall As String
    Dim lon As String
    Dim alma450 As String, alma As String, junta As String
    Dim lperfil As Double
    Dim n4570 As Integer
    Dim n4500 As Integer
    Dim n4100 As Integer
    Dim n4030 As Integer
    Dim n3070 As Integer
    Dim n3000 As Integer
    Dim n1500 As Integer
    Dim nil15000 As Integer
    Dim n900 As Integer, n6070 As Integer, n6000 As Integer, ni15000 As Integer, ni10500 As Integer
    Dim Mp90 As Integer, Mp180 As Integer, Mp450 As Integer, Mp270 As Integer
    Dim lP4570JR As Double
    Dim lP4500 As Double
    Dim lP3070JR As Double
    Dim lP3000 As Double
    Dim lMp90 As Double, lMp270 As Double, lMp450 As Double, lMp180 As Double
    Dim ruta2 As String
    Dim lfija As Double, lbisagra As Double, lP900 As Double
    Dim rutaperf As String, rutamp As String
    Dim repite As Integer
    Dim PAa As Variant, Esqtt As Variant
    Dim GetAngleBetweenLines As Double
    Dim Ptemp(0 To 2) As Double
    Dim lfijaH300 As Double, lfijaH300JR As Double, lfijaH450 As Double, lfijaH600 As Double, lfijaH600JR As Double
    
    On Error GoTo terminar
    
    rutaperf = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\Perfiles\"
    rutamp = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\MSHOR\VIGAS\"
    ruta2 = "C:\Users\" & Environ$("Username") & "\Incye\Ingenieria - Documentos\12_Aplicaciones\MACROS_21\Automaticos_Biblioteca\TORNILLERIA\"
    repite = 1
    
    'Valores fijos
    PI = 4 * Atn(1)
    hbisa = 630
    vbisa = 160
    lbisagra = 630
    lP900 = 900
    lP450 = 4500
    lP600N = 4030
    lP600JR = 4100
    lP300N = 4500
    lP4570JR = 4570
    lP4500 = 4500
    lP3070JR = 3070
    lP3000 = 3000
    lP6070JR = 6070
    lP6000 = 6000
    lP1500 = 1500
    li15000 = 15000
    li10500 = 10500
    lP4030 = 4030
    lP4100 = 4100
    lMp90 = 90
    lMp180 = 180
    lMp270 = 270
    lMp450 = 450
    
    lfijaH300 =  lP900 
    lfijaH300JR = lP3070JR 
    lfijaH450 = lP3000 
    lfijaH600JR = lP4100 
    lfijaH600 = lP4030 
    
    On Error GoTo terminar
	
		' Dibujar la primera línea
        Dim puntoInicioLinea1 As Variant
        Dim puntoFinLinea1 As Variant
		
		' Dibujar la segunda línea
        Dim puntoInicioLinea2 As Variant
        Dim puntoFinLinea2 As Variant
					
            va = rutaperf & "Incye_600JR_4000_AL.dwg"
            v300jr4570 = rutaperf & "Incye_300JR_4570_AL.dwg"
            v300jr3070 = rutaperf & "Incye_300JR_3070_AL.dwg"
            v300jr6070 = rutaperf & "Incye_300JR_6070_AL.dwg"
            v300n6000 = rutaperf & "Incye_300_6000_AL.dwg"
            v300n4500 = rutaperf & "Incye_300_4500_AL.dwg"
            v300jr4500 = rutaperf & "Incye_300_4500_AL.dwg"
            v300n3000 = rutaperf & "Incye_300_3000_AL.dwg"
            v300n1500 = rutaperf & "Incye_300_1500_AL.dwg"
            v300n900 = rutaperf & "Incye_300_900_AL.dwg"
            v450SAjr4500 = rutaperf & "Incye_450SAJR_4500_AL.dwg"
            v450SAjr6000 = rutaperf & "Incye_450SAJR_6000_AL.dwg"
            v450SAjr3000 = rutaperf & "Incye_450SAJR_3000_AL.dwg"
            v450jr4500 = rutaperf & "Incye_450JR_4500_AL.dwg"
            v450jr6000 = rutaperf & "Incye_450JR_6000_AL.dwg"
            v450jr3000 = rutaperf & "Incye_450JR_3000_AL.dwg"
            v600jr4000 = rutaperf & "Incye_600JR_4000_AL.dwg"
            v600n4000 = rutaperf & "Incye_600_4000_AL.dwg"
            Mpshor450 = rutamp & "Mshor450ALZ.dwg"
            Mpshor270 = rutamp & "Mshor270ALZ.dwg"
            Mpshor180 = rutamp & "Mshor180ALZ.dwg"
            Mpshor90 = rutamp & "Mshor90ALZ.dwg"
        	
			kwordList = "MuroA MuroB"
            ThisDrawing.Utility.InitializeUserInput 0, kwordList
            wall = ThisDrawing.Utility.GetKeyword(vbLf & "Cuantos lados: [MuroA/MuroB]")
			
			If wall = "MuroA" Or wall = "" Then
			
				puntoInicioLinea1 = doc.Utility.GetPoint(, "Selecciona el punto de inicio del muro: ")
				puntoFinLinea1 = doc.Utility.GetPoint(puntoInicioLinea1, "Selecciona el punto de fin del muro: ")
				puntoInicioLinea2 = doc.Utility.GetPoint(puntoFinLinea1, "Selecciona el lado del muro: ")
			
				'PA es el punto de inserción de la primera linea
				PA(0) = puntoInicioLinea1(0): PA(1) = puntoInicioLinea1(1): PA(2) = puntoInicioLinea1(2)
				PP1(0) = puntoFinLinea1(0): PP1(1) = puntoFinLinea1(1): PP1(2) = puntoFinLinea1(2)
				
				'Obtener el angulo de las lineas
				DirMuro1 = gcadUtil.AngleFromXAxis(PA, PP1)
				
				'calculo lado A
				xa = PA(0) - PP1(0)
				ya = PA(1) - PP1(1)
				
				Xs = 1
				Ys = 1
				Zs = 1
				distanciaA = Val(Sqr((xa ^ 2 + ya ^ 2)))
				
				'Giro de direccion 90 grados para hallar perpendicular
				DirMuro1a = DirMuro1 - ((PI) / 2)
											
				M30x100_2 = ruta2 & "2-M30x100{10.9}.dwg"
				M30x100_3 = ruta2 & "3-M30x100{10.9}.dwg"
				M20x130_6 = ruta2 & "6-M20x130.dwg"
				M20x90_10 = ruta2 & "10-M20x90.dwg"
				M20x130_12 = ruta2 & "12-M20x130.dwg"
				M20x90_6 = ruta2 & "6-M20x90.dwg"
				
				'Ubicacion tornillo en junta reforzada
				Esqt(0) = PA(0) - 290 * Cos(DirMuro1a): Esqt(1) = PA(1) - 290 * Sin(DirMuro1a): Esqt(2) = PA(2)
				
				kwordList = "300 450 600"
				ThisDrawing.Utility.InitializeUserInput 0, kwordList
				perf = ThisDrawing.Utility.GetKeyword(vbLf & "Viga HEB: [300/450/600]")
			
				If perf = "300" or perf = "" Then
				
					kwordList = "Sí No"
					ThisDrawing.Utility.InitializeUserInput 0, kwordList
					jr3 = ThisDrawing.Utility.GetKeyword(vbLf & "Juntas reforzadas?: [Sí/No]")
					If jr3 = "Sí" Or jr3 = "" Then                
						               
						kwordList = "3070 4570 6070"
						ThisDrawing.Utility.InitializeUserInput 0, kwordList
						lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [3070/4570/6070]")
						
						If lon = "3070" or lon="" Then
						
							If distanciaA < lfijaH300JR Then
							MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300JR & "mm."""
							GoTo terminar
							End If
							
							lperfil = distanciaA
							n3070 = Fix(lperfil / lP3070JR)
							lperfil = lperfil - n3070 * lP3070JR
							Mp450 = Fix(lperfil / lMp450)
							Mp270 = Fix(lperfil / lMp270)
							Mp180 = Fix(lperfil / lMp180)
							Mp90 = Fix(lperfil / lMp90)
                                                        
							If n3070 > 0 Then
							i = 0
							Do While i < n3070
							Set blockRef = gcadModel.InsertBlock(PA, v300jr3070, Xs, Ys, Zs, DirMuro1 )
							blockRef.Layer = "Perfiles INCYE"
							Set blockRef = gcadModel.InsertBlock(PA, M20x130_6, Xs, Ys, Zs, DirMuro1 )
							blockRef.Layer = "Nonplot"
							blockRef.Update
							blockRef.Explode
							blockRef.Delete
							PA(0) = PA(0) + lP3070JR * Cos(DirMuro1): PA(1) = PA(1) + lP3070JR * Sin(DirMuro1): PA(2) = PA(2)
							Set blockRef = gcadModel.InsertBlock(Esqt, M30x100_3, Xs, Ys, Zs, DirMuro1 )
							blockRef.Layer = "Nonplot"
							blockRef.Update
							blockRef.Explode
							blockRef.Delete
							Esqt(0) = Esqt(0) + lP3070JR * Cos(DirMuro1): Esqt(1) = Esqt(1) + lP3070JR * Sin(DirMuro1): Esqt(2) = Esqt(2)
							i = i + 1
							Loop
                    
							End If
                        
                            'Nivelar Megapro
                            PA(0) = PA(0) - 10 * Cos(DirMuro1a): PA(1) = PA(1) - 10 * Sin(DirMuro1a): PA(2) = PA(2)
                                            
                            'Call MegaproLadoA(PA, DirMuro1, Mp90, Mp180, Mp270, Mp450, lperfil)
						
						Elseif lon = "4570" Then
						
							If distanciaA < lfijaH300JR Then
							MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300JR & "mm."""
							GoTo terminar
							End If
						
						Elseif lon = "6070" Then
						
							If distanciaA < lfijaH300JR Then
							MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300JR & "mm."""
							GoTo terminar
							End If
						
						End If
																		
					Else
																						                
						kwordList = "900 1500 3000 4500 6000"
						ThisDrawing.Utility.InitializeUserInput 0, kwordList
						lon = ThisDrawing.Utility.GetKeyword(vbLf & "Longitud?: [900/1500/3000/4500/6000]")
						
						If lon = "900" or lon="" Then
						
							If distanciaA < lfijaH300 Then
							MsgBox "Medida Muro " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300 & "mm."""
							GoTo terminar
							End If
						
						Elseif lon = "1500" Then
						
							If distanciaA < lfijaH300 Then
							MsgBox "Medida Muro " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300 & "mm."""
							GoTo terminar
							End If
						
						Elseif lon = "3000" Then
						
							If distanciaA < lfijaH300 Then
							MsgBox "Medida Muro " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300 & "mm."""
							GoTo terminar
							End If
						
						Elseif lon = "4500" Then
						
							If distanciaA < lfijaH300 Then
							MsgBox "Medida Muro " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300 & "mm."""
							GoTo terminar
							End If
						
						Elseif lon = "6000" Then
						
							If distanciaA < lfijaH300 Then
							MsgBox "Medida Muro " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH300 & "mm."""
							GoTo terminar
							End If
						
						End If
					
					End If
				
				ElseIf perf = "450" Then				
					
					kwordList = "Triple Simple"
					ThisDrawing.Utility.InitializeUserInput 0, kwordList
					alma450 = ThisDrawing.Utility.GetKeyword(vbLf & "Alma?: [Triple/Simple]")
					
					If alma450 = "Triple" Or alma450 = "" Then
						
							If distanciaA < lfijaH450 Then						
							MsgBox "Medida Muro " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH450 & "mm."""						
							GoTo terminar						
							End If
											
					ElseIf alma450 = "Simple" Then
						
							If distanciaA < lfijaH450 Then						
							MsgBox "Medida Muro " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH450 & "mm."""						
							GoTo terminar						
							End If
											
					End If
				
				'HEB600
						
				Else 
				
					kwordList = "Sí No"
					ThisDrawing.Utility.InitializeUserInput 0, kwordList
					jr = ThisDrawing.Utility.GetKeyword(vbLf & "Juntas reforzadas?: [Sí/No]")

					If jr = "Sí" Or jr = "" Then
						
							If distanciaA < lfijaH600JR Then
								MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH600JR & "mm."""
								GoTo terminar
							End If
						
					Else
						
							If distanciaA < lfijaH600 Then
								MsgBox "Medida Muro A  " & distanciaA & "mm, menor que el mínimo necesario de " & lfijaH600 & "mm."""
								GoTo terminar
							End If
						
					End If
				
				End If					
																											
				On Error GoTo terminar
			
								
			Elseif wall = "MuroB" Then
			
				Call bisagra.bisagra
				
				On Error GoTo terminar
			
			Else
			
				Goto terminar
			
			End if
				
				
			
			
        
terminar:
End Sub
