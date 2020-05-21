Attribute VB_Name = "ModCnt"
Option Explicit

'Mapa actual seleccionado para el conectar renderizado
Private SelectConnectMap As Byte

Enum EPantalla
    PConnect = 0
    PCuenta
    PCrearPJ
End Enum

'Indica la posicion donde se va a renderizar los PJ
Private PJPos(1 To 10) As WorldPos

Public Pantalla As EPantalla

Public Sub InicializarPosicionesPJ()
'********************************************
'Autor: Lorwik
'Fecha: 19/05/2020
'Descripcion: Inicia las posiciones donde se van a mostrar los PJ
'********************************************

    PJPos(1).x = 468
    PJPos(1).y = 462
    
    PJPos(2).x = 340
    PJPos(2).y = 456
    
    PJPos(3).x = 570
    PJPos(3).y = 453
    
    PJPos(4).x = 243
    PJPos(4).y = 378
    
    PJPos(5).x = 664
    PJPos(5).y = 408
    
    PJPos(6).x = 223
    PJPos(6).y = 450
    
    PJPos(7).x = 300
    PJPos(7).y = 286
    
    PJPos(8).x = 608
    PJPos(8).y = 286
    
    PJPos(9).x = 747
    PJPos(9).y = 550
    
    PJPos(10).x = 637
    PJPos(10).y = 627
    
End Sub

Public Sub MostrarConnect(Optional ByVal Mostrar As Boolean = False)
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'llama al frmConnect con el mapa, de lo contrario no funcionaria correctamente.
'******************************

    'Seteamos el modo login
    Pantalla = PConnect
    
    If Mostrar = True Then frmConnect.Visible = True
    
    'Sorteamos el mapa a mostrar
    'Nota el mapa 1 es para el crear pj, el 2 para las cuentas
    SelectConnectMap = RandomNumber(3, NumConnectMap)
    Call SwitchMap(MapaConnect(SelectConnectMap).Map)
End Sub

Public Sub MostrarCuenta(Optional ByVal Mostrar As Boolean = False)
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'Cambia al modo cuenta
'******************************

    'Seteamos el modo login
    Pantalla = PCuenta
    
    If Mostrar = True Then frmConnect.Visible = True
    
    'Ponemos el mapa de cuentas
    SelectConnectMap = 2
    Call SwitchMap(MapaConnect(SelectConnectMap).Map)
End Sub

Public Sub MostrarCreacion(Optional ByVal Mostrar As Boolean = False)
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'Cambia al modo cuenta
'******************************

    'Seteamos el modo login
    Pantalla = PCrearPJ
    
    If Mostrar = True Then frmConnect.Visible = True
    
    'Ponemos el mapa de cuentas
    SelectConnectMap = 1
    Call SwitchMap(MapaConnect(SelectConnectMap).Map)
End Sub

Sub RenderConnect()
'******************************
'Autor: Lorwik
'Fecha: 15/05/2020
'Renderiza el screen del conectar
'******************************
    Dim x As Long
    Dim y As Long
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    Static RE As RECT
    
    With RE
        .Left = 0
        .Top = 0
        .Bottom = 768
        .Right = 1024
    End With
    
    Call Engine_BeginScene
     
    For x = 1 To 32
        For y = 1 To 24
            PixelOffsetXTemp = (x - 1) * 32
            PixelOffsetYTemp = (y - 1) * 32
            
            With MapData(x + MapaConnect(SelectConnectMap).x, y + MapaConnect(SelectConnectMap).y)
                'Capa 1
                Call Draw_Grh(.Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                
                'Capa 2
                Call Draw_Grh(.Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
            End With
        Next y
    Next x
        
    'Capa 3
    For x = 1 To 32
        For y = 1 To 24
            PixelOffsetXTemp = (x - 1) * 32
            PixelOffsetYTemp = (y - 1) * 32
            With MapData(x + MapaConnect(SelectConnectMap).x, y + MapaConnect(SelectConnectMap).y)
            
                'Objectos
                If .ObjGrh.GrhIndex <> 0 Then _
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                    
                'Personajes
                Call RenderPJ
                            
                'Capa 3
                Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                
                'Particulas
                If .Particle_Group_Index Then
                    
                    'Solo las renderizamos si estan cerca del area de vision.
                    If EstaDentroDelArea(x, y) Then
                        Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                    End If
                        
                End If
            End With
        Next y
    Next x
    
    For x = 1 To 32
        For y = 1 To 24
            PixelOffsetXTemp = (x - 1) * 32
            PixelOffsetYTemp = (y - 1) * 32
            
            With MapData(x + MapaConnect(SelectConnectMap).x, y + MapaConnect(SelectConnectMap).y)
                'Capa 4
                Call Draw_Grh(.Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
            End With
        Next y
    Next x
    
    'Renderizamos la interfaz
    Call RenderConnectGUI
     
    'Get timing info
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
        
    Call Engine_EndScene(RE, frmConnect.renderer.hWnd)
End Sub

Private Sub RenderConnectGUI()
'******************************
'Autor: Lorwik
'Fecha: 15/05/2020
'Dibuja la interfaz
'******************************
    Select Case Pantalla
    
        Case 0 'Login (frmconnect)
        
        Case 1 'Cuenta
            'Crear PJ
            Call Engine_Draw_Box(850, 500, 150, 50, D3DColorARGB(150, 0, 0, 0))
            Call DrawText(930, 520, "Crear Personaje", -1, True)
        
            'Borrar PJ
            Call Engine_Draw_Box(850, 570, 150, 50, D3DColorARGB(150, 0, 0, 0))
            Call DrawText(930, 585, "Borrar Personaje", -1, True)
            
            'Salir
            Call Engine_Draw_Box(30, 670, 150, 50, D3DColorARGB(150, 0, 0, 0))
            Call DrawText(100, 685, "Salir", -1, True)
            
        Case 2 'Crear PJ
    
    End Select
    
    '<------- Desde aqui lo que siempre se va a mostrar ------->
    'Logo
    Call Draw_GrhIndex(31480, 1, 1, 0, Normal_RGBList(), 0, False)
    
    Call DrawText(10, 750, "WinterAO " & GetVersionOfTheGame() & " Resurrection", Color_Paralisis)
End Sub

Private Sub RenderPJ()
'******************************
'Autor: Lorwik
'Fecha: 15/05/2020
'Dibuja los Personajes
'******************************
    Dim Index As Byte
    
    Select Case Pantalla
        Case 1 'Cuenta
            For Index = 1 To NumberOfCharacters
                With cPJ(Index)
    
                    If .Body <> 0 Then
            
                        Call Draw_Grh(BodyData(.Body).Walk(3), PJPos(Index).x, PJPos(Index).y, 1, Normal_RGBList(), 0)
            
                        If .Head <> 0 Then
                            Call Draw_Grh(HeadData(.Head).Head(3), PJPos(Index).x + BodyData(.Body).HeadOffset.x, PJPos(Index).y + BodyData(.Body).HeadOffset.y, 1, Normal_RGBList(), 0)
                        End If
            
                        If .helmet <> 0 Then
                            Call Draw_Grh(CascoAnimData(.helmet).Head(3), PJPos(Index).x + BodyData(.Body).HeadOffset.x, PJPos(Index).y + BodyData(.Body).HeadOffset.y + OFFSET_HEAD, 1, Normal_RGBList(), 0)
                        End If
            
                        If .weapon <> 0 Then
                            Call Draw_Grh(WeaponAnimData(.weapon).WeaponWalk(3), PJPos(Index).x, PJPos(Index).y, 1, Normal_RGBList(), 0)
                        End If
            
                        If .shield <> 0 Then
                            Call Draw_Grh(ShieldAnimData(.shield).ShieldWalk(3), PJPos(Index).x, PJPos(Index).y, 1, Normal_RGBList(), 0)
                        End If
                        
                        'Nombre
                        Call DrawText(PJPos(Index).x + 16, PJPos(Index).y + 30, .Nombre, -1, True)
                        
                        'Nombre de la cuenta
                        Call DrawText(500, 15, AccountName, -1, True, 2)
                        
                    End If
                
                End With
            Next Index
            
    End Select
End Sub

Public Sub ClickEvent(ByVal TX As Long, ByVal TY As Long)
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'Eventos al realizar clicks en la GUI
'******************************
    Dim i As Integer
    
    Dim Index As Byte
    
    Select Case Pantalla
        Case 1 'Cuenta

            'Conectar a PJ
            For i = 1 To NumberOfCharacters
                With cPJ(i)
                    If (TX >= PJPos(i).x And TX <= PJPos(i).x + 20) And (TY >= PJPos(i).y And TY <= PJPos(i).y - OFFSET_HEAD) Then
    
                        If LenB(.Nombre) <> 0 Then
                            UserName = .Nombre
                            Call WriteLoginExistingChar(i)
                        End If
                    End If
                End With
            Next i
            
            'Crear Nuevo PJ
            If (TX >= 850 And TX <= 950) And (TY >= 500 And TY <= 550) Then Call CrearNuevoPJ

            'Salir cuenta
            If (TX >= 30 And TX <= 180) And (TY >= 670 And TY <= 720) Then
                frmMain.Client.CloseSck
                Call ResetAllInfoAccounts
                Call MostrarConnect
            End If
            
    End Select
    
End Sub

Private Sub CrearNuevoPJ()
    If NumberOfCharacters > 9 Then
        Call MostrarMensaje(JsonLanguage.item("ERROR_DEMASIADOS_PJS").item("TEXTO"))
        Exit Sub
    End If
    
    frmCrearPersonaje.Show
    'Call MostrarCreacion
End Sub
