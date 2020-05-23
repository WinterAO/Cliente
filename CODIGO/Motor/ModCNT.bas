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

Public PJAccSelected As Byte
Private TextSelected As Byte

Public Pantalla As EPantalla

Public Sub InicializarPosicionesPJ()
'********************************************
'Autor: Lorwik
'Fecha: 19/05/2020
'Descripcion: Inicia las posiciones donde se van a mostrar los PJ
'********************************************

    PJPos(1).X = 468
    PJPos(1).Y = 462
    
    PJPos(2).X = 340
    PJPos(2).Y = 456
    
    PJPos(3).X = 570
    PJPos(3).Y = 453
    
    PJPos(4).X = 243
    PJPos(4).Y = 378
    
    PJPos(5).X = 664
    PJPos(5).Y = 408
    
    PJPos(6).X = 223
    PJPos(6).Y = 450
    
    PJPos(7).X = 300
    PJPos(7).Y = 286
    
    PJPos(8).X = 608
    PJPos(8).Y = 286
    
    PJPos(9).X = 747
    PJPos(9).Y = 550
    
    PJPos(10).X = 637
    PJPos(10).Y = 627
    
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
    Dim X As Long
    Dim Y As Long
    
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
     
    For X = 1 To 32
        For Y = 1 To 24
            PixelOffsetXTemp = (X - 1) * 32
            PixelOffsetYTemp = (Y - 1) * 32
            
            With MapData(X + MapaConnect(SelectConnectMap).X, Y + MapaConnect(SelectConnectMap).Y)
                'Capa 1
                Call Draw_Grh(.Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                
                'Capa 2
                Call Draw_Grh(.Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
            End With
        Next Y
    Next X
        
    'Capa 3
    For X = 1 To 32
        For Y = 1 To 24
            PixelOffsetXTemp = (X - 1) * 32
            PixelOffsetYTemp = (Y - 1) * 32
            With MapData(X + MapaConnect(SelectConnectMap).X, Y + MapaConnect(SelectConnectMap).Y)
            
                'Objectos
                If .ObjGrh.GrhIndex <> 0 Then _
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                                
                'Capa 3
                Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                
                'Particulas
                If .Particle_Group_Index Then
                    
                    'Solo las renderizamos si estan cerca del area de vision.
                    If EstaDentroDelArea(X, Y) Then
                        Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                    End If
                        
                End If
            End With
        Next Y
    Next X
    
    'Personajes
    Call RenderPJ
    
    For X = 1 To 32
        For Y = 1 To 24
            PixelOffsetXTemp = (X - 1) * 32
            PixelOffsetYTemp = (Y - 1) * 32
            
            With MapData(X + MapaConnect(SelectConnectMap).X, Y + MapaConnect(SelectConnectMap).Y)
                'Capa 4
                Call Draw_Grh(.Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
            End With
        Next Y
    Next X
    
    'Renderizamos la interfaz
    Call RenderConnectGUI
     
    'Get timing info
    'timerElapsedTime = GetElapsedTime()
    'timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
        
    Call Engine_EndScene(RE, frmConnect.Renderer.hWnd)
End Sub

Private Sub RenderConnectGUI()
'******************************
'Autor: Lorwik
'Fecha: 15/05/2020
'Dibuja la interfaz
'******************************
    Select Case Pantalla
    
        Case 0 'Login (frmconnect)
            
            'Menu Cabecera
            Call Draw_GrhIndex(31487, 805, 400, 0, Normal_RGBList(), 0, False)
                  
            'Menu Cabecera
            Call Draw_GrhIndex(31489, 805, 458, 0, Normal_RGBList(), 0, False)
            
            'Recuperar Cuenta
            Call Draw_GrhIndex(31488, 805, 523, 0, Normal_RGBList(), 0, False)
            
            'Marco
            Call Draw_GrhIndex(31481, 0, 0, 0, Normal_RGBList(), 0, False)
    
            'Logo
            Call Draw_GrhIndex(31480, 1, -140, 0, Normal_RGBList(), 0, False)
            
            'Conectarse
            Call Draw_GrhIndex(31490, 390, 435, 0, Normal_RGBList(), 0, False)
            
            'Conectarse
            Call Draw_GrhIndex(31491, 530, 435, 0, Normal_RGBList(), 0, False)
            
            'Recordar
            Call Draw_GrhIndex(31484, 390, 490, 0, Normal_RGBList(), 0, False)
            
            'User
            Call DrawText(445, 372, frmConnect.txtNombre.Text, -1, False)
            If TextSelected = 1 Then _
                Call Draw_GrhIndex(25319, 446 + Engine_AnchoTexto(1, frmConnect.txtNombre.Text), 372, 0, Normal_RGBList(), 0, False)
            
            'Password
            Call DrawText(445, 406, frmConnect.txtPasswd.Text, -1, False)
            If TextSelected = 2 Then _
                Call Draw_GrhIndex(25319, 446 + Engine_AnchoTexto(1, frmConnect.txtPasswd.Text), 406, 0, Normal_RGBList(), 0, False)
            
        Case 1 'Cuenta
        
            'Marco
            Call Draw_GrhIndex(31481, 0, 0, 0, Normal_RGBList(), 0, False)
            
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
            'Marco
            Call Draw_GrhIndex(31481, 0, 0, 0, Normal_RGBList(), 0, False)
    
    End Select
    
    '<------- Desde aqui lo que siempre se va a mostrar ------->
    
    ' Calculamos los FPS y los mostramos
    Call Engine_Update_FPS
    'If ClientSetup.FPSShow = True Then
    Call DrawText(970, 30, "FPS: " & Mod_TileEngine.FPS, -1, True)
    
    Call DrawText(25, 730, "WinterAO " & GetVersionOfTheGame() & " Resurrection", Color_Paralisis)
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
            
                        Call Draw_Grh(BodyData(.Body).Walk(3), PJPos(Index).X, PJPos(Index).Y, 1, Normal_RGBList(), 0)
            
                        If .Head <> 0 Then
                            Call Draw_Grh(HeadData(.Head).Head(3), PJPos(Index).X + BodyData(.Body).HeadOffset.X, PJPos(Index).Y + BodyData(.Body).HeadOffset.Y, 1, Normal_RGBList(), 0)
                        End If
            
                        If .helmet <> 0 Then
                            Call Draw_Grh(CascoAnimData(.helmet).Head(3), PJPos(Index).X + BodyData(.Body).HeadOffset.X, PJPos(Index).Y + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, 1, Normal_RGBList(), 0)
                        End If
            
                        If .weapon <> 0 Then
                            Call Draw_Grh(WeaponAnimData(.weapon).WeaponWalk(3), PJPos(Index).X, PJPos(Index).Y, 1, Normal_RGBList(), 0)
                        End If
            
                        If .shield <> 0 Then
                            Call Draw_Grh(ShieldAnimData(.shield).ShieldWalk(3), PJPos(Index).X, PJPos(Index).Y, 1, Normal_RGBList(), 0)
                        End If
                        
                        'Nombre
                        Call DrawText(PJPos(Index).X + 16, PJPos(Index).Y + 30, .Nombre, -1, True)
                        
                        'Nombre de la cuenta
                        Call DrawText(500, 15, AccountName, -1, True, 2)
                        
                    End If
                
                End With
            Next Index
            
    End Select

End Sub

Public Sub DobleClickEvent(ByVal TX As Long, ByVal TY As Long)
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'Eventos al realizar doble click en la GUI
'******************************
    Dim i As Integer
    
    Dim Index As Byte
    
    Select Case Pantalla
        Case 1 'Cuenta

            'Con doble click conectamos al PJ
            For i = 1 To NumberOfCharacters
                With cPJ(i)
                    If (TX >= PJPos(i).X And TX <= PJPos(i).X + 20) And (TY >= PJPos(i).Y And TY <= PJPos(i).Y - OFFSET_HEAD) Then
    
                        If LenB(.Nombre) <> 0 Then
                            UserName = .Nombre
                            Call WriteLoginExistingChar
                        End If
                    End If
                End With
            Next i

    End Select
    
End Sub

Public Sub ClickEvent(ByVal TX As Long, ByVal TY As Long)
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'Eventos al realizar click en la GUI
'******************************
    Dim i As Integer
    
    Dim Index As Byte

    Select Case Pantalla
        Case 0 'Conectar
        
            If (TX >= 443 And TX <= 605) And (TY >= 372 And TY <= 384) Then
                frmConnect.txtNombre.SetFocus
                TextSelected = 1
            End If
            
            If (TX >= 443 And TX <= 605) And (TY >= 405 And TY <= 424) Then
                frmConnect.txtPasswd.SetFocus
                TextSelected = 2
            End If
        
        Case 1 'Cuenta

            'Seleccionamos un PJ
            For i = 1 To NumberOfCharacters
                With cPJ(i)
                    If (TX >= PJPos(i).X And TX <= PJPos(i).X + 20) And (TY >= PJPos(i).Y And TY <= PJPos(i).Y - OFFSET_HEAD) Then
    
                        If LenB(.Nombre) <> 0 Then
                            'El PJ seleccionado queda guardado
                            UserName = .Nombre
                            PJAccSelected = i
                        End If
                    End If
                End With
            Next i
            
            'Crear Nuevo PJ
            If (TX >= 850 And TX <= 950) And (TY >= 500 And TY <= 550) Then Call CrearNuevoPJ
            
            'Borrar PJ
            If (TX >= 850 And TX <= 950) And (TY >= 570 And TY <= 620) Then
                If PJAccSelected < 1 Then
                    Call MostrarMensaje(JsonLanguage.item("ERROR_PERSONAJE_NO_SELECCIONADO").item("TEXTO"))
                    Exit Sub
                End If
                    
                frmBorrarPJ.Show
            
            End If

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
