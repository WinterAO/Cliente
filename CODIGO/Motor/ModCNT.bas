Attribute VB_Name = "ModCnt"
'**************************************************************************************
'Autor: Lorwik
'Fecha 20/05/2020
'Descripcion: En este modulo vamos a renderizar todos los referente con
'el conectar, cuentas y crearpj, asi como eventos.
'
'NOTA: No esta programado de la forma mas eficiente, pero por el momento es funcional
'en el futuro se deberia revisar y mejorar en todo lo posible.
'**************************************************************************************

Option Explicit

'Indica que mapa vamos a renderizar en el conectar
Private SelectConnectMap As Byte

'*********************
'Modo de pantalla renderizado
'*********************
Enum EPantalla
    PConnect = 0
    PCuenta
    PCrearPJ
End Enum

Public Pantalla As EPantalla
'*********************

'Indica la posicion donde se va a renderizar los PJ
Private PJPos(1 To 10) As WorldPos

'Indifca el PJ seleccionado
Public PJAccSelected As Byte

'Indica el TextBox selecionado
Private TextSelected As Byte

'Codigo de encriptado de la pass
Public Const AES_PASSWD As String = "illoestapassestodifisi"

'*********************
'Flags
'*********************
Public Conectando As Boolean 'Para evitar mandar varias peticiones al servidor a la hora de conectar
Private botonCrear As Boolean
'*********************

'Velocidad con la que parpadea el cursor de texto
Private Const CursorFlashRate As Long = 450

'*********************
'Creacion de PJ
'*********************
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'Puramente informativo
Private Type tModRaza
    Fuerza As Single
    Agilidad As Single
    Inteligencia As Single
    Carisma As Single
    Constitucion As Single
End Type

Private ModRaza()  As tModRaza
Private lblModRaza(1 To NUMRAZAS) As Integer
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Private lblAtributos(1 To NUMATRIBUTES) As Byte 'Puntos ya asignados
Private lblTotal As Byte 'Total de puntos para asignar

Private SexoSelect(1 To 2) As String 'Sexo seleccionado
'**********************


Public Sub InicializarRndCNT()
'********************************************
'Autor: Lorwik
'Fecha: 19/05/2020
'Descripcion: Inicia todo lo que tiene que ver con el conectar renderizado
'como las posiciones donde se van a mostrar los PJ, pantalla, etc...
'********************************************

    'Cargamos los mapas que mostraremos
    Call CargarConnectMaps

    'Inicializamos las posiciones de los PJ en la pantalla de cuentas
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
    
    Pantalla = PConnect 'Establecemos la pantalla en el conectar
    TextSelected = 1 ' Establecemos el cursor de texto en Nombre
    
    SexoSelect(1) = JsonLanguage.item("FRM_CREARPJ_HOMBRE").item("TEXTO")
    SexoSelect(2) = JsonLanguage.item("FRM_CREARPJ_MUJER").item("TEXTO")
    
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
    
    If frmConnect.txtNombre.Visible = False Then frmConnect.txtNombre.Visible = True
    If frmConnect.txtPasswd.Visible = False Then frmConnect.txtPasswd.Visible = True
    If frmConnect.txtCrearPJNombre.Visible Then frmConnect.txtCrearPJNombre.Visible = False
    
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
    
    If frmConnect.txtNombre.Visible Then frmConnect.txtNombre.Visible = False
    If frmConnect.txtPasswd.Visible Then frmConnect.txtPasswd.Visible = False
    If frmConnect.txtCrearPJNombre.Visible Then frmConnect.txtCrearPJNombre.Visible = False
    
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
    Dim i As Byte
    
    'Seteamos el modo login
    Pantalla = PCrearPJ
    
    If Mostrar = True Then frmConnect.Visible = True
    
    If frmConnect.txtNombre.Visible Then frmConnect.txtNombre.Visible = False
    If frmConnect.txtPasswd.Visible Then frmConnect.txtPasswd.Visible = False
    If frmConnect.txtCrearPJNombre.Visible = False Then frmConnect.txtCrearPJNombre.Visible = True
    
    'Seteamos todos los valores
    UserSexo = Hombre
    UserName = vbNullString
    UserRaza = 0
    UserClase = 0
    lblTotal = 40
    
    For i = 1 To NUMATRIBUTES
        lblAtributos(i) = 6
    Next i
    
    Call DarCuerpoYCabeza
    Call LoadCharInfo

    'Focus al nombre del PJ y lo reseteamos
    frmConnect.txtCrearPJNombre.SetFocus
    frmConnect.txtCrearPJNombre.Text = vbNullString
    
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
                If .Particle_Group_Index Then _
                    Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
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
                  
            'Crear Cuenta
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
            
            If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then
                If TextSelected = 1 Then
                    Call Draw_GrhIndex(25319, 446 + Engine_AnchoTexto(1, frmConnect.txtNombre.Text), 372, 0, Normal_RGBList(), 0, False)
                    
                ElseIf TextSelected = 2 Then
                    Call Draw_GrhIndex(25319, 446 + Engine_AnchoTexto(1, frmConnect.txtPasswd.Text), 406, 0, Normal_RGBList(), 0, False)
                    
                End If
            End If
            
            'Password
            Call DrawText(445, 406, frmConnect.txtPasswd.Text, -1, False)
            
            'Salir
            Call Draw_GrhIndex(31503, 20, 650, 0, Normal_RGBList(), 0, False)
            
        Case 1 'Cuenta
        
            'Marco
            Call Draw_GrhIndex(31481, 0, 0, 0, Normal_RGBList(), 0, False)
            
            'Crear PJ
            Call Draw_GrhIndex(31501, 785, 510, 0, Normal_RGBList(), 0, False)
            
            'Borrar PJ
            Call Draw_GrhIndex(31495, 785, 580, 0, Normal_RGBList(), 0, False)
            
            'Gestion de cuentas
            Call Draw_GrhIndex(31497, 785, 650, 0, Normal_RGBList(), 0, False)
            
            'Desconectar
            Call Draw_GrhIndex(31494, 20, 650, 0, Normal_RGBList(), 0, False)
            
        Case 2 'Crear PJ
            'Marco
            Call Draw_GrhIndex(31481, 0, 0, 0, Normal_RGBList(), 0, False)
            
            'Nombre
            Call Draw_GrhIndex(31482, 350, 650, 0, Normal_RGBList(), 0, False)
            Call DrawText(400, 670, frmConnect.txtCrearPJNombre.Text, -1, False)
            
            If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then _
                Call Draw_GrhIndex(25319, 400 + Engine_AnchoTexto(1, frmConnect.txtCrearPJNombre.Text), 670, 0, Normal_RGBList(), 0, False)
            
            'Volver
            Call Draw_GrhIndex(31500, 20, 650, 0, Normal_RGBList(), 0, False)
            
            'Crear Personaje
            If botonCrear = False Then Call Draw_GrhIndex(31501, 785, 650, 0, Normal_RGBList(), 0, False)
            
            'Seleccion de Sexo
            Call Draw_GrhIndex(31505, 350, 300, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31485, 377, 312, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 600, 312, 0, Normal_RGBList(), 0, False)
            If UserSexo <> 0 Then Call DrawText(505, 320, SexoSelect(UserSexo), -1, True)
            
            'Seleccion de Raza
            Call Draw_GrhIndex(31505, 350, 350, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31485, 377, 362, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 600, 362, 0, Normal_RGBList(), 0, False)
            If UserRaza <> 0 Then Call DrawText(505, 370, ListaRazas(UserRaza), -1, True)
            
            'Seleccion de Clase
            Call Draw_GrhIndex(31505, 350, 400, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31485, 377, 412, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 600, 412, 0, Normal_RGBList(), 0, False)
            If UserClase <> 0 Then Call DrawText(505, 420, ListaClases(UserClase), -1, True)
            
            'Atributos
            Call Engine_Draw_Box(730, 250, 250, 300, D3DColorARGB(100, 0, 0, 0))
            
            'Fuerza
            Call DrawText(780, 285, "Fuerza:", -1, True)
            Call Draw_GrhIndex(31485, 810, 275, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 880, 275, 0, Normal_RGBList(), 0, False)
            Call DrawText(865, 285, lblAtributos(eAtributos.Fuerza), -1, True)
            Call DrawText(940, 285, lblModRaza(eAtributos.Fuerza), -1, True) '
            
            'Agilidad
            Call DrawText(780, 320, "Agilidad:", -1, True)
            Call Draw_GrhIndex(31485, 810, 310, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 880, 310, 0, Normal_RGBList(), 0, False)
            Call DrawText(865, 320, lblAtributos(eAtributos.Agilidad), -1, True)
            Call DrawText(940, 320, lblModRaza(eAtributos.Agilidad), -1, True)
            
            'Inteligencia
            Call DrawText(780, 363, "Inteligencia:", -1, True)
            Call Draw_GrhIndex(31485, 810, 353, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 880, 353, 0, Normal_RGBList(), 0, False)
            Call DrawText(865, 363, lblAtributos(eAtributos.Inteligencia), -1, True)
            Call DrawText(940, 363, lblModRaza(eAtributos.Inteligencia), -1, True)
            
            'Carisma
            Call DrawText(780, 400, "Carisma:", -1, True)
            Call Draw_GrhIndex(31485, 810, 395, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 880, 395, 0, Normal_RGBList(), 0, False)
            Call DrawText(865, 400, lblAtributos(eAtributos.Carisma), -1, True)
            Call DrawText(940, 400, lblModRaza(eAtributos.Carisma), -1, True)
            
            'Constitucion
            Call DrawText(780, 440, "Constitucion:", -1, True)
            Call Draw_GrhIndex(31485, 810, 440, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 880, 440, 0, Normal_RGBList(), 0, False)
            Call DrawText(865, 440, lblAtributos(eAtributos.Constitucion), -1, True)
            Call DrawText(940, 440, lblModRaza(eAtributos.Constitucion), -1, True)
            
            'Total
            Call DrawText(830, 500, lblTotal, -1, True)
            
            'Seleccion de Cabeza
            Call Draw_GrhIndex(31485, 188, 600, 0, Normal_RGBList(), 0, False)
            Call Draw_GrhIndex(31486, 258, 600, 0, Normal_RGBList(), 0, False)
            
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
                            Call Draw_Grh(HeadData(.Head).Head(3), PJPos(Index).X + BodyData(.Body).HeadOffset.X, PJPos(Index).Y + HeadData(.Head).Offset.Y + BodyData(.Body).HeadOffset.Y, 1, Normal_RGBList(), 0)
                        End If
            
                        If .helmet <> 0 Then
                            Call Draw_Grh(CascoAnimData(.helmet).Head(3), PJPos(Index).X + BodyData(.Body).HeadOffset.X, PJPos(Index).Y + CascoAnimData(.helmet).Offset.Y + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, 1, Normal_RGBList(), 0)
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
            
        Case 2 'Crear PJ
        
        If UserBody <> 0 Then
            Call Draw_Grh(BodyData(UserBody).Walk(3), 225, 560, 1, Normal_RGBList(), 0)
                
            If UserHead <> 0 Then _
                Call Draw_Grh(HeadData(UserHead).Head(3), 225 + BodyData(UserBody).HeadOffset.X, 560 + HeadData(UserHead).Offset.Y + BodyData(UserBody).HeadOffset.Y, 1, Normal_RGBList(), 0)
                
            'Nombre
            Call DrawText(225 + 16, 560 + 30, frmConnect.txtCrearPJNombre.Text, -1, True)
            
        End If
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
Debug.Print TX & " - " & TY
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
            
            'Conectar
            If (TX >= 403 And TX <= 513) And (TY >= 444 And TY <= 480) Then Call btnConectar
            
            'Teclas
            If (TX >= 543 And TX <= 647) And (TY >= 444 And TY <= 480) Then Call btnTeclas
            
            'Crear Cuenta
            If (TX >= 823 And TX <= 985) And (TY >= 469 And TY <= 509) Then Call btnGestion
            
            'Recuperar
            If (TX >= 823 And TX <= 985) And (TY >= 547 And TY <= 583) Then Call btnGestion
            
            'Salir
            If (TX >= 28 And TX <= 218) And (TY >= 656 And TY <= 714) Then Call CloseClient
        
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

            If (TX >= 795 And TX <= 987) And (TY >= 659 And TY <= 713) Then Call btnGestion

            'Desconectar
            If (TX >= 28 And TX <= 218) And (TY >= 656 And TY <= 714) Then
                frmMain.Client.CloseSck
                Call ResetAllInfoAccounts
                Call MostrarConnect
            End If
            
        Case 2 'Crear PJ
        
            'Volver
            If (TX >= 28 And TX <= 218) And (TY >= 656 And TY <= 714) Then
                Call Audio.PlayBackgroundMusic("2", MusicTypes.MP3)
                Call MostrarCuenta
            End If
            
            'SexoAnterior
            If (TX >= 338 And TX <= 406) And (TY >= 313 And TY <= 339) Then
                If UserSexo > 1 Then
                    UserSexo = UserSexo - 1
                    Call DarCuerpoYCabeza
                End If
            End If
                
            'SexoSiguiente
            If (TX >= 602 And TX <= 630) And (TY >= 311 And TY <= 339) Then
                If UserSexo < 2 Then
                    UserSexo = UserSexo + 1
                    Call DarCuerpoYCabeza
                End If
            End If
                
            'RazaAnterior
            If (TX >= 380 And TX <= 406) And (TY >= 361 And TY <= 389) Then
                If UserRaza > 1 Then
                    UserRaza = UserRaza - 1
                    Call DarCuerpoYCabeza
                    Call UpdateRazaMod
                End If
            End If
                
            'RazaSiguiente
            If (TX >= 604 And TX <= 630) And (TY >= 363 And TY <= 391) Then
                If UserRaza < NUMRAZAS Then
                    UserRaza = UserRaza + 1
                    Call DarCuerpoYCabeza
                    Call UpdateRazaMod
                End If
            End If
                
            'ClaseAnterior
            If (TX >= 380 And TX <= 406) And (TY >= 413 And TY <= 437) Then
                If UserClase > 1 Then
                    UserClase = UserClase - 1
                    Call DarCuerpoYCabeza
                End If
            End If
                
            'ClaseSiguiente
            If (TX >= 604 And TX <= 632) And (TY >= 413 And TY <= 437) Then
                If UserClase < NUMCLASES Then
                    UserClase = UserClase + 1
                    Call DarCuerpoYCabeza
                End If
            End If
            
            'Fuerza
            If (TX >= 813 And TX <= 841) And (TY >= 276 And TY <= 302) Then Call Menos_Click(eAtributos.Fuerza)
            If (TX >= 883 And TX <= 911) And (TY >= 276 And TY <= 302) Then Call Mas_Click(eAtributos.Fuerza)
            
            'Agilidad
            If (TX >= 813 And TX <= 841) And (TY >= 310 And TY <= 340) Then Call Menos_Click(eAtributos.Agilidad)
            If (TX >= 883 And TX <= 911) And (TY >= 310 And TY <= 340) Then Call Mas_Click(eAtributos.Agilidad)
            
            'Inteligencia
            If (TX >= 813 And TX <= 841) And (TY >= 352 And TY <= 380) Then Call Menos_Click(eAtributos.Inteligencia)
            If (TX >= 883 And TX <= 911) And (TY >= 354 And TY <= 378) Then Call Mas_Click(eAtributos.Inteligencia)
            
            'Carisma
            If (TX >= 813 And TX <= 841) And (TY >= 399 And TY <= 421) Then Call Menos_Click(eAtributos.Carisma)
            If (TX >= 883 And TX <= 911) And (TY >= 397 And TY <= 423) Then Call Mas_Click(eAtributos.Carisma)
            
            'Constitucion
            If (TX >= 813 And TX <= 841) And (TY >= 442 And TY <= 468) Then Call Menos_Click(eAtributos.Constitucion)
            If (TX >= 883 And TX <= 911) And (TY >= 442 And TY <= 466) Then Call Mas_Click(eAtributos.Constitucion)
            
            'Crear PJ
            If (TX >= 793 And TX <= 989) And (TY >= 656 And TY <= 710) Then _
                If botonCrear = False Then Call btnCrear
                
            'Nombre del PJ
            If (TX >= 379 And TX <= 625) And (TY >= 659 And TY <= 689) Then _
                frmConnect.txtCrearPJNombre.SetFocus
                
            'Cabezas
            If (TX >= 192 And TX <= 228) And (TY >= 600 And TY <= 628) Then Call btnHeadPJ(1) 'Menos
            If (TX >= 262 And TX <= 290) And (TY >= 600 And TY <= 628) Then Call btnHeadPJ(0) 'Mas
            
            
    End Select
    
End Sub

'<<<<<----------------------------BOTONES-------------------------------->>>>>>

Private Sub CrearNuevoPJ()
'**************************************
'Autor: Lorwik
'Fecha: 20/05/2020
'Descripcion: Boton de crear personaje
'**************************************
    Call Audio.PlayWave(SND_CLICK)

    If NumberOfCharacters > 9 Then
        Call MostrarMensaje(JsonLanguage.item("ERROR_DEMASIADOS_PJS").item("TEXTO"))
        Exit Sub
    End If
    
    'frmCrearPersonaje.Show
    Call MostrarCreacion
End Sub

Private Sub btnConectar()
'**************************************
'Autor: Lorwik
'Fecha: 23/05/2020
'Descripcion: Boton de conectar cuenta
'**************************************
    Call Audio.PlayWave(SND_CLICK)

    'update user info
    AccountName = frmConnect.txtNombre.Text
    AccountPassword = frmConnect.txtPasswd.Text

    'Clear spell list
    frmMain.hlst.Clear

    If frmConnect.chkRecordar.Checked = False Then
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Remember", "False")
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "UserName", vbNullString)
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Password", vbNullString)
    Else
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Remember", "True")
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "UserName", AccountName)
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Password", Cripto.AesEncryptString(AccountPassword, AES_PASSWD))
    End If

    If CheckUserData() = True Then
        Call Protocol.Connect(E_MODO.Normal)
    End If
End Sub

Private Sub btnTeclas()
'**************************************
'Autor: Lorwik
'Fecha: 20/05/2020
'Descripcion: Boton de Teclas anti Keylogger
'**************************************
    Load frmKeypad
    frmKeypad.Show vbModal
    Unload frmKeypad
    frmConnect.txtPasswd.SetFocus
End Sub

Private Sub btnGestion()
'**************************************
'Autor: Lorwik
'Fecha: 20/05/2020
'Descripcion: Boton de gestion de cuentas
'**************************************
    Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "http://winterao.com.ar/", "", App.path, SW_SHOWNORMAL)
    
End Sub

Private Sub Mas_Click(ByVal Index As Integer)
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Resta un atributo de la asignacion
'**************************************

    Call Audio.PlayWave(SND_CLICK)
    
    If lblAtributos(Index) < 18 And lblTotal > 0 Then
        lblAtributos(Index) = lblAtributos(Index) + 1
        lblTotal = lblTotal - 1
    End If
    
End Sub

Private Sub Menos_Click(ByVal Index As Integer)
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Suma un atributo de la asignacion
'**************************************

    Call Audio.PlayWave(SND_CLICK)
    
    If lblTotal = "40" Then Exit Sub
    If lblAtributos(Index) > 6 Then
        lblAtributos(Index) = lblAtributos(Index) - 1
        lblTotal = lblTotal + 1
    End If
    
End Sub

Private Sub btnHeadPJ(ByVal Index As Integer)

    Select Case Index

        Case 0
            UserHead = CheckCabeza(UserHead + 1)

        Case 1
            UserHead = CheckCabeza(UserHead - 1)

    End Select
    
End Sub

Private Sub btnCrear()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Mandamos la creacion del personaje
'**************************************

    Dim i As Integer
    
    'Nombre de usuario
    UserName = frmConnect.txtCrearPJNombre.Text
            
    '¿El nombre esta vacio y es correcto?
    If Right$(UserName, 1) = " " Then
        UserName = RTrim$(UserName)
       Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(2))

    End If
    
    'Atributos asignados
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Val(lblAtributos(i))
    Next i
    
    'Comprobamos que todo este OK
    If Not CheckData Then Exit Sub
    
    EstadoLogin = E_MODO.CrearNuevoPJ
    
    'Limpio la lista de hechizos
    frmMain.hlst.Clear
        
    'Conexion!!!
    If Not frmMain.Client.State = sckConnected Then
        Call MostrarMensaje(JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO"))
        Call MostrarCuenta
    Else
        'Si ya mandamos el paquete, evitamos que se pueda volver a mandar
        botonCrear = True
        Call Login
        botonCrear = False
    End If
    
    'Mandamos el tutorial de inicio
    bShowTutorial = True

End Sub

'<<<<<--------------------------------------------------------------------->>>>>>
'CREACION DE PJ

Private Sub DarCuerpoYCabeza()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Asignamos un cuerpo y unac abeza segun la raza y el sexo
'**************************************

    Select Case UserSexo
    
        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano
                    UserHead = eCabezas.HUMANO_H_PRIMER_CABEZA
                    UserBody = eCabezas.HUMANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = eCabezas.ELFO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ELFO_H_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = eCabezas.DROW_H_PRIMER_CABEZA
                    UserBody = eCabezas.DROW_H_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = eCabezas.ENANO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ENANO_H_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = eCabezas.GNOMO_H_PRIMER_CABEZA
                    UserBody = eCabezas.GNOMO_H_CUERPO_DESNUDO
                    
                Case eRaza.Orco
                    UserHead = eCabezas.ORCO_H_PRIMER_CABEZA
                    UserBody = eCabezas.ORCO_H_CUERPO_DESNUDO
                    
                Case eRaza.Vampiro
                    UserHead = eCabezas.VAMPIRO_H_PRIMER_CABEZA
                    UserBody = eCabezas.VAMPIRO_H_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano
                    UserHead = eCabezas.HUMANO_M_PRIMER_CABEZA
                    UserBody = eCabezas.HUMANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Elfo
                    UserHead = eCabezas.ELFO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ELFO_M_CUERPO_DESNUDO
                    
                Case eRaza.ElfoOscuro
                    UserHead = eCabezas.DROW_M_PRIMER_CABEZA
                    UserBody = eCabezas.DROW_M_CUERPO_DESNUDO
                    
                Case eRaza.Enano
                    UserHead = eCabezas.ENANO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ENANO_M_CUERPO_DESNUDO
                    
                Case eRaza.Gnomo
                    UserHead = eCabezas.GNOMO_M_PRIMER_CABEZA
                    UserBody = eCabezas.GNOMO_M_CUERPO_DESNUDO
                    
                Case eRaza.Orco
                    UserHead = eCabezas.ORCO_M_PRIMER_CABEZA
                    UserBody = eCabezas.ORCO_M_CUERPO_DESNUDO
                    
                Case eRaza.Vampiro
                    UserHead = eCabezas.VAMPIRO_M_PRIMER_CABEZA
                    UserBody = eCabezas.VAMPIRO_M_CUERPO_DESNUDO
                    
                Case Else
                    UserHead = 0
                    UserBody = 0
            End Select
            
        Case Else
            UserHead = 0
            UserBody = 0
            
    End Select
    
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

On Error GoTo errhandler

    Select Case UserSexo

        Case eGenero.Hombre

            Select Case UserRaza

                Case eRaza.Humano

                    If Head > eCabezas.HUMANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_H_PRIMER_CABEZA + (Head - eCabezas.HUMANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.HUMANO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_H_ULTIMA_CABEZA - (eCabezas.HUMANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Elfo

                    If Head > eCabezas.ELFO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_H_PRIMER_CABEZA + (Head - eCabezas.ELFO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ELFO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_H_ULTIMA_CABEZA - (eCabezas.ELFO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.ElfoOscuro

                    If Head > eCabezas.DROW_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.DROW_H_PRIMER_CABEZA + (Head - eCabezas.DROW_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.DROW_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.DROW_H_ULTIMA_CABEZA - (eCabezas.DROW_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Enano

                    If Head > eCabezas.ENANO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_H_PRIMER_CABEZA + (Head - eCabezas.ENANO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ENANO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_H_ULTIMA_CABEZA - (eCabezas.ENANO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Gnomo

                    If Head > eCabezas.GNOMO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_H_PRIMER_CABEZA + (Head - eCabezas.GNOMO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.GNOMO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_H_ULTIMA_CABEZA - (eCabezas.GNOMO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Orco

                    If Head > eCabezas.ORCO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_H_PRIMER_CABEZA + (Head - eCabezas.ORCO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ORCO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_H_ULTIMA_CABEZA - (eCabezas.ORCO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Vampiro

                    If Head > eCabezas.VAMPIRO_H_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.VAMPIRO_H_PRIMER_CABEZA + (Head - eCabezas.VAMPIRO_H_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.VAMPIRO_H_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.VAMPIRO_H_ULTIMA_CABEZA - (eCabezas.VAMPIRO_H_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case Else
                    CheckCabeza = CheckCabeza(Head)
                    
            End Select
        
        Case eGenero.Mujer

            Select Case UserRaza

                Case eRaza.Humano

                    If Head > eCabezas.HUMANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_M_PRIMER_CABEZA + (Head - eCabezas.HUMANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.HUMANO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.HUMANO_M_ULTIMA_CABEZA - (eCabezas.HUMANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Elfo

                    If Head > eCabezas.ELFO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_M_PRIMER_CABEZA + (Head - eCabezas.ELFO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ELFO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ELFO_M_ULTIMA_CABEZA - (eCabezas.ELFO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.ElfoOscuro

                    If Head > eCabezas.DROW_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.DROW_M_PRIMER_CABEZA + (Head - eCabezas.DROW_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.DROW_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.DROW_M_ULTIMA_CABEZA - (eCabezas.DROW_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Enano

                    If Head > eCabezas.ENANO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_M_PRIMER_CABEZA + (Head - eCabezas.ENANO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ENANO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ENANO_M_ULTIMA_CABEZA - (eCabezas.ENANO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Gnomo

                    If Head > eCabezas.GNOMO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_M_PRIMER_CABEZA + (Head - eCabezas.GNOMO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.GNOMO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.GNOMO_M_ULTIMA_CABEZA - (eCabezas.GNOMO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case eRaza.Orco

                    If Head > eCabezas.ORCO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_M_PRIMER_CABEZA + (Head - eCabezas.ORCO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.ORCO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.ORCO_M_ULTIMA_CABEZA - (eCabezas.ORCO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                    
                Case eRaza.Vampiro

                    If Head > eCabezas.VAMPIRO_M_ULTIMA_CABEZA Then
                        CheckCabeza = eCabezas.VAMPIRO_M_PRIMER_CABEZA + (Head - eCabezas.VAMPIRO_M_ULTIMA_CABEZA) - 1
                    ElseIf Head < eCabezas.VAMPIRO_M_PRIMER_CABEZA Then
                        CheckCabeza = eCabezas.VAMPIRO_M_ULTIMA_CABEZA - (eCabezas.VAMPIRO_M_PRIMER_CABEZA - Head) + 1
                    Else
                        CheckCabeza = Head
                    End If
                
                Case Else
                    CheckCabeza = Head
                    
            End Select

        Case Else
            CheckCabeza = Head
            
    End Select
    
errhandler:

    If Err.number Then
        Call LogError(Err.number, Err.Description, "frmCrearPersonaje.CheckCabeza")
    End If
    
    Exit Function
    
End Function

Public Sub UpdateRazaMod()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Actualiza los modificadores de atributos que otorga cada raza
'**************************************

    If UserRaza > -1 Then
        
        With ModRaza(UserRaza)
            lblModRaza(eAtributos.Fuerza) = IIf(.Fuerza >= 0, "+", vbNullString) & .Fuerza
            lblModRaza(eAtributos.Agilidad) = IIf(.Agilidad >= 0, "+", vbNullString) & .Agilidad
            lblModRaza(eAtributos.Inteligencia) = IIf(.Inteligencia >= 0, "+", vbNullString) & .Inteligencia
            lblModRaza(eAtributos.Carisma) = IIf(.Carisma >= 0, "+", "") & .Carisma
            lblModRaza(eAtributos.Constitucion) = IIf(.Constitucion >= 0, "+", vbNullString) & .Constitucion
        End With
        
    End If
    
End Sub

Private Sub LoadCharInfo()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Carga los modificadores de cada raza
'**************************************

    Dim SearchVar As String
    Dim i         As Integer

    ReDim ModRaza(1 To NUMRAZAS)

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(Game.path(INIT) & "CharInfo_" & Language & ".dat")
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
    
        With ModRaza(i)
            SearchVar = Replace(ListaRazas(i), " ", vbNullString)
        
            .Fuerza = CSng(Lector.GetValue("MODRAZA", SearchVar + "Fuerza"))
            .Agilidad = CSng(Lector.GetValue("MODRAZA", SearchVar + "Agilidad"))
            .Inteligencia = CSng(Lector.GetValue("MODRAZA", SearchVar + "Inteligencia"))
            .Carisma = CSng(Lector.GetValue("MODRAZA", SearchVar + "Carisma"))
            .Constitucion = CSng(Lector.GetValue("MODRAZA", SearchVar + "Constitucion"))
        End With
        
    Next i

End Sub

Private Function CheckData() As Boolean
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Comprobacion antes de crear el PJ
'**************************************
    
    Dim i As Integer
    Dim Suma As Byte
    
    '¿Puso un nombre?
    If LenB(frmConnect.txtCrearPJNombre.Text) = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_NOMBRE_PJ").item("TEXTO"))
        frmConnect.txtCrearPJNombre.SetFocus
        Exit Function
    End If

    '¿Selecciono una raza?
    If UserRaza = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_RAZA").item("TEXTO"))
        Exit Function
    End If
    
    '¿Selecciono el Sexo?
    If UserSexo = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_SEXO").item("TEXTO"))
        Exit Function
    End If
    
    '¿Seleciono la clase?
    If UserClase = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_CLASE").item("TEXTO"))
        Exit Function
    End If

    '¿Estamos intentando crear sin tener el AccountHash?
    If Len(AccountHash) = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_HASH").item("TEXTO"))
        Exit Function
    End If

    'Sumamos los atributos asignados
    For i = 1 To NUMATRIBUTOS
        If Val(UserAtributos(i)) > 18 Then
            Call MostrarMensaje(JsonLanguage.item("VALIDACION_ATRIBUTOS").item("TEXTO"))
            Exit Function
        End If
        
        Suma = Suma + UserAtributos(i)
    Next i

    '¿Los atributos asignados son validos?
    If Suma <> 70 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_ATRIBUTOS").item("TEXTO"))
        Exit Function
    End If
    
    '¿El nombre de usuario supera los 30 caracteres?
    If LenB(UserName) > 30 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(1))
        Exit Function
    End If
    
    CheckData = True

End Function
