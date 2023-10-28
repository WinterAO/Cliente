Attribute VB_Name = "ModConectar"
'********************************************************************************************
'Autor: Lorwik
'Fecha 20/05/2020
'Descripcion: En este modulo vamos a renderizar todos los referente con
'el conectar, cuentas y crearpj, asi como eventos.
'
'NOTA: No esta programado de la forma mas eficiente, pero por el momento es funcional
'en el futuro se deberia revisar y mejorar en todo lo posible.
'********************************************************************************************

Option Explicit

'Conectar renderizado
Private Type tMapaConnect
    Map As Integer
    X As Integer
    Y As Integer
End Type

Public MapaConnect() As tMapaConnect
Public NumConnectMap As Byte 'Numero total de mapas cargados

'Indica que mapa vamos a renderizar en el conectar
Public SelectConnectMap As Byte

'******************************
'Modo de pantalla renderizado
'******************************
Enum EPantalla
    PConnect = 0
    PCuenta
    PCrearPJ
End Enum

Public Pantalla As EPantalla
'*********************

Private Type tButtonsGUI
    X As Integer
    Y As Integer
    PosX As Integer
    PosY As Integer
    GrhNormal As Long
    GrhClarito As Long
   ' color(0 To 3) As Long
End Type

Private ButtonGUI() As tButtonsGUI

Private Const MAXPJACCOUNTS As Byte = 10
'Indica la posicion donde se va a renderizar los PJ
Private PJPos(1 To MAXPJACCOUNTS) As WorldPos

'Indifca el PJ seleccionado
Public PJAccSelected As Byte

'Indica el TextBox selecionado
Private TextSelected As Byte

Private GRHFX_PJ_Selecionado As Grh
Private Const FX_PJ_Seleccionado As Long = 13181

'*********************
'Flags
'*********************
Public Conectando As Boolean 'Para evitar mandar varias peticiones al servidor a la hora de conectar
Private botonCrear As Boolean
'*********************

'*********************
'Creacion de PJ
'*********************
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

Private SexoSelect(1 To 2) As String 'Sexo seleccionado
'**********************

Public Sub InicializarRndCNT()
'********************************************
'Autor: Lorwik
'Fecha: 19/05/2020
'Descripcion: Inicia todo lo que tiene que ver con el conectar renderizado
'como las posiciones donde se van a mostrar los PJ, pantalla, etc...
'********************************************
    
    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim fileBuff  As clsByteBuffer
    Dim i As Integer
    Dim j As Byte
    Dim NumButtons As Integer

    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("GUI.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("GUI.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer

        'INIT
        NumButtons = fileBuff.getInteger
        NumConnectMap = fileBuff.getByte
        ReDim Preserve MapaConnect(NumConnectMap) As tMapaConnect
        
        'Mapas de GUI
        For i = 1 To NumConnectMap
            MapaConnect(i).Map = fileBuff.getInteger
            MapaConnect(i).X = fileBuff.getInteger
            MapaConnect(i).Y = fileBuff.getInteger
        Next i
    
        'Posiciones de los PJ
        For i = 1 To MAXPJACCOUNTS
            PJPos(i).X = fileBuff.getInteger
            PJPos(i).Y = fileBuff.getInteger
        Next i
        
        ReDim ButtonGUI(1 To NumButtons) As tButtonsGUI
        
        'Posiciones de los botones
        For i = 1 To NumButtons
            With ButtonGUI(i)
                .X = fileBuff.getInteger
                .Y = fileBuff.getInteger
                .PosX = fileBuff.getInteger
                .PosY = fileBuff.getInteger
                .GrhNormal = fileBuff.getLong
                
                'For j = 0 To 3
                '    .color(j) = COLOR_WHITE(j)
                'Next j
            End With
        Next i
        
        Set fileBuff = Nothing
        
        Pantalla = PConnect 'Establecemos la pantalla en el conectar
        TextSelected = 1 ' Establecemos el cursor de texto en Nombre
        
        SexoSelect(1) = JsonLanguage.item("FRM_CREARPJ_HOMBRE").item("TEXTO")
        SexoSelect(2) = JsonLanguage.item("FRM_CREARPJ_MUJER").item("TEXTO")
        
        Call InitGrh(GRHFX_PJ_Selecionado, FX_PJ_Seleccionado)
        
    Else
    
        Call MostrarMensaje("No se ha podido inicializar la GUI, si el problema persiste reinstale el juego.")
        Call CloseClient
    End If
    
End Sub

Public Sub MapConnect(ByVal ConnectMap As Byte)
    
    Call SwitchMap(MapaConnect(ConnectMap).Map)

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
    
    If ClientSetup.Remember Then _
        frmConnect.txtNombre = ClientSetup.rUserName
    
    'frmConnect.txtNombre.SetFocus
    'frmConnect.txtNombre.SelStart = Len(frmConnect.txtNombre.Text)
    TextSelected = 1
    
    Battlegrounds = False
    
    EngineRun = False
    
    'LISTA DE SERVIDORES
    Call ListarServidores
    
    SelectConnectMap = RandomNumber(3, 4)

    Call MapConnect(SelectConnectMap)
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
    
    EngineRun = False
    
    SelectConnectMap = 2
    
    'Ponemos el mapa de cuentas
    Call MapConnect(SelectConnectMap)

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
    CurrentUser.UserSexo = Hombre
    CurrentUser.UserName = vbNullString
    CurrentUser.UserRaza = 0
    CurrentUser.UserClase = 0
    
    Call DarCuerpoYCabeza
    Call LoadCharInfo

    'Focus al nombre del PJ y lo reseteamos
    'frmConnect.txtCrearPJNombre.SetFocus
    'frmConnect.txtCrearPJNombre.Text = vbNullString
    'frmConnect.txtCrearPJNombre.SelStart = Len(frmConnect.txtCrearPJNombre.Text)
    
    EngineRun = False
    
    SelectConnectMap = 1
    
    'Ponemos el mapa de cuentas
    Call MapConnect(SelectConnectMap)

End Sub

Sub RenderConnect()
'******************************
'Autor: Lorwik
'Fecha: 15/05/2020
'Renderiza el screen del conectar
'******************************
On Error GoTo ErrorHandler:

    Dim X As Long
    Dim Y As Long
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    Static RE As RECT
    
    With RE
        .Left = 0
        .Top = 0
        .Bottom = frmConnect.ScaleHeight
        .Right = frmConnect.ScaleWidth
    End With
    
    Movement_Speed = 1
    
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
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
        
    Call Engine_EndScene(RE, frmConnect.Renderer.hWnd)
    
ErrorHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        
        Call mDx8_Engine.Engine_DirectX8_Init
        
        Call LoadGraphics
    
    End If
  
End Sub

Private Sub RenderConnectGUI()
'******************************
'Autor: Lorwik
'Fecha: 15/05/2020
'Dibuja la interfaz
'******************************
    Dim asterisco As String
    Dim i As Integer
    
    Select Case Pantalla
    
        Case 0 'Login (frmconnect)
            
            For i = 1 To 11
                With ButtonGUI(i)
                    Call Draw_GrhIndex(.GrhNormal, .X, .Y, 0, COLOR_WHITE, 0, False)
                End With
            Next i
            
            'Server
            Call DrawText(480, 340, Servidor(ServIndSel).Nombre, COLOR_WHITE, False)
            
        Case 1 'Cuenta
        
            'Marco
            Call Draw_GrhIndex(ButtonGUI(2).GrhNormal, ButtonGUI(2).X, ButtonGUI(2).Y, 0, COLOR_WHITE, 0, False)
            
            For i = 12 To 15
                With ButtonGUI(i)
                     Call Draw_GrhIndex(.GrhNormal, .X, .Y, 0, COLOR_WHITE, 0, False)
                End With
            Next i
            
            If CurrentUser.esVIP Then _
                Call DrawText(490, 680, "Cuenta VIP hasta el " & CurrentUser.VIP, COLOR_WHITE, True)
            
            'Conectando
            If ModConectar.Conectando = False Then _
                Call DrawText(490, 620, "Conectando...", COLOR_WHITE, True, 2)

        Case 2 'Crear PJ
        
            'Marco
            Call Draw_GrhIndex(ButtonGUI(2).GrhNormal, ButtonGUI(2).X, ButtonGUI(2).Y, 0, COLOR_WHITE, 0, False)
            
            For i = 16 To 28
                With ButtonGUI(i)
                    Call Draw_GrhIndex(.GrhNormal, .X, .Y, 0, COLOR_WHITE, 0, False)
                End With
            Next i
            
            'If timeGetTime Mod CursorFlashRate * 2 < CursorFlashRate Then _
                Call Draw_GrhIndex(25319, 400 + Engine_AnchoTexto(1, frmConnect.txtCrearPJNombre.Text), 670, 0, .Color(), 0, False)

            'Crear Personaje
            If botonCrear = False Then Call Draw_GrhIndex(ButtonGUI(29).GrhNormal, ButtonGUI(29).X, ButtonGUI(29).Y, 0, COLOR_WHITE, 0, False)
            
            'Textos
            Call DrawText(400, 670, frmConnect.txtCrearPJNombre.Text, COLOR_WHITE, False)
            If CurrentUser.UserSexo <> 0 Then Call DrawText(505, 320, SexoSelect(CurrentUser.UserSexo), COLOR_WHITE, True)
            If CurrentUser.UserRaza <> 0 Then Call DrawText(505, 370, ListaRazas(CurrentUser.UserRaza), COLOR_WHITE, True)
            If CurrentUser.UserClase <> 0 Then Call DrawText(505, 420, ListaClases(CurrentUser.UserClase), COLOR_WHITE, True)

            
            Call Draw_GrhIndex(ButtonGUI(30).GrhNormal, ButtonGUI(30).X, ButtonGUI(30).Y, 0, COLOR_WHITE, 0, False)
            Call DrawText(900, 327, lblModRaza(eAtributos.Fuerza), COLOR_WHITE, True) '
            Call DrawText(900, 364, lblModRaza(eAtributos.Agilidad), COLOR_WHITE, True)
            Call DrawText(900, 400, lblModRaza(eAtributos.Inteligencia), COLOR_WHITE, True)
            Call DrawText(900, 438, lblModRaza(eAtributos.Carisma), COLOR_WHITE, True)
            Call DrawText(900, 477, lblModRaza(eAtributos.Constitucion), COLOR_WHITE, True)
            
    End Select
    
    '<------- Desde aqui lo que siempre se va a mostrar ------->
    
    ' Calculamos los FPS y los mostramos
    Call Engine_Update_FPS
    'If ClientSetup.FPSShow = True Then
    Call DrawText(970, 30, "FPS: " & Mod_TileEngine.FPS, COLOR_WHITE, True)
    
    Call DrawText(25, 730, "WinterAO " & GetVersionOfTheGame() & " Resurrection", COLOR_WHITE)
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
            For Index = 1 To CurrentUser.NumberOfCharacters
                With cPJ(Index)
    
                    If .Body <> 0 Then
            
                        If PJAccSelected = Index Then Call Draw_Grh(GRHFX_PJ_Selecionado, PJPos(Index).X, PJPos(Index).Y + 60, 1, COLOR_WHITE(), 1, True)
                        
                        Call Draw_Grh(BodyData(.Body).Walk(1), PJPos(Index).X, PJPos(Index).Y, 1, COLOR_WHITE(), 0)
            
                        If .Head <> 0 Then _
                            Call DrawHead(.Head, PJPos(Index).X + BodyData(.Body).HeadOffset.X, PJPos(Index).Y + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, COLOR_WHITE(), 1, True)
            
                        If .helmet <> 0 Then _
                            Call DrawHead(.helmet, PJPos(Index).X + BodyData(.Body).HeadOffset.X, PJPos(Index).Y + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, COLOR_WHITE(), 1, False)
            
                        If .weapon <> 0 Then
                            Call Draw_Grh(WeaponAnimData(.weapon).WeaponWalk(1), PJPos(Index).X, PJPos(Index).Y, 1, COLOR_WHITE(), 0)
                        End If
            
                        If .shield <> 0 Then
                            Call Draw_Grh(ShieldAnimData(.shield).ShieldWalk(1), PJPos(Index).X, PJPos(Index).Y, 1, COLOR_WHITE(), 0)
                        End If
                        
                        'Nombre
                        Call DrawText(PJPos(Index).X + 16, PJPos(Index).Y + 30, .Nombre, COLOR_WHITE, True)
                        
                        'Nombre de la cuenta
                        Call DrawText(500, 25, CurrentUser.AccountName, COLOR_WHITE, True, 2)
                        
                        'Nombre del servidor
                        Call DrawText(30, 25, "Servidor " & Servidor(ServIndSel).Nombre, COLOR_WHITE, False)
                        
                    End If
                
                End With
            Next Index
            
        Case 2 'Crear PJ
        
        If CurrentUser.UserBody <> 0 Then
            Call Draw_Grh(BodyData(CurrentUser.UserBody).Walk(1), 225, 560, 1, COLOR_WHITE(), 0)
                
            If CurrentUser.UserHead <> 0 Then _
                Call DrawHead(CurrentUser.UserHead, 225 + BodyData(CurrentUser.UserBody).HeadOffset.X, 527 + BodyData(CurrentUser.UserBody).HeadOffset.Y, COLOR_WHITE(), 1, True)
                
            'Nombre
            'Call DrawText(225 + 16, 560 + 30, frmConnect.txtCrearPJNombre.Text, -1, True)
            
        End If
    End Select

End Sub

Public Sub DobleClickEvent(ByVal tX As Long, ByVal tY As Long)
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
            For i = 1 To CurrentUser.NumberOfCharacters
                With cPJ(i)
                    If (tX >= PJPos(i).X And tX <= PJPos(i).X + 20) And (tY >= PJPos(i).Y And tY <= PJPos(i).Y - OFFSET_HEAD) Then
    
                        If LenB(.Nombre) <> 0 Then
                            CurrentUser.UserName = .Nombre
                            Call ConnectPJ
                        End If
                        
                    End If
                End With
            Next i

    End Select
    
End Sub

Public Sub ClickEvent(ByVal tX As Long, ByVal tY As Long)
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'Eventos al realizar click en la GUI
'******************************
    Dim i As Integer

    Dim Index As Byte
    
    Select Case Pantalla
        Case 0 'Conectar

            'If (TX >= 443 And TX <= 605) And (TY >= 372 And TY <= 384) Then
            '    frmConnect.txtNombre.SetFocus
            '    frmConnect.txtNombre.SelStart = Len(frmConnect.txtNombre.Text)
            '    TextSelected = 1
            'End If
            
            'If (TX >= 443 And TX <= 605) And (TY >= 405 And TY <= 424) Then
            '    frmConnect.txtPasswd.SetFocus
            '    frmConnect.txtPasswd.SelStart = Len(frmConnect.txtPasswd.Text)
            '    TextSelected = 2
            'End If'
            
            'Servers
            If (tX >= ButtonGUI(9).X And tX <= ButtonGUI(9).PosX) And (tY >= ButtonGUI(9).Y And tY <= ButtonGUI(9).PosY) Then
                Call Sound.Sound_Play(SND_CLICK)
                If ServIndSel > LBound(Servidor()) Then ServIndSel = ServIndSel - 1
            End If
            
            If (tX >= ButtonGUI(10).X And tX <= ButtonGUI(10).PosX) And (tY >= ButtonGUI(10).Y And tY <= ButtonGUI(10).PosY) Then
                Call Sound.Sound_Play(SND_CLICK)
                If ServIndSel < UBound(Servidor()) Then ServIndSel = ServIndSel + 1
            End If
            
            'Conectar
            If (tX >= ButtonGUI(6).X And tX <= ButtonGUI(6).PosX) And (tY >= ButtonGUI(6).Y And tY <= ButtonGUI(6).PosY) Then Call btnConectar
            
            'Teclas
            If (tX >= ButtonGUI(7).X And tX <= ButtonGUI(7).PosX) And (tY >= ButtonGUI(7).Y And tY <= ButtonGUI(7).PosY) Then Call btnTeclas
            
            'Crear Cuenta
            If (tX >= ButtonGUI(4).X And tX <= ButtonGUI(4).PosX) And (tY >= ButtonGUI(4).Y And tY <= ButtonGUI(4).PosY) Then Call btnCrearCuenta
            
            'Recuperar
            If (tX >= ButtonGUI(5).X And tX <= ButtonGUI(5).PosX) And (tY >= ButtonGUI(5).Y And tY <= ButtonGUI(5).PosY) Then Call btnGestion
            
            'Salir
            If (tX >= ButtonGUI(11).X And tX <= ButtonGUI(11).PosX) And (tY >= ButtonGUI(11).Y And tY <= ButtonGUI(11).PosY) Then Call CloseClient
        
        Case 1 'Cuenta

            'Seleccionamos un PJ
            For i = 1 To CurrentUser.NumberOfCharacters
                With cPJ(i)
                    If (tX >= PJPos(i).X And tX <= PJPos(i).X + 20) And (tY >= PJPos(i).Y And tY <= PJPos(i).Y - OFFSET_HEAD) Then
    
                        If LenB(.Nombre) <> 0 Then
                            'El PJ seleccionado queda guardado
                            CurrentUser.UserName = .Nombre
                            PJAccSelected = i
                        End If
                    End If
                End With
            Next i
            
            'Crear Nuevo PJ
            If (tX >= ButtonGUI(12).X And tX <= ButtonGUI(12).PosX) And (tY >= ButtonGUI(12).Y And tY <= ButtonGUI(12).PosY) Then Call CrearNuevoPJ
            
            'Borrar PJ
            If (tX >= ButtonGUI(13).X And tX <= ButtonGUI(13).PosX) And (tY >= ButtonGUI(13).Y And tY <= ButtonGUI(13).PosY) Then
                If PJAccSelected < 1 Then
                    Call MostrarMensaje(JsonLanguage.item("ERROR_PERSONAJE_NO_SELECCIONADO").item("TEXTO"))
                    Exit Sub
                End If
                    
                frmBorrarPJ.Show
            
            End If

            If (tX >= ButtonGUI(14).X And tX <= ButtonGUI(14).PosX) And (tY >= ButtonGUI(14).Y And tY <= ButtonGUI(14).PosY) Then Call btnGestion

            'Desconectar
            If (tX >= ButtonGUI(15).X And tX <= ButtonGUI(15).PosX) And (tY >= ButtonGUI(15).Y And tY <= ButtonGUI(15).PosY) Then
                frmMain.Client.CloseSck
                Call ResetAllInfoAccounts
                Call MostrarConnect
            End If
            
        Case 2 'Crear PJ
        
            'Volver
            If (tX >= ButtonGUI(17).X And tX <= ButtonGUI(17).PosX) And (tY >= ButtonGUI(17).Y And tY <= ButtonGUI(17).PosY) Then
                Call Sound.Sound_Play(SND_CLICK)

                If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
                        Sound.NextMusic = MUS_VolverInicio
                        Sound.Fading = 200
                    End If
                End If
                
                Call MostrarCuenta
            End If
            
            'SexoAnterior <
            If (tX >= ButtonGUI(19).X And tX <= ButtonGUI(19).PosX) And (tY >= ButtonGUI(19).Y And tY <= ButtonGUI(19).PosY) Then
                If CurrentUser.UserSexo > 1 Then
                    CurrentUser.UserSexo = CurrentUser.UserSexo - 1
                    Call DarCuerpoYCabeza
                End If
            End If
                
            'SexoSiguiente >
            If (tX >= ButtonGUI(20).X And tX <= ButtonGUI(20).PosX) And (tY >= ButtonGUI(20).Y And tY <= ButtonGUI(20).PosY) Then
                If CurrentUser.UserSexo < 2 Then
                    CurrentUser.UserSexo = CurrentUser.UserSexo + 1
                    Call DarCuerpoYCabeza
                End If
            End If
                
            'RazaAnterior <
            If (tX >= ButtonGUI(22).X And tX <= ButtonGUI(22).PosX) And (tY >= ButtonGUI(22).Y And tY <= ButtonGUI(22).PosY) Then
                If CurrentUser.UserRaza > 1 Then
                    CurrentUser.UserRaza = CurrentUser.UserRaza - 1
                    Call DarCuerpoYCabeza
                    Call UpdateRazaMod
                End If
            End If
                
            'RazaSiguiente >
            If (tX >= ButtonGUI(23).X And tX <= ButtonGUI(23).PosX) And (tY >= ButtonGUI(23).Y And tY <= ButtonGUI(23).PosY) Then
                If CurrentUser.UserRaza < NUMRAZAS Then
                    CurrentUser.UserRaza = CurrentUser.UserRaza + 1
                    Call DarCuerpoYCabeza
                    Call UpdateRazaMod
                End If
            End If
                
            'ClaseAnterior <
            If (tX >= ButtonGUI(25).X And tX <= ButtonGUI(25).PosX) And (tY >= ButtonGUI(25).Y And tY <= ButtonGUI(25).PosY) Then
                If CurrentUser.UserClase > 1 Then
                    CurrentUser.UserClase = CurrentUser.UserClase - 1
                    Call DarCuerpoYCabeza
                End If
            End If
                
            'ClaseSiguiente >
            If (tX >= ButtonGUI(26).X And tX <= ButtonGUI(26).PosX) And (tY >= ButtonGUI(26).Y And tY <= ButtonGUI(26).PosY) Then
                If CurrentUser.UserClase < NUMCLASES Then
                    CurrentUser.UserClase = CurrentUser.UserClase + 1
                    Call DarCuerpoYCabeza
                End If
            End If
            
            'Crear PJ
            If (tX >= ButtonGUI(29).X And tX <= ButtonGUI(29).PosX) And (tY >= ButtonGUI(29).Y And tY <= ButtonGUI(29).PosY) Then _
                If botonCrear = False Then Call btnCrear
                
            'Nombre del PJ
            'If (TX >= 379 And TX <= 625) And (TY >= 659 And TY <= 689) Then
            '    frmConnect.txtCrearPJNombre.SetFocus
            '    frmConnect.txtCrearPJNombre.SelStart = Len(frmConnect.txtCrearPJNombre.Text)
            'End If
                
            'Cabezas
            If (tX >= ButtonGUI(27).X And tX <= ButtonGUI(27).PosX) And (tY >= ButtonGUI(27).Y And tY <= ButtonGUI(27).PosY) Then Call btnHeadPJ(1) 'Menos
            If (tX >= ButtonGUI(28).X And tX <= ButtonGUI(28).PosX) And (tY >= ButtonGUI(28).Y And tY <= ButtonGUI(28).PosY) Then Call btnHeadPJ(0) 'Mas
            
            
    End Select
    
End Sub

Public Sub MouseMove_Event(ByVal tX As Long, ByVal tY As Long)
    Dim i As Integer
    
    Select Case Pantalla
    
        Case 0 'Conectar
        
        Case 1 'Cuenta
        
        Case 2 'Crear PJ
        
    End Select
End Sub

Public Sub TeclaEvent(ByVal KeyCode As Integer)
'**************************************
'Autor: Lorwik
'Fecha: 19/06/2020
'Descripcion: Recibimos la pulsación de una tecla y ejecutamos
'**************************************

    'Si pulsamos Escape salimos
    Select Case KeyCode
    
    Case 27
    
        Call CloseClient
        
    Case 13  'Si pulsamos Enter...
    
        'y estamos en el conectar, entramos a la cuenta
        If Pantalla = PConnect Then
            Call btnConectar
            
        ElseIf Pantalla = PCuenta Then 'y estamos en la cuenta, entramos al pj
            If PJAccSelected <= 0 Or PJAccSelected > 10 Then
                MsgBox "Selecciona un PJ"
                Exit Sub
            End If
            
            Call ConnectPJ
            
        End If
        
    Case 46 'Eliminar PJ si esta dentro de la cuenta
    
        'Si no esta dentro de cuenta...
        If Not Pantalla = PCuenta Then Exit Sub
        
        '¿Tiene un PJ Seleccionado?
        If PJAccSelected < 1 Then
            Call MostrarMensaje(JsonLanguage.item("ERROR_PERSONAJE_NO_SELECCIONADO").item("TEXTO"))
            Exit Sub
        End If
                    
        frmBorrarPJ.Show
        
    End Select
    
End Sub

'<<<<<--------------------------------------------------------------------->>>>>>
'CONECTAR

Private Sub CrearNuevoPJ()
'**************************************
'Autor: Lorwik
'Fecha: 20/05/2020
'Descripcion: Boton de crear personaje
'**************************************
    Call Sound.Sound_Play(SND_CLICK)

    If CurrentUser.NumberOfCharacters > 9 Then
        Call MostrarMensaje(JsonLanguage.item("ERROR_DEMASIADOS_PJS").item("TEXTO"))
        Exit Sub
    End If
    
    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        If ClientSetup.bMusic <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_CrearPersonaje
            Sound.Fading = 500
        End If
    End If

    Call MostrarCreacion
End Sub

Private Sub btnConectar()
'**************************************
'Autor: Lorwik
'Fecha: 23/05/2020
'Descripcion: Boton de conectar cuenta
'**************************************
    Call Sound.Sound_Play(SND_CLICK)

    'Conectamos al servidor seleccionado
    CurServerIp = Servidor(ServIndSel).Ip
    CurServerPort = Servidor(ServIndSel).Puerto

    'update user info
    CurrentUser.AccountName = frmConnect.txtNombre.Text
    CurrentUser.AccountPassword = frmConnect.txtPasswd.Text

    'Clear spell list
    frmMain.hlst.Clear
    
    ClientSetup.rUserName = CurrentUser.AccountName

    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Login", "Remember", IIf(ClientSetup.Remember, "1", "0"))
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Login", "UserName", ClientSetup.rUserName)

    If CheckUserData() = True Then _
        Call Protocol_Handler.Connect(E_MODO.Normal)
        
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
    Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(0, "Open", "https://winterao.com/#/perfil", "", App.Path, SW_SHOWNORMAL)
    
End Sub

Private Sub btnCrearCuenta()
'**************************************
'Autor: Lorwik
'Fecha: 20/05/2020
'Descripcion: Boton de gestion de cuentas
'**************************************
    Call Sound.Sound_Play(SND_CLICK)

    Call ShellExecute(0, "Open", "https://winterao.com/#/registro", "", App.Path, SW_SHOWNORMAL)
    
End Sub

Public Function ListarServidores() As Boolean

    On Error GoTo ListarServidores_Err

    Dim lista()        As String

    Dim Elementos      As Byte
    
    Dim i              As Byte

    Dim responseServer As String
    
    Set Inet = New clsInet
    
    responseServer = Inet.OpenRequest("https://api.winterao.com/server-list.txt", "GET")
    responseServer = Inet.Execute
    responseServer = Inet.GetResponseAsString
    
    lista = Split(responseServer, ";")
    
    ReDim Servidor(0 To UBound(lista())) As Servidores
    
    For i = 0 To UBound(lista())
        Servidor(i).Ip = ReadField(1, lista(i), Asc("|"))
        Servidor(i).Puerto = ReadField(2, lista(i), Asc("|"))
        Servidor(i).Nombre = ReadField(3, lista(i), Asc("|"))
    Next i

    ListarServidores = True

    Exit Function

ListarServidores_Err:
    If Err.number <> 0 Then
       MsgBox ("No se ha podido obtener la lista de servidores. Error " & Err.number & " : " & Err.Description)
       ListarServidores = False
    End If
End Function

Public Sub ConnectPJ()
'**************************************
'Autor: Lorwik
'Fecha: 24/06/2020
'Descripcion: Mandamos el connect PJ
'**************************************

    If Not frmMain.Client.State = sckConnected Then
        MsgBox JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO")
        Call MostrarConnect
        
    Else
        If ModConectar.Conectando Then
            ModConectar.Conectando = False
            Call WriteLoginExistingChar
            
            DoEvents
    
            Call FlushBuffer
        End If
    End If
                
End Sub

'<<<<<--------------------------------------------------------------------->>>>>>
'CREACION DE PJ

Private Sub btnHeadPJ(ByVal Index As Integer)

    Select Case Index

        Case 0
            CurrentUser.UserHead = CheckCabeza(CurrentUser.UserHead + 1)

        Case 1
            CurrentUser.UserHead = CheckCabeza(CurrentUser.UserHead - 1)

    End Select
    
End Sub

Private Sub btnCrear()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Mandamos la creacion del personaje
'**************************************

    Dim i As Integer
    Dim Count As Byte
    
    'Nombre de usuario
    CurrentUser.UserName = LTrim$(frmConnect.txtCrearPJNombre.Text)
            
    '¿El nombre esta vacio y es correcto?
    If Right$(CurrentUser.UserName, 1) = " " Then
        CurrentUser.UserName = RTrim$(CurrentUser.UserName)
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(2))
        Exit Sub
    End If
    
    'Solo permitimos 1 espacio en los nombres
    For i = 1 To Len(CurrentUser.UserName)
        
        If mid$(CurrentUser.UserName, i, 1) = Chr$(32) Then Count = Count + 1
        
    Next i
    If Count > 1 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(3))
        Exit Sub
    End If
    
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

Private Sub DarCuerpoYCabeza()
'**************************************
'Autor: Lorwik
'Fecha: 24/05/2020
'Descripcion: Asignamos un cuerpo y unac abeza segun la raza y el sexo
'**************************************

    With CurrentUser

        Select Case .UserSexo
        
            Case eGenero.Hombre
    
                Select Case .UserRaza
    
                    Case eRaza.Humano
                        .UserHead = eCabezas.HUMANO_H_PRIMER_CABEZA
                        .UserBody = eCabezas.HUMANO_H_CUERPO_DESNUDO
                        
                    Case eRaza.Elfo
                        .UserHead = eCabezas.ELFO_H_PRIMER_CABEZA
                        .UserBody = eCabezas.ELFO_H_CUERPO_DESNUDO
                        
                    Case eRaza.ElfoOscuro
                        .UserHead = eCabezas.DROW_H_PRIMER_CABEZA
                        .UserBody = eCabezas.DROW_H_CUERPO_DESNUDO
                        
                    Case eRaza.Enano
                        .UserHead = eCabezas.ENANO_H_PRIMER_CABEZA
                        .UserBody = eCabezas.ENANO_H_CUERPO_DESNUDO
                        
                    Case eRaza.Gnomo
                        .UserHead = eCabezas.GNOMO_H_PRIMER_CABEZA
                        .UserBody = eCabezas.GNOMO_H_CUERPO_DESNUDO
                        
                    Case eRaza.Orco
                        .UserHead = eCabezas.ORCO_H_PRIMER_CABEZA
                        .UserBody = eCabezas.ORCO_H_CUERPO_DESNUDO
                        
                    Case eRaza.Vampiro
                        .UserHead = eCabezas.VAMPIRO_H_PRIMER_CABEZA
                        .UserBody = eCabezas.VAMPIRO_H_CUERPO_DESNUDO
                        
                    Case Else
                        .UserHead = 0
                        .UserBody = 0
                End Select
                
            Case eGenero.Mujer
    
                Select Case .UserRaza
    
                    Case eRaza.Humano
                        .UserHead = eCabezas.HUMANO_M_PRIMER_CABEZA
                        .UserBody = eCabezas.HUMANO_M_CUERPO_DESNUDO
                        
                    Case eRaza.Elfo
                        .UserHead = eCabezas.ELFO_M_PRIMER_CABEZA
                        .UserBody = eCabezas.ELFO_M_CUERPO_DESNUDO
                        
                    Case eRaza.ElfoOscuro
                        .UserHead = eCabezas.DROW_M_PRIMER_CABEZA
                        .UserBody = eCabezas.DROW_M_CUERPO_DESNUDO
                        
                    Case eRaza.Enano
                        .UserHead = eCabezas.ENANO_M_PRIMER_CABEZA
                        .UserBody = eCabezas.ENANO_M_CUERPO_DESNUDO
                        
                    Case eRaza.Gnomo
                        .UserHead = eCabezas.GNOMO_M_PRIMER_CABEZA
                        .UserBody = eCabezas.GNOMO_M_CUERPO_DESNUDO
                        
                    Case eRaza.Orco
                        .UserHead = eCabezas.ORCO_M_PRIMER_CABEZA
                        .UserBody = eCabezas.ORCO_M_CUERPO_DESNUDO
                        
                    Case eRaza.Vampiro
                        .UserHead = eCabezas.VAMPIRO_M_PRIMER_CABEZA
                        .UserBody = eCabezas.VAMPIRO_M_CUERPO_DESNUDO
                        
                    Case Else
                        .UserHead = 0
                        .UserBody = 0
                End Select
                
            Case Else
                .UserHead = 0
                .UserBody = 0
                
        End Select
    End With
    
End Sub

Private Function CheckCabeza(ByVal Head As Integer) As Integer

On Error GoTo ErrHandler

    Select Case CurrentUser.UserSexo

        Case eGenero.Hombre

            Select Case CurrentUser.UserRaza

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

            Select Case CurrentUser.UserRaza

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
    
ErrHandler:

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

    If CurrentUser.UserRaza > -1 Then
        
        With ModRaza(CurrentUser.UserRaza)
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
    Call Lector.Initialize(Carga.Path(Lenguajes) & "CharInfo_" & Language & ".dat")
    
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
    
    '¿Puso un nombre?
    If LenB(frmConnect.txtCrearPJNombre.Text) = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_NOMBRE_PJ").item("TEXTO"))
        frmConnect.txtCrearPJNombre.SetFocus
        Exit Function
    End If

    '¿Selecciono una raza?
    If CurrentUser.UserRaza = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_RAZA").item("TEXTO"))
        Exit Function
    End If
    
    '¿Selecciono el Sexo?
    If CurrentUser.UserSexo = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_SEXO").item("TEXTO"))
        Exit Function
    End If
    
    '¿Seleciono la clase?
    If CurrentUser.UserClase = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_CLASE").item("TEXTO"))
        Exit Function
    End If

    '¿Estamos intentando crear sin tener el AccountName?
    If Len(CurrentUser.AccountName) = 0 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_HASH").item("TEXTO"))
        Exit Function
    End If
    
    '¿El nombre de usuario supera los 30 caracteres?
    If LenB(CurrentUser.UserName) > 30 Then
        Call MostrarMensaje(JsonLanguage.item("VALIDACION_BAD_NOMBRE_PJ").item("TEXTO").item(1))
        Exit Function
    End If
    
    CheckData = True

End Function
