Attribute VB_Name = "ModCnt"
Option Explicit

'Mapa actual seleccionado para el conectar renderizado
Private SelectConnectMap As Byte

Enum EPantalla
    PConnect = 0
    PCuenta
    PCrearPJ
End Enum

Public Pantalla As EPantalla

Public Sub MostrarConnect()
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'llama al frmConnect con el mapa, de lo contrario no funcionaria correctamente.
'******************************

    'Seteamos el modo login
    Pantalla = PConnect
    
    frmConnect.Visible = True
    
    'Sorteamos el mapa a mostrar
    'Nota el mapa 1 es para el crear pj
    SelectConnectMap = RandomNumber(2, NumConnectMap)
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
                            
                'Capa 3
                Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, .Engine_Light(), 1)
                
                'Particulas
                If .Particle_Group_Index Then
                    
                    'Solo las renderizamos si estan cerca del area de vision.
                    If EstaDentroDelArea(x, y) Then
                        Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetXTemp + 16)
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
        
        Case 2 'Crear PJ
    
    End Select
    
    '<------- Desde aqui lo que siempre se va a mostrar ------->
    'Logo
    Call Draw_GrhIndex(31480, 1, 1, 0, Normal_RGBList(), 0, False)
    
    Call DrawText(10, 750, "WinterAO " & GetVersionOfTheGame() & " Resurrection", Color_Paralisis)
End Sub

Public Sub ClickEvent(ByVal TX As Long, ByVal TY As Long)
'******************************
'Autor: Lorwik
'Fecha: 13/05/2020
'Eventos al realizar clicks en la GUI
'******************************
    Dim x As Integer
    Dim y As Integer
    
    Debug.Print TX & " " & TY
    
    If (TX >= 100 And TX <= 200) And (TY >= 100 And TY <= 200) Then
            MsgBox "Hola Mundo"
        End If
    
End Sub
