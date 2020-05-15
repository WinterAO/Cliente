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

Sub RenderConnect()
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
                        Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetXTemp + 16)
                    End If
                        
                End If
            End With
        Next Y
    Next X
    
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

