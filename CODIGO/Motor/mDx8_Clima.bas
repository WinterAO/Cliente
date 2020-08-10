Attribute VB_Name = "mDx8_Clima"
Option Explicit

'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 15/05/10
'Blisse-AO | Set the Roof Color and Render _
    the Lights.
'***************************************************

Enum e_estados
    Amanecer = 0
    MedioDia
    Tarde
    Noche
    Lluvia
    Nieve
    Niebla
    FogLluvia 'Niebla mas lluvia
End Enum

Public Estados(0 To 8) As D3DCOLORVALUE
Public Estado_Actual As D3DCOLORVALUE
Public Estado_Actual_Date As Byte

Public Sub Init_MeteoEngine()
'***************************************************
'Author: Standelf
'Last Modification: 15/05/10
'Initializate
'***************************************************
    With Estados(e_estados.Amanecer)
        .a = 255
        .r = 230
        .g = 200
        .b = 200
    End With
    
    With Estados(e_estados.MedioDia)
        .a = 255
        .r = 255
        .g = 255
        .b = 255
    End With
    
    With Estados(e_estados.Tarde)
        .a = 255
        .r = 200
        .g = 200
        .b = 200
    End With
  
    With Estados(e_estados.Noche)
        .a = 255
        .r = 165
        .g = 165
        .b = 165
    End With
    
    With Estados(e_estados.Lluvia)
        .a = 255
        .r = 200
        .g = 200
        .b = 200
    End With
    
    With Estados(e_estados.Niebla)
        .a = 255
        .r = 200
        .g = 200
        .b = 200
    End With
    
    With Estados(e_estados.FogLluvia)
        .a = 255
        .r = 200
        .g = 200
        .b = 200
    End With
    
    Estado_Actual_Date = 1
    
End Sub

Public Sub Set_AmbientColor()
    Estado_Actual.a = 255
    Estado_Actual.b = CurMapAmbient.OwnAmbientLight.b
    Estado_Actual.g = CurMapAmbient.OwnAmbientLight.g
    Estado_Actual.r = CurMapAmbient.OwnAmbientLight.r
End Sub

Public Sub Actualizar_Estado(ByVal Estado As Byte)
'***************************************************
'Author: Lorwik
'Last Modification: 09/08/2020
'Actualiza el estado del clima y del dia
'***************************************************
    If Estado < 0 Or Estado > 8 Then Exit Sub
    
    If Estado = 0 Then Estado = e_estados.MedioDia
        
    Estado_Actual = Estados(Estado)
    Estado_Actual_Date = Estado
        
    Dim X As Byte, Y As Byte
    
    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X
        
    Call LightRenderAll
    
    If Estado = (e_estados.Lluvia Or e_estados.FogLluvia) Then
        If Not InMapBounds(UserPos.X, UserPos.Y) Then Exit Sub
    
        bTecho = (MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.BAJOTECHO Or _
            MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.CASA Or _
            MapData(UserPos.X, UserPos.Y).Trigger = eTrigger.ZONASEGURA)
        
    End If
    
End Sub

Public Sub Start_Rampage()
'***************************************************
'Author: Standelf
'Last Modification: 27/05/2010
'Init Rampage
'***************************************************
    Dim X As Byte, Y As Byte, TempColor As D3DCOLORVALUE
    TempColor.a = 255: TempColor.b = 255: TempColor.r = 255: TempColor.g = 255
    
        For X = XMinMapSize To XMaxMapSize
            For Y = YMinMapSize To YMaxMapSize
                Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), TempColor)
            Next Y
        Next X
End Sub

Public Sub End_Rampage()

    '***************************************************
    'Author: Standelf
    'Last Modification: 27/05/2010
    'End Rampage
    '***************************************************
    Dim X As Byte, Y As Byte

    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X

    Call LightRenderAll

End Sub

Public Function bRain() As Boolean
    If Estado_Actual_Date = (e_estados.Lluvia Or e_estados.FogLluvia) Then
        bRain = True
        Exit Function
    End If
    
    bRain = False
End Function
