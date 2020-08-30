Attribute VB_Name = "mDx8_Clima"
Option Explicit

'***************************************************
'Autor: Lorwik
'Descripción: Este sistema es una adaptación del que hice en
'las versiones anteriores de Winter que posteriormente mejore en
'AODrag. El sistema fue adaptado al que trae AOLibre que a su vez
'se basaba en el de Blisse.
'***************************************************

Public Enum e_estados
    Amanecer = 0
    MedioDia
    Tarde
    Noche
    Lluvia
    Niebla
    FogLluvia 'Niebla mas lluvia
End Enum

Public Estados(0 To 8) As D3DCOLORVALUE
Public Estado_Actual As D3DCOLORVALUE
Public Estado_Actual_Date As Byte

'****************************
'Usado para las particulas
'****************************

Private RainParticle As Long
Private NieveParticle As Long
Private ArenaParticle As Long

Public OnRampage As Long
Public OnRampageImg As Long
Public OnRampageImgGrh As Integer

Public Enum eWeather
    Rain
    Nieve
    Arena
End Enum

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
    
    OnRampageImgGrh = 0
    OnRampageImg = 0
    
    Dim X As Byte, Y As Byte

    For X = XMinMapSize To XMaxMapSize
        For Y = YMinMapSize To YMaxMapSize
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual)
        Next Y
    Next X

    Call LightRenderAll

End Sub

Public Function bRain() As Boolean
'*****************************************************************
'Author: Lorwik
'Fecha: 13/08/2020
'Devuelve un True o un False si hay lluvia
'*****************************************************************

    If Estado_Actual_Date = e_estados.FogLluvia Or Estado_Actual_Date = e_estados.Lluvia Then
        bRain = True
        Exit Function
    End If
    
    bRain = False
End Function

Public Sub Engine_Weather_Update()
'*****************************************************************
'Author: Lorwik
'Fecha: 13/08/2020
'Controla los climas, aqui se renderizan la lluvia, nieve, etc.
'*****************************************************************

    '¿Esta lloviendo y no esta en dungeon?
    If bRain And mapInfo.Zona <> "DUNGEON" Then
            
        'Particula segun el terreno...
        Select Case mapInfo.Terreno
        
            Case "BOSQUE"
                If RainParticle <= 0 Then
                    'Creamos las particulas de lluvia
                    Call mDx8_Clima.LoadWeatherParticles(eWeather.Rain)
                ElseIf RainParticle > 0 Then
                    Call mDx8_Particulas.Particle_Group_Render(RainParticle, 250, -1)
                End If
                
                'EXTRA: Relampagos
                If RandomNumber(1, 200000) < 20 Then
                    Call Sound.Sound_Play(SND_RELAMPAGO)
                    Start_Rampage
                    OnRampage = GetTickCount
                    OnRampageImg = OnRampage
                    OnRampageImgGrh = 2837
            End If
            
                If OnRampageImg <> 0 Then
                    If GetTickCount - OnRampageImg > 36 Then
                    
                        OnRampageImgGrh = OnRampageImgGrh + 1
                        If OnRampageImgGrh = 2847 Then OnRampageImgGrh = 2837
            
                        OnRampageImg = GetTickCount
                    End If
                End If
                
                If OnRampage <> 0 Then 'Hay Uno en curso
                    If GetTickCount - OnRampage > 400 Then
                        End_Rampage
                        OnRampage = 0
                    End If
                End If
                
                If OnRampageImgGrh <> 0 Then
                    Call Draw_GrhIndex(OnRampageImgGrh, 0, 0, 0, Normal_RGBList(), , True)
                End If
            
            Case "NIEVE"
            
                If NieveParticle <= 0 Then
                    'Creamos las particulas de nieve
                    Call mDx8_Clima.LoadWeatherParticles(eWeather.Nieve)
                ElseIf NieveParticle > 0 Then
                    Call mDx8_Particulas.Particle_Group_Render(NieveParticle, 250, -1)
                End If
            
            Case "DESIERTO"
        
                If ArenaParticle <= 0 Then
                    'Creamos las particulas de Arena
                    Call mDx8_Clima.LoadWeatherParticles(eWeather.Arena)
                ElseIf ArenaParticle > 0 Then
                    Call mDx8_Particulas.Particle_Group_Render(ArenaParticle, 250, -1)
                End If
        
        End Select
    
    Else '¿No esta lloviendo o dejo de llover?
        
        Call RemoveWeatherParticlesAll
            
    End If
    
    Engine_Weather_UpdateFog 100, 255, 255, 255

End Sub

Public Sub LoadWeatherParticles(ByVal Weather As Byte)
'*****************************************************************
'Author: Lucas Recoaro (Recox)
'Last Modify Date: 19/12/2019
'Crea las particulas de clima.
'*****************************************************************
    Select Case Weather

        Case eWeather.Rain
            RainParticle = mDx8_Particulas.General_Particle_Create(8, -1, -1)
            
        Case eWeather.Nieve
            NieveParticle = mDx8_Particulas.General_Particle_Create(56, -1, -1)
            
        Case eWeather.Arena
            ArenaParticle = mDx8_Particulas.General_Particle_Create(59, -1, -1)
    End Select
End Sub

Public Sub RemoveWeatherParticlesAll()
'*****************************************************************
'Author: Lorwik
'Last Modify Date: 13/08/2020
'Comprobamos si hay alguna particula climatologica activa para eliminarla
'*****************************************************************

    'Si alguna de las siguientes particulas esta cargada, la eliminamos
    If RainParticle > 0 Then
        Call mDx8_Clima.RemoveWeatherParticles(eWeather.Rain)
            
    ElseIf NieveParticle > 0 Then
        Call mDx8_Clima.RemoveWeatherParticles(eWeather.Nieve)
            
    ElseIf ArenaParticle > 0 Then
        Call mDx8_Clima.RemoveWeatherParticles(eWeather.Arena)
    End If
End Sub

Public Sub RemoveWeatherParticles(ByVal Weather As Byte)
'*****************************************************************
'Author: Lorwik
'Fecha: 14/08/2020
'Elimina las particulas climatologicas segun la que reciba
'*****************************************************************
    Select Case Weather

        Case eWeather.Rain
            Particle_Group_Remove (RainParticle)
            RainParticle = 0
            
        Case eWeather.Nieve
            Particle_Group_Remove (NieveParticle)
            NieveParticle = 0
            
        Case eWeather.Arena
            Particle_Group_Remove (ArenaParticle)
            ArenaParticle = 0

    End Select
End Sub

Sub Engine_Weather_UpdateFog(ByVal a As Byte, ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
'*****************************************************************
'Autor: ????
'Fecha: ????
'Descripción: Renderiza la niebla.
'*****************************************************************

    If Estado_Actual_Date = e_estados.Niebla Or Estado_Actual_Date = e_estados.FogLluvia Then
    
        Dim TempGrh As Grh
        Dim i As Long
        Dim X As Long
        Dim Y As Long
        Dim FogColor(3) As Long
    
        'Make sure we have the fog value
        If WeatherFogCount = 0 Then WeatherFogCount = 13
        
        'Update the fog's position
        WeatherFogX1 = WeatherFogX1 + (timerElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
        WeatherFogY1 = WeatherFogY1 + (timerElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
        Do While WeatherFogX1 < -512
            WeatherFogX1 = WeatherFogX1 + 512
        Loop
        Do While WeatherFogY1 < -512
            WeatherFogY1 = WeatherFogY1 + 512
        Loop
        Do While WeatherFogX1 > 0
            WeatherFogX1 = WeatherFogX1 - 512
        Loop
        Do While WeatherFogY1 > 0
            WeatherFogY1 = WeatherFogY1 - 512
        Loop
        
        WeatherFogX2 = WeatherFogX2 - (timerElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
        WeatherFogY2 = WeatherFogY2 - (timerElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
        Do While WeatherFogX2 < -512
            WeatherFogX2 = WeatherFogX2 + 512
        Loop
        Do While WeatherFogY2 < -512
            WeatherFogY2 = WeatherFogY2 + 512
        Loop
        Do While WeatherFogX2 > 0
            WeatherFogX2 = WeatherFogX2 - 512
        Loop
        Do While WeatherFogY2 > 0
            WeatherFogY2 = WeatherFogY2 - 512
        Loop
    
        TempGrh.FrameCounter = 1
        
        'Render fog 2
        TempGrh.GrhIndex = 3193
        
        X = 2
        Y = -1
        For i = 0 To 3
            FogColor(i) = D3DColorARGB(a, r, g, b)
        Next i
        
        For i = 1 To WeatherFogCount
            Call Draw_Grh(TempGrh, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, 1, FogColor(), 1, True)
            X = X + 1
            If X > (1 + (ScreenWidth \ 512)) Then
                X = 0
                Y = Y + 1
            End If
        Next i
                
        'Render fog 1
        TempGrh.GrhIndex = 3194
        X = 0
        Y = 0
        For i = 1 To WeatherFogCount
            Call Draw_Grh(TempGrh, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, 1, FogColor(), 1, True)
            X = X + 1
            If X > (2 + (ScreenWidth \ 512)) Then
                X = 0
                Y = Y + 1
            End If
        Next i
        
    End If

End Sub
