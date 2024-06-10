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
        .A = 255
        .R = 230
        .G = 200
        .B = 200
    End With
    
    With Estados(e_estados.MedioDia)
        .A = 255
        .R = 255
        .G = 255
        .B = 255
    End With
    
    With Estados(e_estados.Tarde)
        .A = 255
        .R = 200
        .G = 200
        .B = 200
    End With
  
    With Estados(e_estados.Noche)
        .A = 255
        .R = 165
        .G = 165
        .B = 165
    End With
    
    With Estados(e_estados.Lluvia)
        .A = 255
        .R = 200
        .G = 200
        .B = 200
    End With
    
    With Estados(e_estados.Niebla)
        .A = 255
        .R = 200
        .G = 200
        .B = 200
    End With
    
    With Estados(e_estados.FogLluvia)
        .A = 255
        .R = 200
        .G = 200
        .B = 200
    End With
    
    Estado_Actual_Date = 1
    
End Sub

Public Sub Actualizar_Estado(Optional ByVal Estado As Byte = 255)
'***************************************************
'Author: Lorwik
'Last Modification: 09/08/2020
'Actualiza el estado del clima y del dia
'***************************************************
    Dim x  As Integer
    Dim y  As Integer

    'Primero actualizamos la imagen del frmmain
    If Estado <> 255 Then _
        Call ActualizarImgClima(Estado)

    '¿Es un estado invalido?
    If Estado < 0 Or Estado > 8 Then Estado = e_estados.MedioDia
            
    Estado_Actual = Estados(Estado)
    Estado_Actual_Date = Estado
            
    For x = XMinMapSize To XMaxMapSize
        For y = YMinMapSize To YMaxMapSize
            
            If MapZonas(MapData(x, y).ZonaIndex).LuzBase <> 0 Then '¿La zona tiene su propia luz?
            
                Call Long_2_RGBAList(MapData(x, y).Light_Value(), MapZonas(MapData(x, y).ZonaIndex).LuzBase)
                
            Else
                Call RGBAList(MapData(x, y).Light_Value(), Estado_Actual.R, Estado_Actual.G, Estado_Actual.B, Estado_Actual.A)
                
            End If
                
        Next y
    Next x
            
    Call LucesRedondas.LightRenderAll
    
    If Estado = (e_estados.Lluvia Or e_estados.FogLluvia) Then
        If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
    
        bTecho = (MapData(UserPos.x, UserPos.y).Trigger = eTrigger.BAJOTECHO Or _
            MapData(UserPos.x, UserPos.y).Trigger = eTrigger.CASA Or _
            MapData(UserPos.x, UserPos.y).Trigger = eTrigger.ZONASEGURA)
        
    End If

End Sub

Private Sub ActualizarImgClima(ByVal Estado As Byte)

    Select Case Estado
    
        Case e_estados.Amanecer
            frmMain.imgClima.Picture = General_Load_Picture_From_Resource("226.gif")
        
        Case e_estados.MedioDia
            frmMain.imgClima.Picture = General_Load_Picture_From_Resource("226.gif")
            
        Case e_estados.Tarde
            frmMain.imgClima.Picture = General_Load_Picture_From_Resource("225.gif")
        
        Case e_estados.Noche
            frmMain.imgClima.Picture = General_Load_Picture_From_Resource("227.gif")
        
        Case e_estados.Lluvia
            frmMain.imgClima.Picture = General_Load_Picture_From_Resource("227.gif")
        
        Case e_estados.Niebla
            frmMain.imgClima.Picture = General_Load_Picture_From_Resource("227.gif")
        
        Case e_estados.FogLluvia
            frmMain.imgClima.Picture = General_Load_Picture_From_Resource("227.gif")
    
    End Select

End Sub

Public Sub Start_Rampage()
    '***************************************************
    'Author: Standelf
    'Last Modification: 27/05/2010
    'Init Rampage
    '***************************************************
    Dim x As Integer, y As Integer
    
    For x = XMinMapSize To XMaxMapSize
        For y = YMinMapSize To YMaxMapSize
            Call RGBAList(MapData(x, y).Light_Value(), 255, 255, 255, 255)
        Next y
    Next x

End Sub

Public Sub End_Rampage()

    '***************************************************
    'Author: Standelf
    'Last Modification: 27/05/2010
    'End Rampage
    '***************************************************
    
    OnRampageImgGrh = 0
    OnRampageImg = 0
    
    Dim x As Integer, y As Integer

    For x = XMinMapSize To XMaxMapSize
        For y = YMinMapSize To YMaxMapSize
            Call RGBAList(MapData(x, y).Light_Value(), Estado_Actual.R, Estado_Actual.G, Estado_Actual.B, Estado_Actual.A)
        Next y
    Next x

    Call LucesRedondas.LightRenderAll

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
    If bRain And MeterologiaEnDungeon Then
            
        'Particula segun el terreno...
        Select Case MapZonas(UserZonaId(UserCharIndex)).Terreno
        
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
                    Call Draw_GrhIndex(OnRampageImgGrh, 0, 0, 0, COLOR_WHITE(), , True)
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

Sub Engine_Weather_UpdateFog(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
'*****************************************************************
'Autor: ????
'Fecha: ????
'Descripción: Renderiza la niebla.
'*****************************************************************

    If MeterologiaEnDungeon = False Then Exit Sub
    
    If Estado_Actual_Date = e_estados.Niebla Or Estado_Actual_Date = e_estados.FogLluvia Then
    
        Dim TempGrh As Grh
        Dim i As Long
        Dim x As Long
        Dim y As Long
        Dim FogColor(3) As RGBA
    
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
        
        x = 2
        y = -1
        
        Call RGBAList(FogColor, R, G, B, A)
        
        For i = 1 To WeatherFogCount
            Call Draw_Grh(TempGrh, (x * 512) + WeatherFogX1, (y * 512) + WeatherFogY1, 1, FogColor(), 1, True)
            x = x + 1
            If x > (1 + (ScreenWidth \ 512)) Then
                x = 0
                y = y + 1
            End If
        Next i
                
        'Render fog 1
        TempGrh.GrhIndex = 3194
        x = 0
        y = 0
        For i = 1 To WeatherFogCount
            Call Draw_Grh(TempGrh, (x * 512) + WeatherFogX1, (y * 512) + WeatherFogY1, 1, FogColor(), 1, True)
            x = x + 1
            If x > (2 + (ScreenWidth \ 512)) Then
                x = 0
                y = y + 1
            End If
        Next i

    End If
    
End Sub

Public Function MeterologiaEnDungeon() As Boolean
'*********************************************
'Autor: Lorwik
'Fecha: 26/10/2020
'Descripcion: Comprueba si hay algun fenomeno meteorologico activo y si esta en dungeon
''*********************************************
    If (Estado_Actual_Date = e_estados.Niebla Or _
        bRain) And MapZonas(UserZonaId(UserCharIndex)).Zona <> "DUNGEON" Then
        
        MeterologiaEnDungeon = True
        
    Else
    
        MeterologiaEnDungeon = False
    
    End If
            
End Function
