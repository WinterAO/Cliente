Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matias Fernando Pequeno
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

'Caminata fluida
Public Movement_Speed As Single

Dim temp_verts(3) As TLVERTEX

Public OffsetCounterX As Single
Public OffsetCounterY As Single
    
Public WeatherFogX1 As Single
Public WeatherFogY1 As Single
Public WeatherFogX2 As Single
Public WeatherFogY2 As Single
Public WeatherFogCount As Byte

Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Long
Public LastOffsetY As Long

'Map sizes in tiles
Public Const XMaxMapSize As Integer = 1000
Public Const XMinMapSize As Integer = 1
Public Const YMaxMapSize As Integer = 1000
Public Const YMinMapSize As Integer = 1

Private Const GrhFogata As Long = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180

'Posicion en un mapa
Public Type Position
    X As Integer
    Y As Integer
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamano y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    speed As Single
    
    Trans As Byte
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Long
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.SOUTH To E_Heading.EAST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.SOUTH To E_Heading.EAST) As Grh
    offset As Position
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.SOUTH To E_Heading.EAST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.SOUTH To E_Heading.EAST) As Grh
End Type

'Lista de las animaciones de ataque
Type AtaqueAnimData
    AtaqueWalk(E_Heading.SOUTH To E_Heading.EAST) As Grh
    HeadOffset As Position
End Type

Public Enum eStatusQuest
    NoAceptada = 0
    EnCurso = 1
    Terminada = 2
    NoTieneQuest = 255
End Enum

'Apariencia del personaje
Public Type Char
    Movement As Boolean
    Active As Byte
    Heading As E_Heading
    Pos As Position
    moved As Boolean
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As Integer
    Casco As Integer
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    Ataque As AtaqueAnimData
    UsandoArma As Boolean
    AuraAnim As Grh
    AuraColor As Long
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Atacable As Byte
    WorldBoss As Boolean
    
    Nombre As String
    Clan As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    attacking As Boolean
    
    ParticleIndex As Integer
    Particle_Count As Long
    Particle_Group() As Long
    
    NPCAttack As Boolean
    EstadoQuest As eStatusQuest
    
    BarTime As Single
    MaxBarTime As Integer
    BarAccion As Byte
    
    Speeding As Single
    
    EsNPC As Boolean
End Type

'Info de un objeto
Public Type obj
    ObjIndex As Integer
    Amount As Integer
    Shadow As Byte
End Type

Public Type tGhost

    Active As Boolean
    Body As Grh
    Head As Integer
    Weapon As Grh
    Helmet As Integer
    Shield As Grh
    Body_Aura As String
    AlphaB As Single
    OffX As Integer
    Offy As Integer
    Heading As Byte

End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    Damage As DList
    
    NPCIndex As Integer
    OBJInfo As obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    
    Light_Value(3) As RGBA
    
    Particle_Index As Integer
    Particle_Group_Index As Long 'Particle Engine
    
    fX As Grh
    FxIndex As Integer
    
    ZonaIndex As Integer
    
    GhostChar As tGhost
End Type

'Info de cada mapa
Public Type tZonaInfo
    Music As Integer
    name As String
    StartPos As WorldPos
    MapVersion As Integer
    Ambient As Integer
    AmbientNight As Integer
    Zona As String
    Terreno As String
    LuzBase As Long
    battle_mode As Boolean
End Type

Public IniPath As String

'Bordes del mapa
Public MinXBorder As Integer
Public MaxXBorder As Integer
Public MinYBorder As Integer
Public MaxYBorder As Integer

'Status del user
Public CurMap As String 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserPos As Position 'Posicion
Public UserPosCuadrante As Position
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Public FPSLastCheck As Long

'Tamano del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Tamano de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer

Private MouseTileX As Integer
Private MouseTileY As Integer

'?????????Graficos???????????
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public AtaqueData() As AtaqueAnimData
'?????????????????????????

'?????????Mapa????????????
Public MapData() As MapBlock ' Mapa
Public MapZonas() As tZonaInfo ' Info acerca del mapa en uso
Public CantZonas As Integer
'?????????????????????????

'   Control de Lluvia
Public bTecho       As Boolean 'hay techo?
Public bFogata       As Boolean

Public charlist(1 To 10000) As Char

Public LastOffset2X As Double
Public LastOffset2Y As Double

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.speed = GrhData(Grh.GrhIndex).speed
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 09/21/2010
' 09/21/2010: C4b3z0n - Changed from Private Funtion tu Public Function.
'***************************************************
    
    With charlist(CharIndex).Pos
    
        EstaPCarea = .X > UserPos.X - MinXBorder And _
                     .X < UserPos.X + MinXBorder And _
                     .Y > UserPos.Y - MinYBorder And _
                     .Y < UserPos.Y + MinYBorder
    End With
    
End Function

Private Function HayFogata(ByRef Location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    Location.X = j
                    Location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Public Sub DrawTransparentGrhtoHdc(ByVal dsthdc As Long, ByVal srchdc As Long, ByRef SourceRect As RECT, ByRef DestRect As RECT, ByVal TransparentColor As Long)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 27/07/2012 - ^[GS]^
'*************************************************************
    Dim color As Long
    Dim X As Long
    Dim Y As Long
    
    For X = SourceRect.Left To SourceRect.Right
        For Y = SourceRect.Top To SourceRect.Bottom
            color = GetPixel(srchdc, X, Y)
            
            If color <> TransparentColor Then
                Call SetPixel(dsthdc, DestRect.Left + (X - SourceRect.Left), DestRect.Top + (Y - SourceRect.Top), color)
            End If
        Next Y
    Next X
End Sub

Public Sub DrawImageInPicture(ByRef PictureBox As PictureBox, ByRef Picture As StdPicture, ByVal x1 As Single, ByVal y1 As Single, Optional Width1, Optional Height1, Optional x2, Optional y2, Optional Width2, Optional Height2)
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 12/28/2009
'Draw Picture in the PictureBox
'*************************************************************

    Call PictureBox.PaintPicture(Picture, x1, y1, Width1, Height1, x2, y2, Width2, Height2)

End Sub

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************
    
    On Error GoTo RenderScreen_Err
    
    Dim Y                As Long     'Keeps track of where on map we are
    Dim X                As Long     'Keeps track of where on map we are
    
    Dim screenminY       As Integer  'Start Y pos on current screen
    Dim screenmaxY       As Integer  'End Y pos on current screen
    Dim screenminX       As Integer  'Start X pos on current screen
    Dim screenmaxX       As Integer  'End X pos on current screen
    
    Dim MinY             As Integer  'Start Y pos on current map
    Dim MaxY             As Integer  'End Y pos on current map
    Dim MinX             As Integer  'Start X pos on current map
    Dim MaxX             As Integer  'End X pos on current map
    
    Dim screenX          As Integer  'Keeps track of where to place tile on screen
    Dim screeny          As Integer  'Keeps track of where to place tile on screen
    
    Dim minXOffset       As Integer
    Dim minYOffset       As Integer
    
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    Dim ElapsedTime      As Single
    Dim ColorFinal(3)    As RGBA
    
    ElapsedTime = Engine_ElapsedTime()
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    MinY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    MinX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If MinY < XMinMapSize Then
        minYOffset = YMinMapSize - MinY
        MinY = YMinMapSize
    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If MinX < XMinMapSize Then
        minXOffset = XMinMapSize - MinX
        MinX = XMinMapSize
    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        screeny = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        screenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    ParticleOffsetX = (Engine_PixelPosX(screenminX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(screenminY) - PixelOffsetY)

    'Draw floor layer
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            
            PixelOffsetXTemp = (screenX - 1) * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = (screeny - 1) * TilePixelHeight + PixelOffsetY
            
            'Layer 1 **********************************
            If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then Call Draw_Grh(MapData(X, Y).Graphic(1), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Light_Value(), 1)
            
            screenX = screenX + 1
        Next
    
        'Reset ScreenX to original value and increment ScreenY
        screenX = screenX - X + screenminX
        screeny = screeny + 1
    Next
    
    '<----- Layer 2 ----->
    screeny = minYOffset - TileBufferSize

    For Y = MinY To MaxY
        
        screenX = minXOffset - TileBufferSize

        For X = MinX To MaxX

            If Map_InBounds(X, Y) Then
            
                PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = screeny * TilePixelHeight + PixelOffsetY
   
                'Layer 2 **********************************
                If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then Call Draw_Grh(MapData(X, Y).Graphic(2), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Light_Value(), 1)
    
            End If
            
            screenX = screenX + 1
        Next X

        screeny = screeny + 1
    Next Y
    
    '<----- Layer Obj, Char, 3 ----->
    screeny = minYOffset - TileBufferSize

    For Y = MinY To MaxY
        
        screenX = minXOffset - TileBufferSize

        For X = MinX To MaxX

            If Map_InBounds(X, Y) Then
            
                PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
                PixelOffsetYTemp = screeny * TilePixelHeight + PixelOffsetY
                
                With MapData(X, Y)
                
                    'Object Layer **********************************
                    If .ObjGrh.GrhIndex <> 0 Then
                        If .OBJInfo.Shadow = 1 Then Call Draw_Grh(.ObjGrh, PixelOffsetXTemp + 5, PixelOffsetYTemp + -9, 1, COLOR_SHADOW(), 0, False, 187, 1, 1.2)
                           
                        Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, .Light_Value(), 1)
                    End If
                    '***********************************************

                    'Char Ghost ***********************************
                    Call RenderGhotChar(X, Y, PixelOffsetXTemp, PixelOffsetYTemp)
                    '**********************************************

                    'Char layer********************************
                    If .CharIndex <> 0 Then Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                    '*************************************************

                    'Layer 3 *****************************************
                    If .Graphic(3).GrhIndex <> 0 Then
                        
                        '¿El Grh tiene propiedades de transparencia?
                        If GrhData(.Graphic(3).GrhIndex).Trans = 1 Then
                            If Abs(UserPos.X - X) < 2 And (Abs(UserPos.Y - Y)) < 5 And (Abs(UserPos.Y) < Y) Then
                                Call Copy_RGBAList(COLOR_ARBOL, ColorFinal)
                            Else
                                Call Copy_RGBAList(.Light_Value, ColorFinal)
                            End If
    
                        Else
                            Call Copy_RGBAList(.Light_Value, ColorFinal)
                        End If
                        
                        Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, ColorFinal(), 1)
                        
                    End If
                    '************************************************
                    
                    'Dibujamos los danos.
                    If .Damage.Activated Then
                        Call mDx8_Dibujado.Damage_Draw(X, Y, PixelOffsetXTemp, PixelOffsetYTemp - 20)
                    End If
                    
                    'Particulas
                    If .Particle_Group_Index Then
                    
                        'Solo las renderizamos si estan cerca del area de vision.
                        If EstaDentroDelArea(X, Y) Then
                            Call mDx8_Particulas.Particle_Group_Render(.Particle_Group_Index, PixelOffsetXTemp + 16, PixelOffsetYTemp + 16)
                        End If
                        
                    End If

                    If Not .FxIndex = 0 Then
                        Call Draw_Grh(.fX, PixelOffsetXTemp + FxData(MapData(X, Y).FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxIndex).OffsetY, 1, .Light_Value(), 1, True)

                        If .fX.Started = 0 Then .FxIndex = 0
                    End If
                    
                End With
                
            End If
            
            screenX = screenX + 1
        Next X

        screeny = screeny + 1
    Next Y
    
    '<----- Layer 4 ----->
    screeny = minYOffset - TileBufferSize

    For Y = MinY To MaxY

        screenX = minXOffset - TileBufferSize

        For X = MinX To MaxX
            
            PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = screeny * TilePixelHeight + PixelOffsetY
            
            'Layer 4
            If MapData(X, Y).Graphic(4).GrhIndex Then
            
                If bTecho Then 'Esta bajo techo
                    Call Draw_Grh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, temp_rgb(), 1)
                    
                Else
                
                    If ColorTecho = 250 Then
                        Call Draw_Grh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, MapData(X, Y).Light_Value(), 1)
                        
                    Else
                        Call Draw_Grh(MapData(X, Y).Graphic(4), PixelOffsetXTemp, PixelOffsetYTemp, 1, temp_rgb(), 1)
                        
                    End If
                    
                End If
                
            End If
            
            'Areas
            If DebugAreas Then
                Call RenderAreas(PixelOffsetXTemp, PixelOffsetYTemp)
            End If
            
            screenX = screenX + 1
            
        Next X

        screeny = screeny + 1
    Next Y

    'Weather Update & Render - Aca se renderiza la lluvia, nieve, etc.
    If ClientSetup.ParticleEngine Then Call mDx8_Clima.Engine_Weather_Update

    If ClientSetup.ProyectileEngine Then
                            
        If LastProjectile > 0 Then
            Dim j As Long ' Long siempre en los bucles es mucho mas rapido
                                
            For j = 1 To LastProjectile

                If ProjectileList(j).Grh.GrhIndex Then
                    Dim angle As Single
                    
                    'Update the position
                    angle = DegreeToRadian * Engine_GetAngle(ProjectileList(j).X, ProjectileList(j).Y, ProjectileList(j).tX, ProjectileList(j).tY)
                    ProjectileList(j).X = ProjectileList(j).X + (Sin(angle) * ElapsedTime * 0.8)
                    ProjectileList(j).Y = ProjectileList(j).Y - (Cos(angle) * ElapsedTime * 0.8)
                    
                    'Update the rotation
                    If ProjectileList(j).RotateSpeed > 0 Then
                        ProjectileList(j).Rotate = ProjectileList(j).Rotate + (ProjectileList(j).RotateSpeed * ElapsedTime * 0.01)

                        Do While ProjectileList(j).Rotate > 360
                            ProjectileList(j).Rotate = ProjectileList(j).Rotate - 360
                        Loop
                    End If
    
                    'Draw if within range
                    X = ((-MinX - 1) * 32) + ProjectileList(j).X + PixelOffsetX + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(j).OffsetX
                    Y = ((-MinY - 1) * 32) + ProjectileList(j).Y + PixelOffsetY + ((10 - TileBufferSize) * 32) - 288 + ProjectileList(j).OffsetY

                    If Y >= -32 Then
                        If Y <= (ScreenHeight + 32) Then
                            If X >= -32 Then
                                If X <= (ScreenWidth + 32) Then
                                    If ProjectileList(j).Rotate = 0 Then
                                        Call Draw_Grh(ProjectileList(j).Grh, X, Y, 0, MapData(50, 50).Light_Value(), 0, True, ProjectileList(j).Rotate + 128)
                                    Else
                                        Call Draw_Grh(ProjectileList(j).Grh, X, Y, 0, MapData(50, 50).Light_Value(), 0, True, ProjectileList(j).Rotate + 128)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                End If
            Next j
            
            'Check if it is close enough to the target to remove
            For j = 1 To LastProjectile

                If ProjectileList(j).Grh.GrhIndex Then
                    If Abs(ProjectileList(j).X - ProjectileList(j).tX) < 20 Then
                        If Abs(ProjectileList(j).Y - ProjectileList(j).tY) < 20 Then
                            Call Engine_Projectile_Erase(j)
                        End If
                    End If
                End If
            Next j
            
        End If
    End If
    
    If colorRender <> 240 Then
        Call Draw_GrhIndex(34027, 462, 110, 1, render_msg())
        Call DrawText(482, 110, renderTextPk, render_msg(), True, 4)
        Call DrawText(482, 60, renderText, render_msg(), True, 3)
    End If
    
    'Call Draw_GrhIndex(34027, 372, 80, 1, COLOR_WHITE())
    'Call DrawText(390, 50, "Nivel 21", -1, True, 2)
    
    '   Set Offsets
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    
    If ClientSetup.PartyMembers Then Call Draw_Party_Members

RenderScreen_Err:

    If Err.number Then
        Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.RenderScreen")
    End If
    
End Sub

Private Sub RenderGhotChar(ByVal X As Integer, _
                           ByVal Y As Integer, _
                           screenX As Integer, _
                           screeny As Integer)
    '****************************************
    'Autor: Lorwik
    'Fecha: 19/09/2024
    'Descripción: Renderizamos los cuerpos que se desvanecen al ser eliminados.
    '****************************************
    
    Dim TempColor(3) As RGBA
    
    With MapData(X, Y)
    
        If .GhostChar.Active Then
                    
            If .GhostChar.AlphaB > 0 Then
                    
                .GhostChar.AlphaB = .GhostChar.AlphaB - (timerTicksPerFrame * 30)
                        
                'Redondeamos a 0 para prevenir errores
                If .GhostChar.AlphaB < 0 Then .GhostChar.AlphaB = 0
                        
                Call Copy_RGBAList_WithAlpha(TempColor, .Light_Value, .GhostChar.AlphaB)
                        
                'Seteamos el color
                If .GhostChar.Heading = 1 Or .GhostChar.Heading = 2 Then
                    Call Draw_Grh(.GhostChar.Shield, screenX, screeny, 1, TempColor(), 1, False)
                    Call Draw_Grh(.GhostChar.Body, screenX, screeny, 1, TempColor(), 1, False)
                    Call DrawHead(.GhostChar.Head, screenX + .GhostChar.OffX, screeny + .GhostChar.Offy, TempColor(), .GhostChar.Heading, True, False)
                    Call DrawHead(.GhostChar.Helmet, screenX + .GhostChar.OffX, screeny + .GhostChar.Offy, TempColor(), .GhostChar.Heading, False, False)
                    Call Draw_Grh(.GhostChar.Weapon, screenX, screeny, 1, TempColor(), 1, False)
                Else
                    Call Draw_Grh(.GhostChar.Body, screenX, screeny, 1, TempColor(), 1, False)
                    Call DrawHead(.GhostChar.Head, screenX + .GhostChar.OffX, screeny + .GhostChar.Offy, TempColor(), .GhostChar.Heading, True, False)
                    Call Draw_Grh(.GhostChar.Shield, screenX, screeny, 1, TempColor(), 1, False)
                    Call DrawHead(.GhostChar.Helmet, screenX + .GhostChar.OffX, screeny + .GhostChar.Offy, TempColor(), .GhostChar.Heading, False, False)
                    Call Draw_Grh(.GhostChar.Weapon, screenX, screeny, 1, TempColor(), 1, False)
                End If

            Else
                .GhostChar.Active = False

            End If

        End If
    
    End With
    
End Sub

Private Sub RenderHUD()
    '****************************************
    'Autor: Lorwik
    'Fecha: 29/04/2020
    'Descripción: Renderizamos información relevante al juego en el screen
    '****************************************

    If Dialogos.NeedRender Then Call Dialogos.Render ' GSZAO
    Call DibujarCartel

    If DialogosClanes.Activo Then Call DialogosClanes.Draw ' GSZAO

    ' Calculamos los FPS y los mostramos
    Call Engine_Update_FPS

    If ClientSetup.FPSShow = True Then
        Call DrawText(940, 5, "FPS: " & Mod_TileEngine.FPS, COLOR_WHITE, True)
        Call DrawText(940, 20, "Zona: " & MapData(UserPos.X, UserPos.Y).ZonaIndex, COLOR_WHITE, True)
    End If
        
    If ClientSetup.HUD Then
        If Not lblHelm = "0/0" And Not lblHelm = "" Then
            Call Draw_GrhIndex(30792, 20, 450, 1, COLOR_WHITE, 0, False)
            Call DrawText(50, 457, lblHelm, COLOR_WHITE, True)
        End If
            
        If Not lblArmor = "0/0" And Not lblArmor = "" Then
            Call Draw_GrhIndex(30793, 20, 490, 1, COLOR_WHITE, 0, False)
            Call DrawText(50, 497, lblArmor, COLOR_WHITE, True)
        End If
            
        If Not lblShielder = "0/0" And Not lblShielder = "" Then
            Call Draw_GrhIndex(30794, 20, 530, 1, COLOR_WHITE, 0, False)
            Call DrawText(50, 537, lblShielder, COLOR_WHITE, True)
        End If
            
        If Not lblWeapon = "0/0" And Not lblWeapon = "" Then
            Call Draw_GrhIndex(30795, 20, 570, 1, COLOR_WHITE, 0, False)
            Call DrawText(50, 573, lblWeapon, COLOR_WHITE, True)
        End If
    End If
        
End Sub

Private Sub RenderAreas(ByVal X As Integer, ByVal Y As Integer)
'****************************************
'Autor: Lorwik
'Fecha: 13/09/2024
'Descripción: Renderiza las areas
'****************************************

On Error GoTo RenderAreas_Err

    Dim MinAreaX As Integer, MaxAreaX As Integer, MinAreaY As Integer, MaxAreaY As Integer
                
    Call CalcularArea(1, MinAreaX, MaxAreaX, MinAreaY, MaxAreaY)

    If Y >= MinAreaY And Y <= MaxAreaY Then
        If X >= MinAreaX And X <= MaxAreaX Then
            Call Draw_GrhIndex(2, X, Y, 1, COLOR_WHITE())
        End If
    End If
                
    Call CalcularArea(2, MinAreaX, MaxAreaX, MinAreaY, MaxAreaY)

    If Y >= MinAreaY And Y <= MaxAreaY Then
        If X >= MinAreaX And X <= MaxAreaX Then
            Call Draw_GrhIndex(2, X, Y, 1, COLOR_WHITE())
        End If
    End If
                
    Call CalcularArea(3, MinAreaX, MaxAreaX, MinAreaY, MaxAreaY)

    If Y >= MinAreaY And Y <= MaxAreaY Then
        If X >= MinAreaX And X <= MaxAreaX Then
            Call Draw_GrhIndex(2, X, Y, 1, COLOR_WHITE())
        End If
    End If
                
    Call CalcularArea(4, MinAreaX, MaxAreaX, MinAreaY, MaxAreaY)

    If Y >= MinAreaY And Y <= MaxAreaY Then
        If X >= MinAreaX And X <= MaxAreaX Then
            Call Draw_GrhIndex(2, X, Y, 1, COLOR_WHITE())
        End If
    End If
    
RenderAreas_Err:
    If Err.number Then
        Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.RenderArea")
    End If
    
End Sub

Sub DoPasosFx(ByVal CharIndex As Integer)
Static TerrenoDePaso As TipoPaso

    With charlist(CharIndex)
        If Not CurrentUser.UserNavegando Then
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
             
                    If Not Char_Big_Get(CharIndex) Then
                        TerrenoDePaso = GetTerrenoDePaso(.Pos.X, .Pos.Y)
                    Else
                        TerrenoDePaso = CONST_PESADO
                    End If

                    If .pie = 0 Then
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(1), , Sound.Calculate_Volume(.Pos.X, .Pos.Y), Sound.Calculate_Pan(.Pos.X, .Pos.Y))
                    Else
                        Call Sound.Sound_Play(Pasos(TerrenoDePaso).Wav(2), , Sound.Calculate_Volume(.Pos.X, .Pos.Y), Sound.Calculate_Pan(.Pos.X, .Pos.Y))
                    End If
            End If
        End If
    End With
End Sub

Public Sub InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer)
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
'Configures the engine to start running.
'***************************************************

On Error GoTo ErrorHandler:

    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = Round(frmMain.MainViewPic.Height / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.Width / 32, 0)
    
    IniPath = Carga.Path(Script)
    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    Call CalcularAreas(HalfWindowTileWidth, HalfWindowTileHeight)

    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

    Exit Sub
    
ErrorHandler:

    Call RegistrarError(Err.number, Err.Description, "Mod_TileEngine.InitTileEngine")
    
    Call CloseClient
    
End Sub

Public Sub LoadGraphics()
    Call SurfaceDB.Initialize(DirectD3D8, ClientSetup.byMemory)
End Sub

Sub ShowNextFrame(ByVal MouseViewX As Integer, _
                  ByVal MouseViewY As Integer)

    On Error GoTo ErrorHandler:

    If EngineRun Then
        
        Call Engine_BeginScene
            
        Call DesvanecimientoTechos
        Call DesvanecimientoMsg
            
        If UserMoving Then
            
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                LastOffset2X = ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame * charlist(UserCharIndex).Speeding
                OffsetCounterX = OffsetCounterX - LastOffset2X

                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    LastOffset2X = 0
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If

            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                LastOffset2Y = ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame * charlist(UserCharIndex).Speeding
                OffsetCounterY = OffsetCounterY - LastOffset2Y

                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    LastOffset2Y = 0
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
    
        End If
            
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
            
        '****** Update screen ******
        If CurrentUser.UserCiego Then
            Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
        Else
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        End If
            
        Call RenderHUD
    
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_BaseSpeed
            
        Call Engine_EndScene(MainScreenRect, 0)
        
        Call Inventario.DrawDragAndDrop
    
    End If
    
ErrorHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        
        Call mDx8_Engine.Engine_DirectX8_Init
        
        Call LoadGraphics
    
    End If
  
End Sub

Public Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 16/09/2010 (Zama)
    'Draw char's to screen without offcentering them
    '16/09/2010: ZaMa - Ya no se dibujan los bodies cuando estan invisibles.
    '***************************************************
    
    Dim moved As Boolean
    Dim AuraColorFinal(3) As RGBA
    Dim ColorFinal(3) As RGBA
    Dim TempGrh As Grh
        
    With charlist(CharIndex)
        
        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * .Speeding
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * .Speeding
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        If .Heading = 0 Then Exit Sub
        
        'if is attacking we set the attack anim
        If .attacking And .Arma.WeaponWalk(.Heading).Started = 0 Then
            .Arma.WeaponWalk(.Heading).Started = 1
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
        
            'if the anim has ended or we are no longer attacking end the animation
        ElseIf .Arma.WeaponWalk(.Heading).FrameCounter > 4 And .attacking Then
            .attacking = False 'this is just for testing, it shouldnt be done here

        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            
            '//Evito runtime
            If Not .Heading <> 0 Then .Heading = EAST
            
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            '//Movimiento del arma y el escudo
            If Not .Movement And Not .attacking Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
                
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1

            End If
            
            If .NPCAttack = False Then
                .Ataque.AtaqueWalk(.Heading).Started = 0
                .Ataque.AtaqueWalk(.Heading).FrameCounter = 1
            End If
            
            .Moving = False
            
        Else
            .NPCAttack = False

        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If Not .muerto Then
            Call Copy_RGBAList(MapData(.Pos.X, .Pos.Y).Light_Value(), ColorFinal())

        Else

            If esGM(Val(CharIndex)) Then
                Call RGBAList(ColorFinal(), 200, 200, 0, 150)
            Else

                If .Criminal Then
                    Call RGBAList(ColorFinal(), 255, 100, 100, 100)
                Else
                    Call RGBAList(ColorFinal(), 128, 255, 255, 100)
                End If

            End If

        End If
                
        Movement_Speed = 0.5
                
        If Not .invisible Then
        
            'Sombras
            If ClientSetup.UsarSombras And .AuraAnim.GrhIndex = 0 Then Call RenderSombras(CharIndex, PixelOffsetX, PixelOffsetY)
            
            'Auras
            If .AuraAnim.GrhIndex > 0 And ClientSetup.UsarAuras Then
                Call Long_2_RGBAList(AuraColorFinal(), .AuraColor)
                Call Draw_Grh(.AuraAnim, PixelOffsetX, PixelOffsetY + 35, 1, AuraColorFinal(), 1, True)
            End If
            
            If .NPCAttack = True And .Ataque.AtaqueWalk(.Heading).GrhIndex > 0 Then
                Call Draw_Grh(.Ataque.AtaqueWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal, 1)
                
            Else
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal, 1)
                    
            End If
            
            'Draw name when navigating
            If Len(.Nombre) > 0 Then
                If Nombres Then
                    If .iHead = 0 And .iBody > 0 Then
                        Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
                End If
            End If
            
            'Draw Head
            If .Head Then _
                Call DrawHead(.Head, PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, ColorFinal(), .Heading, True)
                
            'Draw Helmet
            If .Casco Then _
                Call DrawHead(.Casco, PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, ColorFinal(), .Heading, False)
                
            'Draw Weapon
            If .Arma.WeaponWalk(.Heading).GrhIndex Then
                Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
            End If
                
            'Draw Shield
            If .Escudo.ShieldWalk(.Heading).GrhIndex Then
                Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1)
            End If
            
            If ClientSetup.ParticleEngine Then

                Call RenderCharParticles(CharIndex, PixelOffsetX + 17, PixelOffsetY + 10)

            End If
            
            'Draw name over head
            If LenB(.Nombre) > 0 Then
                If Nombres Then
                    Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                End If
            End If
            
        ElseIf CharIndex = UserCharIndex Or (.Clan <> vbNullString And .Clan = charlist(UserCharIndex).Clan) Then
            
            'Auras
            If .AuraAnim.GrhIndex > 0 And ClientSetup.UsarAuras Then
                Call Long_2_RGBAList(AuraColorFinal(), .AuraColor)
                Call Draw_Grh(.AuraAnim, PixelOffsetX, PixelOffsetY + 35, 1, AuraColorFinal(), 1, True)
            End If
            
            'Draw Transparent Body
            If .Body.Walk(.Heading).GrhIndex Then
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
            End If

            'Draw Transparent Head
            If .Head Then _
                Call DrawHead(.Head, PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, ColorFinal(), .Heading, True, True)
                
            'Draw Transparent Helmet
            If .Casco Then _
                Call DrawHead(.Casco, PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, ColorFinal(), .Heading, False, True)
                
            'Draw Transparent Weapon
            If .Arma.WeaponWalk(.Heading).GrhIndex Then
                Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
            End If
                
            'Draw Transparent Shield
            If .Escudo.ShieldWalk(.Heading).GrhIndex Then
                Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, True)
            End If
            
            If LenB(.Nombre) > 0 Then
                If Nombres Then
                    Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY, True)
                End If
            End If
            
        End If
        
        'Update dialogs - 34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo.
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex)
        
        Movement_Speed = 1
        
        'Draw FX
        If .FxIndex <> 0 Then
            Call Draw_Grh(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, MapData(.Pos.X, .Pos.Y).Light_Value(), 1, True)
            
            'Check if animation is over
            If .fX.Started = 0 Then .FxIndex = 0
            
        End If
        
        '************Draw Pasos************
        If CharIndex = UserCharIndex Then
            If Not CurrentUser.UserEquitando Then
                If MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 7704 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 7719 Or _
                    MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 1315 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 1330 Or _
                        MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex >= 30120 And MapData(.Pos.X, .Pos.Y).Graphic(1).GrhIndex <= 30439 Then
                    
                    Select Case .Heading
                    
                        Case E_Heading.SOUTH 'Arriba
                            Call General_Particle_Create(65, .Pos.X, .Pos.Y, 250)
                            
                        Case E_Heading.NORTH  'Abajo
                            Call General_Particle_Create(122, .Pos.X, .Pos.Y, 250)
                            
                        Case E_Heading.WEST  'Izquierda
                            Call General_Particle_Create(124, .Pos.X, .Pos.Y, 250)
                            
                        Case E_Heading.EAST  'Derecha
                            Call General_Particle_Create(123, .Pos.X, .Pos.Y, 250)
                        
                    End Select
                End If
            End If
        End If
        
        '¿Tiene quest?
        If .EstadoQuest <> eStatusQuest.NoTieneQuest Then
            Dim GrhQuest As Long
            
            Select Case .EstadoQuest
                Case eStatusQuest.NoAceptada
                    GrhQuest = 558
                Case eStatusQuest.EnCurso
                    GrhQuest = 559
                Case eStatusQuest.Terminada
                    GrhQuest = 2637
            End Select
            
            Call Draw_GrhIndex(GrhQuest, PixelOffsetX, PixelOffsetY + OFFSET_HEAD - 23, 1, ColorFinal())
        End If
        
        'Barra de tiempo
        If .BarTime < .MaxBarTime And Not .invisible Then
            Call InitGrh(TempGrh, 4637)

            Call Draw_Grh(TempGrh, PixelOffsetX + 1 + .Body.HeadOffset.X, PixelOffsetY - 55 + .Body.HeadOffset.Y, 1, COLOR_WHITE(), False)

            Engine_Draw_Box PixelOffsetX + 5 + .Body.HeadOffset.X, PixelOffsetY - 28 + .Body.HeadOffset.Y, .BarTime / .MaxBarTime * 26, 4, D3DColorARGB(3, 214, 166, 120) ', RGBA_From_Comp(0, 0, 0, 255)

            .BarTime = .BarTime + (timerElapsedTime / 1000)
            'Debug.Print .BarTime
            If .BarTime >= .MaxBarTime Then
                charlist(CharIndex).BarTime = 0
                charlist(CharIndex).BarAccion = 99
                charlist(CharIndex).MaxBarTime = 0
            End If

        End If
        
        
    End With
    
End Sub

Private Sub RenderSombras(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'****************************************************
' Renderizamos las sombras sobre el char
'****************************************************
   
    With charlist(CharIndex)
        
        'Shadow Body & Shadow Head
        
        If (.iHead > 0) And (.iBody = 617 Or .iBody = 612 Or .iBody = 614 Or .iBody = 616) Then
        
            'Si estÃ¡ montando se dibuja de esta manera
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 8, PixelOffsetY - 14, 1, COLOR_SHADOW(), 0, False, 187, 1, 1.2) ' Shadow Body
            Call DrawHead(.Head, PixelOffsetX + .Body.HeadOffset.X + 12, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD - 13, COLOR_SHADOW(), .Heading, True, False, 187, 1, 1.2)
            Call DrawHead(.Casco, PixelOffsetX + .Body.HeadOffset.X + 12, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD - 13, COLOR_SHADOW(), .Heading, False, False, 187, 1, 1.2)
        
        'Si estÃ¡ navegando se dibuja de esta manera
        ElseIf ((.iHead = 0) And (HayAgua(.Pos.X, .Pos.Y + 1) Or HayAgua(.Pos.X + 1, .Pos.Y) Or HayAgua(.Pos.X, .Pos.Y - 1) Or HayAgua(.Pos.X - 1, .Pos.Y))) Then
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 5, PixelOffsetY - 26, 1, COLOR_SHADOW(), 0, False, 186, 1, 1.33) ' Shadow Body
        
        Else
        
            'Si NO estÃ¡ montando ni navegando se dibuja de esta manera
            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 8, PixelOffsetY - 11, 1, COLOR_SHADOW(), 0, False, 195, 1, 1.2) ' Shadow Body
            Call DrawHead(.Head, PixelOffsetX + .Body.HeadOffset.X + 12, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD - 13, COLOR_SHADOW(), .Heading, True, False, 195, 1, 1.2)
            Call DrawHead(.Casco, PixelOffsetX + .Body.HeadOffset.X + 12, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD - 13, COLOR_SHADOW(), .Heading, False, False, 195, 1, 1.2)

        End If
                
        'Shadow Weapon
        If .Arma.WeaponWalk(.Heading).GrhIndex Then
            Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX + 9, PixelOffsetY - 12, 1, COLOR_SHADOW(), 0, False, 195, 1, 1.2)
        End If
                
        'Shadow Shield
        If .Escudo.ShieldWalk(.Heading).GrhIndex Then
            Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX + 9, PixelOffsetY - 12, 1, COLOR_SHADOW(), 0, False, 195, 1, 1.2)
        End If
        
    End With
    
End Sub

Private Sub RenderCharParticles(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'****************************************************
' Renderizamos las particulas fijadas en el char
'****************************************************

    Dim i As Integer
    
    With charlist(CharIndex)

        If .Particle_Count > 0 Then

            For i = 1 To .Particle_Count
                        
                If .Particle_Group(i) > 0 Then
                    Call mDx8_Particulas.Particle_Group_Render(.Particle_Group(i), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY)
                End If
                            
            Next i
    
        End If
    
    End With
  
End Sub

Private Sub RenderName(ByVal CharIndex As Long, _
                       ByVal X As Integer, _
                       ByVal Y As Integer, _
                       Optional ByVal Invi As Boolean = False)
    Dim Pos      As Integer
    Dim line     As String
    Dim color(3) As RGBA
   
    With charlist(CharIndex)
        Pos = getTagPosition(.Nombre)
    
        If .priv = 0 Then
            If .muerto Then
                Call RGBAList(color, 220, 220, 255, 255)
                
            Else

                If .WorldBoss = True Then 'WorldBoss
                    Call RGBAList(color, ColoresPJ(8).R, ColoresPJ(8).G, ColoresPJ(8).B)
                    
                Else

                    If .Criminal Then 'Criminal
                        Call RGBAList(color, ColoresPJ(50).R, ColoresPJ(50).G, ColoresPJ(50).B)
                        
                    Else 'Ciudadano
                        Call RGBAList(color, ColoresPJ(49).R, ColoresPJ(49).G, ColoresPJ(49).B)
                        
                    End If
                End If
            End If
        Else
            Call RGBAList(color, ColoresPJ(.priv).R, ColoresPJ(.priv).G, ColoresPJ(.priv).B)
            
        End If
    
        If Invi Then _
            Call RGBAList(color, 150, 180, 220, 180)

        'Nick
        line = Left$(.Nombre, Pos - 2)
        Call DrawText(X + 16, Y + 30, line, color, True)
            
        'Clan
        
        If .priv = 2 Or .priv = 3 Then
            line = "<Game Master>"
        ElseIf .priv = 4 Then
            line = "<Administrador>"
        Else
            line = mid$(.Nombre, Pos)
        End If
        
        Call DrawText(X + 16, Y + 45, line, color, True)

    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
            .fX.Loops = Loops
        End If
        
    End With
    
End Sub

Public Sub Device_Textured_Render(ByVal X As Single, ByVal Y As Single, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  ByVal tex As Long, _
                                  ByRef color() As RGBA, _
                                  Optional ByVal Alpha As Boolean = False, _
                                  Optional ByVal angle As Single = 0, _
                                  Optional ByVal ScaleX As Single = 1!, _
                                  Optional ByVal ScaleY As Single = 1!)

        Dim Texture As Direct3DTexture8
        
        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
        
        With SpriteBatch

                Call .SetTexture(Texture)
                    
                Call .SetAlpha(Alpha)
                
                If TextureWidth <> 0 And TextureHeight <> 0 Then
                    Call .Draw(X, Y, Width * ScaleX, Height * ScaleY, color, sX / TextureWidth, sY / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
                Else
                    Call .Draw(X, Y, TextureWidth * ScaleX, TextureHeight * ScaleY, color, , , , , angle)
                End If
                
        End With
        
End Sub

Public Sub RenderItem(ByVal hWndDest As Long, ByVal GrhIndex As Long)
    Dim DR As RECT
    
    With DR
        .Left = 0
        .Top = 0
        .Right = 32
        .Bottom = 32
    End With
    
    Call Engine_BeginScene

    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, COLOR_WHITE(), 0, False)
        
    Call Engine_EndScene(DR, hWndDest)
    
End Sub

Sub Draw_GrhIndex(ByVal GrhIndex As Long, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As RGBA, Optional ByVal angle As Single = 0, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        'Draw
        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha)
    End With
    
End Sub

Sub Draw_Grh(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As RGBA, ByVal Animate As Byte, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    
On Error GoTo Error

    Dim CurrentGrhIndex As Long
    
    If Grh.GrhIndex = 0 Then Exit Sub
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.speed) * Movement_Speed

            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - (.pixelWidth * ScaleX - TilePixelWidth) \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        Call Device_Textured_Render(X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color_List(), Alpha, angle, ScaleX, ScaleY)
        
    End With
    
Exit Sub

Error:
    If Err.number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        'Call Log_Engine("Error in Draw_Grh, " & Err.Description & ", (" & Err.number & ")")
        MsgBox "Error en el Engine Grafico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Debug.Print "Error en el Engine Grafico. Grh: " & Grh.GrhIndex
        Call CloseClient
    End If
End Sub

Public Sub DrawHead(ByVal Head As Integer, ByVal X As Integer, ByVal Y As Integer, Light() As RGBA, ByVal Heading As Byte, Optional ByVal EsCabeza As Boolean = True, Optional ByVal Alpha As Boolean = False, Optional ByVal angle As Single = 0, Optional ByVal ScaleX As Single = 1!, Optional ByVal ScaleY As Single = 1!)

    Dim textureX1 As Integer
    Dim textureX2 As Integer
    Dim textureY1 As Integer
    Dim textureY2 As Integer
    Dim OffsetX As Integer
    Dim OffsetY As Integer
    Dim Texture As Long

    If EsCabeza Then
        If heads(Head).Texture <= 0 Then Exit Sub
        Texture = heads(Head).Texture
    Else
        If Cascos(Head).Texture <= 0 Then Exit Sub
        Texture = Cascos(Head).Texture
    End If
    
    textureX2 = 27
    textureY2 = 32
 
    If EsCabeza Then
        textureX1 = heads(Head).startX - textureX2
        textureY1 = ((Heading - 2) * textureY2) + heads(Head).startY
    Else
        textureX1 = Cascos(Head).startX - textureX2 + 1
        textureY1 = ((Heading - 2) * textureY2) + Cascos(Head).startY + 2
    End If
    
    Device_Textured_Render X - OffsetX + 3, Y - OffsetY + 4, textureX2, textureY2, (textureX2 + textureX1), (textureY2 + textureY1), Texture, Light, Alpha, angle, ScaleX, ScaleY

End Sub

Public Function GrhCheck(ByVal GrhIndex As Long) As Boolean
        '**************************************************************
        'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
        'Last Modify Date: 1/04/2003
        '
        '**************************************************************
        'check grh_index

        If GrhIndex > 0 And GrhIndex <= UBound(GrhData()) Then
                GrhCheck = GrhData(GrhIndex).NumFrames
        End If

End Function

Public Sub GrhUninitialize(Grh As Grh)
        '*****************************************************************
        'Author: Aaron Perkins
        'Last Modify Date: 1/04/2003
        'Resets a Grh
        '*****************************************************************

        With Grh
        
                'Copy of parameters
                .GrhIndex = 0
                .Started = False
                .Loops = 0
        
                'Set frame counters
                .FrameCounter = 0
                .speed = 0
                
        End With

End Sub

Private Sub DesvanecimientoTechos()
 
    If bTecho Then
        If Not Val(ColorTecho) = 50 Then ColorTecho = ColorTecho - 1
    Else
        If Not Val(ColorTecho) = 250 Then ColorTecho = ColorTecho + 1
    End If
    
    If Not Val(ColorTecho) = 250 Then
        Call RGBAList(temp_rgb(), ColorTecho, ColorTecho, ColorTecho, ColorTecho)
    End If
    
End Sub

Public Sub DesvanecimientoMsg()
'*****************************************************************
'Author: FrankoH
'Last Modify Date: 04/09/2019
'DESVANECIMIENTO DE LOS TEXTOS DEL RENDER
'*****************************************************************
    Static lastmovement As Long
    
    If GetTickCount - lastmovement > 1 Then
        lastmovement = GetTickCount
    Else
        Exit Sub
    End If

    If LenB(renderText) Then
        If Not Val(colorRender) = 0 Then colorRender = colorRender - 1
    ElseIf LenB(renderText) = 0 Then
        Exit Sub
    Else
        If Not Val(colorRender) = 240 Then colorRender = colorRender + 1
    End If
    
    If Not Val(colorRender) = 240 Then
        Call RGBAList(render_msg(), 255, 255, 255, colorRender)
    End If
    
    If colorRender = 0 Then renderMsgReset
    
End Sub

Public Sub renderMsgReset()

    renderFont = 1
    renderText = vbNullString
    renderTextPk = vbNullString

End Sub

Public Function Char_Pos_Get(ByVal char_index As Integer, ByRef map_x As Integer, ByRef map_y As Integer) As Boolean
'************************************
'Autor: Lorwik
'Fecha: ???
'***********************************

   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        map_x = charlist(char_index).Pos.X
        map_y = charlist(char_index).Pos.Y
        Char_Pos_Get = True
    End If
End Function
