Attribute VB_Name = "ModAreas"
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

' WyroX: Pequenia modificacion para que el tamanio de las areas se calcule automaticamente
' en base al tamanio del render y de un valor arbitrario para el buffer (tiles extra)

Option Explicit

' Tamanio de las areas
Private AreasX As Byte
Private AreasY As Byte

' Area actual
Private CurAreaX As Integer
Private CurAreaY As Integer

Public DebugAreas As Boolean

Public Sub CalcularAreas(HalfWindowTileWidth As Integer, HalfWindowTileHeight As Integer)
    AreasX = HalfWindowTileWidth + TileBufferSize
    AreasY = HalfWindowTileHeight + TileBufferSize
End Sub

' Elimina todo fuera del area del usuario
Public Sub CambioDeArea(ByVal x As Integer, _
                        ByVal y As Integer, _
                        ByVal Heading As E_Heading)

    Dim CharIndex As Integer
    
    Dim MinX As Integer
    Dim MaxX As Integer
    Dim MinY As Integer
    Dim MaxY As Integer
    
    Dim loopX     As Integer
    Dim loopY     As Integer
    
    ' Calculamos el area actual al que pertenece
    CurAreaX = x \ AreasX
    CurAreaY = y \ AreasY
    
    For loopX = UserPos.x - 50 To UserPos.x + 50
        For loopY = UserPos.y - 50 To UserPos.y + 50

            ' Si el tile esta fuera del area
            If Not EstaDentroDelArea(loopX, loopY) Then

                ' Borrar char
                CharIndex = Char_MapPosExits(loopX, loopY)

                If (CharIndex > 0) Then
                    If (CharIndex <> UserCharIndex) Then
                        Call Char_Erase(CharIndex)
                    End If
                End If

                ' Borrar objeto
                If (Map_PosExitsObject(loopX, loopY) > 0) Then
                    Call Map_DestroyObject(loopX, loopY)
                End If

            End If

        Next loopY
    Next loopX

End Sub

Public Sub CalcularArea(ByVal Heading As Byte, ByRef MinX As Integer, ByRef MaxX As Integer, ByRef MinY As Integer, ByRef MaxY As Integer)

    Dim MinAreaX  As Integer
    Dim MinAreaY  As Integer
    Dim MaxAreaX  As Integer
    Dim MaxAreaY  As Integer

    Select Case Heading

        Case E_Heading.SOUTH
            ' 3 areas nuevas arriba
            MinAreaX = CurAreaX - 1
            MinAreaY = CurAreaY - 1
            MaxAreaX = CurAreaX + 1
            MaxAreaY = CurAreaY - 1

        Case E_Heading.WEST
            ' 3 areas nuevas a la derecha
            MinAreaX = CurAreaX + 1
            MinAreaY = CurAreaY - 1
            MaxAreaX = CurAreaX + 1
            MaxAreaY = CurAreaY + 1

        Case E_Heading.NORTH
            ' 3 areas nuevas abajo
            MinAreaX = CurAreaX - 1
            MinAreaY = CurAreaY + 1
            MaxAreaX = CurAreaX + 1
            MaxAreaY = CurAreaY + 1

        Case E_Heading.EAST
            ' 3 areas nuevas a la izquierda
            MinAreaX = CurAreaX - 1
            MinAreaY = CurAreaY - 1
            MaxAreaX = CurAreaX - 1
            MaxAreaY = CurAreaY + 1

        Case Else ' Heading = USER_NUEVO (cambio de mapa, recien logueado, etc.)
            ' 9 areas nuevas alrededor del usuario (3x3)
            MinAreaX = CurAreaX - 1
            MinAreaY = CurAreaY - 1
            MaxAreaX = CurAreaX + 1
            MaxAreaY = CurAreaY + 1

    End Select

    ' Convertimos de areas a tiles
    MinX = MinAreaX * AreasX
    MinY = MinAreaY * AreasY
    MaxX = (MaxAreaX + 1) * AreasX - 1
    MaxY = (MaxAreaY + 1) * AreasY - 1

    ' Comprobamos que este dentro del mapa
    If MinX < XMinMapSize Then MinX = XMinMapSize
    If MinY < YMinMapSize Then MinY = YMinMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    'Debug.Print "MinX: " & MinX & " MaxX: " & MaxX & " MinY: " & MinY & " MaxY: " & MaxY & " Heading: " & Heading

End Sub

' Calcula si la posicion se encuentra dentro del area del usuario
Public Function EstaDentroDelArea(ByVal x As Integer, ByVal y As Integer) As Boolean
    EstaDentroDelArea = (Abs(CurAreaX - x \ AreasX) <= 1) And (Abs(CurAreaY - y \ AreasY) <= 1)
End Function

Public Sub LimpiarAreas()
    Dim loopX As Integer
    Dim loopY As Integer
    Dim CharIndex As Integer
    
    For loopX = XMinMapSize To XMaxMapSize
        For loopY = YMinMapSize To YMaxMapSize

            ' Borrar char
            CharIndex = Char_MapPosExits(loopX, loopY)

            If (CharIndex > 0) Then
                Call Char_Erase(CharIndex)
            End If

            ' Borrar objeto
            If (Map_PosExitsObject(loopX, loopY) > 0) Then
                Call Map_DestroyObject(loopX, loopY)
            End If

        Next loopY
    Next loopX

    Debug.Print "Limpieza de areas completas"
    
End Sub
