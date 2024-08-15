Attribute VB_Name = "Areas"
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

Public Sub CalcularAreas(HalfWindowTileWidth As Integer, HalfWindowTileHeight As Integer)
    AreasX = HalfWindowTileWidth + TileBufferSize
    AreasY = HalfWindowTileHeight + TileBufferSize
End Sub

' Elimina todo fuera del area del usuario
Public Sub CambioDeArea(ByVal x As Integer, ByVal y As Integer, ByVal Heading As E_Heading)

    Dim loopX     As Integer
    Dim loopY     As Integer
    Dim CharIndex As Integer
    Dim MinX      As Integer
    Dim MinY      As Integer
    Dim MaxX      As Integer
    Dim MaxY      As Integer

    CurAreaX = x \ AreasX
    CurAreaY = y \ AreasY
    
    Select Case Heading

        Case E_Heading.SOUTH
            MinX = x - AreasX
            MaxX = x + AreasX
            MinY = y - AreasY - 1
            MaxY = y

        Case E_Heading.NORTH
            MinX = x - AreasX
            MaxX = x + AreasX
            MinY = y + AreasY + 1
            MaxY = y

        Case E_Heading.EAST
            MinX = x - AreasX - 1
            MaxX = x
            MinY = y - AreasY
            MaxY = y + AreasY

        Case E_Heading.WEST
            MinX = x + AreasX + 1
            MaxX = x
            MinY = y - AreasY
            MaxY = y + AreasY

    End Select

    If MinX > MaxX Then
        Dim tempX As Integer
        tempX = MinX
        MinX = MaxX
        MaxX = tempX
    End If

    If MinY > MaxY Then
        Dim tempY As Integer
        tempY = MinY
        MinY = MaxY
        MaxY = tempY
    End If

    If MinX < 1 Then MinX = 1
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize

    If MinY < 1 Then MinY = 1
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    Debug.Print "MinX: " & MinX & " MaxX: " & MaxX & " MinY: " & MinY & " MaxY: " & MaxY & " Heading: " & Heading
    
    For loopX = MinX To MaxX
        For loopY = MinY To MaxY

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

' Calcula si la posicion se encuentra dentro del area del usuario
Public Function EstaDentroDelArea(ByVal x As Integer, ByVal y As Integer) As Boolean
    ' Calcula si está dentro del área
    EstaDentroDelArea = (Abs(CurAreaX - x \ AreasX) <= 1) And (Abs(CurAreaY - y \ AreasY) <= 1)
    
    ' Log para depuración
    'Debug.Print "EstaDentroDelArea - Posición: (" & x & ", " & y & ") - Resultado: " & EstaDentroDelArea & " (CurAreaX: " & CurAreaX & ", CurAreaY: " & CurAreaY & ")"
End Function

Public Sub LimpiarArea()
Dim x As Integer
Dim y As Integer

    For x = UserPos.x - AreasX * 2 To UserPos.x + AreasX * 2
        For y = UserPos.y - AreasY * 2 To UserPos.y + AreasY * 2
        
            If InMapBounds(x, y) Then
                If MapData(x, y).CharIndex > 0 Then
                    If MapData(x, y).CharIndex <> UserCharIndex Then
                        Call Char_Erase(MapData(x, y).CharIndex)
                    End If
                End If

                ' Borrar objeto
                If (Map_PosExitsObject(x, y) > 0) Then
                    Call Map_DestroyObject(x, y)
                End If
            End If
        Next y
    Next x
End Sub
