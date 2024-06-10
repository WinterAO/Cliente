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

' Cantidad de tiles buffer
' (para que graficos grandes se vean desde fuera de la pantalla)
' (debe coincidir con el mismo valor en el server - areas)
Public Const TilesBuffer As Byte = 8

' Tamanio de las areas
Private AreasX As Byte
Private AreasY As Byte

' Area actual
Private CurAreaX As Integer
Private CurAreaY As Integer

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public Const MargenX As Integer = 16
Public Const MargenY As Integer = 14

Public Sub CalcularAreas(HalfWindowTileWidth As Integer, HalfWindowTileHeight As Integer)
    AreasX = HalfWindowTileWidth + TileBufferSize
    AreasY = HalfWindowTileHeight + TileBufferSize
End Sub

' Elimina todo fuera del area del usuario
Public Sub CambioDeArea(ByVal x As Integer, ByVal y As Integer, ByVal Head As Byte)

    Dim loopX     As Integer
    Dim loopY     As Integer
    Dim CharIndex As Integer
    Dim MinX      As Integer
    Dim MinY      As Integer
    Dim MaxX      As Integer
    Dim MaxY      As Integer

    CurAreaX = x \ AreasX
    CurAreaY = y \ AreasY

    MinX = x
    MinY = y
    MaxX = x
    MaxY = y

    Select Case Head
    
        Case E_Heading.SOUTH
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY - MargenY - 1
            MaxY = MinY
            
        Case Head = E_Heading.NORTH
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY + MargenY + 1
            MaxY = MinY
        
        Case Head = E_Heading.EAST
            MinX = MinX - MargenX - 1
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        
        Case Head = E_Heading.WEST
            MinX = MinX + MargenX + 1
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
    
    End Select
    
    If MinY < 1 Then MinY = 1
    If MinX < 1 Then MinX = 1
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize

    ' Recorremos el mapa entero (TODO: Se puede optimizar si el server nos enviara la direccion del area que nos movimos)
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
    EstaDentroDelArea = (Abs(CurAreaX - x \ AreasX) <= 1) And (Abs(CurAreaY - y \ AreasY) <= 1)
End Function
