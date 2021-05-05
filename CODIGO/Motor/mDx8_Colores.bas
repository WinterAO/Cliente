Attribute VB_Name = "mDx8_Colores"
Option Explicit

' Desvanecimiento en Techos
Public ColorTecho As Byte
Public temp_rgb(3) As Long

' Titulos en el render (Nombre de mapa, subida de lvl, etc)
Public renderText As String
Public renderTextPk As String
Public renderFont As Integer
Public colorRender As Byte
Public render_msg(3) As Long

'Colores de PJ (nicks y demas)
Public Const MAXCOLORES As Byte = 56
Public ColoresPJ(0 To MAXCOLORES) As Long

'Colores del mapa
Public Normal_RGBList(3) As Long
Public Color_Shadow(3) As Long
Public NoUsa_RGBList(3) As Long
Public Color_Arbol(3) As Long
Public Color_Paralisis As Long

Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef Color As D3DCOLORVALUE)
    Dim dest(3) As Byte
    CopyMemory dest(0), ARGB, 4
    Color.a = dest(3)
    Color.r = dest(2)
    Color.g = dest(1)
    Color.b = dest(0)
End Function

Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal a As Long) As Long
        
    Dim c As Long
        
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    
    ARGB = c

End Function

Public Sub Engine_D3DColor_To_RGB_List(rgb_list() As Long, Color As D3DCOLORVALUE)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Set a D3DColorValue to a RGB List
'***************************************************
    rgb_list(0) = D3DColorARGB(Color.a, Color.r, Color.g, Color.b)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
'***************************************************
'Author: Ezequiel Juarez (Standelf)
'Last Modification: 16/05/10
'Blisse-AO | Set a Long Color to a RGB List
'***************************************************
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Sub ConvertLongToRGB(ByVal value As Long, r As Byte, g As Byte, b As Byte)
    r = value Mod 256
    g = Int(value / 256) Mod 256
    b = Int(value / 256 / 256) Mod 256
End Sub

Public Function SetARGB_Alpha(rgb_list() As Long, Alpha As Byte) As Long()

    '***************************************************
    'Author: Juan Manuel Couso (Cucsifae)
    'Last Modification: 29/08/18
    'Obtiene un ARGB list le modifica el alpha y devuelve una copia
    '***************************************************
    Dim TempColor        As D3DCOLORVALUE
    Dim tempARGB(0 To 3) As Long

    'convertimos el valor del rgb list a D3DCOLOR
    Call ARGBtoD3DCOLORVALUE(rgb_list(1), TempColor)

    'comprobamos ue no se salga del rango permitido
    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    
    'seteamos el alpha
    TempColor.a = Alpha
    
    'generamos el nuevo RGB_List
    Call Engine_D3DColor_To_RGB_List(tempARGB(), TempColor)

    SetARGB_Alpha = tempARGB()

End Function

