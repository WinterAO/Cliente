Attribute VB_Name = "mDx8_Luces"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Light Engine, Read the _
    #LightEngine to Set the type of Lights
'***************************************************

Option Base 0

Private Type tLight
    RGBcolor As D3DCOLORVALUE
    active As Boolean
    map_x As Integer
    map_y As Integer
    range As Byte
End Type
 
Private Light_List() As tLight
Private NumLights As Integer

Public Function Create_Light_To_Map(ByVal map_x As Integer, ByVal map_y As Integer, Optional range As Byte = 3, Optional ByVal Red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255, Optional ByVal Render As Boolean = True)
    NumLights = NumLights + 1
   
    ReDim Preserve Light_List(1 To NumLights) As tLight
   
    Light_List(NumLights).RGBcolor.r = Red
    Light_List(NumLights).RGBcolor.g = Green
    Light_List(NumLights).RGBcolor.b = Blue
    Light_List(NumLights).RGBcolor.a = 255
    Light_List(NumLights).range = range
    Light_List(NumLights).active = True
    Light_List(NumLights).map_x = map_x
    Light_List(NumLights).map_y = map_y
   
    If Render Then _
        Call LightRender(NumLights)
End Function

Public Function Delete_Light_To_Map(ByVal X As Integer, ByVal Y As Integer)
   
    Dim i As Long
   
    For i = 1 To NumLights
        If Light_List(i).map_x = X And Light_List(i).map_y = Y Then
            Delete_Light_To_Index i
            
            Exit Function
        End If
    Next i
 
End Function

#If LightEngine = 1 Then '   Luces Radiales

Public Function Delete_Light_To_Index(ByVal light_index As Integer, Optional ByVal Actualizar As Boolean = True)
'************************************
'Autor: Lorwik
'Fecha: 14/08/2020
'Descripción: Primero desactivamos una luz concreta y luego reordenamos el array
'************************************

    'Borramos la luz
    Light_List(light_index).active = False
    
    'Reordamos el Aray
    If light_index = NumLights Then
    
        Do Until Light_List(NumLights).active
            NumLights = NumLights - 1
            If NumLights = 0 Then
                Exit Function
            End If
        Loop
        ReDim Preserve Light_List(1 To NumLights)
    
    End If
 
    If Actualizar Then _
        Call Actualizar_Estado(Estado_Actual_Date)

End Function

Private Sub LightRender(ByVal light_index As Integer)
 
 On Local Error Resume Next
 
    If light_index = 0 Then Exit Sub
    If Light_List(light_index).active = False Then Exit Sub
   
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Ya As Integer
    Dim Xa As Integer
   
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
   
    Dim XCoord As Integer
    Dim YCoord As Integer
   
    AmbientColor.a = Estado_Actual.a
    AmbientColor.r = Estado_Actual.r
    AmbientColor.g = Estado_Actual.g
    AmbientColor.b = Estado_Actual.b

    LightColor = Light_List(light_index).RGBcolor
       
    min_x = Light_List(light_index).map_x - Light_List(light_index).range
    max_x = Light_List(light_index).map_x + Light_List(light_index).range
    min_y = Light_List(light_index).map_y - Light_List(light_index).range
    max_y = Light_List(light_index).map_y + Light_List(light_index).range
    
    Dim TEMP_COLOR As D3DCOLORVALUE
    
    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                XCoord = Xa
                YCoord = Ya
                Call Engine_Get_ARGB(MapData(Xa, Ya).Engine_Light(0), TEMP_COLOR)
                
                MapData(Xa, Ya).Engine_Light(0) = LightCalculate(Light_List(light_index).range, Light_List(light_index).map_x, Light_List(light_index).map_y, XCoord, YCoord, MapData(Xa, Ya).Engine_Light(0), LightColor, TEMP_COLOR)
 
                XCoord = Xa + 1
                YCoord = Ya
                Call Engine_Get_ARGB(MapData(Xa, Ya).Engine_Light(3), TEMP_COLOR)
                                
                MapData(Xa, Ya).Engine_Light(3) = LightCalculate(Light_List(light_index).range, Light_List(light_index).map_x, Light_List(light_index).map_y, XCoord, YCoord, MapData(Xa, Ya).Engine_Light(3), LightColor, TEMP_COLOR)
                       
                XCoord = Xa
                YCoord = Ya + 1
                Call Engine_Get_ARGB(MapData(Xa, Ya).Engine_Light(1), TEMP_COLOR)
                MapData(Xa, Ya).Engine_Light(1) = LightCalculate(Light_List(light_index).range, Light_List(light_index).map_x, Light_List(light_index).map_y, XCoord, YCoord, MapData(Xa, Ya).Engine_Light(1), LightColor, TEMP_COLOR)
   
                XCoord = Xa
                YCoord = Ya
                Call Engine_Get_ARGB(MapData(Xa, Ya).Engine_Light(2), TEMP_COLOR)
                MapData(Xa, Ya).Engine_Light(2) = LightCalculate(Light_List(light_index).range, Light_List(light_index).map_x, Light_List(light_index).map_y, XCoord, YCoord, MapData(Xa, Ya).Engine_Light(2), LightColor, TEMP_COLOR)
               
            End If
        Next Xa
    Next Ya
End Sub

Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoordenadas As Integer, ByVal YCoordenadas As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim DistanciaX As Single
    Dim DistanciaY As Single
    Dim DistanciaVertex As Single
    Dim Radio As Integer
    
    Dim CurrentColor As D3DCOLORVALUE
    
    Radio = cRadio
    
    DistanciaX = LightX + 0.5 - XCoordenadas
    DistanciaY = LightY + 0.5 - YCoordenadas
    
    DistanciaVertex = Sqr(DistanciaX * DistanciaX + DistanciaY * DistanciaY)
    
    If DistanciaVertex <= Radio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, DistanciaVertex / Radio)
        LightCalculate = D3DColorXRGB(CurrentColor.r, CurrentColor.g, CurrentColor.b)
        If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function

#Else 'Luces Normales

Private Sub LightRender(ByVal light_index As Integer)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim ia As Single
    Dim i As Integer
    Dim Color(3) As Long
    Dim Ya As Integer
    Dim Xa As Integer

    Dim XCoord As Integer
    Dim YCoord As Integer
    
    Color(0) = D3DColorARGB(255, Light_List(light_index).RGBcolor.r, Light_List(light_index).RGBcolor.g, Light_List(light_index).RGBcolor.b)
    Color(1) = Color(0)
    Color(2) = Color(0)
    Color(3) = Color(0)

    'Set up light borders
    min_x = Light_List(light_index).map_x - Light_List(light_index).range
    min_y = Light_List(light_index).map_y - Light_List(light_index).range
    max_x = Light_List(light_index).map_x + Light_List(light_index).range
    max_y = Light_List(light_index).map_y + Light_List(light_index).range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).Engine_Light(3) = Color(3)
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).Engine_Light(2) = Color(2)
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).Engine_Light(0) = Color(0)
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).Engine_Light(1) = Color(1)
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).Engine_Light(2) = Color(2)
            MapData(X, min_y).Engine_Light(3) = Color(3)
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).Engine_Light(1) = Color(1)
            MapData(X, max_y).Engine_Light(0) = Color(0)
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).Engine_Light(1) = Color(1)
            MapData(min_x, Y).Engine_Light(3) = Color(3)
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).Engine_Light(0) = Color(0)
            MapData(max_x, Y).Engine_Light(2) = Color(2)
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).Engine_Light(0) = Color(0)
                MapData(X, Y).Engine_Light(1) = Color(1)
                MapData(X, Y).Engine_Light(2) = Color(2)
                MapData(X, Y).Engine_Light(3) = Color(3)
            End If
        Next Y
    Next X
    
End Sub

Private Sub Delete_Light_To_Index(ByVal light_index As Integer)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim colorz As Long

    colorz = D3DColorARGB(Estado_Actual.a, Estado_Actual.r, Estado_Actual.g, Estado_Actual.b)
    
    'Set up light borders
    min_x = Light_List(light_index).map_x - Light_List(light_index).range
    min_y = Light_List(light_index).map_y - Light_List(light_index).range
    max_x = Light_List(light_index).map_x + Light_List(light_index).range
    max_y = Light_List(light_index).map_y + Light_List(light_index).range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).Engine_Light(3) = colorz
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).Engine_Light(2) = colorz
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).Engine_Light(0) = colorz
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).Engine_Light(1) = colorz
    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).Engine_Light(2) = colorz
            MapData(X, min_y).Engine_Light(3) = colorz
        End If
    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).Engine_Light(1) = colorz
            MapData(X, max_y).Engine_Light(0) = colorz
        End If
    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).Engine_Light(1) = colorz
            MapData(min_x, Y).Engine_Light(3) = colorz
        End If
    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).Engine_Light(0) = colorz
            MapData(max_x, Y).Engine_Light(2) = colorz
        End If
    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).Engine_Light(0) = colorz
                MapData(X, Y).Engine_Light(1) = colorz
                MapData(X, Y).Engine_Light(2) = colorz
                MapData(X, Y).Engine_Light(3) = colorz
            End If
        Next Y
    Next X
    
    'Reducimos en 1 el numero de luces y redimensionamos el Array
    NumLights = NumLights - 1
    ReDim Preserve Light_List(1 To NumLights) As tLight
    
End Sub

#End If 'Terminamos de Seleccionar las luces

Public Sub DeInit_LightEngine()
    'Kill Font's
    Erase Light_List()
    
    'Exit, The works is done.
    Exit Sub
End Sub

Public Function LightRenderAll()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim i As Long

    If NumLights = 0 Then Exit Function

    For i = 1 To UBound(Light_List)
        LightRender i
    Next i

End Function

Public Sub LightRemoveAll(ByVal Actualizar As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************

    Dim i As Long
    
    If NumLights = 0 Then Exit Sub
    Debug.Print "Numero de luces a destruir: " & NumLights & "-" & UBound(Light_List)
    For i = 1 To UBound(Light_List)
        Delete_Light_To_Index i, Actualizar
    Next i

End Sub


