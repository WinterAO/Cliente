Attribute VB_Name = "mPooChar"
'---------------------------------------------------------------------------------------
' Module    : Mod_PooChar
' Author    :  Miqueas
' Date      : 02/02/2014
' Purpose   :  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'---------------------------------------------------------------------------------------

Option Explicit
 
Public Sub Char_Erase(ByVal CharIndex As Integer)
    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************
 
    With charlist(CharIndex)

        If (CharIndex = 0) Then Exit Sub
        If (CharIndex > LastChar) Then Exit Sub
                
        If Map_InBounds(.Pos.x, .Pos.y) Then  '// Posicion valida
            MapData(.Pos.x, .Pos.y).CharIndex = 0  '// Borramos el user
        End If
       
        'Update lastchar
        If CharIndex = LastChar Then
 
            Do Until charlist(LastChar).Heading > 0
               
                LastChar = LastChar - 1
 
                If LastChar = 0 Then
                                
                    NumChars = 0

                    Exit Sub

                End If
                       
            Loop
 
        End If
   
        Call Char_ResetInfo(CharIndex)
                
        'Remove char's dialog
        Call Dialogos.RemoveDialog(CharIndex)
                
        'Update NumChars
        NumChars = NumChars - 1
        
        Debug.Print "Char_Erase: Borrado el CharIndex: " & CharIndex & " NumChars: " & NumChars
 
        Exit Sub
 
    End With
 
End Sub
 
Private Sub Char_ResetInfo(ByVal CharIndex As Integer)

    '*****************************************************************
    'Author: Ao 13.0
    'Last Modify Date: 13/12/2013
    'Reset Info User
    '*****************************************************************

    With charlist(CharIndex)
        'Remove particles
        Call Char_Particle_Group_Remove_All(CharIndex)
            
        .Active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
            
        .Moving = 0
        .muerto = False
        .Nombre = vbNullString
        .Clan = vbNullString
        .pie = False
        .Pos.x = 0
        .Pos.y = 0
        .UsandoArma = False
        .attacking = False
        .NPCAttack = False
        .BarTime = 0
        .BarAccion = 0
        .MaxBarTime = 0
        .EsNPC = False
    End With
 
End Sub
 
Private Sub Char_MapPosGet(ByVal CharIndex As Long, ByRef x As Integer, ByRef y As Integer)
                                
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 13/12/2013
    '// By Miqueas150
    '
    '*****************************************************************
        
    'Make sure it's a legal char_index
      
    With charlist(CharIndex)
                  
        'Get map pos
        x = .Pos.x
        y = .Pos.y
        
    End With
 
End Sub
 
Public Sub Char_MapPosSet(ByVal x As Integer, ByVal y As Integer)

    'Sets the user postion

    If (Map_InBounds(x, y)) Then  '// Posicion valida
        
        UserPos.x = x
        UserPos.y = y
                        
        'Set char
        MapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex
        charlist(UserCharIndex).Pos = UserPos
        
        Exit Sub
 
    End If

End Sub
 
Public Function Char_Techo() As Boolean

    '// Autor : Marcos Zeni
    '// Nueva forma de establecer si el usuario esta bajo un techo

    Char_Techo = False
 
    With charlist(UserCharIndex)
      
        If (Map_InBounds(.Pos.x, .Pos.y)) Then '// Posicion valida
                       
            If (MapData(.Pos.x, .Pos.y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.x, .Pos.y).Trigger = eTrigger.CASA) Then
                Char_Techo = True

            End If
                               
        End If
   
    End With

End Function
 
Public Function Char_MapPosExits(ByVal x As Integer, ByVal y As Integer) As Integer
 
    '*****************************************************************
    'Checks to see if a tile position has a char_index and return it
    '*****************************************************************
   
    If (Map_InBounds(x, y)) Then
        Char_MapPosExits = MapData(x, y).CharIndex
    Else
        Char_MapPosExits = 0
    End If
  
End Function
 
Public Sub Char_UserPos()

    '// Author Miqueas
    '// Actualizamo el lbl de la posicion del usuario
 
    Dim x As Integer

    Dim y As Integer
     
    If Char_Check(UserCharIndex) Then
        
        '// Damos valor a las variables asi sacamos la pos del usuario.
        Call Char_MapPosGet(UserCharIndex, x, y)
                
        bTecho = Char_Techo '// Pos : Techo :P
        
        Call CheckZona(UserCharIndex)
               
        Call frmMain.ActualizarCoordenadas(x, y)

        Call ActualizarMiniMapa
 
        Exit Sub
 
    End If

End Sub
 
Public Sub Char_UserIndexSet(ByVal CharIndex As Integer)
 
    UserCharIndex = CharIndex
 
    With charlist(UserCharIndex)
 
        'Nueva posicion para el usuario.
        UserPos = .Pos
         
        Exit Sub
 
    End With
         
End Sub
 
Public Function Char_Check(ByVal CharIndex As Integer) As Boolean
       
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
    'Last Modify by Miqueas150 Date: 24/02/2013
    'Chequeamos el Char
    '**************************************************************
       
    'check char_index
 
    If CharIndex > 0 And CharIndex <= LastChar Then
 
        With charlist(CharIndex) '// check char_index
            Char_Check = (.Heading > 0)

        End With
 
    End If
   
End Function
 
Public Sub Char_SetInvisible(ByVal CharIndex As Integer, ByVal value As Boolean)
       
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
    'Last Modify by Miqueas150 Date: 24/02/2013
 
    '**************************************************************
       
    If Char_Check(CharIndex) Then
 
        With charlist(CharIndex)
 
            .invisible = value '// User invisible o no ?
                        
            Exit Sub
 
        End With
 
    End If
 
End Sub
 
Public Sub Char_SetBody(ByVal CharIndex As Integer, ByVal BodyIndex As Integer)
 
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify by Miqueas150 Date: 24/02/2013
    'Seteamos el CharBody
    '**************************************************************

    If BodyIndex < LBound(BodyData()) Or BodyIndex > UBound(BodyData()) Then
        charlist(CharIndex).Body = BodyData(0)
        charlist(CharIndex).iBody = 0

        Exit Sub

    End If

    If Char_Check(CharIndex) Then

        With charlist(CharIndex)
               
            .Body = BodyData(BodyIndex)
            .iBody = BodyIndex
                        
            Exit Sub
 
        End With
 
    End If
 
End Sub
 
Public Sub Char_SetHead(ByVal CharIndex As Integer, ByVal HeadIndex As Integer)
 
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify by Miqueas150 Date: 24/02/2013
    'Seteamos el CharHead
    '**************************************************************
 
    If HeadIndex < 1 Or HeadIndex > NumHeads Then
        charlist(CharIndex).Head = 0
        charlist(CharIndex).iHead = 0

        Exit Sub

    End If

    If Char_Check(CharIndex) Then
 
        With charlist(CharIndex)
            .Head = HeadIndex
            .iHead = HeadIndex
                               
            .muerto = (HeadIndex = eCabezas.CASPER_HEAD)
                     
            Exit Sub
 
        End With
 
    End If
 
End Sub
 
Public Sub Char_SetHeading(ByVal CharIndex As Long, ByVal Heading As Byte)
 
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify by Miqueas150 Date: 24/02/2013
    'Changes the character heading
    '*****************************************************************
    
    'Make sure it's a legal char_index
 
    If Char_Check(CharIndex) Then
 
        With charlist(CharIndex)
               
            .Heading = Heading
 
            Exit Sub
 
        End With
 
    End If
 
End Sub

Public Sub Char_SetName(ByVal CharIndex As Integer, ByVal name As String)
 
    '**************************************************************
    'Author: Miqueas150
    'Last Modify Date: 04/12/2013
    '
    '**************************************************************
 
    If (Len(name) = 0) Then

        Exit Sub

    End If

    If Char_Check(CharIndex) Then
 
        With charlist(CharIndex)
               
            .Nombre = name
            .Clan = mid$(.Nombre, getTagPosition(.Nombre))
 
            Exit Sub
 
        End With
 
    End If
 
End Sub
 
Public Sub Char_SetWeapon(ByVal CharIndex As Integer, ByVal WeaponIndex As Integer)
 
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify by Miqueas150 Date: 24/02/2013
    '
    '**************************************************************
 
    If WeaponIndex > UBound(WeaponAnimData()) Or WeaponIndex < LBound(WeaponAnimData()) Then

        Exit Sub

    End If

    If Char_Check(CharIndex) Then
 
        With charlist(CharIndex)
               
            .Arma = WeaponAnimData(WeaponIndex)
 
            Exit Sub
 
        End With
 
    End If
 
End Sub
 
Public Sub Char_SetShield(ByVal CharIndex As Integer, ByVal ShieldIndex As Integer)
 
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify by Miqueas150 Date: 24/02/2013
    '
    '**************************************************************
 
    If ShieldIndex > UBound(ShieldAnimData()) Or ShieldIndex < LBound(ShieldAnimData()) Then

        Exit Sub

    End If

    If Char_Check(CharIndex) Then
 
        With charlist(CharIndex)
   
            .Escudo = ShieldAnimData(ShieldIndex)
                        
            Exit Sub
 
        End With
 
    End If
 
End Sub
 
Public Sub Char_SetCasco(ByVal CharIndex As Integer, ByVal CascoIndex As Integer)
 
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify by Miqueas150 Date: 24/02/2013
    '
    '**************************************************************
 
    If CascoIndex > NumCascos Or CascoIndex < 1 Then

        Exit Sub

    End If

    If Char_Check(CharIndex) Then
 
        With charlist(CharIndex)
               
            .Casco = CascoIndex
 
            Exit Sub
 
        End With
 
    End If
     
End Sub
 
Public Sub Char_SetFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
 
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
  
    If (Char_Check(CharIndex)) Then
        
        With charlist(CharIndex)

            .FxIndex = fX
        
            If .FxIndex > 0 Then
                        
                Call InitGrh(.fX, FxData(fX).Animacion)
                .fX.Loops = Loops
                                
            End If

        End With
        
    End If
   
End Sub
 
Public Sub Char_SetAura(ByVal CharIndex As Integer, ByVal AuraAnim As Long, ByVal AuraColor As Long)
    '***************************************************
    'Autor: Lorwik
    'Fecha: 20/06/2020
    '***************************************************
    
        If (Char_Check(CharIndex)) Then
        
        With charlist(CharIndex)
                        
            Call InitGrh(.AuraAnim, AuraAnim)
            .AuraColor = AuraColor
                                
        End With
        
    End If
End Sub
 
Public Sub Char_Make(ByVal CharIndex As Integer, _
                     ByVal Body As Integer, _
                     ByVal Head As Integer, _
                     ByVal Heading As Byte, _
                     ByVal x As Integer, _
                     ByVal y As Integer, _
                     ByVal Arma As Integer, _
                     ByVal Escudo As Integer, _
                     ByVal Casco As Integer, _
                     ByVal Ataque As Integer, _
                     ByVal AuraAnim As Long, _
                     ByVal AuraColor As Long)
 
    'Apuntamos al ultimo Char
 
    If CharIndex > LastChar Then
        LastChar = CharIndex

    End If
 
    NumChars = NumChars + 1

    If Arma = 0 Then Arma = 2
    If Escudo = 0 Then Escudo = 2
    If Casco = 0 Then Casco = 2
        
    With charlist(CharIndex)
       
        'If the char wasn't allready active (we are rewritting it) don't increase char count
                
        .iHead = Head
        .iBody = Body
                
        .Head = Head
        .Body = BodyData(Body)
                
        .Arma = WeaponAnimData(Arma)
        .Escudo = ShieldAnimData(Escudo)
        .Casco = Casco
        .Ataque = AtaqueData(Ataque)

        Call InitGrh(.AuraAnim, AuraAnim)
        .AuraColor = AuraColor
        
        .Heading = Heading
         
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
                
        'attack state
        .attacking = False
        .NPCAttack = False
       
        'Update position
        .Pos.x = x
        .Pos.y = y
           
    End With
   
    'Plot on map
    MapData(x, y).CharIndex = CharIndex
       
End Sub

Sub Char_MovebyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)

    Dim x        As Integer
    Dim y        As Integer
    Dim addx     As Integer
    Dim addy     As Integer
    Dim nHeading As E_Heading
    
    If (CharIndex <= 0) Then Exit Sub

    With charlist(CharIndex)
        
        x = .Pos.x
        y = .Pos.y
                
        '// Miqueas : Agrego este parchesito para evitar un run time
        If Not (Map_InBounds(x, y)) Then Exit Sub

        MapData(x, y).CharIndex = 0
        
        addx = nX - x
        addy = nY - y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        
        ElseIf Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        
        ElseIf Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        
        ElseIf Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH

        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or _
           .FxIndex = FxMeditar.GRANDE Or _
           .FxIndex = FxMeditar.MEDIANO Or _
           .FxIndex = FxMeditar.XGRANDE Or _
           .FxIndex = FxMeditar.XXGRANDE Then
            
            .FxIndex = 0

        End If

    End With

    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If Not EstaDentroDelArea(nX, nY) Then
        Call Char_Erase(CharIndex)
    End If
    
End Sub

Sub Char_MoveScreen(ByVal nHeading As E_Heading)
    '******************************************
    'Starts the screen moving in a direction
    '******************************************

    Dim x  As Integer
    Dim y  As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move

    Select Case nHeading

        Case E_Heading.NORTH
            y = -1
        
        Case E_Heading.EAST
            x = 1
        
        Case E_Heading.SOUTH
            y = 1
        
        Case E_Heading.WEST
            x = -1

    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y

    'Check to see if its out of bounds
    If (tX < MinXBorder) Or (tX > MaxXBorder) Or (tY < MinYBorder) Or (tY > MaxYBorder) Then

        Exit Sub

    Else
        
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
                
        bTecho = Char_Techo
               
    End If

End Sub

Sub Char_MovebyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************

    Dim addx As Integer
    Dim addy As Integer
        
    Dim x    As Integer
    Dim y    As Integer
        
    Dim nX   As Integer
    Dim nY   As Integer
    
    If (CharIndex <= 0) Then Exit Sub
    
    With charlist(CharIndex)
        
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move

        Select Case nHeading

            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
                                
        End Select
        
        nX = x + addx
        nY = y + addy
                
        '// Miqueas : Agrego este parchesito para evitar un run time
        If Not (Map_InBounds(nX, nY)) Then Exit Sub

        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.y = nY
        
        MapData(x, y).CharIndex = 0
         
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy

    End With
    
    If (CurrentUser.UserEstado = 0) Then
        Call DoPasosFx(CharIndex)
    End If

    If CharIndex <> UserCharIndex Then
        If Not EstaDentroDelArea(nX, nY) Then
            Call Char_Erase(CharIndex)
        End If
    End If

End Sub

Sub Char_CleanAll()
    
    '// Borramos los obj y char que esten

    Dim x         As Long, y As Long
    Dim CharIndex As Integer, obj As Long
    
    For x = XMinMapSize To XMaxMapSize
        For y = YMinMapSize To YMaxMapSize
          
            'Erase NPCs
            CharIndex = Char_MapPosExits(CInt(x), CInt(y))
 
            If (CharIndex > 0) Then
                Call Char_Erase(CharIndex)
            End If
                        
            'Erase OBJs
            obj = Map_PosExitsObject(CInt(x), CInt(y))

            If (obj > 0) Then
                Call Map_DestroyObject(CInt(x), CInt(y))
            End If

        Next y
    Next x

End Sub



