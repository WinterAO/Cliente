Attribute VB_Name = "mPooMap"
'---------------------------------------------------------------------------------------
' Module    : Mod_PooMap
' Author    :  Miqueas
' Date      : 02/02/2014
' Purpose   :  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
'---------------------------------------------------------------------------------------

Option Explicit

Private Const GrhFogata As Long = 1521

Public Sub Map_RemoveOldUser()

      With MapData(UserPos.x, UserPos.y)

            If (.CharIndex = UserCharIndex) Then
                  .CharIndex = 0
            End If

      End With

End Sub

Public Sub Map_CreateObject(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Long, ByVal ParticulaIndex As Integer, ByVal Shadow As Byte)

      'Dim objgrh As Integer
        
    '¿El objeto no tiene un Grh valido ni particula?
      If Not GrhCheck(GrhIndex) And ParticulaIndex = 0 Then
            Exit Sub

      End If
                        
      If (Map_InBounds(x, y)) Then

            With MapData(x, y)

                  'If (Map_PosExitsObject(x, y) > 0) Then
                  '      Call Map_DestroyObject(x, y)
                  'End If
                  
                  .OBJInfo.Shadow = Shadow
                  If ParticulaIndex > 0 And .Particle_Group_Index = 0 Then .Particle_Group_Index = General_Particle_Create(ParticulaIndex, x, y)
                  Call InitGrh(.ObjGrh, GrhIndex)
            End With

      End If

End Sub

Public Sub Map_DestroyObject(ByVal x As Integer, ByVal y As Integer)
    Dim ParticulaObject As Long
    
      If (Map_InBounds(x, y)) Then

            With MapData(x, y)
                  '.objgrh.GrhIndex = 0
                  .OBJInfo.ObjIndex = 0
                  .OBJInfo.Amount = 0

                  Call GrhUninitialize(.ObjGrh)
        
            End With

      End If

End Sub

Public Function Map_PosExitsObject(ByVal x As Integer, ByVal y As Integer) As Long
 
      '*****************************************************************
      'Checks to see if a tile position has a char_index and return it
      '*****************************************************************

      If (Map_InBounds(x, y)) Then
        If MapData(x, y).ObjGrh.GrhIndex > 0 Then
            Map_PosExitsObject = MapData(x, y).ObjGrh.GrhIndex
        ElseIf MapData(x, y).Particle_Group_Index > 0 Then
            Map_PosExitsObject = MapData(x, y).Particle_Group_Index
        End If
      Else
            Map_PosExitsObject = 0
      End If
 
End Function

Public Function Map_GetBlocked(ByVal x As Integer, ByVal y As Integer) As Boolean
      '*****************************************************************
      'Author: Aaron Perkins - Modified by Juan Martin Sotuyo Dodero
      'Last Modify Date: 10/07/2002
      'Checks to see if a tile position is blocked
      '*****************************************************************

      If (Map_InBounds(x, y)) Then
            Map_GetBlocked = (MapData(x, y).Blocked)
      End If

End Function

Public Sub Map_SetBlocked(ByVal x As Integer, ByVal y As Integer, ByVal block As Byte)

      If (Map_InBounds(x, y)) Then
            MapData(x, y).Blocked = block
      End If

End Sub

Sub Map_MoveTo(ByVal Direccion As E_Heading)
      '***************************************************
      'Author: Alejandro Santos (AlejoLp)
      'Last Modify Date: 06/28/2008
      'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
      ' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
      ' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
      ' 06/28/2008: NicoNZ - Saque lo que impedia que si el usuario estaba paralizado se ejecute el sub.
      '***************************************************

      Dim LegalOk As Boolean
      Static lastmovement As Long
      
      If Cartel Then Cartel = False
    
      Select Case Direccion

            Case E_Heading.NORTH
                  LegalOk = Map_LegalPos(UserPos.x, UserPos.y - 1)

            Case E_Heading.EAST
                  LegalOk = Map_LegalPos(UserPos.x + 1, UserPos.y)

            Case E_Heading.SOUTH
                  LegalOk = Map_LegalPos(UserPos.x, UserPos.y + 1)

            Case E_Heading.WEST
                  LegalOk = Map_LegalPos(UserPos.x - 1, UserPos.y)
                        
      End Select

      If LegalOk And Not CurrentUser.UserParalizado And Not CurrentUser.UserDescansar And Not CurrentUser.UserMeditar Then
          Call WriteWalk(Direccion)
          Call ActualizarMiniMapa
          
          Call MainTimer.Restart(TimersIndex.Walk)

          Call Char_MovebyHead(UserCharIndex, Direccion)
          Call Char_MoveScreen(Direccion)
      
      Else
      
        If (charlist(UserCharIndex).Heading <> Direccion) Then
            If MainTimer.Check(TimersIndex.ChangeHeading) Then
                Call WriteChangeHeading(Direccion)
                Call Char_SetHeading(UserCharIndex, Direccion)
            End If
        End If
                
      End If
  
      ' Esto es un parche por que por alguna razon si el pj esta meditando y nos movemos el juego explota por eso cambie
      ' Las validaciones en la linea 131 y agregue esto para arreglarlo (Recox)
      If CurrentUser.UserMeditar Then
        CurrentUser.UserMeditar = Not CurrentUser.UserMeditar
      End If

      If CurrentUser.UserDescansar Then
        CurrentUser.UserDescansar = Not CurrentUser.UserDescansar
      End If
        
End Sub

Function Map_LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
      '*****************************************************************
      'Author: ZaMa
      'Last Modification: 06/04/2020
      'Checks to see if a tile position is legal, including if there is a casper in the tile
      '10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
      '01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
      '12/01/2020: Recox - Now we manage monturas.
      '06/04/2020: FrankoH298 - Si estamos montados, no nos deja ingresar a las casas.
      '*****************************************************************

      Dim CharIndex As Integer
    
      'Limites del mapa

      If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then

            Exit Function

      End If
    
      'Tile Bloqueado?

      If (Map_GetBlocked(x, y)) Then
         
            Exit Function

      End If
    
      CharIndex = (Char_MapPosExits(CInt(x), CInt(y)))
        
      'Hay un personaje?

      If (CharIndex > 0) Then
    
            If (Map_GetBlocked(UserPos.x, UserPos.y)) Then
                
                  Exit Function

            End If
        
            With charlist(CharIndex)
                  ' Si no es casper, no puede pasar

                  If .iHead <> eCabezas.CASPER_HEAD And .iBody <> eCabezas.FRAGATA_FANTASMAL Then
                              
                        Exit Function

                  Else
                        ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)

                        If (Map_CheckWater(UserPos.x, UserPos.y)) Then
                              If Not (Map_CheckWater(x, y)) Then
                                            
                                    Exit Function

                              End If

                        Else
                              ' No puedo intercambiar con un casper que este en la orilla (Lado agua)

                              If (Map_CheckWater(x, y)) Then
                                             
                                    Exit Function

                              End If
                                        
                        End If
                
                        ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles

                        If (esGM(UserCharIndex)) Then

                              If (charlist(UserCharIndex).invisible) Then
                                             
                                    Exit Function

                              End If
                                        
                        End If
                  End If

            End With

      End If
   
      If (CurrentUser.UserNavegando <> Map_CheckWater(x, y)) Then
               
            Exit Function

      End If

      'Esta el usuario Equitando bajo un techo?
      If CurrentUser.UserEquitando And MapData(x, y).Trigger = eTrigger.BAJOTECHO Or CurrentUser.UserEquitando And MapData(x, y).Trigger = eTrigger.CASA Then
            Exit Function
      End If
      
      If CurrentUser.UserEvento Then Exit Function
      
    
      Map_LegalPos = True
End Function

Function Map_InBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
      '*****************************************************************
      'Checks to see if a tile position is in the maps bounds
      '*****************************************************************

      If (x < XMinMapSize) Or (x > XMaxMapSize) Or (y < YMinMapSize) Or (y > YMaxMapSize) Then
            Map_InBounds = False

            Exit Function

      End If
    
      Map_InBounds = True
End Function

Public Function Map_CheckBonfire(ByRef Location As Position) As Boolean

      Dim j As Long
      Dim k As Long
    
      For j = UserPos.x - 8 To UserPos.x + 8
            For k = UserPos.y - 6 To UserPos.y + 6

                  If Map_InBounds(j, k) Then
                        If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                              Location.x = j
                              Location.y = k
                              Map_CheckBonfire = True

                              Exit Function

                        End If
                  End If

            Next k
      Next j

End Function

Function Map_CheckWater(ByVal x As Integer, ByVal y As Integer) As Boolean

      If Map_InBounds(x, y) Then

            With MapData(x, y)

                  If ((.Graphic(1).GrhIndex >= 1505 And .Graphic(1).GrhIndex <= 1520) Or (.Graphic(1).GrhIndex >= 5665 And .Graphic(1).GrhIndex <= 5680) Or (.Graphic(1).GrhIndex >= 13547 And .Graphic(1).GrhIndex <= 13562)) And .Graphic(2).GrhIndex = 0 Then
                        Map_CheckWater = True
                  Else
                        Map_CheckWater = False
                  End If

            End With

      End If
                  
End Function

