Attribute VB_Name = "modUtils"
Option Explicit

Public Function GetTerrenoDePaso(ByVal X As Integer, ByVal Y As Integer) As TipoPaso
    With MapData(X, Y).Graphic(1)
        If .GrhIndex >= 6000 And .GrhIndex <= 6307 Then
            GetTerrenoDePaso = CONST_BOSQUE
            Exit Function
        ElseIf .GrhIndex >= 7501 And .GrhIndex <= 7507 Or .GrhIndex >= 7508 And .GrhIndex <= 2508 Then
            GetTerrenoDePaso = CONST_DUNGEON
            Exit Function
        ElseIf (.GrhIndex >= 30120 And .GrhIndex <= 30375) Then
            GetTerrenoDePaso = CONST_NIEVE
            Exit Function
        Else
            GetTerrenoDePaso = CONST_PISO
        End If
    End With
End Function

Public Function Char_Big_Get(ByVal CharIndex As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
   On Error GoTo ErrorHandler


   'Make sure it's a legal char_index
    If Char_Check(CharIndex) Then
        Char_Big_Get = (GrhData(charlist(CharIndex).Body.Walk(charlist(CharIndex).Heading).GrhIndex).TileWidth > 4)
    End If

    Exit Function

ErrorHandler:

End Function

Public Sub Ghost_Create(ByVal CharIndex As Integer)
    
    On Error GoTo GhostCreate_Err
    
    With charlist(CharIndex)

        If .Body.Walk(.Heading).GrhIndex = 0 Then Exit Sub

        MapData(.Pos.X, .Pos.Y).GhostChar.Body.GrhIndex = .Body.Walk(.Heading).GrhIndex
        MapData(.Pos.X, .Pos.Y).GhostChar.Head = .Head
        MapData(.Pos.X, .Pos.Y).GhostChar.Weapon.GrhIndex = .Arma.WeaponWalk(.Heading).GrhIndex
        MapData(.Pos.X, .Pos.Y).GhostChar.Helmet = .Casco
        MapData(.Pos.X, .Pos.Y).GhostChar.Shield.GrhIndex = .Escudo.ShieldWalk(.Heading).GrhIndex
        MapData(.Pos.X, .Pos.Y).GhostChar.Body_Aura = .AuraAnim.GrhIndex
        MapData(.Pos.X, .Pos.Y).GhostChar.AlphaB = 255
        MapData(.Pos.X, .Pos.Y).GhostChar.Active = True
        MapData(.Pos.X, .Pos.Y).GhostChar.OffX = .Body.HeadOffset.X
        MapData(.Pos.X, .Pos.Y).GhostChar.Offy = .Body.HeadOffset.Y
        MapData(.Pos.X, .Pos.Y).GhostChar.Heading = .Heading

    End With
    
    Exit Sub

GhostCreate_Err:
    Call RegistrarError(Err.number, Err.Description, "ModUtils.GhostCreate", Erl)
    Resume Next
    
End Sub

