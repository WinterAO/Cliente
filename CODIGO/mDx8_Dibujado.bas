Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

' Dano en Render
Private Const DAMAGE_TIME As Integer = 1000
Private Const DAMAGE_OFFSET As Integer = 20
Private Const DAMAGE_FONT_S As Byte = 12
 
Private Enum EDType
     edPunal = 1    'Apunalo.
     edNormal = 2   'Hechizo o golpe com�n.
     edCritico = 3  'Golpe Critico
     edFallo = 4    'Fallo el ataque
     edCurar = 5    'Curacion a usuario
     edTrabajo = 6  'Cantidad de items obtenidas a partir del trabajo realizado
End Enum
 
Private DNormalFont    As New StdFont
 
Type DList
     DamageVal      As Integer      'Cantidad de da�o.
     ColorRGB(3)    As RGBA         'Color.
     DamageType     As EDType       'Tipo, se usa para saber si es apu o no.
     DamageFont     As New StdFont  'Efecto del apu.
     StartedTime    As Long         'Cuando fue creado.
     Downloading    As Byte         'Contador para la posicion Y.
     Activated      As Boolean      'Si esta activado..
End Type

Private DrawBuffer As cDIBSection

Sub DrawGrhtoHdc(ByRef Pic As PictureBox, _
                 ByVal GrhIndex As Long, _
                 ByRef DestRect As RECT)

    '*****************************************************************
    'Draws a Grh's portion to the given area of any Device Context
    '*****************************************************************
         
    DoEvents
    
    Pic.AutoRedraw = False
        
    'Clear the inventory window
    Call Engine_BeginScene
        
    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, COLOR_WHITE())
        
    Call Engine_EndScene(DestRect, Pic.hWnd)
    
    Call DrawBuffer.LoadPictureBlt(Pic.hDC)

    Pic.AutoRedraw = True

    Call DrawBuffer.PaintPicture(Pic.hDC, 0, 0, Pic.Width, Pic.Height, 0, 0, vbSrcCopy)

    Pic.Picture = Pic.Image
        
End Sub

Public Sub PrepareDrawBuffer()
    Set DrawBuffer = New cDIBSection
    'El tamanio del buffer es arbitrario = 1024 x 1024
    Call DrawBuffer.Create(1024, 1024)
End Sub

Public Sub CleanDrawBuffer()
    Set DrawBuffer = Nothing
End Sub

Sub Damage_Initialize()

    ' Inicializamos el dano en render
    With DNormalFont
        .Size = 20
        .italic = False
        .bold = False
        .name = "Tahoma"
    End With

End Sub

Sub Damage_Create(ByVal X As Integer, _
                  ByVal Y As Integer, _
                  ByVal DamageValue As Integer, _
                  ByVal edMode As Byte)
 
    ' @ Agrega un nuevo dano.
 
    With MapData(X, Y).Damage
     
        .Activated = True
        
        .DamageType = edMode
        .DamageVal = DamageValue
        .StartedTime = GetTickCount
        .Downloading = 0
     
        Select Case .DamageType
        
            Case EDType.edPunal

                With .DamageFont
                    .Size = Val(DAMAGE_FONT_S)
                    .name = "Tahoma"
                    .bold = False
                    Exit Sub

                End With
            
        End Select
     
        .DamageFont = DNormalFont
        .DamageFont.Size = 14
     
    End With
 
End Sub

Private Function EaseOutCubic(Time As Double)
    Time = Time - 1
    EaseOutCubic = Time * Time * Time + 1
End Function
 
Sub Damage_Draw(ByVal X As Integer, _
                ByVal Y As Integer, _
                ByVal PixelX As Integer, _
                ByVal PixelY As Integer)
 
    ' @ Dibuja un dano
 
    With MapData(X, Y).Damage
     
        If (Not .Activated) Or (Not .DamageVal <> 0) Then Exit Sub
        
        Dim ElapsedTime As Long
        ElapsedTime = GetTickCount - .StartedTime
        
        If ElapsedTime < DAMAGE_TIME Then
           
            .Downloading = EaseOutCubic(ElapsedTime / DAMAGE_TIME) * DAMAGE_OFFSET
           
            Select Case .DamageType
                   
                Case EDType.edPunal
                    Call RGBAList(.ColorRGB, ColoresPJ(52).R, ColoresPJ(52).G, ColoresPJ(52).B)
                    
                Case EDType.edFallo
                    Call RGBAList(.ColorRGB, ColoresPJ(54).R, ColoresPJ(54).G, ColoresPJ(54).B)
                    
                Case EDType.edCurar
                    Call RGBAList(.ColorRGB, ColoresPJ(55).R, ColoresPJ(55).G, ColoresPJ(55).B)
                
                Case EDType.edTrabajo
                    Call RGBAList(.ColorRGB, ColoresPJ(56).R, ColoresPJ(56).G, ColoresPJ(56).B)
                    
                Case Else 'EDType.edNormal
                    Call RGBAList(.ColorRGB, ColoresPJ(51).R, ColoresPJ(51).G, ColoresPJ(51).B)
                    
            End Select
           
            'Efectito para el apu
            If .DamageType = EDType.edPunal Then
                .DamageFont.Size = Damage_NewSize(ElapsedTime)

            End If
               
            'Dibujo
            Select Case .DamageType
            
                Case EDType.edCritico
                    Call DrawText(PixelX, PixelY - .Downloading, .DamageVal & "!!", .ColorRGB)
                
                Case EDType.edCurar
                    Call DrawText(PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB)
                
                Case EDType.edTrabajo
                    Call DrawText(PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB)
                    
                Case EDType.edFallo
                    Call DrawText(PixelX, PixelY - .Downloading, "Fallo", .ColorRGB)
                    
                Case Else 'EDType.edNormal
                    Call DrawText(PixelX, PixelY - .Downloading, "-" & .DamageVal, .ColorRGB)
                    
            End Select
            
        'Si llego al tiempo lo limpio
        Else
            Damage_Clear X, Y
           
        End If
       
    End With
 
End Sub
 
Sub Damage_Clear(ByVal X As Integer, ByVal Y As Integer)
 
    ' @ Limpia todo.
 
    With MapData(X, Y).Damage
        .Activated = False
        .DamageVal = 0
        .StartedTime = 0

    End With
 
End Sub
 
Function Damage_NewSize(ByVal ElapsedTime As Integer) As Byte
 
    ' @ Se usa para el "efecto" del apu.

    ' Nos basamos en la constante DAMAGE_TIME
    Select Case ElapsedTime
 
        Case Is <= DAMAGE_TIME / 5
            Damage_NewSize = 14
       
        Case Is <= DAMAGE_TIME * 2 / 5
            Damage_NewSize = 13
           
        Case Is <= DAMAGE_TIME * 3 / 5
            Damage_NewSize = 12
           
        Case Else
            Damage_NewSize = 11
       
    End Select
 
End Function

Public Sub DibujarMenuMacros(Optional ActualizarCual As Byte = 0)
'************************************
'Autor: Lorwik
'Fecha: 07/03/2021
'Descripcion: Dibuja los macros del frmmain
'***********************************

    Dim i As Integer
    
    If ActualizarCual <= 0 Then
    
        For i = 1 To NUMMACROS
            Select Case MacrosKey(i).TipoAccion
                Case 1 'Envia comando
                    Call Mod_TileEngine.RenderItem(frmMain.picMacro(i - 1), 37605)
                    frmMain.picMacro(i - 1).ToolTipText = "Enviar comando: " & MacrosKey(i).Comando
                    
                Case 2 'Lanza hechizo
                    Call Mod_TileEngine.RenderItem(frmMain.picMacro(i - 1), 609)
                    frmMain.picMacro(i - 1).ToolTipText = "Lanzar hechizo: " & MacrosKey(i).SpellName
                    
                Case 3 'Equipa
                    If MacrosKey(i).InvGrh > 0 Then
                        Call Mod_TileEngine.RenderItem(frmMain.picMacro(i - 1), MacrosKey(i).InvGrh)
                        frmMain.picMacro(i - 1).ToolTipText = "Equipar objeto: " & MacrosKey(i).invName
                    End If
                    
                Case 4 'Usa
                    If MacrosKey(i).InvGrh > 0 Then
                        Call Mod_TileEngine.RenderItem(frmMain.picMacro(i - 1), MacrosKey(i).InvGrh)
                        frmMain.picMacro(i - 1).ToolTipText = "Usar objeto: " & MacrosKey(i).invName
                    End If
                    
                End Select
        Next i
    
    Else
    
        Select Case MacrosKey(ActualizarCual).TipoAccion
            Case 1 'Envia comando
                Call Mod_TileEngine.RenderItem(frmMain.picMacro(ActualizarCual - 1), 37605)
                frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Enviar comando: " & MacrosKey(ActualizarCual).Comando
                
            Case 2 'Lanza hechizo
                Call Mod_TileEngine.RenderItem(frmMain.picMacro(ActualizarCual - 1), 609)
                frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Lanzar hechizo: " & MacrosKey(ActualizarCual).SpellName
                
            Case 3 'Equipa
                If MacrosKey(ActualizarCual).InvGrh > 0 Then
                    Call Mod_TileEngine.RenderItem(frmMain.picMacro(ActualizarCual - 1), MacrosKey(ActualizarCual).InvGrh)
                    frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Equipar objeto: " & MacrosKey(ActualizarCual).invName
                End If
                
            Case 4 'Usa
                If MacrosKey(ActualizarCual).InvGrh > 0 Then
                    Call Mod_TileEngine.RenderItem(frmMain.picMacro(ActualizarCual - 1), MacrosKey(ActualizarCual).InvGrh)
                    frmMain.picMacro(ActualizarCual - 1).ToolTipText = "Usar objeto: " & MacrosKey(ActualizarCual).invName
                End If
        End Select
    
        frmMain.picMacro(ActualizarCual - 1).Refresh
    
    End If

End Sub

