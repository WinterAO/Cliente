VERSION 5.00
Begin VB.Form frmHerrero 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantItems 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5175
      MaxLength       =   5
      TabIndex        =   13
      Text            =   "1"
      Top             =   2940
      Width           =   1050
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   5430
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   1545
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   5400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   2340
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   5430
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   3135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   5430
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   3930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   870
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   3930
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1710
      ScaleHeight     =   480
      ScaleWidth      =   1440
      TabIndex        =   7
      Top             =   3930
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   870
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   3135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1710
      ScaleHeight     =   480
      ScaleWidth      =   1440
      TabIndex        =   5
      Top             =   3135
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   870
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   2340
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1710
      ScaleHeight     =   480
      ScaleWidth      =   1440
      TabIndex        =   3
      Top             =   2340
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLingotes0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1710
      ScaleHeight     =   480
      ScaleWidth      =   1440
      TabIndex        =   2
      Top             =   1545
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   870
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1545
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   4
      Left            =   1560
      Top             =   3780
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   3
      Left            =   1560
      Top             =   2985
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   2
      Left            =   1560
      Top             =   2190
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   1
      Left            =   5280
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   2
      Left            =   5280
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   3
      Left            =   5280
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoUpgrade 
      Height          =   780
      Index           =   4
      Left            =   5280
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   4
      Left            =   720
      Top             =   3780
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   3
      Left            =   720
      Top             =   2985
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   2
      Left            =   720
      Top             =   2190
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   2760
      Top             =   4650
      Width           =   1455
   End
   Begin VB.Image picConstruir2 
      Height          =   420
      Left            =   3360
      Top             =   3180
      Width           =   1710
   End
   Begin VB.Image picConstruir1 
      Height          =   420
      Left            =   3360
      Top             =   2370
      Width           =   1710
   End
   Begin VB.Image picConstruir3 
      Height          =   420
      Left            =   3360
      Top             =   3960
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   1
      Left            =   720
      Top             =   1395
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoLingotes 
      Height          =   780
      Index           =   1
      Left            =   1560
      Top             =   1395
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image picConstruir0 
      Height          =   420
      Left            =   3360
      Top             =   1560
      Width           =   1710
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private picCheck As Picture
Private picRecuadroItem As Picture
Private picRecuadroLingotes As Picture

Private cPicCerrar As clsGraphicalButton
Private cPicConstruir(0 To 3) As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Cargando As Boolean

Private clsFormulario As clsFormMovementManager

Private Sub CargarImagenes()
    Dim ImgPath As String
    Dim Index As Integer
    
    ImgPath = Carga.Path(Interfaces)

    ' TODO: Traducir los textos de las imagenes via labels en visual basic, para que en el futuro si se quiere se pueda traducir a mas idiomas
    ' No ando con mas ganas/tiempo para hacer eso asi que se traducen las imagenes asi tenemos el juego en ingles.
    ' Tambien usar los controles uAObuttons para los botones, usar de ejemplo frmCambiaMotd.frm
    If Language = "spanish" Then
      Me.Picture = LoadPicture(ImgPath & "VentanaHerreriaArmas_spanish.jpg")
    Else
      Me.Picture = LoadPicture(ImgPath & "VentanaHerreriaArmas_english.jpg")
    End If

    Set picCheck = LoadPicture(ImgPath & "CheckBoxHerreria.jpg")
    
    Set picRecuadroItem = LoadPicture(ImgPath & "RecuadroItemsHerreria.jpg")
    Set picRecuadroLingotes = LoadPicture(ImgPath & "RecuadroLingotes.jpg")
    
    For Index = 1 To MAX_LIST_ITEMS
        imgMarcoItem(Index).Picture = picRecuadroItem
        imgMarcoLingotes(Index).Picture = picRecuadroLingotes
    Next Index
    
    Set cPicCerrar = New clsGraphicalButton
    Set cPicConstruir(0) = New clsGraphicalButton
    Set cPicConstruir(1) = New clsGraphicalButton
    Set cPicConstruir(2) = New clsGraphicalButton
    Set cPicConstruir(3) = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton
    
    Call cPicCerrar.Initialize(imgCerrar, ImgPath & "BotonCerrarHerreria.jpg", ImgPath & "BotonCerrarRolloverHerreria.jpg", ImgPath & "BotonCerrarClickHerreria.jpg", Me)
    Call cPicConstruir(0).Initialize(picConstruir0, ImgPath & "BotonConstruirHerreria.jpg", ImgPath & "BotonConstruirRolloverHerreria.jpg", ImgPath & "BotonConstruirClickHerreria.jpg", Me)
    Call cPicConstruir(1).Initialize(picConstruir1, ImgPath & "BotonConstruirHerreria.jpg", ImgPath & "BotonConstruirRolloverHerreria.jpg", ImgPath & "BotonConstruirClickHerreria.jpg", Me)
    Call cPicConstruir(2).Initialize(picConstruir2, ImgPath & "BotonConstruirHerreria.jpg", ImgPath & "BotonConstruirRolloverHerreria.jpg", ImgPath & "BotonConstruirClickHerreria.jpg", Me)
    Call cPicConstruir(3).Initialize(picConstruir3, ImgPath & "BotonConstruirHerreria.jpg", ImgPath & "BotonConstruirRolloverHerreria.jpg", ImgPath & "BotonConstruirClickHerreria.jpg", Me)
    
    Me.MouseIcon = picMouseIcon
    
End Sub

Private Sub ConstruirItem(ByVal Index As Integer)
    Dim ItemIndex As Integer
    
    If Scroll.Visible = True Then ItemIndex = Scroll.value
    ItemIndex = ItemIndex + Index
    
    Call WriteCraftBlacksmith(ObjetoHerrero(ItemIndex).objindex, txtCantItems.Text)
            
    Unload Me

End Sub

Private Sub Form_Load()
    Dim MaxConstItem As Integer
    Dim i As Integer
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    CargarImagenes
    
    Cargando = True
    
    MaxConstItem = CInt((UserLvl - 2) * 0.2)
    MaxConstItem = IIf(MaxConstItem < 1, 1, MaxConstItem)
    
    Cargando = False

End Sub

Private Sub Form_Activate()
On Error Resume Next

    InvLingosHerreria(1).DrawInventory
    InvLingosHerreria(2).DrawInventory
    InvLingosHerreria(3).DrawInventory
    InvLingosHerreria(4).DrawInventory

End Sub

Private Sub Form_GotFocus()
On Error Resume Next

    InvLingosHerreria(1).DrawInventory
    InvLingosHerreria(2).DrawInventory
    InvLingosHerreria(3).DrawInventory
    InvLingosHerreria(4).DrawInventory

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    InvLingosHerreria(1).DrawInventory
    InvLingosHerreria(2).DrawInventory
    InvLingosHerreria(3).DrawInventory
    InvLingosHerreria(4).DrawInventory

End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer, Optional ByVal Upgrading As Boolean = False)
    Dim i As Integer

    picLingotes0.Visible = (NumItems >= 1)
    picLingotes1.Visible = (NumItems >= 2)
    picLingotes2.Visible = (NumItems >= 3)
    picLingotes3.Visible = (NumItems >= 4)
    
    For i = 1 To MAX_LIST_ITEMS
        picItem(i).Visible = (NumItems >= i)
        imgMarcoItem(i).Visible = (NumItems >= i)
        imgMarcoLingotes(i).Visible = (NumItems >= i)
    Next i
    
    picConstruir0.Visible = (NumItems >= 1 And Not Upgrading)
    picConstruir1.Visible = (NumItems >= 2 And Not Upgrading)
    picConstruir2.Visible = (NumItems >= 3 And Not Upgrading)
    picConstruir3.Visible = (NumItems >= 4 And Not Upgrading)
    
    txtCantItems.Visible = Not Upgrading
    
    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Cargando = True
        Scroll.Max = NumItems - MAX_LIST_ITEMS
        Cargando = False
    Else
        Scroll.Visible = False
    End If
    
End Sub

Private Sub RenderItem(ByRef Pic As PictureBox, ByVal GrhIndex As Long)
On Error Resume Next

    Dim DR As RECT
    
    With DR
        .Right = 32
        .Bottom = 32
    End With

    Call DrawGrhtoHdc(Pic, GrhIndex, DR)

End Sub

Public Sub RenderList(ByVal Inicio As Integer)

On Error Resume Next

    Dim i As Long
    Dim j As Byte
    Dim NumItems As Integer
    Dim ObjHerrero() As tItemsConstruibles
    
    ObjHerrero = ObjetoHerrero
    
    NumItems = UBound(ObjHerrero)
    Inicio = Inicio - 1
    
    For i = 1 To MAX_LIST_ITEMS
    
        If i + Inicio <= NumItems Then
        
            With ObjHerrero(i + Inicio)
            
                ' Agrego el item
                Call RenderItem(picItem(i), .GrhIndex)
                picItem(i).ToolTipText = .name
                
                 ' Inventariode lingotes
                For j = 1 To MAXMATERIALES
                    Call InvLingosHerreria(i).SetItem(j, 0, .CantMateriales(j), 0, .Materiales(j), 0, 0, 0, 0, 0, 0, .NameMateriales(j))
                Next j
                
            End With
            
        End If
        
    Next i
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvLingosHerreria(i) = Nothing
    Next i
    
    MirandoHerreria = False
End Sub

Private Sub imgCantidadCiclo_Click()

End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub picConstruir0_Click()
    Call ConstruirItem(1)
End Sub

Private Sub picConstruir1_Click()
    Call ConstruirItem(2)
End Sub

Private Sub picConstruir2_Click()
    Call ConstruirItem(3)
End Sub

Private Sub picConstruir3_Click()
    Call ConstruirItem(4)
End Sub

Private Sub picLingotes0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub picLingotes1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub picLingotes2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub picLingotes3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Scroll_Change()
    Dim i As Long
    
    If Cargando Then Exit Sub
    
    i = Scroll.value
    ' Cargo inventarios e imagenes
    
    Call RenderList(i + 1)
            
End Sub

Private Sub txtCantItems_Change()
On Error GoTo errhandler
    If Val(txtCantItems.Text) < 0 Then
        txtCantItems.Text = 1
    End If
    
    If Val(txtCantItems.Text) > MAX_INVENTORY_OBJS Then
        txtCantItems.Text = MAX_INVENTORY_OBJS
    End If
    
    Exit Sub
    
errhandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    txtCantItems.Text = MAX_INVENTORY_OBJS
End Sub

Private Sub txtCantItems_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

