VERSION 5.00
Begin VB.Form frmArtesano 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Artesano"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   4
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   3855
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picObj3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1725
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   7
      Top             =   3855
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   3
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   3060
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picObj2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1725
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   5
      Top             =   3060
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   2
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   2265
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picObj1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1725
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   2265
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picObj0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   1725
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   1
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1470
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Costo 
      BackStyle       =   0  'Transparent
      Caption         =   "999.999.999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   4740
      Width           =   1545
   End
   Begin VB.Image imgConstruir0 
      Height          =   420
      Left            =   4200
      Top             =   1515
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir1 
      Height          =   420
      Left            =   4200
      Top             =   2325
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir2 
      Height          =   420
      Left            =   4200
      Top             =   3135
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgConstruir3 
      Height          =   420
      Left            =   4200
      Top             =   3915
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   1
      Left            =   690
      Top             =   1320
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   2
      Left            =   690
      Top             =   2115
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   3
      Left            =   690
      Top             =   2910
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoItem 
      Height          =   780
      Index           =   4
      Left            =   690
      Top             =   3705
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image imgMarcoReqItem 
      Height          =   780
      Index           =   1
      Left            =   1590
      Top             =   1320
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image imgMarcoReqItem 
      Height          =   780
      Index           =   2
      Left            =   1590
      Top             =   2115
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image imgMarcoReqItem 
      Height          =   780
      Index           =   3
      Left            =   1590
      Top             =   2910
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image imgMarcoReqItem 
      Height          =   780
      Index           =   4
      Left            =   1590
      Top             =   3705
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   720
      Top             =   4605
      Width           =   1455
   End
End
Attribute VB_Name = "frmArtesano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private picRecuadroItem         As Picture
Private picRecuadroReqItems     As Picture

Private cBotonCerrar            As clsGraphicalButton
Private cBotonConstruir(0 To 4) As clsGraphicalButton

Public LastButtonPressed        As clsGraphicalButton
Public ArtesaniaCosto           As Long

Private Sub Form_Load()
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Scroll.value = 0
    
    Me.Picture = General_Load_Picture_From_Resource("16.bmp", False)
    
    Call LoadButtons
    
End Sub

Private Sub Form_Activate()
On Error Resume Next

    InvObjArtesano(1).DrawInventory
    InvObjArtesano(2).DrawInventory
    InvObjArtesano(3).DrawInventory
    InvObjArtesano(4).DrawInventory

End Sub

Private Sub LoadButtons()
    Dim Index   As Long

    Set picRecuadroItem = General_Load_Picture_From_Resource("14.bmp", False)
    Set picRecuadroReqItems = General_Load_Picture_From_Resource("15.bmp", False)

    For Index = 1 To MAX_LIST_ITEMS
        imgMarcoItem(Index).Picture = picRecuadroItem
        imgMarcoReqItem(Index).Picture = picRecuadroReqItems
    Next Index

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonConstruir(0) = New clsGraphicalButton
    Set cBotonConstruir(1) = New clsGraphicalButton
    Set cBotonConstruir(2) = New clsGraphicalButton
    Set cBotonConstruir(3) = New clsGraphicalButton

    Set LastButtonPressed = New clsGraphicalButton

    Call cBotonCerrar.Initialize(imgCerrar, "8.bmp", "9.bmp", "10.bmp", Me)

    Call cBotonConstruir(0).Initialize(imgConstruir0, "11.bmp", "12.bmp", "13.bmp", Me)
    Call cBotonConstruir(1).Initialize(imgConstruir1, "11.bmp", "12.bmp", "13.bmp", Me)
    Call cBotonConstruir(2).Initialize(imgConstruir2, "11.bmp", "12.bmp", "13.bmp", Me)
    Call cBotonConstruir(3).Initialize(imgConstruir3, "11.bmp", "12.bmp", "13.bmp", Me)

    Costo.Caption = Format$(ArtesaniaCosto, "##,##")

End Sub

Private Sub Construir(ByVal Index As Integer)

    Dim ItemIndex As Integer

    If Scroll.Visible = True Then ItemIndex = Scroll.value
    ItemIndex = ItemIndex + Index

    Call WriteCraftsmanCreate(ItemIndex)

    Unload Me

End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer)
    Dim i As Integer
    
    picObj0.Visible = (NumItems >= 1)
    picObj1.Visible = (NumItems >= 2)
    picObj2.Visible = (NumItems >= 3)
    picObj3.Visible = (NumItems >= 4)
    
    imgConstruir0.Visible = (NumItems >= 1)
    imgConstruir1.Visible = (NumItems >= 2)
    imgConstruir2.Visible = (NumItems >= 3)
    imgConstruir3.Visible = (NumItems >= 4)

    For i = 1 To MAX_LIST_ITEMS
        picItem(i).Visible = (NumItems >= i)
        imgMarcoItem(i).Visible = (NumItems >= i)
        imgMarcoReqItem(i).Visible = (NumItems >= i)
    Next i
    
    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Scroll.Max = NumItems - MAX_LIST_ITEMS
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

    Dim i        As Integer
    Dim j        As Integer
    Dim NumItems As Integer
    
    NumItems = UBound(ObjArtesano)
    Inicio = Inicio - 1
    
    For i = 1 To MAX_LIST_ITEMS

        If i + Inicio <= NumItems Then

            With ObjArtesano(i + Inicio)
            
                ' Agrego el item
                Call RenderItem(picItem(i), .GrhIndex)
                picItem(i).ToolTipText = .name

                ' Items requeridos
                For j = 1 To UBound(.ItemsCrafteo)
                    Call InvObjArtesano(i).SetItem(j, .ItemsCrafteo(j).ObjIndex, .ItemsCrafteo(j).Amount, 0, .ItemsCrafteo(j).GrhIndex, 0, 0, 0, 0, 0, 0, .ItemsCrafteo(j).name)
                Next j
                
            End With
            
        End If
        
    Next i
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvObjArtesano(i) = Nothing
    Next i

End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgConstruir0_Click()
    Call Construir(1)
End Sub

Private Sub imgConstruir1_Click()
    Call Construir(2)
End Sub

Private Sub imgConstruir2_Click()
    Call Construir(3)
End Sub

Private Sub imgConstruir3_Click()
    Call Construir(4)
End Sub

Private Sub Scroll_Change()
    Call RenderList(Scroll.value + 1)
End Sub


