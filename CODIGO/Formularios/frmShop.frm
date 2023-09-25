VERSION 5.00
Begin VB.Form frmShop 
   BorderStyle     =   0  'None
   Caption         =   "Tienda"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmShop.frx":0000
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PictureItemShop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4440
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2790
      Width           =   495
   End
   Begin VB.ListBox lstItemsShop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000000&
      Height          =   3930
      Left            =   765
      TabIndex        =   0
      Top             =   2100
      Width           =   2325
   End
   Begin WinterAOR_Client.uAOButton imgComprar 
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   4380
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   661
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmShop.frx":1DB47
      PICF            =   "frmShop.frx":1DB63
      PICH            =   "frmShop.frx":1DB7F
      PICV            =   "frmShop.frx":1DB9B
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1890
      TabIndex        =   5
      Top             =   810
      Width           =   2625
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   4650
      TabIndex        =   3
      Top             =   1980
      Width           =   165
   End
   Begin VB.Image imgCross 
      Height          =   450
      Left            =   5580
      Top             =   300
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recomentamos que una vez realizada la transacción, reloguee su personaje por seguridad"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1020
      TabIndex        =   2
      Top             =   6270
      Width           =   4665
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager

Private cBotonCruz As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()

    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    'Cargamos la interfaz
    Me.Picture = General_Load_Picture_From_Resource("228.gif", False)
    
    Call LoadTextsForm
    Call LoadButtons
End Sub

Private Sub LoadTextsForm()
    imgComprar.Caption = JsonLanguage.item("FRMCOMERCIAR_COMPRAR").item("TEXTO")
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    GrhPath = Carga.Path(Interfaces)
    
    'Lo dejamos solo para que no explote, habria que sacar estos LastButtonPressed
    Set LastButtonPressed = New clsGraphicalButton

    Set cBotonCruz = New clsGraphicalButton
    
    Call cBotonCruz.Initialize(imgCross, "", _
                                    "171.gif", _
                                    "171.gif", Me)

End Sub

Private Sub imgComprar_Click()

    Call WriteBuyShop(lstItemsShop.ListIndex + 1)
    
End Sub

Private Sub imgCross_Click()
    Unload Me
End Sub

Private Sub lstItemsShop_Click()
    Dim DR As RECT
    
    With DR
        .Right = 32
        .Bottom = 32
    End With
    
    lblNombre.Caption = ShopObject(lstItemsShop.ListIndex + 1).Nombre & " - Precio: " & ShopObject(lstItemsShop.ListIndex + 1).valor
    Call DrawGrhtoHdc(PictureItemShop, ShopObject(lstItemsShop.ListIndex + 1).ObjIndex, DR)
    
End Sub
