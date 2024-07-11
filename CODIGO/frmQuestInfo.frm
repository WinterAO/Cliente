VERSION 5.00
Begin VB.Form frmQuestInfo 
   BorderStyle     =   0  'None
   Caption         =   "Informacion de la mision"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   6525
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   885
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1500
      Width           =   4815
   End
   Begin WinterAOR_Client.uAOButton Aceptar 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5115
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmQuestInfo.frx":0000
      PICF            =   "frmQuestInfo.frx":001C
      PICH            =   "frmQuestInfo.frx":0038
      PICV            =   "frmQuestInfo.frx":0054
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
   Begin WinterAOR_Client.uAOButton Rechazar 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5115
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmQuestInfo.frx":0070
      PICF            =   "frmQuestInfo.frx":008C
      PICH            =   "frmQuestInfo.frx":00A8
      PICV            =   "frmQuestInfo.frx":00C4
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
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Informacion de la mision:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "frmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.Picture = General_Load_Picture_From_Resource("188.bmp", False)
    
    Call LoadTextsForm

End Sub

Private Sub LoadTextsForm()
    Me.lblDesc.Caption = JsonLanguage.item("FRM_QUEST_DESC").item("TEXTO")
    Me.Aceptar.Caption = JsonLanguage.item("FRM_QUEST_ACCEPT").item("TEXTO")
    Me.Rechazar.Caption = JsonLanguage.item("FRM_QUEST_EXIT").item("TEXTO")
End Sub

Private Sub Aceptar_Click()
    Comerciando = False
    Call WriteQuestAccept
    Unload Me
End Sub

Private Sub Rechazar_Click()
    Comerciando = False
    Unload Me
End Sub
