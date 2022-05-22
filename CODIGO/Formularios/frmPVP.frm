VERSION 5.00
Begin VB.Form frmPVP 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "PVP"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13485
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   899
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinterAOR_Client.uAOProgress uAOProgressExperiencePVP 
      Height          =   390
      Left            =   570
      TabIndex        =   0
      ToolTipText     =   "Experiencia necesaria para pasar de nivel"
      Top             =   540
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   688
      BackColor       =   192
      BorderColor     =   0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin WinterAOR_Client.uAOButton btnRetos 
      Height          =   525
      Left            =   870
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   926
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmPVP.frx":0000
      PICF            =   "frmPVP.frx":001C
      PICH            =   "frmPVP.frx":0038
      PICV            =   "frmPVP.frx":0054
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
   Begin VB.Label lblPVP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   12450
      TabIndex        =   1
      Top             =   510
      Width           =   660
   End
End
Attribute VB_Name = "frmPVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call LoadTextsForm
End Sub

Private Sub LoadTextsForm()
    btnRetos.Caption = JsonLanguage.item("LBL_RETOS").item("TEXTO")
End Sub
