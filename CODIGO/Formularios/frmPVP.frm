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
   Picture         =   "frmPVP.frx":0000
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   899
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   630
      Left            =   11700
      TabIndex        =   3
      Top             =   5520
      Width           =   1590
   End
   Begin WinterAOR_Client.uAOProgress uAOProgressExperiencePVP 
      Height          =   390
      Left            =   1770
      TabIndex        =   0
      ToolTipText     =   "Experiencia necesaria para pasar de nivel"
      Top             =   780
      Width           =   8745
      _ExtentX        =   15425
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
      Left            =   1410
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2340
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   926
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmPVP.frx":26A1D
      PICF            =   "frmPVP.frx":26A39
      PICH            =   "frmPVP.frx":26A55
      PICV            =   "frmPVP.frx":26A71
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
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   12000
      TabIndex        =   1
      Top             =   810
      Width           =   660
   End
End
Attribute VB_Name = "frmPVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadTextsForm
    
    Me.Picture = General_Load_Picture_From_Resource("230.gif")
    
End Sub

Public Sub IniciarLabels()
    lblPVP.Caption = CurrentUser.UserNivelPVP
    uAOProgressExperiencePVP.Max = CurrentUser.UserEXPPVP
    uAOProgressExperiencePVP.value = CurrentUser.UserEXPPVP
End Sub

Private Sub LoadTextsForm()
    btnRetos.Caption = JsonLanguage.item("LBL_RETOS").item("TEXTO")
End Sub
