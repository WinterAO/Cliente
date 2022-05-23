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
   Begin WinterAOR_Client.uAOProgress uAOProgressExperiencePVP 
      Height          =   390
      Left            =   1770
      TabIndex        =   0
      ToolTipText     =   "Experiencia necesaria para pasar de nivel"
      Top             =   780
      Width           =   8745
      _extentx        =   15425
      _extenty        =   688
      backcolor       =   192
      bordercolor     =   0
      font            =   "frmPVP.frx":26A1D
   End
   Begin WinterAOR_Client.uAOButton btnRetos 
      Height          =   525
      Left            =   1410
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2340
      Width           =   3675
      _extentx        =   6482
      _extenty        =   926
      tx              =   ""
      enab            =   -1
      fcol            =   16777215
      ocol            =   16777215
      pice            =   "frmPVP.frx":26A45
      picf            =   "frmPVP.frx":26A61
      pich            =   "frmPVP.frx":26A7D
      picv            =   "frmPVP.frx":26A99
      font            =   "frmPVP.frx":26AB5
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   9990
      Top             =   5280
      Width           =   2775
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

Private cBotonCerrar As clsGraphicalButton
Private clsFormulario As clsFormMovementManager
Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Load_Picture_From_Resource("230.gif")
    
    Call LoadButtons
    Call LoadTextsForm
    
End Sub

Public Sub IniciarLabels()
    lblPVP.Caption = CurrentUser.UserNivelPVP
    uAOProgressExperiencePVP.Max = CurrentUser.UserELVPVP
    uAOProgressExperiencePVP.value = CurrentUser.UserEXPPVP
End Sub

Private Sub LoadTextsForm()
    btnRetos.Caption = JsonLanguage.item("LBL_RETOS").item("TEXTO")
End Sub

Private Sub LoadButtons()

    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    ' Load pictures
    Call cBotonCerrar.Initialize(imgCerrar, "57.gif", _
                                    "58.gif", _
                                    "59.gif", Me)
    
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub
