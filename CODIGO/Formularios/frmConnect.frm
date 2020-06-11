VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "WinterAO Resurrection"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   5.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtCrearPJNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
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
      Height          =   330
      Left            =   6000
      MaxLength       =   30
      TabIndex        =   4
      Top             =   10020
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtNombre 
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
      Height          =   225
      Left            =   6675
      MaxLength       =   23
      TabIndex        =   0
      Top             =   5595
      Width           =   2460
   End
   Begin VB.TextBox txtPasswd 
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   6675
      MaxLength       =   23
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   6120
      Width           =   2460
   End
   Begin VB.PictureBox Renderer 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   11520
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   768
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   15360
      Begin WinterAO.uAOCheckbox chkRecordar 
         Height          =   345
         Left            =   6000
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   7440
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         CHCK            =   0   'False
         ENAB            =   -1  'True
         PICC            =   "frmConnect.frx":000C
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Lector As clsIniManager

Public MouseX              As Long
Public MouseY              As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call CloseClient
    End If
End Sub

Private Sub Renderer_Click()
    Call ModCnt.ClickEvent(MouseX, MouseY)
End Sub

Private Sub Renderer_DblClick()
    Call ModCnt.DobleClickEvent(MouseX, MouseY)
End Sub

Private Sub Renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - Renderer.Left
    MouseY = Y - Renderer.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > Renderer.Width Then
        MouseX = Renderer.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > Renderer.Height Then
        MouseY = Renderer.Height
    End If
    
End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
  '  If KeyAscii = vbKeyReturn Then btnConectarse_Click
End Sub
