VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online Libre"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19845
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
   ScaleWidth      =   1323
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox PortTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   16590
      TabIndex        =   9
      Text            =   "7666"
      Top             =   2160
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   17460
      TabIndex        =   8
      Text            =   "localhost"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   16560
      TabIndex        =   7
      Top             =   2880
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
      Left            =   16560
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3360
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15360
      Begin WinterAO.uAOButton btnTeclas 
         Height          =   375
         Left            =   7560
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         TX              =   "Teclas"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":000C
         PICF            =   "frmConnect.frx":0A36
         PICH            =   "frmConnect.frx":16F8
         PICV            =   "frmConnect.frx":268A
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin WinterAO.uAOButton btnConectarse 
         Height          =   375
         Left            =   6240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         TX              =   "Conectarse"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":358C
         PICF            =   "frmConnect.frx":3FB6
         PICH            =   "frmConnect.frx":4C78
         PICV            =   "frmConnect.frx":5C0A
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin WinterAO.uAOCheckbox chkRecordar 
         Height          =   345
         Left            =   6480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   5760
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         CHCK            =   0   'False
         ENAB            =   -1  'True
         PICC            =   "frmConnect.frx":6B0C
      End
      Begin WinterAO.uAOButton btnSalir 
         Height          =   375
         Left            =   8520
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   10320
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         TX              =   "Salir"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":7BF2
         PICF            =   "frmConnect.frx":861C
         PICH            =   "frmConnect.frx":92DE
         PICV            =   "frmConnect.frx":A270
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin WinterAO.uAOButton btnRecuperar 
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   10320
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         TX              =   "Recuperar Pass"
         ENAB            =   -1  'True
         FCOL            =   7314354
         OCOL            =   16777215
         PICE            =   "frmConnect.frx":B172
         PICF            =   "frmConnect.frx":BB9C
         PICH            =   "frmConnect.frx":C85E
         PICV            =   "frmConnect.frx":D7F0
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmConnect"
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
'
'Matias Fernando Pequeno
'matux@fibertel.com.ar
'www.noland-studios.com.ar
'Acoyte 678 Piso 17 Dto B
'Capital Federal, Buenos Aires - Republica Argentina
'Codigo Postal 1405

Option Explicit

' Animacion de los Controles...
Private Type tAnimControl
    Activo As Boolean
    Velocidad As Double
    Top As Integer
End Type

Private Lector As clsIniManager

Private Const AES_PASSWD As String = "tumamaentanga"

Public MouseX              As Long
Public MouseY              As Long

Private Sub btnConectarse_Click()
    'update user info
    AccountName = txtNombre.Text
    AccountPassword = txtPasswd.Text

    'Clear spell list
    frmMain.hlst.Clear

    If Me.chkRecordar.Checked = False Then
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Remember", "False")
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "UserName", vbNullString)
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Password", vbNullString)
    Else
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Remember", "True")
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "UserName", AccountName)
        Call WriteVar(Game.path(INIT) & "Config.ini", "Login", "Password", Cripto.AesEncryptString(AccountPassword, AES_PASSWD))
    End If

    If CheckUserData() = True Then
        Call Protocol.Connect(E_MODO.Normal)
    End If
End Sub

Private Sub btnSalir_Click()
    Call CloseClient
End Sub

Private Sub btnTeclas_Click()
    Load frmKeypad
    frmKeypad.Show vbModal
    Unload frmKeypad
    txtPasswd.SetFocus
End Sub

Private Sub Form_Activate()

    If CBool(GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "Remember")) = True Then
        Me.txtNombre = GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "UserName")
        Me.txtPasswd = Cripto.AesDecryptString(GetVar(Game.path(INIT) & "Config.ini", "LOGIN", "Password"), AES_PASSWD)
        Me.chkRecordar.Checked = True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Call CloseClient
    End If
End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    EngineRun = False
    '[END]

    Call LoadTextsForm
    'Call LoadAOCustomControlsPictures(Me)
    'Todo: Poner la carga de botones como en el frmCambiaMotd.frm para mantener coherencia con el resto de la aplicacion
    'y poder borrar los frx de este archivo

End Sub

Private Sub LoadTextsForm()
    btnConectarse.Caption = JsonLanguage.item("BTN_CONECTARSE").item("TEXTO")
    btnRecuperar.Caption = JsonLanguage.item("BTN_RECUPERAR").item("TEXTO")
    'lblRecordarme.Caption = JsonLanguage.item("LBL_RECORDARME").item("TEXTO")
    btnSalir.Caption = JsonLanguage.item("BTN_SALIR").item("TEXTO")
    btnTeclas.Caption = JsonLanguage.item("LBL_TECLAS").item("TEXTO")
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
    If KeyAscii = vbKeyReturn Then btnConectarse_Click
End Sub
