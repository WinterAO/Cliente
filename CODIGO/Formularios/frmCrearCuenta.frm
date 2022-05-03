VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Crear Cuenta"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   4365
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FraCrearCuenta 
      BackColor       =   &H00535353&
      Caption         =   "Crear Cuenta Winter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4065
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   3645
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   360
         Left            =   210
         TabIndex        =   6
         Top             =   3510
         Width           =   1590
      End
      Begin VB.TextBox txtNombreCuenta 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   330
         TabIndex        =   5
         Top             =   600
         Width           =   3045
      End
      Begin VB.TextBox txtEmail 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   330
         TabIndex        =   4
         Top             =   1410
         Width           =   3045
      End
      Begin VB.TextBox txtPass 
         BorderStyle     =   0  'None
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   330
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2250
         Width           =   3015
      End
      Begin VB.TextBox txtRePass 
         BorderStyle     =   0  'None
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   330
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3000
         Width           =   3045
      End
      Begin VB.CommandButton cmdCrearCuenta 
         Caption         =   "Crear Cuenta"
         Height          =   360
         Left            =   1890
         TabIndex        =   1
         Top             =   3510
         Width           =   1515
      End
      Begin VB.Label lblNombreDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de cuenta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   390
         TabIndex        =   9
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label lblContraseña 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   330
         TabIndex        =   8
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label lblRepitaLa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repita la Contraseña:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   330
         TabIndex        =   7
         Top             =   2790
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private botonCrear As Boolean

Private Sub cmdCerrar_Click()

    CurrentUser.AccountName = vbNullString
    CurrentUser.AccountPassword = vbNullString
    CurrentUser.AccountMail = vbNullString

    Unload Me
End Sub

Private Sub cmdCrearCuenta_Click()

    If Not checkCuenta Then Exit Sub
        
    CurrentUser.AccountName = txtNombreCuenta.Text
    CurrentUser.AccountPassword = txtPass.Text
    CurrentUser.AccountMail = txtEmail.Text
        
    'Conexion!!!
    If Not frmMain.Client.State = sckConnected Then
        Call MostrarMensaje(JsonLanguage.item("ERROR_CONN_LOST").item("TEXTO"))
        Call cmdCerrar_Click
    Else
        'Si ya mandamos el paquete, evitamos que se pueda volver a mandar
        botonCrear = True
        Call Login
        botonCrear = False
    End If
    
End Sub

Private Function checkCuenta() As Boolean

    If txtNombreCuenta.Text = vbNullString Then
        MsgBox "El campo de nombre de cuenta esta vacio."
        checkCuenta = False
        Exit Function
    End If
    
    If LenB(txtNombreCuenta.Text) > 24 Or Len(txtNombreCuenta.Text) < 4 Then
        MsgBox "El nombre de cuenta debe tener un minimo de 4 caracteres y un maximo de 24."
        checkCuenta = False
        Exit Function
    End If
    
    If LenB(txtEmail.Text) = 0 Then
        MsgBox "El campo de Email esta vacio."
        checkCuenta = False
        Exit Function
    End If
    
    If txtPass.Text = vbNullString Then
        MsgBox "El campo de contraseña esta vacio."
        checkCuenta = False
        Exit Function
    End If
    
    If Len(txtPass.Text) < 6 Then
        MsgBox "La contraseña debe contener al menos 6 caracteres."
        checkCuenta = False
        Exit Function
    End If
    
    If txtPass.Text <> txtRePass.Text Then
        MsgBox "Las contraseñas no coinciden"
        checkCuenta = False
        Exit Function
    End If
    
    checkCuenta = True
    

End Function

