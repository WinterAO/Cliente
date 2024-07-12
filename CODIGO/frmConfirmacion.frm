VERSION 5.00
Begin VB.Form frmConfirmacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCancelar 
      Height          =   525
      Left            =   240
      Top             =   2595
      Width           =   1695
   End
   Begin VB.Image imgAceptar 
      Height          =   525
      Left            =   2040
      Top             =   2595
      Width           =   1695
   End
   Begin VB.Label msg 
      BackStyle       =   0  'Transparent
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
      Height          =   1995
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmConfirmacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Load_Picture_From_Resource("mensaje.bmp", False)
    
    Call LoadButtons
End Sub

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub LoadButtons()

    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    If Language = "spanish" Then
        Call cBotonAceptar.Initialize(imgAceptar, _
            Carga.Path(Interfaces) & "btnaceptar_es.bmp", _
            Carga.Path(Interfaces) & "btnaceptar-over_es.bmp", _
            Carga.Path(Interfaces) & "btnaceptar-down_es.bmp", Me)
    Else
        Call cBotonAceptar.Initialize(imgAceptar, _
            Carga.Path(Interfaces) & "btnaceptar_en.bmp", _
            Carga.Path(Interfaces) & "btnaceptar-over_en.bmp", _
            Carga.Path(Interfaces) & "btnaceptar-down_en.bmp", Me)

    End If
                                     
    If Language = "spanish" Then
        Call cBotonCancelar.Initialize(imgCancelar, _
            Carga.Path(Interfaces) & "btncancelar_es.bmp", _
            Carga.Path(Interfaces) & "btncancelar-over_es.bmp", _
            Carga.Path(Interfaces) & "btncancelar-down_es.bmp", Me)
    Else
        Call cBotonCancelar.Initialize(imgCancelar, _
            Carga.Path(Interfaces) & "btncancelar_en.bmp", _
            Carga.Path(Interfaces) & "btncancelar-over_en.bmp", _
            Carga.Path(Interfaces) & "btncancelar-down_en.bmp", Me)

    End If
                                     
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgAceptar_Click()
    Call WriteRespuestaInstruccion(True)
    Unload Me
End Sub

Private Sub imgCancelar_Click()
    Call WriteRespuestaInstruccion(False)
    Unload Me
End Sub

Private Sub msg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

