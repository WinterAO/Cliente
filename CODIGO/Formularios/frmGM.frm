VERSION 5.00
Begin VB.Form frmGM 
   BorderStyle     =   0  'None
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton OptDenuncia 
      BackColor       =   &H00000000&
      Caption         =   "Denuncia"
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
      Height          =   195
      Index           =   3
      Left            =   3330
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      BackColor       =   &H00000000&
      Caption         =   "Sugerencia"
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
      Height          =   195
      Index           =   2
      Left            =   2055
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton optConsulta 
      BackColor       =   &H00000000&
      Caption         =   " Bug"
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
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.OptionButton optConsulta 
      BackColor       =   &H00000000&
      Caption         =   "Soporte"
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
      Height          =   195
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   2760
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox TXTMessage 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3075
      Width           =   4215
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   315
      TabIndex        =   4
      Top             =   600
      Width           =   4170
   End
   Begin VB.Image CMDSalir 
      Height          =   255
      Left            =   2400
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image CMDEnviar 
      Height          =   255
      Left            =   960
      Top             =   5400
      Width           =   1455
   End
End
Attribute VB_Name = "frmGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'By Lorwik
'Form encargado de enviar el mensaje al GM
'************************************************************

Private Sub CMDEnviar_Click()

    Call Sound.Sound_Play(SND_CLICK)
    
    If TXTMessage.Text = "" Then
        MsgBox "Debes de escribir el motivo de tu consulta."
        Exit Sub
    End If
    
    If optConsulta(0).value = True Then
        Call WriteGMRequest(0, TXTMessage.Text)
        Unload Me
        Exit Sub
        
    ElseIf optConsulta(1).value = True Then
        Call WriteGMRequest(1, TXTMessage.Text)
        Unload Me
        Exit Sub
        
    ElseIf optConsulta(2).value = True Then
        Call WriteGMRequest(2, TXTMessage.Text)
        Unload Me
        Exit Sub
        
    ElseIf optConsulta(3).value = True Then
        Call WriteGMRequest(3, TXTMessage.Text)
        Unload Me
        Exit Sub
    End If
    
    
End Sub

Private Sub CMDSalir_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Picture = General_Load_Picture_From_Resource("210.gif")
    lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("SOPORTE")
End Sub

Private Sub optConsulta_Click(Index As Integer)
    Select Case Index
        Case 0
            lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("SOPORTE")
        Case 1
            lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("BUG")
        Case 2
            lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("SUGERENCIA")
        Case 3
            lblInfo.Caption = JsonLanguage.item("VENTANAGM").item("DENUNCIA")
    End Select

End Sub
