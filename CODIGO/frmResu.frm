VERSION 5.00
Begin VB.Form frmResu 
   BorderStyle     =   0  'None
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
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
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgAceptar 
      Height          =   540
      Left            =   3360
      Tag             =   "1"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image imgCerrar 
      Height          =   540
      Left            =   840
      Tag             =   "1"
      Top             =   990
      Width           =   1695
   End
   Begin VB.Label lblConfirmacion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   435
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5265
   End
End
Attribute VB_Name = "frmResu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton
Private cBotonAceptar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Load_Picture_From_Resource("215.bmp", False)
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()

   ' GrhPath = Carga.path(Interfaces)

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonAceptar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    If Language = "spanish" Then

        Call cBotonCerrar.Initialize(imgCerrar, "4.bmp", _
                                          "213.bmp", _
                                          "214.bmp", Me)
                                          
        Call cBotonAceptar.Initialize(imgAceptar, "3.bmp", _
                                          "216.bmp", _
                                          "217.bmp", Me)
    Else
    
        Call cBotonCerrar.Initialize(imgCerrar, "4.bmp", _
                                          "211.bmp", _
                                          "212.bmp", Me)
                                          
        Call cBotonAceptar.Initialize(imgAceptar, "6.bmp", _
                                          "218.bmp", _
                                          "219.bmp", Me)
        
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

