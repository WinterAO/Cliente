VERSION 5.00
Begin VB.Form frmAvisoBeta 
   BorderStyle     =   0  'None
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
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
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCerrar 
      Height          =   405
      Left            =   1440
      Tag             =   "1"
      Top             =   4800
      Width           =   1695
   End
End
Attribute VB_Name = "frmAvisoBeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonCerrar As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Load_Picture_From_Resource("209.gif", False)
    
    Call LoadButtons
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub LoadButtons()
    Dim boton As String
    
   ' GrhPath = Carga.path(Interfaces)

    Set cBotonCerrar = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    boton = "122.gif"
    
    Call cBotonCerrar.Initialize(imgCerrar, boton, _
                                      "125.gif", _
                                      "123.gif", Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub
