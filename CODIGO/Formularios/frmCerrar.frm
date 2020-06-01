VERSION 5.00
Begin VB.Form frmCerrar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cerrar"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmCerrar.frx":0000
   ScaleHeight     =   212
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin WinterAO.uAOButton cRegresar 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":9B51
      PICF            =   "frmCerrar.frx":9B6D
      PICH            =   "frmCerrar.frx":9B89
      PICV            =   "frmCerrar.frx":9BA5
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
   Begin WinterAO.uAOButton cSalir 
      CausesValidation=   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":9BC1
      PICF            =   "frmCerrar.frx":9BDD
      PICH            =   "frmCerrar.frx":9BF9
      PICV            =   "frmCerrar.frx":9C15
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
   Begin WinterAO.uAOButton cCancelQuit 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmCerrar.frx":9C31
      PICF            =   "frmCerrar.frx":9C4D
      PICH            =   "frmCerrar.frx":9C69
      PICV            =   "frmCerrar.frx":9C85
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
End
Attribute VB_Name = "frmCerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub cCancelQuit_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Set clsFormulario = Nothing
    Unload Me
End Sub

Private Sub cRegresar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    Set clsFormulario = Nothing
    
    If UserParalizado Then 'Inmo

        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NO_SALIR").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
        End With
        
        Exit Sub
        
    End If
    
    ' Nos desconectamos y lo mando al Panel de la Cuenta
    Call WriteQuit
    
    Call Unload(Me)
    
End Sub

Private Sub cSalir_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Set clsFormulario = Nothing
    Call CloseClient
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Set clsFormulario = Nothing
        Call Unload(Me)
    End If
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    Call clsFormulario.Initialize(Me)
    
    Me.Picture = LoadPicture(Carga.Path(Interfaces) & "frmCerrar.jpg")
    'Call LoadAOCustomControlsPictures(Me)
    'Todo: Poner la carga de botones como en el frmCambiaMotd.frm para mantener coherencia con el resto de la aplicacion
    'y poder borrar los frx de este archivo

    Call LoadFormTexts
End Sub

Private Sub LoadFormTexts()
    cRegresar.Caption = JsonLanguage.item("CERRAR").item("TEXTOS").item(1)
    cSalir.Caption = JsonLanguage.item("CERRAR").item("TEXTOS").item(2)
    cCancelQuit.Caption = JsonLanguage.item("CERRAR").item("TEXTOS").item(3)
End Sub

