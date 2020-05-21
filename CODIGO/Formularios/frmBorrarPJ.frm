VERSION 5.00
Begin VB.Form frmBorrarPJ 
   BorderStyle     =   0  'None
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
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
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVolver 
      Caption         =   "Volver"
      Height          =   360
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   360
      Left            =   4920
      TabIndex        =   3
      Top             =   1920
      Width           =   1710
   End
   Begin VB.TextBox txtconfirmacion 
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
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label lblATENCIÓNESTAS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡ATENCIÓN ESTAS A PUNTO DE BORRAR UN PERSONAJE!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7035
   End
   Begin VB.Label lblEstasSeguro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escribe ""BORRAR XXXXX"" para eliminarlo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   5955
   End
End
Attribute VB_Name = "frmBorrarPJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBorrar_Click()
'*************************************
'Autor: Lorwik
'Fecha: 21/05/2020
'Descripcion: Preguntamos si desea eliminar el PJ y en caso de afirmacion mandamos eliminar
'*************************************

    '¿Paso la verificacion?
    If CheckBorrarData = False Then Exit Sub

    Select Case MsgBox(JsonLanguage.item("FRMPANELACCOUNT_CONFIRMAR_BORRAR_PJ").item("TEXTO"), vbYesNo + vbExclamation, JsonLanguage.item("FRMPANELACCOUNT_CONFIRMAR_BORRAR_PJ_TITULO").item("TEXTO"))
    
        Case vbYes
            
            Call WriteDeleteChar
            
            'Salimos del form
            Call cmdVolver_Click
            
        Case vbNo
            cmdVolver_Click
            Exit Sub
            
    End Select
End Sub

Private Sub cmdVolver_Click()
'*************************************
'Autor: Lorwik
'Fecha: 21/05/2020
'Descripcion: Reseteamos y salimos del form
'*************************************

    lblEstasSeguro.Caption = vbNullString
    
    Unload Me
End Sub

Private Function CheckBorrarData() As Boolean
'*************************************
'Autor: Lorwik
'Fecha: 21/05/2020
'Descripcion: Checkeamos antes de borrar
'*************************************

    '¿El indice es correcto?
    If PJAccSelected < 1 Or PJAccSelected > 10 Then
        Call MostrarMensaje("Error al borrar el PJ. Intentelo de nuevo o contacte con un Administrador.")
        CheckBorrarData = False
        Exit Function
    End If
    
    'Escribio la palabra magica?
    If Not txtconfirmacion.Text = "BORRAR " & cPJ(PJAccSelected).Nombre Then
        Call MostrarMensaje("Escribe BORRAR " & cPJ(PJAccSelected).Nombre & " para eliminar el personaje.")
        CheckBorrarData = False
        Exit Function
    End If
    
    CheckBorrarData = True
End Function

Private Sub Form_Load()
    lblEstasSeguro.Caption = "Escribe 'BORRAR " & cPJ(PJAccSelected).Nombre & "' para eliminarlo"
End Sub
