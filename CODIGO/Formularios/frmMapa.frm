VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Mapa"
   ClientHeight    =   12165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMapa.frx":0000
   ScaleHeight     =   811
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   849
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgCerrar 
      Height          =   345
      Left            =   12240
      MouseIcon       =   "frmMapa.frx":7C3A1
      Tag             =   "1"
      Top             =   165
      Width           =   345
   End
   Begin VB.Image imgToogleMap 
      Height          =   255
      Index           =   1
      Left            =   3840
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgToogleMap 
      Height          =   255
      Index           =   0
      Left            =   3960
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   735
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
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
'**************************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager

Private Enum eMaps
    ieGeneral
    ieDungeon
End Enum

Private picMaps(1)           As Picture
Private cBotonCerrar         As clsGraphicalButton
Public LastButtonPressed     As clsGraphicalButton

Private CurrentMap           As eMaps

''
' This form is used to show the world map.
' It has two levels. The world map and the dungeons map.
' You can toggle between them pressing the arrows
'
' @file     frmMapa.frm
' @author Marco Vanotti (MarKoxX) marcovanotti15@gmail.com
' @version 1.0.0
' @date 20080724

''
' Checks what Key is down. If the key is const vbKeyDown or const vbKeyUp, it toggles the maps, else the form unloads.
'
' @param KeyCode Specifies the key pressed
' @param Shift Specifies if Shift Button is pressed
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 24/07/08
'
'*************************************************

    Select Case KeyCode
        Case vbKeyDown, vbKeyUp 'Cambiamos el "nivel" del mapa, al estilo Zelda ;D
            ToggleImgMaps
        Case Else
            Unload Me
    End Select
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

''
' Toggle which image is visible.
'
Private Sub ToggleImgMaps()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 24/07/08
'
'*************************************************

    imgToogleMap(CurrentMap).Visible = False
    
    If CurrentMap = eMaps.ieGeneral Then
        imgCerrar.Visible = False
        CurrentMap = eMaps.ieDungeon
    Else
        imgCerrar.Visible = True
        CurrentMap = eMaps.ieGeneral
    End If
    
    imgToogleMap(CurrentMap).Visible = True
    Me.Picture = picMaps(CurrentMap)
End Sub

''
' Load the images. Resizes the form, adjusts image's left and top and set lblTexto's Top and Left.
'
Private Sub Form_Load()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 24/07/08
'
'*************************************************

On Error GoTo Error
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    Call LoadButtons
        
    'Cargamos las imagenes de los mapas
    Set picMaps(eMaps.ieGeneral) = General_Load_Picture_From_Resource("17.gif", False)
    Set picMaps(eMaps.ieDungeon) = General_Load_Picture_From_Resource("18.gif", False)
    
    ' Imagen de fondo
    CurrentMap = eMaps.ieGeneral
    Me.Picture = picMaps(CurrentMap)
    
    imgCerrar.MouseIcon = picMouseIcon
    imgToogleMap(0).MouseIcon = picMouseIcon
    imgToogleMap(1).MouseIcon = picMouseIcon
    
    Exit Sub
Error:
    MsgBox Err.Description, vbInformation, JsonLanguage.item("ERROR").item("TEXTO") & ": " & Err.number
    Unload Me
End Sub

Private Sub LoadButtons()

    Set LastButtonPressed = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(imgCerrar, "Cerrar.gif", _
                                "cerrar_hover.gif", _
                                "cerrar_click.gif", Me)
    
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgToogleMap_Click(Index As Integer)
    ToggleImgMaps
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgCerrar.Tag = 1 Then
        imgCerrar.Picture = General_Load_Picture_From_Resource("cerrar_click.gif", False)
        imgCerrar.Tag = 0
    End If

End Sub
