VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16050
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0152
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Audio"
      ForeColor       =   &H00000000&
      Height          =   3540
      Left            =   7560
      TabIndex        =   13
      Top             =   120
      Width           =   4215
      Begin VB.HScrollBar scrMusic 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   20
         Top             =   3060
         Width           =   2895
      End
      Begin VB.HScrollBar scrVolume 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   19
         Top             =   1890
         Width           =   2895
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonido habilitado"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Música habilitada"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Sonidos Ambientales habilitado"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   2985
      End
      Begin VB.CheckBox chkop 
         Caption         =   "Invertir los canales (L/R)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   2985
      End
      Begin VB.HScrollBar scrAmbient 
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   14
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de música"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   23
         Top             =   2850
         Width           =   2865
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de sonidos ambientales"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   22
         Top             =   2280
         Width           =   2865
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Volúmen de audio"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   2835
      End
   End
   Begin VB.Frame FraSkins 
      Caption         =   "Skins"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   5040
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
      Begin VB.ComboBox cmdLenguajesComboBox 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ComboBox cmdSkinsComboBox 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmOpciones.frx":281E3
         Left            =   240
         List            =   "frmOpciones.frx":281E5
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox ComMouseHechizos 
         Height          =   315
         ItemData        =   "frmOpciones.frx":281E7
         Left            =   240
         List            =   "frmOpciones.frx":281F1
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox ComMouseGeneral 
         Height          =   315
         ItemData        =   "frmOpciones.frx":2820A
         Left            =   240
         List            =   "frmOpciones.frx":28217
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblLenguaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lenguaje"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label lblSkinDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skin de Interfaces"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lblMouseHechizos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Grafico de Hechizos"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1920
      End
      Begin VB.Label lblMouseGrafico 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Grafico General"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1620
      End
   End
   Begin VB.TextBox txtCantMensajes 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "5"
      Top             =   2415
      Width           =   255
   End
   Begin VB.TextBox txtLevel 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "40"
      Top             =   4110
      Width           =   255
   End
   Begin VB.Label lblDesactivarHUD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desactivar HUD"
      Height          =   195
      Left            =   5280
      TabIndex        =   12
      Top             =   480
      Width           =   1140
   End
   Begin VB.Image chkHud 
      Height          =   225
      Left            =   4920
      Top             =   480
      Width           =   210
   End
   Begin VB.Image chkLimitarFPS 
      Height          =   225
      Left            =   4920
      Top             =   120
      Width           =   210
   End
   Begin VB.Label lblLimitarFPS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Limitar FPS"
      Height          =   195
      Left            =   5280
      TabIndex        =   11
      Top             =   120
      Width           =   780
   End
   Begin VB.Image imgChkDesactivarFragShooter 
      Height          =   225
      Left            =   435
      Top             =   4740
      Width           =   210
   End
   Begin VB.Image imgChkAlMorir 
      Height          =   225
      Left            =   435
      Top             =   4425
      Width           =   210
   End
   Begin VB.Image imgChkRequiredLvl 
      Height          =   225
      Left            =   435
      Top             =   4110
      Width           =   210
   End
   Begin VB.Image imgChkNoMostrarNews 
      Height          =   225
      Left            =   2475
      Top             =   3315
      Width           =   210
   End
   Begin VB.Image imgChkMostrarNews 
      Height          =   225
      Left            =   435
      Top             =   3315
      Width           =   210
   End
   Begin VB.Image imgChkPantalla 
      Height          =   225
      Left            =   1950
      Top             =   2430
      Width           =   210
   End
   Begin VB.Image imgChkConsola 
      Height          =   225
      Left            =   435
      Top             =   2430
      Width           =   210
   End
   Begin VB.Image imgTutorial 
      Height          =   285
      Left            =   2520
      Top             =   6240
      Width           =   2010
   End
   Begin VB.Image imgSoporte 
      Height          =   285
      Left            =   360
      Top             =   6240
      Width           =   2010
   End
   Begin VB.Image imgManual 
      Height          =   285
      Left            =   360
      Top             =   5880
      Width           =   2010
   End
   Begin VB.Image imgMapa 
      Height          =   285
      Left            =   360
      Top             =   5520
      Width           =   2010
   End
   Begin VB.Image imgConfigTeclas 
      Height          =   285
      Left            =   360
      Top             =   5160
      Width           =   2010
   End
   Begin VB.Image imgSalir 
      Height          =   285
      Left            =   1440
      Top             =   6600
      Width           =   2010
   End
End
Attribute VB_Name = "frmOpciones"
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

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonConfigTeclas As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonSoporte As clsGraphicalButton
Private cBotonTutorial As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastButtonPressed As clsGraphicalButton

Private picCheckBox As Picture

Private bMusicActivated As Boolean
Private bSoundActivated As Boolean
Private bSoundEffectsActivated As Boolean

Private loading As Boolean

Private Sub chkHud_Click()
'***************************************************
'Author: Lorwik
'Last Modification: 30/04/2020
'30/04/2020: Lorwik - Desactivamos el HUD
'***************************************************
    ClientSetup.LimiteFPS = Not ClientSetup.LimiteFPS
    
    If ClientSetup.LimiteFPS Then
        chkLimitarFPS.Picture = picCheckBox
    Else
        Set chkLimitarFPS.Picture = Nothing
    End If
End Sub

Private Sub chkLimitarFPS_Click()
'***************************************************
'Author: Lorwik
'Last Modification: 28/04/2020
'20/04/2020: Lorwik - Seteamos el Limite de FPS
'***************************************************
    ClientSetup.LimiteFPS = Not ClientSetup.LimiteFPS
    
    If ClientSetup.LimiteFPS Then
        chkLimitarFPS.Picture = picCheckBox
    Else
        Set chkLimitarFPS.Picture = Nothing
    End If

End Sub

Private Sub chkop_Click(Index As Integer)
'***************************************************
'Author: Lorwik
'Fecha: 30/05/2020
'Descripcion: Activa/Desactiva los sonidos
'***************************************************

    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
    
        Case 0 'Musica
                    
            If chkop(Index).value = vbUnchecked Then
                Sound.Music_Stop
                ClientSetup.bMusic = CONST_DESHABILITADA
                scrMusic.Enabled = False
            Else
                ClientSetup.bMusic = CONST_MP3
                scrMusic.Enabled = True
            End If
        
        Case 1 'Sonido
    
            If chkop(Index).value = vbUnchecked Then
                'scrAmbient.Enabled = False
                scrVolume.Enabled = False
                ClientSetup.bSound = 0
            Else
                ClientSetup.bSound = 1
                scrVolume.Enabled = True
            End If
            
        Case 3 'Ambiente
            
            If chkop(Index).value = vbUnchecked Then
                ClientSetup.bAmbient = 0
                Call Sound.Sound_Stop_All
            Else
                ClientSetup.bAmbient = 1
                scrAmbient.Enabled = True
                Call Sound.Ambient_Load(Sound.AmbienteActual, ClientSetup.AmbientVol)
                Call Sound.Ambient_Play
            End If
            
    End Select
End Sub

Private Sub scrMusic_Change()
'***************************************************
'Author: Lorwik
'Fecha: 30/05/2020
'Descripcion: Setea el volumen de la musica
'***************************************************

    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        Sound.Music_Volume_Set scrMusic.value
        Sound.VolumenActualMusicMax = scrMusic.value
        ClientSetup.MusicVolume = Sound.VolumenActualMusicMax
    End If

End Sub

Private Sub scrAmbient_Change()
'***************************************************
'Author: Lorwik
'Fecha: 30/05/2020
'Descripcion: Setea el volumen del sonido ambiente
'***************************************************

    If ClientSetup.bAmbient = 1 Then
        Sound.VolumenActualAmbient_set scrAmbient.value
        ClientSetup.AmbientVol = Sound.VolumenActualAmbient
    End If
    
End Sub

Private Sub scrVolume_Change()
'***************************************************
'Author: Lorwik
'Fecha: 30/05/2020
'Descripcion: Setea el volumen de los sonidos
'***************************************************

    If ClientSetup.bSound = 1 Then
        Sound.VolumenActual = scrVolume.value
        ClientSetup.SoundVolume = Sound.VolumenActual
    End If

End Sub

Private Sub cmdLenguajesComboBox_Click()
'***************************************************
'Author: Recox
'Last Modification: 01/04/2019
'10/11/2019: Recox - Seteamos el lenguaje del juego
'***************************************************
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "Language", cmdLenguajesComboBox.Text)
    MsgBox ("Debe reiniciar el juego aplicar el cambio de idioma. Idioma Seleccionado: " & cmdLenguajesComboBox.Text)
End Sub

Private Sub cmdSkinsComboBox_Click()
'***************************************************
'Author: Recox
'Last Modification: 01/04/2019
'08/11/2019: Recox - Seteamos el skin
'***************************************************
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "SkinSelected", cmdSkinsComboBox.Text)
    MsgBox ("Debe reiniciar el juego aplicar el cambio de skin. Skin Seleccionado: " & cmdSkinsComboBox.Text)
End Sub

Private Sub ComMouseGeneral_Click()
'***************************************************
'Author: Lorwik
'Last Modification: 26/04/2020
'26/04/2020: Lorwik - Seteamos el mouse general
'***************************************************
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "MOUSEGENERAL", ComMouseGeneral.ListIndex)
    MsgBox ("Debe reiniciar el juego aplicar el cambio de mouse. Mouse Seleccionado: " & ComMouseGeneral.Text)
End Sub

Private Sub ComMouseHechizos_Click()
'***************************************************
'Author: Lorwik
'Last Modification: 26/04/2020
'26/04/2020: Lorwik - Seteamos el mouse baston
'***************************************************
    Call WriteVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "MOUSEBASTON", ComMouseHechizos.ListIndex)
    MsgBox ("Debe reiniciar el juego aplicar el cambio de mouse. Mouse Seleccionado: " & ComMouseHechizos.Text)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub imgChkAlMorir_Click()
    ClientSetup.bDie = Not ClientSetup.bDie
    
    If ClientSetup.bDie Then
        imgChkAlMorir.Picture = picCheckBox
    Else
        Set imgChkAlMorir.Picture = Nothing
    End If
End Sub

Private Sub imgChkDesactivarFragShooter_Click()
    ClientSetup.bActive = Not ClientSetup.bActive
    
    If ClientSetup.bActive Then
        Set imgChkDesactivarFragShooter.Picture = Nothing
    Else
        imgChkDesactivarFragShooter.Picture = picCheckBox
    End If
End Sub

Private Sub imgChkRequiredLvl_Click()
    ClientSetup.bKill = Not ClientSetup.bKill
    
    If ClientSetup.bKill Then
        imgChkRequiredLvl.Picture = picCheckBox
    Else
        Set imgChkRequiredLvl.Picture = Nothing
    End If
End Sub

Private Sub txtCantMensajes_Change()
    txtCantMensajes.Text = Val(txtCantMensajes.Text)
    
    If txtCantMensajes.Text > 0 Then
        DialogosClanes.CantidadDialogos = txtCantMensajes.Text
    Else
        txtCantMensajes.Text = 5
    End If
End Sub

Private Sub txtLevel_Change()
    If Not IsNumeric(txtLevel) Then txtLevel = 0
    txtLevel = Trim$(txtLevel)
    ClientSetup.byMurderedLevel = CByte(txtLevel)
End Sub

Private Sub imgChkConsola_Click()
    DialogosClanes.Activo = False
    
    imgChkConsola.Picture = picCheckBox
    Set imgChkPantalla.Picture = Nothing
End Sub

Private Sub imgChkMostrarNews_Click()
    ClientSetup.bGuildNews = True
    
    imgChkMostrarNews.Picture = picCheckBox
    Set imgChkNoMostrarNews.Picture = Nothing
End Sub

Private Sub imgChkNoMostrarNews_Click()
    ClientSetup.bGuildNews = False
    
    imgChkNoMostrarNews.Picture = picCheckBox
    Set imgChkMostrarNews.Picture = Nothing
End Sub

Private Sub imgChkPantalla_Click()
    DialogosClanes.Activo = True
    
    imgChkPantalla.Picture = picCheckBox
    Set imgChkConsola.Picture = Nothing
End Sub

Private Sub imgConfigTeclas_Click()
    If Not loading Then _
        Call Sound.Sound_Play(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub imgManual_Click()
    If Not loading Then _
        Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(0, "Open", "http://winterao.com.ar/wiki/", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub imgMapa_Click()
    Call frmMapa.Show(vbModal, Me)
End Sub

Private Sub imgSalir_Click()
    Call Carga.GuardarConfiguracion
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub imgSoporte_Click()
    
    If Not loading Then _
        Call Sound.Sound_Play(SND_CLICK)
    
    Call ShellExecute(0, "Open", "https://github.com/ao-libre/ao-cliente/issues", "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub imgTutorial_Click()
    frmTutorial.Show vbModeless
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    ' TODO: Traducir los textos de las imagenes via labels en visual basic, para que en el futuro si se quiere se pueda traducir a mas idiomas
    ' No ando con mas ganas/tiempo para hacer eso asi que se traducen las imagenes asi tenemos el juego en ingles.
    ' Tambien usar los controles uAObuttons para los botones, usar de ejemplo frmCambiaMotd.frm
    If Language = "spanish" Then
      Me.Picture = LoadPicture(Carga.Path(Interfaces) & "VentanaOpciones_spanish.jpg")
    Else
      Me.Picture = LoadPicture(Carga.Path(Interfaces) & "VentanaOpciones_english.jpg")
    End If

    LoadButtons
    LoadSkinsInComboBox
    LoadLenguajesInComboBox
    
    loading = True      'Prevent sounds when setting check's values
    LoadUserConfig
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub LoadSkinsInComboBox()
    Dim sFileName As String
    sFileName = Dir$(Carga.Path(Graficos) & "\Skins\", vbDirectory)
    
    Do While sFileName > vbNullString
        Call cmdSkinsComboBox.AddItem(sFileName)
        sFileName = Dir$()
    Loop
    
    'Boorramos los 2 primeros items por que son . y ..
    Call cmdSkinsComboBox.RemoveItem(0)
    Call cmdSkinsComboBox.RemoveItem(0)
End Sub

Private Sub LoadLenguajesInComboBox()
    Dim sFileName As String
    sFileName = Dir$(App.Path & "\Lenguajes\", vbArchive)
    
    Do While sFileName > vbNullString
        sFileName = Replace(sFileName, ".json", vbNullString)
        Call cmdLenguajesComboBox.AddItem(sFileName)
        sFileName = Dir$()
    Loop

End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    
    GrhPath = Carga.Path(Interfaces)

    Set cBotonConfigTeclas = New clsGraphicalButton
    Set cBotonMapa = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonSoporte = New clsGraphicalButton
    Set cBotonTutorial = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton
    
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonConfigTeclas.Initialize(imgConfigTeclas, GrhPath & "BotonConfigurarTeclas.jpg", _
                                    GrhPath & "BotonConfigurarTeclasRollover.jpg", _
                                    GrhPath & "BotonConfigurarTeclasClick.jpg", Me)
                                    
    Call cBotonMapa.Initialize(imgMapa, GrhPath & "BotonMapaAo.jpg", _
                                    GrhPath & "BotonMapaAoRollover.jpg", _
                                    GrhPath & "BotonMapaAoClick.jpg", Me)
                                    
    Call cBotonManual.Initialize(imgManual, GrhPath & "BotonManualAo.jpg", _
                                    GrhPath & "BotonManualAoRollover.jpg", _
                                    GrhPath & "BotonManualAoClick.jpg", Me)
                                    
    Call cBotonSoporte.Initialize(imgSoporte, GrhPath & "BotonSoporte.jpg", _
                                    GrhPath & "BotonSoporteRollover.jpg", _
                                    GrhPath & "BotonSoporteClick.jpg", Me)
                                    
    Call cBotonTutorial.Initialize(imgTutorial, GrhPath & "BotonTutorial.jpg", _
                                    GrhPath & "BotonTutorialRollover.jpg", _
                                    GrhPath & "BotonTutorialClick.jpg", Me)
                                    
    Call cBotonSalir.Initialize(imgSalir, GrhPath & "BotonSalirOpciones.jpg", _
                                    GrhPath & "BotonSalirRolloverOpciones.jpg", _
                                    GrhPath & "BotonSalirClickOpciones.jpg", Me)
                                    
    Set picCheckBox = LoadPicture(GrhPath & "CheckBoxOpciones.jpg")
End Sub

Private Sub LoadUserConfig()

    'Musica
    If ClientSetup.bMusic = CONST_DESHABILITADA Then
        chkop(0).value = 0
        scrMusic.value = ClientSetup.MusicVolume
    Else
        chkop(0).value = 1
        scrMusic.value = ClientSetup.MusicVolume
    End If
    
    'Sonidos
    If ClientSetup.bSound = 1 Then
        chkop(1).value = vbChecked
        chkop(4).value = IIf(ClientSetup.Invertido = True, 1, 0)
        scrVolume.value = ClientSetup.SoundVolume
    Else
        chkop(1).value = vbUnchecked
        chkop(4).value = IIf(ClientSetup.Invertido = True, 1, 0)
        chkop(4).Enabled = False
        scrVolume.value = ClientSetup.SoundVolume
        scrVolume.Enabled = False
    End If
    
    'Ambiente
    If ClientSetup.bAmbient = 1 Then
        chkop(3).value = vbChecked
        scrAmbient.value = ClientSetup.AmbientVol
    Else
        chkop(3).value = vbUnchecked
        scrAmbient.value = ClientSetup.AmbientVol
    End If

    txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)
    
    If DialogosClanes.Activo Then
        imgChkPantalla.Picture = picCheckBox
    Else
        imgChkConsola.Picture = picCheckBox
    End If
    
    If ClientSetup.bGuildNews Then
        imgChkMostrarNews.Picture = picCheckBox
    Else
        imgChkNoMostrarNews.Picture = picCheckBox
    End If
        
    If ClientSetup.bKill Then imgChkRequiredLvl.Picture = picCheckBox
    If ClientSetup.bDie Then imgChkAlMorir.Picture = picCheckBox
    If Not ClientSetup.bActive Then imgChkDesactivarFragShooter.Picture = picCheckBox
    
    txtLevel = ClientSetup.byMurderedLevel
    
    If ClientSetup.LimiteFPS Then chkLimitarFPS.Picture = picCheckBox
    If ClientSetup.HUD Then chkHud.Picture = picCheckBox
    
End Sub
