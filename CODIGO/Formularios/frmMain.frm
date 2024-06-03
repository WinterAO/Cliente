VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   360
   ClientTop       =   -3300
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   6  'Inside Solid
   FillColor       =   &H00008080&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004080&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":1A041
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Sendtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00202020&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Left            =   1440
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1935
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.ListBox ListAmigos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      Left            =   12000
      TabIndex        =   24
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin WinterAOR_Client.uAOProgress uAOProgressExperienceLevel 
      Height          =   180
      Left            =   11520
      TabIndex        =   13
      ToolTipText     =   "Experiencia necesaria para pasar de nivel"
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   318
      BackColor       =   8421376
      BorderColor     =   0
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox MiniMapa 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   9450
      MouseIcon       =   "frmMain.frx":25A085
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   345
      Width           =   1500
      Begin VB.Shape UserAreaMinimap 
         BackColor       =   &H80000004&
         BorderColor     =   &H80000002&
         FillColor       =   &H000080FF&
         Height          =   315
         Left            =   555
         Top             =   585
         Width           =   375
      End
      Begin VB.Shape UserM 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H0000FFFF&
         Height          =   75
         Left            =   720
         Top             =   720
         Width           =   75
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   12000
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   2595
      Width           =   2400
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   9120
      Top             =   2880
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1665
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   240
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   2937
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":25A1D7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3150
      Left            =   11760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   9090
      Left            =   240
      MousePointer    =   99  'Custom
      ScaleHeight     =   606
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   720
      TabIndex        =   10
      Top             =   2310
      Width           =   10800
      Begin VB.Frame fMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   2685
         Left            =   9120
         TabIndex        =   16
         Top             =   5640
         Visible         =   0   'False
         Width           =   1575
         Begin WinterAOR_Client.uAOButton btnMapa 
            Height          =   255
            Left            =   120
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   ""
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":25A254
            PICF            =   "frmMain.frx":25A270
            PICH            =   "frmMain.frx":25A28C
            PICV            =   "frmMain.frx":25A2A8
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
         Begin WinterAOR_Client.uAOButton btnGrupo 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   ""
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":25A2C4
            PICF            =   "frmMain.frx":25A2E0
            PICH            =   "frmMain.frx":25A2FC
            PICV            =   "frmMain.frx":25A318
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
         Begin WinterAOR_Client.uAOButton btnEstadisticas 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   ""
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":25A334
            PICF            =   "frmMain.frx":25A350
            PICH            =   "frmMain.frx":25A36C
            PICV            =   "frmMain.frx":25A388
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
         Begin WinterAOR_Client.uAOButton btnClanes 
            Height          =   255
            Left            =   120
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   ""
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":25A3A4
            PICF            =   "frmMain.frx":25A3C0
            PICH            =   "frmMain.frx":25A3DC
            PICV            =   "frmMain.frx":25A3F8
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
         Begin WinterAOR_Client.uAOButton btnOpciones 
            Height          =   255
            Left            =   120
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   ""
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":25A414
            PICF            =   "frmMain.frx":25A430
            PICH            =   "frmMain.frx":25A44C
            PICV            =   "frmMain.frx":25A468
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
         Begin WinterAOR_Client.uAOButton btnQuest 
            Height          =   255
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   ""
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":25A484
            PICF            =   "frmMain.frx":25A4A0
            PICH            =   "frmMain.frx":25A4BC
            PICV            =   "frmMain.frx":25A4D8
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
         Begin WinterAOR_Client.uAOButton btnPvP 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            TX              =   ""
            ENAB            =   -1  'True
            FCOL            =   16777215
            OCOL            =   16777215
            PICE            =   "frmMain.frx":25A4F4
            PICF            =   "frmMain.frx":25A510
            PICH            =   "frmMain.frx":25A52C
            PICV            =   "frmMain.frx":25A548
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
   End
   Begin VB.Label lblInvisible 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   5760
      TabIndex        =   36
      Top             =   -30
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblBuscarNpc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Npc/Obj"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   4140
      TabIndex        =   35
      Top             =   -30
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPanelGM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PanelGM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Left            =   2340
      TabIndex        =   34
      Top             =   -30
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgClima 
      Height          =   360
      Left            =   13575
      Top             =   10410
      Width           =   1140
   End
   Begin VB.Image picSM 
      Height          =   345
      Index           =   1
      Left            =   14250
      Top             =   9960
      Width           =   375
   End
   Begin VB.Image picSM 
      Height          =   345
      Index           =   0
      Left            =   13665
      Top             =   9960
      Width           =   375
   End
   Begin VB.Label lblHour 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   13920
      TabIndex        =   33
      Top             =   10740
      Width           =   465
   End
   Begin VB.Image btnShop 
      Height          =   360
      Left            =   11700
      MouseIcon       =   "frmMain.frx":25A564
      Tag             =   "1"
      Top             =   10365
      Width           =   1410
   End
   Begin VB.Label LbLChat 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. Normal"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmMain.frx":25A6B6
      TabIndex        =   31
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Image MensajeAmigo 
      Height          =   360
      Left            =   12810
      Top             =   5625
      Width           =   375
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H80000018&
      Height          =   210
      Left            =   12480
      TabIndex        =   29
      Top             =   9240
      Width           =   450
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H80000018&
      Height          =   210
      Left            =   11880
      TabIndex        =   28
      Top             =   9240
      Width           =   510
   End
   Begin VB.Image shpSed 
      Height          =   330
      Left            =   13575
      Top             =   8880
      Width           =   120
   End
   Begin VB.Image shpHambre 
      Height          =   300
      Left            =   14010
      Top             =   8880
      Width           =   465
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   11910
      TabIndex        =   27
      Top             =   8430
      Width           =   2775
   End
   Begin VB.Image shpEnergia 
      Height          =   210
      Left            =   11925
      Picture         =   "frmMain.frx":25A808
      Top             =   8415
      Width           =   2775
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   11970
      TabIndex        =   26
      Top             =   8070
      Width           =   2745
   End
   Begin VB.Image shpMana 
      Height          =   210
      Left            =   11925
      Picture         =   "frmMain.frx":25C6B4
      Top             =   8070
      Width           =   2775
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   11940
      TabIndex        =   25
      Top             =   7620
      Width           =   2715
   End
   Begin VB.Image shpVida 
      Height          =   210
      Left            =   11940
      Picture         =   "frmMain.frx":25E560
      Top             =   7620
      Width           =   2775
   End
   Begin VB.Image btnSolapa 
      Height          =   585
      Index           =   2
      Left            =   14400
      MouseIcon       =   "frmMain.frx":26040C
      ToolTipText     =   "Social"
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image BorrarAmigo 
      Height          =   300
      Left            =   13800
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image AgregarAmigo 
      Height          =   300
      Left            =   13320
      Top             =   5625
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image btnInfo 
      Height          =   495
      Left            =   13575
      MouseIcon       =   "frmMain.frx":26055E
      MousePointer    =   99  'Custom
      Top             =   5940
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image btnLanzar 
      Height          =   540
      Left            =   11640
      MouseIcon       =   "frmMain.frx":2606B0
      MousePointer    =   99  'Custom
      Top             =   5940
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   165
      Left            =   14055
      TabIndex        =   23
      Top             =   9240
      Width           =   465
   End
   Begin VB.Label lblGems 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   165
      Left            =   13920
      TabIndex        =   22
      Top             =   6750
      Width           =   105
   End
   Begin VB.Image btnMenu 
      Height          =   360
      Left            =   11700
      MouseIcon       =   "frmMain.frx":260802
      Tag             =   "1"
      Top             =   9885
      Width           =   1410
   End
   Begin VB.Image btnSolapa 
      Height          =   600
      Index           =   1
      Left            =   12975
      MouseIcon       =   "frmMain.frx":260954
      ToolTipText     =   "Hechizos"
      Top             =   1920
      Width           =   1380
   End
   Begin VB.Image btnSolapa 
      Height          =   600
      Index           =   0
      Left            =   11550
      MouseIcon       =   "frmMain.frx":260AA6
      ToolTipText     =   "Inventario"
      Top             =   1920
      Width           =   1380
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡Nivel Máximo!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   12105
      TabIndex        =   15
      Top             =   900
      Width           =   2190
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del mapa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   9420
      TabIndex        =   12
      Top             =   2010
      Width           =   1575
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13800
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label lblMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   14565
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   14880
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   -120
      Width           =   375
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   360
      Index           =   0
      Left            =   14760
      MouseIcon       =   "frmMain.frx":260BF8
      MousePointer    =   99  'Custom
      Top             =   2925
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   360
      Index           =   1
      Left            =   14760
      MouseIcon       =   "frmMain.frx":260D4A
      MousePointer    =   99  'Custom
      Top             =   2580
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image xz 
      Height          =   255
      Index           =   0
      Left            =   14760
      Top             =   -120
      Width           =   255
   End
   Begin VB.Image xzz 
      Height          =   195
      Index           =   1
      Left            =   14805
      Top             =   -120
      Width           =   225
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del pj"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11520
      TabIndex        =   14
      Top             =   540
      Width           =   3345
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   13035
      TabIndex        =   6
      Top             =   1470
      Width           =   405
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   165
      Left            =   12120
      TabIndex        =   3
      Top             =   6750
      Width           =   105
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   11400
      TabIndex        =   1
      Top             =   11205
      Width           =   3675
   End
   Begin VB.Image InvEqu 
      Height          =   4530
      Left            =   11400
      Top             =   1920
      Width           =   3645
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   165
      Left            =   13440
      TabIndex        =   2
      Top             =   9240
      Width           =   465
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmMain
'    Project    : ARGENTUM
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
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

Public tX                  As Integer
Public tY                  As Integer
Public MouseX              As Long
Public MouseY              As Long
Public MouseBoton          As Long
Public MouseShift          As Long
Private clicX              As Long
Private clicY              As Long

Private clsFormulario      As clsFormMovementManager
Private cBotonShop         As clsGraphicalButton
Private cBotonMenu         As clsGraphicalButton

Public LastButtonPressed   As clsGraphicalButton

Public WithEvents Client   As clsSocket
Attribute Client.VB_VarHelpID = -1

Private ChangeHechi        As Boolean, ChangeHechiNum As Integer

Private FirstTimeChat      As Boolean

'Usado para controlar que no se dispare el binding de la tecla CTRL cuando se usa CTRL+Tecla.
Dim CtrlMaskOn             As Boolean

Private Const NEWBIE_USER_GOLD_COLOR As Long = vbCyan
Private Const USER_GOLD_COLOR As Long = vbYellow

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call Protocol_Write.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
End Sub

Private Sub btnGrupo_Click()
    Call WriteRequestPartyForm
End Sub

Private Sub btnMenu_Click()
    fMenu.Visible = Not fMenu.Visible
End Sub

Private Sub btnPvP_Click()
    Call WriteInitPVP
End Sub

Private Sub btnShop_Click()
    Call WriteShopInit
End Sub


Private Sub btnQuest_Click()
    Call WriteQuestListRequest
End Sub

Private Sub btnSolapa_Click(Index As Integer)
Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
    
        Case 0 'Inventario
            InvEqu.Picture = General_Load_Picture_From_Resource("4.gif", True)
            'btnSolapa(0).Picture = General_Load_Picture_From_Resource("7.gif", True)
            'btnSolapa(1).Picture = General_Load_Picture_From_Resource("10.gif", True)
            'btnSolapa(2).Picture = General_Load_Picture_From_Resource("12.gif", True)
            
            ' Activo controles de inventario
            picInv.Visible = True
        
            ' Desactivo controles de hechizo y amigos
            hlst.Visible = False
            btnInfo.Visible = False
            btnLanzar.Visible = False
            
            cmdMoverHechi(0).Visible = False
            cmdMoverHechi(1).Visible = False
            
            ListAmigos.Visible = False
            AgregarAmigo.Visible = False
            BorrarAmigo.Visible = False
            
            DoEvents
            Call Inventario.DrawInventory
        
        Case 1 'Hechizos
            InvEqu.Picture = General_Load_Picture_From_Resource("5.gif", True)
            'btnSolapa(0).Picture = General_Load_Picture_From_Resource("10.gif", True)
            'btnSolapa(1).Picture = General_Load_Picture_From_Resource("8.gif", True)
            'btnSolapa(2).Picture = General_Load_Picture_From_Resource("12.gif", True)
            'btnLanzar.Picture = General_Load_Picture_From_Resource("13.gif", True)
            'btnInfo.Picture = General_Load_Picture_From_Resource("14.gif", True)
            
            ' Activo controles de hechizos
            hlst.Visible = True
            btnInfo.Visible = True
            btnLanzar.Visible = True
            
            cmdMoverHechi(0).Visible = True
            cmdMoverHechi(1).Visible = True
            
            ' Desactivo controles de inventario y amigos
            picInv.Visible = False
            
            ListAmigos.Visible = False
            AgregarAmigo.Visible = False
            BorrarAmigo.Visible = False
    
        Case 2 'Amigos
            InvEqu.Picture = General_Load_Picture_From_Resource("6.gif", True)
            'btnSolapa(0).Picture = General_Load_Picture_From_Resource("10.gif", True)
            'btnSolapa(1).Picture = General_Load_Picture_From_Resource("11.gif", True)
            'btnSolapa(2).Picture = General_Load_Picture_From_Resource("9.gif", True)
            
            ListAmigos.Visible = True
            AgregarAmigo.Visible = True
            BorrarAmigo.Visible = True
            
            ' Desactivo controles de inventario y hechizos
            picInv.Visible = False
            
            hlst.Visible = False
            btnInfo.Visible = False
            btnLanzar.Visible = False
            
            cmdMoverHechi(0).Visible = False
            cmdMoverHechi(1).Visible = False
            
    End Select
End Sub

Private Sub Form_Activate()

    Call Inventario.DrawInventory

End Sub

Private Sub Form_Load()
    ClientSetup.SkinSeleccionado = GetVar(Carga.Path(Init) & CLIENT_FILE, "Parameters", "SkinSelected")
    
    Me.Picture = General_Load_Picture_From_Resource("es_main.bmp", True)
    cmdMoverHechi(1).Picture = General_Load_Picture_From_Resource("btnspell_top.bmp", True)
    cmdMoverHechi(0).Picture = General_Load_Picture_From_Resource("btnspell_down.bmp", True)
    btnLanzar.Picture = General_Load_Picture_From_Resource("es_btnlaunch.bmp", True)
    btnInfo.Picture = General_Load_Picture_From_Resource("es_btninfo.bmp", True)
    
    InvEqu.Picture = General_Load_Picture_From_Resource("inventorybackground.bmp", True) 'Fondo inventario
    
    btnSolapa(0).Picture = General_Load_Picture_From_Resource("es_btninventory_on.bmp", True) 'Inventario
    btnSolapa(1).Picture = General_Load_Picture_From_Resource("es_btnspells_off.bmp", True) 'Hechizos
    btnSolapa(2).Picture = General_Load_Picture_From_Resource("es_btnsocial_off.bmp", True) 'Social
    
    shpVida.Picture = General_Load_Picture_From_Resource("lifebar.bmp", True)
    shpMana.Picture = General_Load_Picture_From_Resource("manabar.bmp", True)
    shpEnergia.Picture = General_Load_Picture_From_Resource("staminabar.bmp", True)
    shpHambre.Picture = General_Load_Picture_From_Resource("foodbar.bmp", True)
    shpSed.Picture = General_Load_Picture_From_Resource("waterbar.bmp", True)
    
    If Not ResolucionCambiada Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        Call clsFormulario.Initialize(Me, 120)
    End If
        
    Call LoadButtons
    
    ' Seteamos el caption
    Me.Caption = Form_Caption
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)

    ' Reseteamos el tamanio de la ventana para que no queden bordes blancos
    Me.Width = 15360
    Me.Height = 11520

    Call LoadTextsForm
        
    ' Detect links in console
    Call EnableURLDetect(RecTxt.hWnd, Me.hWnd)
    
    ' Make the console transparent
    'Call SetWindowLong(RecTxt.hWnd, -20, &H20&)
    RecTxt.BackColor = RGB(24, 23, 21)
    
    CtrlMaskOn = False
    
    FirstTimeChat = True
    SendingType = 1
    
End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    Dim i As Integer

    Set LastButtonPressed = New clsGraphicalButton
    
    lblDropGold.MouseIcon = picMouseIcon
    lblCerrar.MouseIcon = picMouseIcon
    lblMinimizar.MouseIcon = picMouseIcon
    
    Set cBotonMenu = New clsGraphicalButton
    Set cBotonShop = New clsGraphicalButton
    
    Call cBotonMenu.Initialize(btnMenu, "btnmenu.bmp", _
                                     "btnmenu_over.bmp", _
                                     "btnmenu_down.bmp", Me, , , , , True)
    
    Call cBotonShop.Initialize(btnShop, "btnshop.bmp", _
                                     "btnshop_over.bmp", _
                                     "btnshop_down.bmp", Me, , , , , True)

End Sub

Private Sub LoadTextsForm()
    btnMapa.Caption = JsonLanguage.item("LBL_MAPA").item("TEXTO")
    btnGrupo.Caption = JsonLanguage.item("LBL_GRUPO").item("TEXTO")
    fMenu.Caption = JsonLanguage.item("LBL_GRUPO").item("TEXTO")
    btnOpciones.Caption = JsonLanguage.item("LBL_OPCIONES").item("TEXTO")
    btnEstadisticas.Caption = JsonLanguage.item("LBL_ESTADISTICAS").item("TEXTO")
    btnQuest.Caption = JsonLanguage.item("LBL_QUEST").item("TEXTO")
    btnClanes.Caption = JsonLanguage.item("LBL_CLANES").item("TEXTO")
End Sub
Private Sub cmdMoverHechi_Click(Index As Integer)

    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
    
        Select Case Index

            Case 1 'subir

                If hlst.ListIndex = 0 Then Exit Sub

            Case 0 'bajar

                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
    
        Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)
        
        Select Case Index

            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1

            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 18/11/2010
    '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
    '18/11/2010: Amraphen - Agregue el handle correspondiente para las nuevas configuraciones de teclas (CTRL+0..9).
    '***************************************************
    If (Not SendTxt.Visible) Then
        
        If KeyCode = vbKeyControl Then

            'Chequeo que no se haya usado un CTRL + tecla antes de disparar las bindings.
            If CtrlMaskOn Then
                CtrlMaskOn = False
                Exit Sub
            End If
        End If
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If ClientSetup.bMusic = CONST_MP3 Then
                        Sound.Music_Stop
                        ClientSetup.bMusic = CONST_DESHABILITADA
                    Else
                        ClientSetup.bMusic = CONST_MP3
                    End If
                        
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    'Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
                    ClientSetup.FPSShow = Not ClientSetup.FPSShow
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyQuests)
                    Call WriteQuestListRequest
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)

                    If CurrentUser.UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)

                    If CurrentUser.UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)

                    If CurrentUser.UserEstado = 1 Then

                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)

                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleCombatSafe)
                    Call WriteCombatToggle
                    
            End Select
            
        End If
        
    
        Select Case KeyCode
            Case CustomKeys.BindedKey(eKeyType.mKeyChatNormal)
                SendingType = 1
                If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
                lblChat.Caption = "1.Normal"
            
            Case CustomKeys.BindedKey(eKeyType.mKeyChatGritar)
                SendingType = 2
                If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
                lblChat.Caption = "2.Gritar"
            
            Case CustomKeys.BindedKey(eKeyType.mKeyChatPrivado)
                sndPrivateTo = InputBox("Nombre del destinatario:", vbNullString)
    
                If sndPrivateTo <> vbNullString Then
                    SendingType = 3
                    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
                Else
                    MsgBox "¡Escribe un nombre."
                End If
                lblChat.Caption = "3.Privado"
            
            Case CustomKeys.BindedKey(eKeyType.mKeyChatClan)
                SendingType = 4
                If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
                lblChat.Caption = "4.Clan"
            
            Case CustomKeys.BindedKey(eKeyType.mKeyChatGrupo)
                SendingType = 5
                If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
                lblChat.Caption = "5.Party"
            
            Case CustomKeys.BindedKey(eKeyType.mKeyChatGlobal)
                SendingType = 6
                If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
                lblChat.Caption = "6.Global"
            
            Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
                Call Mod_General.Client_Screenshot(frmMain.hDC, frmMain.ScaleWidth, frmMain.ScaleHeight)
                    
            Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
                Call frmOpciones.Show(vbModeless, frmMain)
                
            Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
    
                Call WriteQuit
                
            Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
    
                If Shift <> 0 Then Exit Sub
                
                If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                    If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
                Else
    
                    If Not MainTimer.Check(TimersIndex.Attack) Or CurrentUser.UserDescansar Or CurrentUser.UserMeditar Then Exit Sub
                End If
                
                If frmCustomKeys.Visible Then Exit Sub 'Chequeo si esta visible la ventana de configuracion de teclas.
                
                Call WriteAttack
                
            Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
                
                If (Not Comerciando) And (Not MirandoAsignarSkills) And (Not frmMSG.Visible) And (Not MirandoForo) And (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) And (Not frmPVP.Visible) And (Not frmShop.Visible) Then
                    Call CompletarEnvioMensajes
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                Else
                    Call Enviar_SendTxt
                End If
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionUno)
                If Len(ClientSetup.Funcion(1)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(1))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionDos)
                If Len(ClientSetup.Funcion(2)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(2))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionTres)
                If Len(ClientSetup.Funcion(3)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(3))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionCuatro)
                If Len(ClientSetup.Funcion(4)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(4))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionCinco)
                If Len(ClientSetup.Funcion(5)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(5))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionSeis)
                If Len(ClientSetup.Funcion(6)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(6))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionSiete)
                If Len(ClientSetup.Funcion(7)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(7))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionOcho)
                If Len(ClientSetup.Funcion(8)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(8))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionNueve)
                If Len(ClientSetup.Funcion(9)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(9))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionDiez)
                If Len(ClientSetup.Funcion(10)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(10))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionOnce)
                If Len(ClientSetup.Funcion(11)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(11))
                
            Case CustomKeys.BindedKey(eKeyType.mKeyFuncionDoce)
                If Len(ClientSetup.Funcion(12)) > 0 Then _
                    Call ParseUserCommand(ClientSetup.Funcion(12))
                
                
        End Select
     End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    clicX = x
    clicY = y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call DisableURLDetect
    
End Sub

Private Sub btnClanes_Click()
    
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub btnEstadisticas_Click()

    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer

    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    
    Alocados = SkillPoints
    frmEstadisticas.lblLibres.Caption = SkillPoints
    
    Call frmEstadisticas.MostrarAsignacion
    
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    
End Sub

Private Sub btnMapa_Click()
    
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub btnOpciones_Click()
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub InvEqu_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub lblScroll_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub lblCerrar_Click()
    Call Sound.Sound_Play(SND_CLICK)
    
    If CurrentUser.UserParalizado Then 'Inmo

        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NO_SALIR").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
        End With
        
        Exit Sub
        
    End If
    
    ' Nos desconectamos y lo mando al Panel de la Cuenta
    Call WriteQuit
End Sub

Private Sub LbLChat_Click()
    frmMensaje.PopupMenuMensaje
End Sub

Private Sub lblMana_Click()

   Call ParseUserCommand("/MEDITAR")
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub lblPanelGM_Click()
    On Error GoTo lblPanelGM_Click_Err
    
    'frmPanelGm.Width = 4860
    Call WriteSOSShowList
    Call WriteGMPanel(0)
    
    Me.SetFocus
    
    Exit Sub

lblPanelGM_Click_Err:
    Call LogError(Err.number, Err.Description, "frmMain.lblpanelGM_Click", Erl)
    Resume Next
End Sub

Private Sub lblBuscarNpc_Click()
    On Error GoTo lblBuscarNpc_Click_Err
    
    Call WriteGMPanel(1)
    Exit Sub

lblBuscarNpc_Click_Err:
    Call LogError(Err.number, Err.Description, "frmMain.lblBuscarNpc_Click", Erl)
    Resume Next
End Sub

Private Sub lblInvisible_Click()
    On Error GoTo lblInvisible_Click_Err
    
    Call ParseUserCommand("/INVISIBLE")
    
    Me.SetFocus
    
    Exit Sub

lblInvisible_Click_Err:
    Call LogError(Err.number, Err.Description, "frmMain.lblInvisible_Click", Erl)
    Resume Next
End Sub


Private Sub MensajeAmigo_Click()
    If ListAmigos.List(ListAmigos.ListIndex) = "" Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NO_AMIGOS").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
        End With
        Exit Sub
    End If
    
    SendTxt.Visible = True
    SendTxt.Text = ("\" & ListAmigos.List(ListAmigos.ListIndex) & " ")
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_AUTO_CAST_SPELL").item("COLOR").item(3), _
                        False, False, True)
End Sub

Private Sub Coord_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("TEXTO"), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(1), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(2), _
                                          JsonLanguage.item("MENSAJE_INFO_COORDENADAS").item("COLOR").item(3), _
                          False, False, True)
End Sub

Private Sub picSM_Click(Index As Integer)
    Select Case Index
    
        Case 0 'Modo combate
            Call WriteCombatToggle
            
        Case 1 'Seguro
            Call WriteSafeToggle
    
    End Select
End Sub

Private Sub RecTxt_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
    StartCheckingLinks
End Sub

Private Sub SendTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Para borrar el mensaje de fondo
    If FirstTimeChat Then
        SendTxt.Text = vbNullString
        FirstTimeChat = False
        ' Cambiamos el color de texto al original
        SendTxt.ForeColor = &HE0E0E0
    End If
    
errhandler:
    
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = vbNullString
        SendTxt.Text = vbNullString
        KeyCode = 0
        SendTxt.Visible = False
        
        If picInv.Visible Then
            picInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        Else
            ListAmigos.SetFocus
        End If
    End If
End Sub

Private Sub Second_Timer()

    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
    
    Call ActualizarHora
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()

    If CurrentUser.UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
    Else

        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else

                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()

    If CurrentUser.UserEstado = 1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()

    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call WriteUseItem(Inventario.SelectedItem)
    End If
    
End Sub

Private Sub EquiparItem()

    If CurrentUser.UserEstado = 1 Then
    
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
        End With
        
    Else
    
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
            Call WriteEquipItem(Inventario.SelectedItem)
        End If
        
    End If
End Sub

Private Sub btnLanzar_Click()
    
    If hlst.List(hlst.ListIndex) <> JsonLanguage.item("NADA").item("TEXTO") And MainTimer.Check(TimersIndex.Work, False) Then
        If CurrentUser.UserEstado = 1 Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
    
End Sub

Private Sub btnLanzar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub btnInfo_Click()
    
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
    
End Sub

Private Sub AgregarAmigo_Click()

    Dim SendName As String
        SendName = InputBox("Escriba el nombre del usuario a agregar.", "Agregar Amigo")

    If LenB(Trim$(SendName)) Then
        
        If MsgBox("Seguro desea agregar a " & SendName & "?", vbYesNo, "Agregar Amigo") = vbYes Then
           Call WriteAddAmigo(SendName, 1)
        End If
        
    Else

        With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
            Call ShowConsoleMsg("Nombre Invalido", .Red, .Green, .Blue, .bold, .italic)
        End With

    End If

End Sub

Private Sub BorrarAmigo_Click()

    If LenB(ListAmigos.List(ListAmigos.ListIndex)) = 0 Then Exit Sub
    
    If MsgBox("Seguro desea borrar a " & ListAmigos.List(ListAmigos.ListIndex) & "?", vbYesNo, "Borrar Amigo") = vbYes Then
        Call WriteDelAmigo(ListAmigos.ListIndex + 1)
    End If
    
End Sub

Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
    '¿Hizo click derecho?
    If Button = 2 Then
        If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
            Call WriteAccionClick(tX, tY)
        End If
    End If
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    MouseX = x
    MouseY = y
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    clicX = x
    clicY = y
End Sub

Private Sub MainViewPic_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteAccionClick(tX, tY)
    End If
    
End Sub

Private Sub MainViewPic_Click()

    'Si el menu esta abierto, lo cerramos.
    If fMenu.Visible Then fMenu.Visible = False

    If Cartel Then Cartel = False
    
    Dim MENSAJE_ADVERTENCIA As String
    Dim VAR_LANZANDO        As String
    
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        If Not InGameArea() Then Exit Sub
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then

                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1

                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If

                'Invitando party
                If InvitandoParty = True Then
                    frmMain.MousePointer = vbDefault
                    Call WriteInvitarPartyClick(tX, tY)
                    InvitandoParty = False
                    Exit Sub
                End If
    
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0

                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            VAR_LANZANDO = JsonLanguage.item("PROYECTILES").item("TEXTO")
                            MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                            MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                            
                            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0

                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                VAR_LANZANDO = JsonLanguage.item("PROYECTILES").item("TEXTO")
                                MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                                MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                                
                                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0

                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    VAR_LANZANDO = JsonLanguage.item("HECHIZOS").item("TEXTO")
                                    MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                                    MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                                    
                                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else

                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0

                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    VAR_LANZANDO = JsonLanguage.item("HECHIZOS").item("TEXTO")
                                    MENSAJE_ADVERTENCIA = JsonLanguage.item("MENSAJE_MACRO_ADVERTENCIA").item("TEXTO")
                                    MENSAJE_ADVERTENCIA = Replace$(MENSAJE_ADVERTENCIA, "VAR_LANZADO", VAR_LANZANDO)
                                    
                                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ADVERTENCIA, .Red, .Green, .Blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                'Call WriteRightClick(tx, tY) 'Proximamnete lo implementaremos..
                Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then

            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", CurrentUser.UserMap, tX, tY, False)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteAccionClick(tX, tY)
    End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseX = x - MainViewPic.Left
    MouseY = y - MainViewPic.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewPic.Width Then
        MouseX = MainViewPic.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewPic.Height Then
        MouseY = MainViewPic.Height
    End If
    
    LastButtonPressed.ToggleToNormal
    
    ' Disable links checking (not over consola)
    StopCheckingLinks
    
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub lblDropGold_Click()

    Inventario.SelectGold

    If CurrentUser.UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
    
End Sub

Private Sub picInv_DblClick()
'**********************************************
'Autor: Lorwik
'Fecha: 14/07/2020
'Descripcion: DobleClick sobre el inventario
'**********************************************
    'Esta validacion es para que el juego no rompa si hacemos doble click
    If MirandoTrabajo > 0 Then Exit Sub
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    '¿Es un slot valido?
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        Call WriteAccionInventario(Inventario.SelectedItem)
    End If
    
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Sound.Sound_Play(SND_CLICK)
End Sub

Private Sub RecTxt_Change()

    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar

    If Not Application.IsAppActive() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    
    ElseIf (Not Comerciando) And _
           (Not MirandoAsignarSkills) And _
           (Not frmMSG.Visible) And _
           (Not MirandoForo) And _
           (Not frmEstadisticas.Visible) And _
           (Not frmCantidad.Visible) And _
           (Not frmPVP.Visible) And _
           (Not frmShop.Visible) And _
           (Not MirandoParty) Then

        If picInv.Visible Then
            picInv.SetFocus
                        
        ElseIf hlst.Visible Then
            hlst.SetFocus

        End If

    End If

End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If

End Sub

Private Sub SendTxt_Change()

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedi se inserten caracteres no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = JsonLanguage.item("MENSAJE_SOY_CHEATER").item("TEXTO")
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i         As Long
        Dim tempstr   As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))

            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)

    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then KeyAscii = 0
End Sub

Private Sub CompletarEnvioMensajes()

    Select Case SendingType
        Case 1
            SendTxt.Text = vbNullString
        Case 2
            SendTxt.Text = "-"
        Case 3
            SendTxt.Text = ("\" & sndPrivateTo & " ")
        Case 4
            SendTxt.Text = "/CMSG "
        Case 5
            SendTxt.Text = "/PMSG "
        Case 6
            SendTxt.Text = "; "
    End Select
    
    stxtbuffer = SendTxt.Text
    SendTxt.SelStart = Len(SendTxt.Text)

End Sub

Private Sub Enviar_SendTxt()
    
    Dim str1 As String
    Dim str2 As String
    
    If Len(stxtbuffer) > 255 Then stxtbuffer = mid$(stxtbuffer, 1, 255)
    
    'Send text
    If Left$(stxtbuffer, 1) = "/" Then
        Call ParseUserCommand(stxtbuffer)

    'Shout
    ElseIf Left$(stxtbuffer, 1) = "-" Then
        If Right$(stxtbuffer, Len(stxtbuffer) - 1) <> vbNullString Then Call ParseUserCommand(stxtbuffer)
        SendingType = 2
        
    'Global
    ElseIf Left$(stxtbuffer, 1) = ";" Then
        If LenB(Right$(stxtbuffer, Len(stxtbuffer) - 1)) > 0 And InStr(stxtbuffer, ">") = 0 Then Call ParseUserCommand(stxtbuffer)
        SendingType = 6

    'Privado
    ElseIf Left$(stxtbuffer, 1) = "\" Then
        str1 = Right$(stxtbuffer, Len(stxtbuffer) - 1)
        str2 = ReadField(1, str1, 32)
        If LenB(str1) > 0 And InStr(str1, ">") = 0 Then Call ParseUserCommand("\" & str1)
        sndPrivateTo = str2
        SendingType = 3
                
    'Say
    Else
        If LenB(stxtbuffer) > 0 Then Call ParseUserCommand(stxtbuffer)
        SendingType = 1
    End If

    stxtbuffer = vbNullString
    SendTxt.Text = vbNullString
    SendTxt.Visible = False
    
End Sub

Private Sub AbrirMenuViewPort()
    'TODO: No usar variable de compilacion y acceder a esto desde el config.ini
    #If (ConMenuseConextuales = 1) Then

        If tX >= MinXBorder And tY >= MinYBorder And tY <= MaxYBorder And tX <= MaxXBorder Then

            If MapData(tX, tY).CharIndex > 0 Then
                If charlist(MapData(tX, tY).CharIndex).invisible = False Then
        
                    Dim m As frmMenuseFashion
                    Set m = New frmMenuseFashion
            
                    Load m
                    m.SetCallback Me
                    m.SetMenuId 1
                    m.ListaInit 2, False
            
                    If LenB(charlist(MapData(tX, tY).CharIndex).Nombre) <> 0 Then
                        m.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
                    Else
                        m.ListaSetItem 0, "<NPC>", True
                    End If
                    m.ListaSetItem 1, JsonLanguage.item("COMERCIAR").item("TEXTO")
            
                    m.ListaFin
                    m.Show , Me

                End If
            End If
        End If

    #End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)

    Select Case MenuId

        Case 0 'Inventario

            Select Case Sel

                Case 0

                Case 1

                Case 2 'Tirar
                    Call TirarItem

                Case 3 'Usar

                    If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
                        Call UsarItem
                    End If

                Case 3 'equipar
                    Call EquiparItem
            End Select
    
        Case 1 'Menu del ViewPort del engine

            Select Case Sel

                Case 0 'Nombre
                    Call WriteLeftClick(tX, tY)
        
                Case 1 'Comerciar
                    Call WriteLeftClick(tX, tY)
                    Call WriteCommerceStart
            End Select
    End Select
End Sub
 
''''''''''''''''''''''''''''''''''''''
'     WINDOWS API                    '
''''''''''''''''''''''''''''''''''''''
Private Sub Client_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)
    
    Second.Enabled = True
    
    'Actualizams la hora
    Call ActualizarHora
    
    Select Case EstadoLogin

        Case E_MODO.CrearNuevoPJ, E_MODO.Normal
            Call Login

        Case E_MODO.Dados
            Call MostrarCreacion
            
    End Select
 
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
    Dim RD     As String
    Dim Data() As Byte
    
    Client.GetData RD, vbByte, bytesTotal
    Data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
    
End Sub

Private Sub Client_CloseSck()
    
    Debug.Print "Cerrando la conexion via API de Windows..."

    If frmMain.Visible = True Then frmMain.Visible = False
    Call ResetAllInfo
    ModConectar.Conectando = True
    Call MostrarConnect(True)
End Sub

Private Sub Client_Error(ByVal number As Integer, _
                         Description As String, _
                         ByVal sCode As Long, _
                         ByVal source As String, _
                         ByVal HelpFile As String, _
                         ByVal HelpContext As Long, _
                         CancelDisplay As Boolean)
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    
    frmConnect.MousePointer = 1
    
    Second.Enabled = False
 
    If Client.State <> sckClosed Then Client.CloseSck
    
    ModConectar.Conectando = True
    Call MostrarConnect
 
End Sub

Private Function InGameArea() As Boolean
'********************************************************************
'Author: NicoNZ
'Last Modification: 29/09/2019
'Checks if last click was performed within or outside the game area.
'********************************************************************
    If clicX < 0 Or clicX > frmMain.MainViewPic.Width Then Exit Function
    If clicY < 0 Or clicY > frmMain.MainViewPic.Height Then Exit Function
    
    InGameArea = True
End Function

Private Sub hlst_Click()
    
    With hlst
    
        If ChangeHechi Then
    
            Dim NewLugar As Integer: NewLugar = .ListIndex
            Dim AntLugar As String: AntLugar = .List(NewLugar)
            
            Call WriteDragAndDropHechizos(ChangeHechiNum + 1, NewLugar + 1)
        
            .BackColor = vbBlack
            .List(NewLugar) = .List(ChangeHechiNum)
            .List(ChangeHechiNum) = AntLugar
        
            ChangeHechi = False
            ChangeHechiNum = 0

        End If

        .BackColor = vbBlack

    End With

End Sub

Private Sub hlst_DblClick()
    ChangeHechi = True
    ChangeHechiNum = hlst.ListIndex
    hlst.BackColor = vbRed

End Sub

Private Sub Minimapa_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)

    If Button = vbRightButton Then
        Call WriteWarpChar("YO", CurrentUser.UserMap, CByte(x - 1), CByte(y - 1), False)
        Call ActualizarMiniMapa
        
    ElseIf Button = vbLeftButton Then
        frmMapa.Show vbModeless, Me
        
    End If
    
End Sub
    'fin Incorporado ReyarB

Public Sub UpdateProgressExperienceLevelBar(ByVal UserExp As Long)

    If CurrentUser.UserLvl = STAT_MAXELV Then
        frmMain.lblPorcLvl.Caption = "¡Nivel Máximo!"

        'Si no tiene mas niveles que subir ponemos la barra al maximo.
        frmMain.uAOProgressExperienceLevel.Max = 100
        frmMain.uAOProgressExperienceLevel.value = 100
    Else
        frmMain.lblPorcLvl.Caption = Round(CDbl(UserExp) * CDbl(100) / CDbl(CurrentUser.UserPasarNivel), 2) & "%"
        frmMain.uAOProgressExperienceLevel.Max = CurrentUser.UserPasarNivel
        frmMain.uAOProgressExperienceLevel.value = UserExp
    End If
End Sub

Public Sub SetGoldColor()

    If CurrentUser.UserGLD >= CLng(CurrentUser.UserLvl) * 10000 And CurrentUser.UserLvl > 12 Then 'Si el nivel es mayor de 12, es decir, no es newbie.
        'Changes color
        frmMain.GldLbl.ForeColor = USER_GOLD_COLOR
    Else
        'Changes color
        frmMain.GldLbl.ForeColor = NEWBIE_USER_GOLD_COLOR
    End If

End Sub

Private Sub btnRetos_Click()
    Call FrmRetos.Show(vbModeless, frmMain)
End Sub

Private Sub btnMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If btnMenu.Tag = 1 Then
        btnMenu.Picture = General_Load_Picture_From_Resource("24.gif", True)
        btnMenu.Tag = 0
    End If

End Sub

Private Sub btnShop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If btnShop.Tag = 1 Then
        btnShop.Picture = General_Load_Picture_From_Resource("27.gif", True)
        btnShop.Tag = 0
    End If

End Sub

Public Sub ActualizarHora()
'**********************************
'Autor: Lorwik
'Fecha: 11/08/2020
'Descripcion: Actualiza la hora del lbl del frmmain
'**********************************

    If ReadField(1, lblHour.Caption, Asc(":")) <> Minute(Now) Then
        lblHour.Caption = Hour(Now) & ":" & Minute(Now)
    End If
        
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
    
    Select Case Index
    
        Case eSMType.sCombatMode
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("COLOR").item(3), _
                                     True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO")
                
                picSM(Index).Picture = General_Load_Picture_From_Resource("221.gif")
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("TEXTO"), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_COMBAT_OFF").item("COLOR").item(3), _
                                     True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_COMBAT_ON").item("TEXTO")
                
                picSM(Index).Picture = General_Load_Picture_From_Resource("223.gif")
            End If
            
        Case eSMType.sSafemode
            
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, UCase$(JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO").item(1)), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("COLOR").item(3), _
                                      True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_ACTIVADO").item("TEXTO").item(2)
                
                picSM(Index).Picture = General_Load_Picture_From_Resource("222.gif")
                
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, UCase$(JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO").item(1)), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(1), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(2), _
                                                      JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("COLOR").item(3), _
                                      True, False, True)
                                        
                picSM(Index).ToolTipText = JsonLanguage.item("MENSAJE_SEGURO_DESACTIVADO").item("TEXTO").item(2)
                
                picSM(Index).Picture = General_Load_Picture_From_Resource("224.gif")
            End If
        
    End Select
    
End Sub

Public Sub ActualizarCoordenadas(ByVal tX As Integer, ByVal tY As Integer)
'*****************************************************
'Autor: Lorwik
'Fecha: 03/04/2021
'Descripción: Actualiza las coordenadas ya sean totales o por cuadrantes
'*****************************************************

    Dim cx As Integer
    Dim cy As Integer
    Dim AnchoMap As Byte
    Dim CurrentCuadrante As Integer
    
    'Guardamos el cuadrante antes del posible cambio
    CurrentCuadrante = CurrentUser.UserCuadrante
    AnchoMap = 10
    
    cx = Fix((tX / 100))
    cy = Fix((tY / 100))
    
    CurrentUser.UserCuadrante = ((cy) * AnchoMap) + cx + 1
    UserPosCuadrante.x = tX - (cx * 100)
    UserPosCuadrante.y = tY - (cy * 100)
    
    'Si cambiamos de cuadrante cambiamos el minimapa
    If CurrentCuadrante <> CurrentUser.UserCuadrante Then _
        Call DibujarMinimapa
    
    If ClientSetup.VerCuadrantes Then

        Coord.Caption = "Cuadrante: " & CurrentUser.UserCuadrante & " X: " & UserPosCuadrante.x & " Y: " & UserPosCuadrante.y
    
    Else
        Coord.Caption = "Map:" & CurrentUser.UserMap & " X:" & tX & " Y:" & tY
        
    End If
End Sub
