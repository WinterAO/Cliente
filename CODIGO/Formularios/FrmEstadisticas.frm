VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Estadisticas"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7005
   ClipControls    =   0   'False
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmEstadisticas.frx":000C
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   21
      Left            =   6000
      Top             =   5670
      Width           =   570
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   20
      Left            =   6000
      Top             =   5385
      Width           =   570
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   21
      Left            =   5640
      TabIndex        =   37
      Top             =   5655
      Width           =   315
   End
   Begin VB.Image imgCerrar 
      Height          =   240
      Left            =   120
      Tag             =   "1"
      Top             =   6480
      Width           =   6810
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   19
      Left            =   5160
      Top             =   3510
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   18
      Left            =   5160
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   17
      Left            =   5160
      Top             =   3930
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   16
      Left            =   6000
      Top             =   5070
      Width           =   570
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   15
      Left            =   4140
      Top             =   5985
      Width           =   570
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   14
      Left            =   6000
      Top             =   5985
      Width           =   570
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   13
      Left            =   4140
      Top             =   5370
      Width           =   570
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   12
      Left            =   4140
      Top             =   5670
      Width           =   570
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   11
      Left            =   5160
      Top             =   3090
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   10
      Left            =   5160
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   9
      Left            =   4140
      Top             =   5070
      Width           =   570
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   8
      Left            =   5160
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   7
      Left            =   5160
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   6
      Left            =   5160
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   5
      Left            =   5160
      Top             =   2250
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   4
      Left            =   5160
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   3
      Left            =   5160
      Top             =   1845
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   2
      Left            =   5160
      Top             =   1650
      Width           =   1095
   End
   Begin VB.Shape shpSkillsBar 
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   120
      Index           =   1
      Left            =   5160
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   5
      Left            =   2325
      TabIndex        =   36
      Top             =   6150
      Width           =   435
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   4
      Left            =   930
      TabIndex        =   35
      Top             =   5910
      Width           =   1785
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   3
      Left            =   1875
      TabIndex        =   34
      Top             =   5640
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   2
      Left            =   1830
      TabIndex        =   33
      Top             =   5400
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   32
      Top             =   5160
      Width           =   585
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Index           =   0
      Left            =   2010
      TabIndex        =   31
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   20
      Left            =   5730
      TabIndex        =   30
      Top             =   5355
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   19
      Left            =   3840
      TabIndex        =   29
      Top             =   3495
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   18
      Left            =   3930
      TabIndex        =   28
      Top             =   3300
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   17
      Left            =   3675
      TabIndex        =   27
      Top             =   3870
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   16
      Left            =   5625
      TabIndex        =   26
      Top             =   5040
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   15
      Left            =   3675
      TabIndex        =   25
      Top             =   5970
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   14
      Left            =   3840
      TabIndex        =   24
      Top             =   5640
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   13
      Left            =   3615
      TabIndex        =   23
      Top             =   5340
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   12
      Left            =   5325
      TabIndex        =   22
      Top             =   5985
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   7
      Left            =   960
      TabIndex        =   21
      Top             =   3885
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   6
      Left            =   960
      TabIndex        =   20
      Top             =   3660
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   5
      Left            =   990
      TabIndex        =   19
      Top             =   3420
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   4
      Left            =   1065
      TabIndex        =   18
      Top             =   3195
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   2
      Left            =   1140
      TabIndex        =   17
      Top             =   2985
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   1
      Left            =   1125
      TabIndex        =   16
      Top             =   2745
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   11
      Left            =   3930
      TabIndex        =   15
      Top             =   3075
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   10
      Left            =   3225
      TabIndex        =   14
      Top             =   3690
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   9
      Left            =   3450
      TabIndex        =   13
      Top             =   5040
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   8
      Left            =   3450
      TabIndex        =   12
      Top             =   2850
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   7
      Left            =   3195
      TabIndex        =   11
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   6
      Left            =   3165
      TabIndex        =   10
      Top             =   2430
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   5
      Left            =   3075
      TabIndex        =   9
      Top             =   2205
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   4
      Left            =   3900
      TabIndex        =   8
      Top             =   2025
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   3
      Left            =   3900
      TabIndex        =   7
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   1620
      Width           =   315
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   1
      Left            =   3045
      TabIndex        =   5
      Top             =   1410
      Width           =   315
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   5
      Left            =   1440
      TabIndex        =   4
      Top             =   1710
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   4
      Left            =   1200
      TabIndex        =   3
      Top             =   1500
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   3
      Left            =   1440
      TabIndex        =   2
      Top             =   1260
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   2
      Left            =   1200
      TabIndex        =   1
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   795
      Width           =   210
   End
End
Attribute VB_Name = "frmEstadisticas"
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

Private cBotonCerrar As clsGraphicalButton
Public LastButtonPressed As clsGraphicalButton

Private Const ANCHO_BARRA As Byte = 73 'pixeles
Private Const BAR_LEFT_POS As Integer = 361 'pixeles

Public Sub Iniciar_Labels()
    'Iniciamos los labels con los valores de los atributos y los skills
    Dim i As Integer
    Dim Ancho As Integer
    
    For i = 1 To NUMATRIBUTOS
        Atri(i).Caption = UserAtributos(i)
    Next
    
    For i = 1 To NUMSKILLS
        Skills(i).Caption = UserSkills(i)
        Ancho = IIf(PorcentajeSkills(i) = 0, ANCHO_BARRA, (100 - PorcentajeSkills(i)) / 100 * ANCHO_BARRA)
        shpSkillsBar(i).Width = Ancho
        shpSkillsBar(i).Left = BAR_LEFT_POS + ANCHO_BARRA - Ancho
    Next
    
    
    Label4(1).Caption = UserReputacion.AsesinoRep
    Label4(2).Caption = UserReputacion.BandidoRep
    'Label4(3).Caption = "Burgues: " & UserReputacion.BurguesRep
    Label4(4).Caption = UserReputacion.LadronesRep
    Label4(5).Caption = UserReputacion.NobleRep
    Label4(6).Caption = UserReputacion.PlebeRep
    
    If UserReputacion.Promedio < 0 Then
        Label4(7).ForeColor = vbRed
        Label4(7).Caption = "Criminal"
    Else
        Label4(7).ForeColor = vbBlue
        Label4(7).Caption = "Ciudadano"
    End If
    
    With UserEstadisticas
        Label6(0).Caption = .CriminalesMatados
        Label6(1).Caption = .CiudadanosMatados
        Label6(2).Caption = .UsuariosMatados
        Label6(3).Caption = .NpcsMatados
        Label6(4).Caption = .Clase
        Label6(5).Caption = .PenaCarcel
    End With

End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    ' TODO: Traducir los textos de las imagenes via labels en visual basic, para que en el futuro si se quiere se pueda traducir a mas idiomas
    ' No ando con mas ganas/tiempo para hacer eso asi que se traducen las imagenes asi tenemos el juego en ingles.
    ' Tambien usar los controles uAObuttons para los botones, usar de ejemplo frmCambiaMotd.frm
    If Language = "spanish" Then
      Me.Picture = LoadPicture(Carga.Path(Interfaces) & "VentanaEstadisticas_spanish.jpg")
    Else
      Me.Picture = LoadPicture(Carga.Path(Interfaces) & "VentanaEstadisticas_english.jpg")
    End If
    
    Call LoadButtons
End Sub

Private Sub LoadButtons()
    
    Dim GrhPath As String
    
    GrhPath = Carga.Path(Interfaces)
    
    Set cBotonCerrar = New clsGraphicalButton
    Set LastButtonPressed = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(imgCerrar, GrhPath & "BotonCerrarEstadisticas.jpg", _
                                    GrhPath & "BotonCerrarRolloverEstadisticas.jpg", _
                                    GrhPath & "BotonCerrarClickEstadisticas.jpg", Me)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If imgCerrar.Tag = 1 Then
        imgCerrar.Picture = LoadPicture(Carga.Path(Interfaces) & "BotonCerrarApretadoEstadisticas.jpg")
        imgCerrar.Tag = 0
    End If

End Sub

