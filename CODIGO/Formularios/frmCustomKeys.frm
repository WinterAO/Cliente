VERSION 5.00
Begin VB.Form frmCustomKeys 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Controles"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10185
   ClipControls    =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   679
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton imgGuardar 
      Caption         =   "Guardar"
      Height          =   480
      Left            =   600
      TabIndex        =   63
      Top             =   4440
      Width           =   2130
   End
   Begin VB.CommandButton imgDefaultKeys 
      Caption         =   "Teclas por Defecto"
      Height          =   480
      Left            =   600
      TabIndex        =   62
      Top             =   3720
      Width           =   2130
   End
   Begin VB.Frame FraMiscelanea 
      Caption         =   "Miscelanea"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   3360
      TabIndex        =   41
      Top             =   0
      Width           =   3495
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   13
         Left            =   1680
         TabIndex        =   58
         Top             =   3240
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   14
         Left            =   1680
         TabIndex        =   57
         Top             =   3600
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   22
         Left            =   1680
         TabIndex        =   56
         Top             =   2880
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   21
         Left            =   1680
         TabIndex        =   54
         Top             =   2520
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   1680
         TabIndex        =   53
         Top             =   2160
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   46
         Top             =   360
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   1680
         TabIndex        =   45
         Top             =   1080
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   1680
         TabIndex        =   44
         Top             =   1440
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   43
         Top             =   720
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   23
         Left            =   1680
         TabIndex        =   42
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label lblSeguroDe 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seguro de Resu..."
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   61
         Top             =   3600
         Width           =   1320
      End
      Begin VB.Label lblSalir 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modo Seguro"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   60
         Top             =   3240
         Width           =   945
      End
      Begin VB.Label lblSalir 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salir"
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   59
         Top             =   2880
         Width           =   525
      End
      Begin VB.Label lblMostrarOpciones 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar Opciones"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   55
         Top             =   2520
         Width           =   1365
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Captura de Pantalla"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   52
         Top             =   2160
         Width           =   1425
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar FPS"
         Height          =   195
         Index           =   10
         Left            =   720
         TabIndex        =   51
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido"
         Height          =   195
         Index           =   9
         Left            =   1080
         TabIndex        =   50
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres"
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   49
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Corregir posición"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   48
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Música"
         Height          =   195
         Index           =   6
         Left            =   1080
         TabIndex        =   47
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.Frame FraAcciones 
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3135
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   960
         TabIndex        =   30
         Top             =   2880
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   17
         Left            =   960
         TabIndex        =   29
         Top             =   2520
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   960
         TabIndex        =   28
         Top             =   2160
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   15
         Left            =   960
         TabIndex        =   27
         Top             =   1800
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   12
         Left            =   960
         TabIndex        =   26
         Top             =   1440
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   960
         TabIndex        =   25
         Top             =   1080
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   960
         TabIndex        =   24
         Top             =   720
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   960
         TabIndex        =   23
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lblAtacar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atacar"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   38
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label lblRobar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Robar"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   37
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblUsar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usar"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   36
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label lblTirar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tirar"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   35
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblOcultar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ocultar"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblDomar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domar"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblEquipar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Equipar"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblAgarrar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agarrar"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame FraMovimiento 
      Caption         =   "Movimiento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6960
      TabIndex        =   13
      Top             =   3240
      Width           =   3135
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   960
         TabIndex        =   17
         Top             =   1440
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   16
         Top             =   1080
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lblDerecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Derecha"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblIzquierda 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Izquierda"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblAbajo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abajo"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblArriba 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arriba"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modo de habla"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   6960
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   840
         TabIndex        =   39
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   25
         Left            =   840
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   840
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   840
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   840
         TabIndex        =   2
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   840
         TabIndex        =   1
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblHablar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hablar"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   40
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   870
         Width           =   495
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gritar"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   1215
         Width           =   405
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Privado"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1590
         Width           =   540
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clan"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   1950
         Width           =   315
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo"
         Height          =   195
         Index           =   4
         Left            =   345
         TabIndex        =   8
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Global"
         Height          =   195
         Index           =   5
         Left            =   345
         TabIndex        =   7
         Top             =   2670
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmCustomKeys"
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

''
'frmCustomKeys - Allows the user to customize keys.
'Implements class clsCustomKeys
'
'@author Rapsodius
'@date 20070805
'@version 1.0.0
'@see clsCustomKeys

Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    Dim i As Long
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadTextsForm

    For i = 1 To CustomKeys.KeyCount
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
End Sub

Private Sub LoadTextsForm()
    imgDefaultKeys.Caption = JsonLanguage.item("FRM_CUSTOMKEYS_DEFAULTKEYS").item("TEXTO")
    imgGuardar.Caption = JsonLanguage.item("FRM_CUSTOMKEYS_GUARDAR").item("TEXTO")
End Sub

Private Sub imgDefaultKeys_Click()
   Unload Me
   frmKeysConfigurationSelect.Visible = True
End Sub

Private Sub imgGuardar_Click()

    Dim i As Long
    Dim sMsg As String
    
    For i = 1 To CustomKeys.KeyCount
        If LenB(Text1(i).Text) = 0 Then
            Call MsgBox(JsonLanguage.item("CUSTOMKEYS_TECLA_INVALIDA").item("TEXTO"), vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Winter AO Resurrection")
            Exit Sub
        End If
    Next i
    
    Unload Me
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long

    If LenB(CustomKeys.ReadableName(KeyCode)) = 0 Then Exit Sub
    'If key is not valid, we exit
    Debug.Print "3"
    Text1(Index).Text = CustomKeys.ReadableName(KeyCode)
    Text1(Index).SelStart = Len(Text1(Index).Text)

    For i = 1 To CustomKeys.KeyCount
        If i <> Index Then
            If CustomKeys.BindedKey(i) = KeyCode Then
                Text1(Index).Text = vbNullString 'If the key is already assigned, simply reject it
                Call Beep 'Alert the user
                KeyCode = 0
                
                Exit Sub
            End If
        End If
    Next i
    
    CustomKeys.BindedKey(Index) = KeyCode
    
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Text1_KeyDown(Index, KeyCode, Shift)
End Sub

Private Sub ShowConfig()

    Dim i As Long

    For i = 1 To CustomKeys.KeyCount
        Text1(i).Text = CustomKeys.ReadableName(CustomKeys.BindedKey(i))
    Next i
    
End Sub
