VERSION 5.00
Begin VB.Form frmMonitor 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WinterAO: Monitor de Recursos"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7890
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
   ScaleHeight     =   2895
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDebug 
      Interval        =   10000
      Left            =   7410
      Top             =   60
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   360
      Left            =   2580
      TabIndex        =   0
      Top             =   2430
      Width           =   2670
   End
   Begin VB.Label lblMemLoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MemLoad"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   360
      Width           =   4425
   End
   Begin VB.Label lblTotalPhysMem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TotalPhysMem"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   660
      Width           =   4335
   End
   Begin VB.Label lblAvailPhysMem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AvailPhysMem"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   960
      Width           =   4635
   End
   Begin VB.Label lblWorkingSet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WorkingSet"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1260
      Width           =   4440
   End
   Begin VB.Label lblPagefileUsage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PagefileUsage"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   1560
      Width           =   4440
   End
   Begin VB.Label lblCPUUsage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPUUsage"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   1860
      Width           =   4275
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Configura el Timer para que dispare cada 1000 ms (1 segundos)
    tmrDebug.Interval = 1000
    tmrDebug.Enabled = True
    
    Call InitializeCPUUsage
    
    Call tmrDebug_Timer
End Sub

Private Sub cmdCerrar_Click()
    tmrDebug.Enabled = False
    
    Unload Me
    
End Sub

Private Sub tmrDebug_Timer()
    Debug.Print "******* TIMER DEBUG *******"
    ' Llamar a las funciones de monitoreo cada vez que se dispare el Timer
    GetMemoryStatus
    GetProcessMemoryUsage
    GetCPUUsage
End Sub

