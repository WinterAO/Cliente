VERSION 5.00
Begin VB.Form frmPanelGm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9465
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
   ScaleHeight     =   7335
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   60
      TabIndex        =   26
      Top             =   5250
      Width           =   5205
      Begin VB.CommandButton cmdIr 
         BackColor       =   &H8000000A&
         Caption         =   "Ira"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   660
         Width           =   1440
      End
      Begin VB.CommandButton cmdInfo 
         BackColor       =   &H8000000A&
         Caption         =   "Info"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1500
         Width           =   1440
      End
      Begin VB.CommandButton cmdSummon 
         BackColor       =   &H8000000A&
         Caption         =   "Sum"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1080
         Width           =   1440
      End
      Begin VB.CommandButton cmdIrCerca 
         BackColor       =   &H8000000A&
         Caption         =   "Ir Cerca"
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   30
      TabIndex        =   20
      Top             =   960
      Width           =   5235
      Begin VB.CommandButton cmdEliminarSeguimiento 
         BackColor       =   &H8000000A&
         Caption         =   "Eliminar Seguimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3930
         Width           =   5025
      End
      Begin VB.TextBox txtNuevoUsuario 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   4995
      End
      Begin VB.CommandButton cmdAddFollow 
         BackColor       =   &H8000000A&
         Caption         =   "Agregar Seguimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3600
         Width           =   5025
      End
      Begin VB.TextBox txtNuevaDescrip 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   120
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   690
         Width           =   4995
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   660
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdActualiza 
      BackColor       =   &H8000000A&
      Caption         =   "Actualizar"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   60
      Width           =   5085
   End
   Begin VB.ComboBox cboListaUsus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   540
      Width           =   5055
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Index           =   0
      Left            =   5310
      TabIndex        =   1
      Top             =   30
      Width           =   4065
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H8000000A&
         Caption         =   "ACTUALIZAR"
         Height          =   495
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2190
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddObs 
         BackColor       =   &H8000000A&
         Caption         =   "Agregar Observacion"
         Height          =   375
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4980
         Width           =   3735
      End
      Begin VB.TextBox txtObs 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   3870
         Width           =   3765
      End
      Begin VB.TextBox txtDescrip 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2970
         Width           =   3735
      End
      Begin VB.TextBox txtCreador 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1710
         Width           =   1695
      End
      Begin VB.TextBox txtTimeOn 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1170
         Width           =   1695
      End
      Begin VB.TextBox txtIP 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   1695
      End
      Begin VB.ListBox lstUsers 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   180
         TabIndex        =   2
         Top             =   330
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Height          =   195
         Left            =   2220
         TabIndex        =   16
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   3690
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
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
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   2790
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Creador"
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
         Height          =   375
         Left            =   2220
         TabIndex        =   9
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Logueado Hace:"
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
         Height          =   375
         Left            =   2220
         TabIndex        =   7
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "IP:"
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
         Left            =   2220
         TabIndex        =   5
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   2940
         TabIndex        =   4
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios Marcados"
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
         Left            =   150
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H8000000A&
      Caption         =   "Cerrar"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6780
      Width           =   3465
   End
   Begin VB.Menu mnuWorld 
      Caption         =   "World"
      Begin VB.Menu cmdLLUVIA 
         Caption         =   "Meteo"
      End
      Begin VB.Menu cmdLIMPIAR 
         Caption         =   "Limpiar"
      End
      Begin VB.Menu cmdCC 
         Caption         =   "CC"
      End
      Begin VB.Menu cmdCT 
         Caption         =   "CT"
      End
      Begin VB.Menu cmdCI 
         Caption         =   "CI"
      End
      Begin VB.Menu cmdPISO 
         Caption         =   "PISO"
      End
      Begin VB.Menu cmdDE 
         Caption         =   "DE"
      End
      Begin VB.Menu cmdDT 
         Caption         =   "DT"
      End
      Begin VB.Menu cmdDEST 
         Caption         =   "DEST"
      End
      Begin VB.Menu cmdMASSDEST 
         Caption         =   "MASSDEST"
      End
   End
   Begin VB.Menu mnuMessage 
      Caption         =   "Message"
      Begin VB.Menu cmdTOGGLEGLOBAL 
         Caption         =   "Activar/Desactivar Global"
      End
      Begin VB.Menu cmdHORA 
         Caption         =   "Hora"
      End
      Begin VB.Menu cmdMOTDCAMBIA 
         Caption         =   "Cambiar Motd"
      End
      Begin VB.Menu cmdTALKAS 
         Caption         =   "Talkas"
      End
      Begin VB.Menu cmdGMSG 
         Caption         =   "GMSG"
      End
      Begin VB.Menu cmdRMSG 
         Caption         =   "RMSG"
      End
      Begin VB.Menu cmdSMSG 
         Caption         =   "SMSG"
      End
      Begin VB.Menu cmdREALMSG 
         Caption         =   "REALMSG"
      End
      Begin VB.Menu cmdCAOSMSG 
         Caption         =   "CAOSMSG"
      End
      Begin VB.Menu cmdCIUMSG 
         Caption         =   "CIUMSG"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin"
      Begin VB.Menu cmdIP2NICK 
         Caption         =   "IP2NICK"
      End
      Begin VB.Menu cmdNICK2IP 
         Caption         =   "NICK2IP"
      End
      Begin VB.Menu mnuban 
         Caption         =   "Banear"
         Begin VB.Menu cmdBAN 
            Caption         =   "...Banear PJ"
         End
         Begin VB.Menu cmdUNBAN 
            Caption         =   "...Desbanear PJ"
         End
         Begin VB.Menu cmdBANIP 
            Caption         =   "...Banear Ip"
         End
         Begin VB.Menu cmdUNBANIP 
            Caption         =   "...Desbanear IP"
         End
         Begin VB.Menu mnuBanSerial 
            Caption         =   "...Banear Serial + Mac"
         End
         Begin VB.Menu mnudesbanserial 
            Caption         =   "...Desbanear Serial + Mac"
         End
         Begin VB.Menu mnubantemp 
            Caption         =   "...Baneo temporal"
         End
         Begin VB.Menu mnubanlife 
            Caption         =   "...Banear de la vida"
         End
      End
      Begin VB.Menu cmdLASTEMAIL 
         Caption         =   "LASTEMAIL"
      End
      Begin VB.Menu cmdBANIPLIST 
         Caption         =   "BANIPLIST"
      End
      Begin VB.Menu cmdBANIPRELOAD 
         Caption         =   "BANIPRELOAD"
      End
      Begin VB.Menu cmdLASTIP 
         Caption         =   "LASTIP"
      End
      Begin VB.Menu cmdSHOWCMSG 
         Caption         =   "SHOWCMSG"
      End
      Begin VB.Menu MIEMBROSCLAN 
         Caption         =   "MIEMBROSCLAN"
      End
      Begin VB.Menu cmdBANCLAN 
         Caption         =   "BANCLAN"
      End
      Begin VB.Menu cmdADVERTENCIA 
         Caption         =   "ADVERTENCIA"
      End
      Begin VB.Menu cmdCARCEL 
         Caption         =   "CARCEL"
      End
      Begin VB.Menu cmdBORRARPENA 
         Caption         =   "BORRARPENA"
      End
      Begin VB.Menu cmdSilenciar 
         Caption         =   "SILENCIAR"
      End
   End
   Begin VB.Menu mnuMundo 
      Caption         =   "Mundo"
      Begin VB.Menu cmdSHOW_SOS 
         Caption         =   "SHOW SOS"
      End
      Begin VB.Menu cmdBORRAR_SOS 
         Caption         =   "BORRAR SOS"
      End
      Begin VB.Menu cmdTRABAJANDO 
         Caption         =   "TRABAJANDO"
      End
      Begin VB.Menu cmdOCULTANDO 
         Caption         =   "OCULTANDO"
      End
      Begin VB.Menu cmdNENE 
         Caption         =   "NENE"
      End
      Begin VB.Menu cmdONLINEMAP 
         Caption         =   "ONLINEMAP"
      End
      Begin VB.Menu cmdONLINEREAL 
         Caption         =   "ONLINEREAL"
      End
      Begin VB.Menu cmdONLINECAOS 
         Caption         =   "ONLINECAOS"
      End
      Begin VB.Menu cmdONLINEGM 
         Caption         =   "ONLINEGM"
      End
   End
   Begin VB.Menu mnuMe 
      Caption         =   "Me"
      Begin VB.Menu cmdINVISIBLE 
         Caption         =   "INVISIBLE"
      End
      Begin VB.Menu cmdIGNORADO 
         Caption         =   "IGNORADO"
      End
      Begin VB.Menu cmdNAVE 
         Caption         =   "NAVE"
      End
      Begin VB.Menu cmdCHATCOLOR 
         Caption         =   "CHATCOLOR"
      End
      Begin VB.Menu cmdREM 
         Caption         =   "REM"
      End
      Begin VB.Menu cmdSHOWNAME 
         Caption         =   "SHOWNAME"
      End
      Begin VB.Menu cmdSETDESC 
         Caption         =   "SETDESC"
      End
   End
   Begin VB.Menu mnuJugador 
      Caption         =   "Jugador"
      Begin VB.Menu cmdTELEP 
         Caption         =   "TELEP"
      End
      Begin VB.Menu cmdDONDE 
         Caption         =   "DONDE"
      End
      Begin VB.Menu cmdConsulta 
         Caption         =   "Consulta"
      End
      Begin VB.Menu cmdSTAT 
         Caption         =   "STAT"
      End
      Begin VB.Menu cmdBAL 
         Caption         =   "BAL"
      End
      Begin VB.Menu cmdINV 
         Caption         =   "INV"
      End
      Begin VB.Menu cmdBOV 
         Caption         =   "BOV"
      End
      Begin VB.Menu cmdSKILLS 
         Caption         =   "SKILLS"
      End
      Begin VB.Menu cmdPENAS 
         Caption         =   "PENAS"
      End
      Begin VB.Menu cmdECHAR 
         Caption         =   "ECHAR"
      End
      Begin VB.Menu cmdRAJARCLAN 
         Caption         =   "RAJARCLAN"
      End
      Begin VB.Menu cmdREVIVIR 
         Caption         =   "REVIVIR"
      End
      Begin VB.Menu cmdEJECUTAR 
         Caption         =   "EJECUTAR"
      End
      Begin VB.Menu cmdCONDEN 
         Caption         =   "CONDEN"
      End
      Begin VB.Menu cmdPERDON 
         Caption         =   "PERDON"
      End
      Begin VB.Menu cmdRAJAR 
         Caption         =   "RAJAR"
      End
      Begin VB.Menu cmdESTUPIDO 
         Caption         =   "ESTUPIDO"
      End
      Begin VB.Menu cmdNOESTUPIDO 
         Caption         =   "NOESTUPIDO"
      End
      Begin VB.Menu cmdNOCAOS 
         Caption         =   "NOCAOS"
      End
      Begin VB.Menu cmdNOREAL 
         Caption         =   "NOREAL"
      End
      Begin VB.Menu cmdACEPTCONSE 
         Caption         =   "ACEPTCONSE"
      End
      Begin VB.Menu cmdACEPTCONSECAOS 
         Caption         =   "ACEPTCONSECAOS"
      End
      Begin VB.Menu cmdKICKCONSE 
         Caption         =   "KICKCONSE"
      End
      Begin VB.Menu cmdSilenciarGlobal 
         Caption         =   "SILENCIARGLOBAL"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "Debug"
      Begin VB.Menu mnuDebugAreas 
         Caption         =   "Areas"
      End
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmPanelGm.frm
'
'**************************************************************

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

''
' IMPORTANT!!!
' To prevent the combo list of usernames from closing when a conole message arrives, the Validate event allways
' sets the Cancel arg to True. This, combined with setting the CausesValidation of the RichTextBox to True
' makes the trick. However, in order to be able to use other commands, ALL OTHER controls in this form must have the
' CuasesValidation parameter set to false (unless you want to code your custom flag system to know when to allow or not the loose of focus).

Private Sub cboListaUsus_Validate(Cancel As Boolean)
    Cancel = True
End Sub

Private Sub cmdACEPTCONSE_Click()
    '/ACEPTCONSE
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea aceptar a " & Nick & " como consejero real?", vbYesNo, "Atencion!") = vbYes Then Call WriteAcceptRoyalCouncilMember(Nick)
End Sub

Private Sub cmdACEPTCONSECAOS_Click()
    '/ACEPTCONSECAOS
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea aceptar a " & Nick & " como consejero del caos?", vbYesNo, "Atencion!") = vbYes Then Call WriteAcceptChaosCouncilMember(Nick)
End Sub

Private Sub cmdAddFollow_Click()
    Dim i As Long

    For i = 0 To lstUsers.ListCount

        If UCase$(lstUsers.List(i)) = UCase$(txtNuevoUsuario.Text) Then
            Call MsgBox("El usuario ya esta en la lista!", vbOKOnly + vbExclamation)
            Exit Sub
        End If
    Next i
            
    If LenB(txtNuevoUsuario.Text) = 0 Then
        Call MsgBox("Escribe el nombre de un usuario!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    
    If LenB(txtNuevaDescrip.Text) = 0 Then
        Call MsgBox("Escribe el motivo del seguimiento!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    
    Call WriteRecordAdd(txtNuevoUsuario.Text, txtNuevaDescrip.Text)
    
    txtNuevoUsuario.Text = vbNullString
    txtNuevaDescrip.Text = vbNullString
End Sub

Private Sub cmdAddObs_Click()
    Dim Obs As String
    
    Obs = InputBox("Ingrese la observacion", "Nueva Observacion")
    
    If LenB(Obs) = 0 Then
        Call MsgBox("Escribe una observacion!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    
    If lstUsers.ListIndex = -1 Then
        Call MsgBox("Seleccione un seguimiento!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    
    Call WriteRecordAddObs(lstUsers.ListIndex + 1, Obs)
End Sub

Private Sub cmdADVERTENCIA_Click()
    '/ADVERTENCIA
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
        
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & Nick)
                
        If LenB(tStr) <> 0 Then
            'We use the Parser to control the command format
            Call ParseUserCommand("/ADVERTENCIA " & Nick & "@" & tStr)
        End If
    End If
End Sub

Private Sub cmdBAL_Click()
    '/BAL
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharGold(Nick)
End Sub

Private Sub cmdBAN_Click()
    '/BAN
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo del ban.", "BAN a " & Nick)
                
        If LenB(tStr) <> 0 Then If MsgBox("Seguro desea banear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteBanChar(Nick, tStr)
    End If
End Sub

Private Sub cmdBANCLAN_Click()
    '/BANCLAN
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan.", "Banear clan")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea banear al clan " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteGuildBan(tStr)
End Sub

Private Sub cmdBANIP_Click()
    '/BANIP
    Dim tStr   As String
    Dim reason As String
    
    tStr = InputBox("Escriba el ip o el nick del PJ.", "Banear IP")
    
    reason = InputBox("Escriba el motivo del ban.", "Banear IP")
    
    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea banear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/BANIP " & tStr & " " & reason) 'We use the Parser to control the command format
End Sub

Private Sub cmdBANIPLIST_Click()
    '/BANIPLIST
    Call WriteBannedIPList
End Sub

Private Sub cmdBANIPRELOAD_Click()
    '/BANIPRELOAD
    Call WriteBannedIPReload
End Sub

Private Sub cmdBORRAR_SOS_Click()

    '/BORRAR SOS
    If MsgBox("Seguro desea borrar el SOS?", vbYesNo, "Atencion!") = vbYes Then Call WriteCleanSOS
End Sub

Private Sub cmdBORRARPENA_Click()
    '/BORRARPENA
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Indique el numero de la pena a borrar.", "Borrar pena")

        If LenB(tStr) <> 0 Then If MsgBox("Seguro desea borrar la pena " & tStr & " a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/BORRARPENA " & Nick & "@" & tStr) 'We use the Parser to control the command format
    End If
End Sub

Private Sub cmdBOV_Click()
    '/BOV
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharBank(Nick)
End Sub

Private Sub cmdCAOSMSG_Click()
    '/CAOSMSG
    Dim tStr As String
    
    tStr = InputBox(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"), "Mensaje por consola LegionOscura")

    If LenB(tStr) <> 0 Then Call WriteChaosLegionMessage(tStr)
End Sub

Private Sub cmdCARCEL_Click()
    '/CARCEL
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la pena.", "Carcel a " & Nick)
                
        If LenB(tStr) <> 0 Then
            tStr = tStr & "@" & InputBox("Indique el tiempo de condena (entre 0 y 60 minutos).", "Carcel a " & Nick)
            'We use the Parser to control the command format
            Call ParseUserCommand("/CARCEL " & Nick & "@" & tStr)
        End If
    End If
End Sub

Private Sub cmdCC_Click()
    '/CC
    Call WriteSpawnListRequest
End Sub

Private Sub cmdCHATCOLOR_Click()
    '/CHATCOLOR
    Dim tStr As String
    
    tStr = InputBox("Defina el color (R G B). Deje en blanco para usar el default.", "Cambiar color del chat")
    
    Call ParseUserCommand("/CHATCOLOR " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdCI_Click()
    '/CI
    Dim tStr As String
    
    tStr = InputBox("Indique el numero del objeto a crear.", "Crear Objeto")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea crear el objeto " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/CI " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdCIUMSG_Click()
    '/CIUMSG
    Dim tStr As String
    
    tStr = InputBox(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"), "Mensaje por consola Ciudadanos")

    If LenB(tStr) <> 0 Then Call WriteCitizenMessage(tStr)
End Sub

Private Sub cmdCONDEN_Click()
    '/CONDEN
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea volver criminal a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteTurnCriminal(Nick)
End Sub

Private Sub cmdConsulta_Click()
    WriteConsultation
End Sub

Private Sub cmdCT_Click()
    '/CT
    Dim tStr As String
    
    tStr = InputBox("Indique la posicion donde lleva el portal (MAPA X Y).", "Crear Portal")

    If LenB(tStr) <> 0 Then Call ParseUserCommand("/CT " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdDE_Click()

    '/DE
    If MsgBox("Seguro desea destruir el Tile Exit?", vbYesNo, "Atencion!") = vbYes Then Call WriteExitDestroy
End Sub

Private Sub cmdDEST_Click()

    '/DEST
    If MsgBox("Seguro desea destruir el objeto sobre el que esta parado?", vbYesNo, "Atencion!") = vbYes Then Call WriteDestroyItems
End Sub

Private Sub cmdDONDE_Click()
    '/DONDE
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteWhere(Nick)
End Sub

Private Sub cmdDT_Click()

    'DT
    If MsgBox("Seguro desea destruir el portal?", vbYesNo, "Atencion!") = vbYes Then Call WriteTeleportDestroy
End Sub

Private Sub cmdECHAR_Click()
    '/ECHAR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteKick(Nick)
End Sub

Private Sub cmdEJECUTAR_Click()
    '/EJECUTAR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea ejecutar a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteExecute(Nick)
End Sub

Private Sub cmdEliminarSeguimiento_Click()

    With lstUsers

        If .ListIndex = -1 Then
            Call MsgBox("Seleccione un usuario para remover el seguimiento!", vbOKOnly + vbExclamation)
            Exit Sub
        End If
        
        If MsgBox("Desea eliminar el seguimiento al personaje " & .List(.ListIndex) & "?", vbYesNo) = vbYes Then
            Call WriteRecordRemove(.ListIndex + 1)
            Call ClearRecordDetails
        End If
    End With
End Sub

Private Sub cmdESTUPIDO_Click()
    '/ESTUPIDO
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteMakeDumb(Nick)
End Sub

Private Sub cmdGMSG_Click()
    '/GMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")

    If LenB(tStr) <> 0 Then Call WriteGMMessage(tStr)
End Sub

Private Sub cmdHORA_Click()
    '/HORA
    Call Protocol_Write.WriteServerTime
End Sub

Private Sub cmdIGNORADO_Click()
    '/IGNORADO
    Call WriteIgnored
End Sub

Private Sub cmdINV_Click()
    '/INV
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharInventory(Nick)
End Sub

Private Sub cmdINVISIBLE_Click()
    '/INVISIBLE
    Call WriteInvisible
End Sub

Private Sub cmdIP2NICK_Click()
    '/IP2NICK
    Dim tStr As String
    
    tStr = InputBox("Escriba la ip.", "IP to Nick")

    If LenB(tStr) <> 0 Then Call ParseUserCommand("/IP2NICK " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdIr_Click()

    '/IR
    With lstUsers

        If .ListIndex <> -1 Then
            Call WriteGoToChar(.List(.ListIndex))
        End If
    End With
    
End Sub

Private Sub cmdInfo_Click()

    '/INFO
    With lstUsers

        If .ListIndex <> -1 Then
            Call WriteRequestCharInfo(.List(.ListIndex))
        End If
    End With
    
End Sub

Private Sub cmdIrCerca_Click()

    '/IRCERCA
    With lstUsers

        If .ListIndex <> -1 Then
            Call WriteGoNearby(.List(.ListIndex))
        End If
    End With
End Sub

Private Sub cmdKICKCONSE_Click()
    'KICKCONSE
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea destituir a " & Nick & " de su cargo de consejero?", vbYesNo, "Atencion!") = vbYes Then Call WriteCouncilKick(Nick)
End Sub

Private Sub cmdLASTEMAIL_Click()
    '/LASTEMAIL
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharMail(Nick)
End Sub

Private Sub cmdLASTIP_Click()
    '/LASTIP
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteLastIP(Nick)
End Sub

Private Sub cmdLIMPIAR_Click()
    '/LIMPIARMUNDO
    Call WriteLimpiarMundo
End Sub

Private Sub cmdMETEO_Click()
    '/METEO
    Dim tBte As Byte
    
    tBte = InputBox("Escriba el fenomeno, 0: Random, 1: Lluvia, 2: Niebla, 3: Niebla + Lluvia.", "Seleccion meteorologico")
    
    If Val(tBte) < 0 And Val(tBte) > 250 Then Call WriteMeteoToggle(Val(tBte))
End Sub

Private Sub cmdMASSDEST_Click()

    '/MASSDEST
    If MsgBox("Seguro desea destruir todos los items del mapa?", vbYesNo, "Atencion!") = vbYes Then Call WriteDestroyAllItemsInArea
End Sub

Private Sub cmdMIEMBROSCLAN_Click()
    '/MIEMBROSCLAN
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan.", "Lista de miembros del clan")

    If LenB(tStr) <> 0 Then Call WriteGuildMemberList(tStr)
End Sub

Private Sub cmdMOTDCAMBIA_Click()
    '/MOTDCAMBIA
    Call WriteChangeMOTD
End Sub

Private Sub cmdNAVE_Click()
    '/NAVE
    Call WriteNavigateToggle
End Sub

Private Sub cmdNENE_Click()
    '/NENE
    Dim tStr As String
    
    tStr = InputBox("Indique el mapa.", "Numero de NPCs enemigos.")

    If LenB(tStr) <> 0 Then Call ParseUserCommand("/NENE " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub cmdNICK2IP_Click()
    '/NICK2IP
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteNickToIP(Nick)
End Sub

Private Sub cmdNOCAOS_Click()
    '/NOCAOS
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea expulsar a " & Nick & " de la legion oscura?", vbYesNo, "Atencion!") = vbYes Then Call WriteChaosLegionKick(Nick)
End Sub

Private Sub cmdNOESTUPIDO_Click()
    '/NOESTUPIDO
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteMakeDumbNoMore(Nick)
End Sub

Private Sub cmdNOREAL_Click()
    '/NOREAL
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea expulsar a " & Nick & " de la armada real?", vbYesNo, "Atencion!") = vbYes Then Call WriteRoyalArmyKick(Nick)
End Sub

Private Sub cmdOCULTANDO_Click()
    '/OCULTANDO
    Call WriteHiding
End Sub

Private Sub cmdONLINECAOS_Click()
    '/ONLINECAOS
    Call WriteOnlineChaosLegion
End Sub

Private Sub cmdONLINEGM_Click()
    '/ONLINEGM
    Call WriteOnlineGM
End Sub

Private Sub cmdONLINEMAP_Click()
    '/ONLINEMAP
    Call WriteOnlineMap(CurrentUser.UserMap)
End Sub

Private Sub cmdONLINEREAL_Click()
    '/ONLINEREAL
    Call WriteOnlineRoyalArmy
End Sub

Private Sub cmdPENAS_Click()
    '/PENAS
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WritePunishments(Nick)
End Sub

Private Sub cmdPERDON_Click()
    '/PERDON
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteForgive(Nick)
End Sub

Private Sub cmdPISO_Click()
    '/PISO
    Call WriteItemsInTheFloor
End Sub

Private Sub cmdRAJAR_Click()
    '/RAJAR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea resetear la faccion de " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteResetFactions(Nick)
End Sub

Private Sub cmdRAJARCLAN_Click()
    '/RAJARCLAN
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea expulsar a " & Nick & " de su clan?", vbYesNo, "Atencion!") = vbYes Then Call WriteRemoveCharFromGuild(Nick)
End Sub

Private Sub cmdREALMSG_Click()
    '/REALMSG
    Dim tStr As String
    
    tStr = InputBox(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"), "Mensaje por consola ArmadaReal")

    If LenB(tStr) <> 0 Then Call WriteRoyalArmyMessage(tStr)
End Sub

Private Sub cmdRefresh_Click()
    Call ClearRecordDetails
    Call WriteRecordListRequest
End Sub

Private Sub cmdREM_Click()
    '/REM
    Dim tStr As String
    
    tStr = InputBox("Escriba el comentario.", "Comentario en el logGM")

    If LenB(tStr) <> 0 Then Call WriteComment(tStr)
End Sub

Private Sub cmdREVIVIR_Click()
    '/REVIVIR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteReviveChar(Nick)
End Sub

Private Sub cmdRMSG_Click()
    '/RMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de RoleMaster")

    If LenB(tStr) <> 0 Then Call WriteServerMessage(tStr)
End Sub

Private Sub cmdSETDESC_Click()
    '/SETDESC
    Dim tStr As String
    
    tStr = InputBox("Escriba una DESC.", "Set Description")

    If LenB(tStr) <> 0 Then Call WriteSetCharDescription(tStr)
End Sub

Private Sub cmdSHOW_SOS_Click()
    '/SHOW SOS
    Call WriteSOSShowList
End Sub

Private Sub cmdSHOWCMSG_Click()
    '/SHOWCMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan que desea escuchar.", "Escuchar los mensajes del clan")

    If LenB(tStr) <> 0 Then Call WriteShowGuildMessages(tStr)
End Sub

Private Sub cmdSHOWNAME_Click()
    '/SHOWNAME
    Call WriteShowName
End Sub

Private Sub cmdSILENCIAR_Click()
    '/SILENCIAR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteSilence(Nick)
End Sub

Private Sub cmdSilenciarGlobal_Click()
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea silenciar del global a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteSilenciarGlobal(Nick)
End Sub

Private Sub cmdSKILLS_Click()
    '/SKILLS
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharSkills(Nick)
End Sub

Private Sub cmdSMSG_Click()
    '/SMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje de sistema")

    If LenB(tStr) <> 0 Then Call WriteSystemMessage(tStr)
End Sub

Private Sub cmdSTAT_Click()
    '/STAT
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharStats(Nick)
End Sub

Private Sub cmdSummon_Click()
    '/SUM
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteSummonChar(Nick)
End Sub

Private Sub cmdTALKAS_Click()
    '/TALKAS
    Dim tStr As String
    
    tStr = InputBox(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"), "Hablar por NPC")

    If LenB(tStr) <> 0 Then Call WriteTalkAsNPC(tStr)
End Sub

Private Sub cmdTELEP_Click()
    '/TELEP
    Dim tStr As String
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Indique la posicion (MAPA X Y).", "Transportar a " & Nick)

        If LenB(tStr) <> 0 Then Call ParseUserCommand("/TELEP " & Nick & " " & tStr) 'We use the Parser to control the command format
    End If
End Sub

Private Sub cmdTOGGLEGLOBAL_Click()
    Call WriteToggleGlobal
End Sub

Private Sub cmdTRABAJANDO_Click()
    '/TRABAJANDO
    Call WriteWorking
End Sub

Private Sub cmdUNBAN_Click()
    '/UNBAN
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea unbanear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteUnbanChar(Nick)
End Sub

Private Sub cmdUNBANIP_Click()
    '/UNBANIP
    Dim tStr As String
    
    tStr = InputBox("Escriba el ip.", "Unbanear IP")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea unbanear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/UNBANIP " & tStr) 'We use the Parser to control the command format
End Sub

Private Sub Form_Load()
    
    'Actualiza los usuarios online
    Call cmdActualiza_Click
    
    'Actualiza los seguimientos
    Call cmdRefresh_Click
    
End Sub

Private Sub cmdActualiza_Click()
    Call WriteRequestUserList
    Call FlushBuffer
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub lstUsers_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    If lstUsers.ListIndex <> -1 Then
        Call ClearRecordDetails
        Call WriteRecordDetailsRequest(lstUsers.ListIndex + 1)
    End If
End Sub

Private Sub ClearRecordDetails()
    txtIP.Text = vbNullString
    txtCreador.Text = vbNullString
    txtDescrip.Text = vbNullString
    txtObs.Text = vbNullString
    txtTimeOn.Text = vbNullString
    lblEstado.Caption = vbNullString
End Sub

Private Sub mnuBanSerial_Click()
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre de usuario.", "Banear Serial")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea banear el serial de " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteBanSerial(tStr)
End Sub

Private Sub mnubantemp_Click()
    On Error Resume Next
    
    Dim tStr  As String
    Dim razon As String
    Dim dias  As Integer
    
    tStr = InputBox("Escriba el nombre de usuario.", "Banear Temporal")
    razon = InputBox("Escriba el motivo del baneo.", "Banear Temporal")
    dias = InputBox("¿Cuantos dias quieres banear al usuario?", "Banear Temporal")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea banear al usuario " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteBanTemporal(tStr, razon, dias)
End Sub

Private Sub mnuDebugAreas_Click()
    DebugAreas = Not DebugAreas
End Sub

Private Sub mnudesbanserial_Click()
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre de usuario.", "Desbanear Serial")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea desbanear el serial de " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteUnBanSerial(tStr)
End Sub
