Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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

'Caption de los Forms del proyecto
Public Form_Caption As String

Public MacAdress        As String
Public HDserial         As Long

Public Battlegrounds    As Boolean

'#######################
'HUD
'#######################
Public lblWeapon As String
Public lblArmor As String
Public lblShielder As String
Public lblHelm As String

'#######################
'SERVIDORES
'#######################
Public CurServerIp As String
Public CurServerPort As Integer

Public Type Servidores
    Nombre As String
    Ip As String
    Puerto As Integer
End Type

Public Servidor() As Servidores
Public ServIndSel As Byte

'#######################
'Objetos publicos
'#######################
Public DialogosClanes As clsGuildDlg
Public Dialogos As clsDialogs
Public Sound As clsSoundEngine
Public Inventario As clsGraphicalInventory
Public InvBanco(1) As clsGraphicalInventory

'Inventarios de comercio con usuario
Public InvComUsu As clsGraphicalInventory  ' Inventario del usuario visible en el comercio
Public InvOroComUsu(2) As clsGraphicalInventory  ' Inventarios de oro (ambos usuarios)
Public InvOfferComUsu(1) As clsGraphicalInventory  ' Inventarios de ofertas (ambos usuarios)

Public InvComNpc As clsGraphicalInventory  ' Inventario con los items que ofrece el npc

'Inventarios de herreria
Public Const MAX_LIST_ITEMS As Byte = 4
Public InvMaterialTrabajo(1 To MAX_LIST_ITEMS) As clsGraphicalInventory
Public InvObjArtesano(1 To MAX_LIST_ITEMS) As clsGraphicalInventory

Public Const MAX_ITEMS_CRAFTEO As Byte = 4

Public Const MAXMATERIALES As Byte = 4 '4 materiales maximo para construir un item

Public CustomKeys As clsCustomKeys

Public incomingData As clsByteQueue
Public outgoingData As clsByteQueue

'The main timer of the Carga.
Public MainTimer As clsTimer

Public Inet As clsInet

'#######################
'CODIGO DE ERRORES
'#######################
Public Enum eSockError
   TOO_FAST = 24036
   REFUSED = 24061
   TIME_OUT = 24060
End Enum

'#######################
'CONSTANTES DE SONIDO
'#######################

'Pasos
Public Enum TipoPaso
    CONST_BOSQUE = 1
    CONST_NIEVE = 2
    CONST_CABALLO = 3
    CONST_DUNGEON = 4
    CONST_PISO = 5
    CONST_DESIERTO = 6
    CONST_PESADO = 7
End Enum

Public Type tPaso
    CantPasos As Byte
    Wav() As Integer
End Type

Public Const NUM_PASOS As Byte = 7
Public Pasos() As tPaso

'Sonidos
Public Const SND_CLICK As String = 190
Public Const SND_MSG As String = 84
Public Const SND_FUEGO As Integer = 79
Public Const GRH_FOGATA As Integer = 1521
Public Const SND_RELAMPAGO As Byte = 97

'Musicas
Public Const MUS_Inicio As String = "1"
Public Const MUS_VolverInicio As String = "2"
Public Const MUS_CrearPersonaje As String = "3"

'Ambiental
Public Const SND_LLUVIAIN As Byte = 17
Public Const SND_LLUVIAOUT As Byte = 18
Public Const SND_VIENTO As Byte = 14
Public Const SND_AMBIENTE_NOCHE As Byte = 7
Public Const SND_AMBIENTE_NOCHE_CIU As Byte = 3

'###########################
' Constantes de intervalo
'###########################
Public Enum eIntervalos
    INT_ATTACK = 1400        'Atacar
    INT_ARROWS = 900        'Flechas
    INT_CAST_SPELL = 500    'Hechizos
    INT_CAST_ATTACK = 580   'Combo
    INT_WORK = 700
    INT_USEITEMU = 250      'Usar Item
    INT_USEITEMDCK = 250    'Usar Item dobleclick
    INT_SENTRPU = 2000      'Manda L
    INT_CHANGE_HEADING = 300
End Enum

'#######################
'CUERPOS Y CABEZAS
'#######################
Public Const iCuerpoMuerto As Integer = 8

Public Enum eCabezas
    CASPER_HEAD = 500
    FRAGATA_FANTASMAL = 87
    
    HUMANO_H_PRIMER_CABEZA = 1
    HUMANO_H_ULTIMA_CABEZA = 40 'En verdad es hasta la 51, pero como son muchas estas las dejamos no seleccionables
    HUMANO_H_CUERPO_DESNUDO = 21
    
    ELFO_H_PRIMER_CABEZA = 101
    ELFO_H_ULTIMA_CABEZA = 122
    ELFO_H_CUERPO_DESNUDO = 210
    
    DROW_H_PRIMER_CABEZA = 201
    DROW_H_ULTIMA_CABEZA = 221
    DROW_H_CUERPO_DESNUDO = 32
    
    ENANO_H_PRIMER_CABEZA = 301
    ENANO_H_ULTIMA_CABEZA = 319
    ENANO_H_CUERPO_DESNUDO = 53
    
    GNOMO_H_PRIMER_CABEZA = 401
    GNOMO_H_ULTIMA_CABEZA = 416
    GNOMO_H_CUERPO_DESNUDO = 222
    
    ORCO_H_PRIMER_CABEZA = 501
    ORCO_H_ULTIMA_CABEZA = 530
    ORCO_H_CUERPO_DESNUDO = 631
    
    VAMPIRO_H_PRIMER_CABEZA = 623
    VAMPIRO_H_ULTIMA_CABEZA = 633
    VAMPIRO_H_CUERPO_DESNUDO = 633
    
    '<--MUJER-->
    
    HUMANO_M_PRIMER_CABEZA = 70
    HUMANO_M_ULTIMA_CABEZA = 89
    HUMANO_M_CUERPO_DESNUDO = 39
    
    ELFO_M_PRIMER_CABEZA = 170
    ELFO_M_ULTIMA_CABEZA = 188
    ELFO_M_CUERPO_DESNUDO = 259
    
    DROW_M_PRIMER_CABEZA = 270
    DROW_M_ULTIMA_CABEZA = 288
    DROW_M_CUERPO_DESNUDO = 40
    
    ENANO_M_PRIMER_CABEZA = 370
    ENANO_M_ULTIMA_CABEZA = 384
    ENANO_M_CUERPO_DESNUDO = 60
    
    GNOMO_M_PRIMER_CABEZA = 470
    GNOMO_M_ULTIMA_CABEZA = 484
    GNOMO_M_CUERPO_DESNUDO = 260
    
    ORCO_M_PRIMER_CABEZA = 570
    ORCO_M_ULTIMA_CABEZA = 599
    ORCO_M_CUERPO_DESNUDO = 632
    
    VAMPIRO_M_PRIMER_CABEZA = 634
    VAMPIRO_M_ULTIMA_CABEZA = 643
    VAMPIRO_M_CUERPO_DESNUDO = 634
End Enum

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Enum ePartesCuerpo
    bCabeza = 1
    bPiernaIzquierda = 2
    bPiernaDerecha = 3
    bBrazoDerecho = 4
    bBrazoIzquierdo = 5
    bTorso = 6
End Enum

Public ObjetoTrabajo() As tItemsConstruibles
Public ObjArtesano() As tItemArtesano

Public UsaMacro As Boolean
Public CnTd As Byte

Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

Public TradingUserName As String

Public Tips() As String * 255

'Direcciones
Public Enum E_Heading
    SOUTH = 1
    NORTH = 2
    WEST = 3
    EAST = 4
End Enum

'##############
'INVENTARIO
'##############

' Cantidad de "slots" en el inventario sin mochila
Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 25

' Cantidad de "slots" en el inventario con mochila chica
Public Const MAX_MOCHILA_CHICA_INVENTORY_SLOTS As Byte = 30

' Cantidad de "slots" en el inventario con alforja
Public Const MAX_INVENTORY_SLOTS        As Byte = 35

Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50

Public Const INV_OFFER_SLOTS As Byte = 20
Public Const INV_GOLD_SLOTS As Byte = 1

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1
Public Const GOLD_OFFER_SLOT As Integer = INV_OFFER_SLOTS + 1

Public Const GRH_SLOT_INVENTARIO_NEGRO As Integer = 26095
Public Const GRH_SLOT_INVENTARIO_ROJO As Integer = 26096
Public Const GRH_SLOT_INVENTARIO_VIOLETA As Integer = 6834
Public Const GRH_SLOT_INVENTARIO_DORADO As Integer = 6840

Public MaxInventorySlots As Byte

Public Enum eClass
    Mage = 1      'Mago
    Cleric = 2    'Clerigo
    Warrior = 3   'Guerrero
    Assasin = 4   'Asesino
    Thief = 5     'Ladron
    Bard = 6      'Bardo
    Druid = 7     'Druida
    Bandit = 8    'Bandido
    Paladin = 9   'Paladin
    Hunter = 10   'Cazador
    Brujo = 11    'Brujo
    Pirate = 12   'Pirata
End Enum

Enum eRaza
    Humano = 1
    Elfo = 2
    ElfoOscuro = 3
    Gnomo = 4
    Enano = 5
    Orco = 6
    Vampiro = 7
End Enum

Public Enum eSkill
    Magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apunalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Defensa = 9
    Proyectiles = 10
    Wrestling = 11
    Comerciar = 12
    Domar = 13
    '<--Fijos-->
    Liderazgo = 14
    Navegacion = 15
    Equitacion = 16
    Talar = 17
    pesca = 18
    Mineria = 19
    Carpinteria = 20
    Herreria = 21
    Sastreria = 22
    Herboristeria = 23
    Alquimia = 24
End Enum

'¿Esta constante es una skill? La poco aqui por las dudas
Public Const FundirMetal As Integer = 88

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer = 2
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType

    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otescudo = 16
    otcasco = 17
    otFlechas = 32

End Enum

'###################
'CONSTANTES MAXIMAS
'###################
Public Const NUMCIUDADES As Byte = 5
Public Const NUMSKILLS As Byte = 24
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 12
Public Const NUMRAZAS As Byte = 7
Public Const MAXSKILLPOINTS As Byte = 100
Public Const NUMATRIBUTES As Byte = 5
Public Const MAXHECHI As Byte = 35

' Determina el color del nick
Public Enum eNickColor
    ieCriminal = &H1
    ieCiudadano = &H2
    ieAtacable = &H4
End Enum

'###################
' Mensajes
'###################

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_2 As String = "!!"
Public Const MENSAJE_22 As String = "!"

Public Enum eMessages
    NPCSwing
    NPCKillUser
    BlockedWithShieldUser
    BlockedWithShieldOther
    UserSwing
    SafeModeOn
    SafeModeOff
    CombatSafeOff
    CombatSafeOn
    NobilityLost
    CantUseWhileMeditating
    NPCHitUser
    UserHitNPC
    UserAttackedSwing
    UserHittedByUser
    UserHittedUser
    WorkRequestTarget
    HaveKilledUser
    UserKill
    EarnExp
    Home
    FinishHome
    
    '//Mensajes nuevos
    UserMuerto
    NpcInmune
    Hechizo_HechiceroMSG_NOMBRE
    Hechizo_HechiceroMSG_ALGUIEN
    Hechizo_HechiceroMSG_CRIATURA
 
    Hechizo_PropioMSG
    Hechizo_TargetMSG
End Enum

'Inventario
Type Inventory
    ObjIndex As Integer
    name As String
    GrhIndex As Long
    Amount As Long
    Equipped As Byte
    valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer 'Budi
    MaxHit As Integer
    MinHit As Integer
    NoUsa As Boolean
End Type

Type NpCinV
    ObjIndex As Integer
    name As String
    GrhIndex As Long
    Amount As Integer
    valor As Single
    OBJType As Integer
    MaxDef As Integer
    MinDef As Integer
    MaxHit As Integer
    MinHit As Integer
    NoUsa As Boolean
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Type tItemsConstruibles
    name As String
    ObjIndex As Integer
    GrhIndex As Long
    Materiales(1 To 4) As Integer
    CantMateriales(1 To 4) As Integer
    NameMateriales(1 To 4) As String
    PrecioConstruccion As Long
End Type

Type tItemCrafteo
    name As String
    ObjIndex As Integer
    GrhIndex As Long
    Amount As Integer
End Type

Type tItemArtesano
    name As String
    ObjIndex As Integer
    GrhIndex As Long
    
    ItemsCrafteo() As tItemCrafteo
End Type

Public Nombres As Boolean

Public UserHechizos(1 To MAXHECHI) As Integer

Public Type PjCuenta
    Nombre      As String
    Head        As Integer
    Body        As Integer
    shield      As Byte
    helmet      As Byte
    weapon      As Byte
    Mapa        As Integer
    Class       As Byte
    Race        As Byte
    Map         As Integer
    Level       As Byte
    Criminal    As Boolean
    Dead        As Boolean
    GameMaster  As Boolean
End Type

Public cPJ() As PjCuenta

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV

Public Type tCurrentUser
    UserMap As Integer
    UserCuadrante As Integer
    UserMeditar As Boolean
    UserName As String
    AccountName As String
    AccountPassword As String
    AccountMail As String 'Solo usado en la creacion de cuenta
    NumberOfCharacters As Byte
    UserMaxHP As Integer
    UserMinHP As Integer
    UserMaxMAN As Integer
    UserMinMAN As Integer
    UserMaxSTA As Integer
    UserMinSTA As Integer
    UserMaxAGU As Byte
    UserMinAGU As Byte
    UserMaxHAM As Byte
    UserMinHAM As Byte
    UserGLD As Long
    UserLvl As Integer
    UserPort As Integer
    UserEstado As Byte '0 = Vivo & 1 = Muerto
    UserPasarNivel As Long
    UserExp As Long
    UserELO As Long
    UserNivelPVP As Long
    UserEXPPVP As Long
    UserELVPVP As Long
    UserReputacion As tReputacion
    UserEstadisticas As tEstadisticasUsu
    UserDescansar As Boolean
    UserParalizado As Boolean
    UserInvisible As Boolean
    UserNavegando As Boolean
    UserEquitando As Boolean
    UserEvento As Boolean
    UserFuerza As Byte
    UserAgilidad As Byte
    UserWeaponEqpSlot As Byte
    UserArmourEqpSlot As Byte
    UserHelmEqpSlot As Byte
    UserShieldEqpSlot As Byte
    UserCiego As Boolean
    UserEstupido As Boolean
    
    UserClase As eClass
    UserSexo As eGenero
    UserRaza As eRaza
    UserEmail As String
    UserBody As Integer
    UserHead As Integer
    
    UserSkills(1 To NUMSKILLS) As Byte
    PorcentajeSkills(1 To NUMSKILLS) As Byte
    UserAtributos(1 To NUMATRIBUTOS) As Byte
    
    VIP As String
    esVIP As Boolean
End Type

Public CurrentUser As tCurrentUser

Public bShowTutorial As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
Public MirandoForo As Boolean
Public MirandoAsignarSkills As Boolean
Public MirandoEstadisticas As Boolean
Public MirandoParty As Boolean
Public MirandoTrabajo As Byte
'<-------------------------NUEVO-------------------------->

Public SkillsNames(1 To NUMSKILLS) As String
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public SendingType As Byte
Public sndPrivateTo As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer

Public UsingSkill As Integer
Public InvitandoParty As Boolean

Public pingTime As Long

Public EsPartyLeader As Boolean

Public Enum E_MODO
    Normal = 1
    CrearNuevoPJ = 2
    Dados = 3
    CrearCuenta = 4
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Vida
    eo_Poss
    eo_Speed
End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param CASA dentro de una casa de las que se compran, para evitar limpiar items
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    nada = 0
    BAJOTECHO = 1
    CASA = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public Connected As Boolean 'True when connected to server

'Control
Public prgRun As Boolean 'When true the program ends

'
'********** FUNCIONES API ***********
'

'******Mouse Cursor*********
'Esto es para poder usar iconos de mouse .ani
'https://www.gs-zone.org/temas/cursor-ani.45555/#post-375757
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 
Public Const GLC_HCURSOR = (-12)
Public hSwapCursor As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
'******End Mouse Cursor****

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el browser y programas externos
Public Const SW_SHOWNORMAL As Long = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'CopyMemory Kernel Function
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'Lista de cabezas
Public Type tHead
    Std As Byte
    Texture As Integer
    startX As Integer
    startY As Integer
End Type

Public heads() As tHead
Public Cascos() As tHead

Public Type tIndiceCuerpo
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceAtaque
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer
End Type

Public Type tIndiceArmas
    weapon(1 To 4) As Long
End Type

Public Type tIndiceEscudos
    shield(1 To 4) As Long
End Type

Public EsperandoLevel As Boolean

' Tipos de mensajes
Public Enum eForumMsgType
    ieGeneral
    ieGENERAL_STICKY
    ieREAL
    ieREAL_STICKY
    ieCAOS
    ieCAOS_STICKY
End Enum

' Indica los privilegios para visualizar los diferentes foros
Public Enum eForumVisibility
    ieGENERAL_MEMBER = &H1
    ieREAL_MEMBER = &H2
    ieCAOS_MEMBER = &H4
End Enum

' Indica el tipo de foro
Public Enum eForumType
    ieGeneral
    ieREAL
    ieCAOS
End Enum

' Limite de posts
Public Const MAX_STICKY_POST As Byte = 5
Public Const MAX_GENERAL_POST As Byte = 30
Public Const STICKY_FORUM_OFFSET As Byte = 50

' Estructura contenedora de mensajes
Public Type tForo
    StickyTitle(1 To MAX_STICKY_POST) As String
    StickyPost(1 To MAX_STICKY_POST) As String
    StickyAuthor(1 To MAX_STICKY_POST) As String
    GeneralTitle(1 To MAX_GENERAL_POST) As String
    GeneralPost(1 To MAX_GENERAL_POST) As String
    GeneralAuthor(1 To MAX_GENERAL_POST) As String
End Type

' 1 foro general y 2 faccionarios
Public Foros(0 To 2) As tForo

' Forum info handler
Public clsForos As clsForum

'FragShooter variables
Public FragShooterCapturePending As Boolean
Public FragShooterNickname As String
Public FragShooterKilledSomeone As Boolean

Public bShowGuildNews As Boolean
Public GuildNames() As String
Public GuildMembers() As String

Public Const OFFSET_HEAD As Integer = -34

Public Enum eSMType
    sCombatMode = 0
    sSafemode
End Enum

'Hardcoded grhs and items

Public Const ORO_INDEX As Long = 12
Public Const ORO_GRH As Long = 511

Public picMouseIcon As Picture

Public Enum eMoveType
    Inventory = 1
    Bank
End Enum

'MundoSeleccionado desde la propiedad Mundo en sinfo.dat / World selected from sinfo.dat file
Public MundoSeleccionado As String

' * Configuracion de estilo de controles

Public JsonTips As Object

'Nivel Maximo
Public Const STAT_MAXELV As Byte = 50
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer

Public Security As New clsSecurity

Public Const SPEED_NORMAL As Double = 0.018

Public Type ShopObj
    ObjIndex As Integer
    Nombre As String
    Amount As Integer
    valor As Long
End Type

Public ShopObject() As ShopObj
