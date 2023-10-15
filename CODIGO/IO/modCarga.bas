Attribute VB_Name = "Carga"
' ***********************************************
'   Nueva carga de configuracion mediante .INI
' ***********************************************

Option Explicit

Public Type tCabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Enum ePath
    Script
    Init
    Graficos
    Interfaces
    skins
    Lenguajes
    recursos
End Enum

Public Enum E_SISTEMA_MUSICA
    CONST_DESHABILITADA = 0
    CONST_MP3 = 1
End Enum

Public Type tSetupMods

    ' VIDEO
    byMemory    As Integer
    ProyectileEngine As Boolean
    PartyMembers As Boolean
    UsarSombras As Boolean
    UsarReflejos As Boolean
    UsarAuras As Boolean
    ParticleEngine As Boolean
    HUD As Boolean
    LimiteFPS As Boolean
    bNoRes      As Boolean
    FPSShow      As Boolean
    OverrideVertexProcess As Byte
    
    ' AUDIO
    bMusic    As E_SISTEMA_MUSICA
    bSound    As Byte
    bAmbient As Byte
    Invertido As Byte
    MusicVolume As Long
    SoundVolume As Long
    AmbientVol As Long
    
    ' GUILDS
    bGuildNews  As Boolean
    bGldMsgConsole As Boolean
    bCantMsgs   As Byte
    
    ' FRAGSHOOTER
    bActive     As Boolean
    bDie        As Boolean
    bKill       As Boolean
    byMurderedLevel As Byte
    
    ' LOGIN
    Remember    As Boolean
    rUserName   As String
    
    ' OTHER
    MostrarTips As Byte
    MostrarBindKeysSelection As Byte
    BloqueoMovimiento As Boolean
    VerCuadrantes As Boolean
    
    'MOUSE
    MouseGeneral As Byte
    MouseBaston As Byte
    SkinSeleccionado As String
    
    'Funciones
    Funcion(1 To 12) As String
End Type

Public ClientSetup As tSetupMods
Public MiCabecera As tCabecera

Private Lector As clsIniManager
Public Const CLIENT_FILE As String = "Config.ini"

'********************************
'Load Map with .CSM format
'********************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroParticulas As Long
    NumeroLuces As Long
    NumeroZonas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
    NumeroData As Integer
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    GrhIndex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosZonas
    X As Integer
    Y As Integer
    Zona As Integer
End Type

Public Type tDatosLuces
    r As Integer
    G As Integer
    b As Integer
    range As Byte
    X As Integer
    Y As Integer
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    ObjIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Public MapSize As tMapSize
'********************************
'END - Load Map with .CSM format
'********************************

Private FileManager As clsIniManager

Public NumHeads As Integer
Public NumCascos As Integer
Public NumEscudosAnims As Integer
Public NumAtaques As Integer

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
        .CRC = Rnd * 245
        .MagicWord = Rnd * 92
    End With
    
End Sub

Public Function Path(ByVal PathType As ePath) As String

    Select Case PathType
            
        Case ePath.Init
            Path = App.Path & "\INIT\"
        
        Case ePath.Graficos
            Path = "..\Shared\Graficos\"
        
        Case ePath.skins
            Path = App.Path & "\Recursos\Skins\"
            
        Case ePath.Lenguajes
            Path = App.Path & "\Recursos\Lenguajes\"
               
        Case ePath.recursos
            Path = App.Path & "\Recursos"
    
    End Select

End Function

Public Sub LeerConfiguracion()
    On Local Error GoTo fileErr:
    
    Dim i As Byte

    Call IniciarCabecera

    Set Lector = New clsIniManager
    Call Lector.Initialize(Carga.Path(Init) & CLIENT_FILE)

    With ClientSetup
        ' VIDEO
        .byMemory = CInt(Lector.GetValue("VIDEO", "DynamicMemory"))
        .bNoRes = CBool(Lector.GetValue("VIDEO", "DisableResolutionChange"))
        .ProyectileEngine = CBool(Lector.GetValue("VIDEO", "ProjectileEngine"))
        .PartyMembers = CBool(Lector.GetValue("VIDEO", "PartyMembers"))
        .UsarSombras = CBool(Lector.GetValue("VIDEO", "Sombras"))
        .UsarReflejos = CBool(Lector.GetValue("VIDEO", "Reflejos"))
        .UsarAuras = CBool(Lector.GetValue("VIDEO", "Auras"))
        .ParticleEngine = CBool(Lector.GetValue("VIDEO", "ParticleEngine"))
        .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LimitarFPS"))
        .HUD = CBool(Lector.GetValue("VIDEO", "HUD"))
        .OverrideVertexProcess = CByte(Lector.GetValue("VIDEO", "VertexProcessingOverride"))
        
        ' AUDIO
        .bMusic = CByte(Lector.GetValue("AUDIO", "MUSICA"))
        .bSound = CByte(Lector.GetValue("AUDIO", "SONIDO"))
        .bAmbient = CByte(Lector.GetValue("AUDIO", "AMBIENT"))
        .MusicVolume = CLng(Lector.GetValue("AUDIO", "VOLMUSICA"))
        .SoundVolume = CLng(Lector.GetValue("AUDIO", "VOLAUDIO"))
        .AmbientVol = CLng(Lector.GetValue("AUDIO", "VOLAMBIENT"))

        ' GUILD
        .bGuildNews = CBool(Lector.GetValue("GUILD", "NEWS"))
        .bGldMsgConsole = CBool(Lector.GetValue("GUILD", "MESSAGES"))
        .bCantMsgs = CByte(Lector.GetValue("GUILD", "MAX_MESSAGES"))
        
        ' FRAGSHOOTER
        .bDie = CBool(Lector.GetValue("FRAGSHOOTER", "DIE"))
        .bKill = CBool(Lector.GetValue("FRAGSHOOTER", "KILL"))
        .byMurderedLevel = CByte(Lector.GetValue("FRAGSHOOTER", "MURDERED_LEVEL"))
        .bActive = CBool(Lector.GetValue("FRAGSHOOTER", "ACTIVE"))
        
        ' LOGIN
        .Remember = CBool(Lector.GetValue("LOGIN", "REMEMBER"))
        .rUserName = Trim$(Lector.GetValue("LOGIN", "USERNAME"))
        
        ' OTHER
        .MostrarTips = CBool(Lector.GetValue("OTHER", "MOSTRAR_TIPS"))
        .MostrarBindKeysSelection = CBool(Lector.GetValue("OTHER", "MOSTRAR_BIND_KEYS_SELECTION"))
        .BloqueoMovimiento = CBool(Lector.GetValue("OTHER", "BLOQUEOMOV"))
        .VerCuadrantes = CBool(Lector.GetValue("OTHER", "CUADRANTES"))
        
        ' FUNCION
        For i = 1 To 12
            .Funcion(i) = Trim$(CStr(Lector.GetValue("FUNCION", "F" & i)))
        Next i
        
    End With

  Exit Sub
  
fileErr:

    If Err.number <> 0 Then
       MsgBox ("Ha ocurrido un error al cargar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
       End 'Usar "End" en vez del Sub CloseClient() ya que todavia no se inicializa nada.
    End If
End Sub

Public Sub GuardarConfiguracion()
    On Local Error GoTo fileErr:
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(Carga.Path(Init) & CLIENT_FILE)

    With ClientSetup
        
        ' VIDEO
        Call Lector.ChangeValue("VIDEO", "DynamicMemory", .byMemory)
        Call Lector.ChangeValue("VIDEO", "DisableResolutionChange", IIf(.bNoRes, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "ParticleEngine", IIf(.ProyectileEngine, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "PartyMembers", IIf(.PartyMembers, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "Sombras", IIf(.UsarSombras, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "Reflejos", IIf(.UsarReflejos, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "Auras", IIf(.UsarAuras, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "ParticleEngine", IIf(.ParticleEngine, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "LimitarFPS", IIf(.LimiteFPS, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "HUD", IIf(.HUD, "1", "0"))
        Call Lector.ChangeValue("VIDEO", "VertexProcessingOverride", .OverrideVertexProcess)
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "MUSICA", .bMusic)
        Call Lector.ChangeValue("AUDIO", "SONIDO", .bSound)
        Call Lector.ChangeValue("AUDIO", "AMBIENT", .bAmbient)
        Call Lector.ChangeValue("AUDIO", "VOLMUSICA", .MusicVolume)
        Call Lector.ChangeValue("AUDIO", "VOLAUDIO", .SoundVolume)
        Call Lector.ChangeValue("AUDIO", "VOLAMBIENT", .AmbientVol)
        
        ' GUILD
        Call Lector.ChangeValue("GUILD", "NEWS", IIf(.bGuildNews, "1", "0"))
        Call Lector.ChangeValue("GUILD", "MESSAGES", IIf(DialogosClanes.Activo, "1", "0"))
        Call Lector.ChangeValue("GUILD", "MAX_MESSAGES", CByte(DialogosClanes.CantidadDialogos))
        
        ' FRAGSHOOTER
        Call Lector.ChangeValue("FRAGSHOOTER", "DIE", IIf(.bDie, "1", "0"))
        Call Lector.ChangeValue("FRAGSHOOTER", "KILL", IIf(.bKill, "1", "0"))
        Call Lector.ChangeValue("FRAGSHOOTER", "MURDERED_LEVEL", CByte(.byMurderedLevel))
        Call Lector.ChangeValue("FRAGSHOOTER", "ACTIVE", IIf(.bActive, "1", "0"))
        
        ' OTHER
        ' Lo comento por que no tiene por que setearse aqui esto.
        ' Al menos no al hacer click en el boton Salir del formulario opciones (Recox)
        ' Call Lector.ChangeValue("OTHER", "MOSTRAR_TIPS", CBool(.MostrarTips))
        Call Lector.ChangeValue("OTHER", "BLOQUEOMOV", IIf(.bActive, "1", "0"))
        Call Lector.ChangeValue("OTHER", "CUADRANTES", IIf(.VerCuadrantes, "1", "0"))
    End With
    
    Call Lector.DumpFile(Carga.Path(Init) & CLIENT_FILE)
fileErr:

    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
    End If
End Sub

Public Sub GuardarFunciones()
'*************************************
'Autor: Lorwik
'Fecha: 13/09/2020
'Descripción: Guarda la configuración de funciones
'*************************************

    On Local Error GoTo fileErr:
    
    Dim i As Byte
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(Carga.Path(Init) & CLIENT_FILE)
    
    For i = 1 To 12
        Call Lector.ChangeValue("FUNCION", "F" & i, ClientSetup.Funcion(i))
    Next i
    
    Call Lector.DumpFile(Carga.Path(Init) & CLIENT_FILE)
fileErr:

    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion de funciones del cliente. Error " & Err.number & " : " & Err.Description)
    End If
End Sub

''
' Loads grh data using the new file format.
'

Public Sub LoadGrhData()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Graficos
'*************************************
On Error GoTo ErrorHandler:

    Dim Grh         As Long
    Dim Frame       As Long
    Dim grhCount    As Long
    Dim fileVersion As Long
    Dim LaCabecera  As tCabecera
    Dim fileBuff    As clsByteBuffer
    Dim InfoHead    As INFOHEADER
    Dim buffer()    As Byte
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("Graficos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Graficos.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        fileVersion = fileBuff.getLong
        
        grhCount = fileBuff.getLong
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Grh < grhCount
            Grh = fileBuff.getLong

            With GrhData(Grh)
            
                '.active = True
                .NumFrames = fileBuff.getInteger
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        .Frames(Frame) = fileBuff.getLong
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    .speed = fileBuff.getSingle
                    If .speed <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                    
                Else
                    
                    .FileNum = fileBuff.getLong
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = fileBuff.getInteger
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .pixelHeight = fileBuff.getInteger
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .sX = fileBuff.getInteger
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    .sY = fileBuff.getInteger
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    '.Trans = fileBuff.getByte
                    'If .Trans < 0 Then GoTo ErrorHandler
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
Exit Sub

ErrorHandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarCabezas()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cabezas
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim i           As Integer
    Dim LaCabecera  As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("Head.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Head.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
        
        NumHeads = fileBuff.getInteger()  'cantidad de cabezas
    
        ReDim heads(0 To NumHeads) As tHead
                
        For i = 1 To NumHeads
            heads(i).Std = fileBuff.getByte()
            heads(i).Texture = fileBuff.getInteger()
            heads(i).startX = fileBuff.getInteger()
            heads(i).startY = fileBuff.getInteger()
        Next i
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Head.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarCascos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cascos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Integer
    Dim LaCabecera  As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("Helmet.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Helmet.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        NumCascos = fileBuff.getInteger()   'cantidad de cascos
             
        ReDim Cascos(0 To NumCascos) As tHead
             
        For i = 1 To NumCascos
            Cascos(i).Std = fileBuff.getByte()
            Cascos(i).Texture = fileBuff.getInteger()
            Cascos(i).startX = fileBuff.getInteger()
            Cascos(i).startY = fileBuff.getInteger()
        Next i
         
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Helmet.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarCuerpos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Cuerpos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    Dim LaCabecera As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("Personajes.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Personajes.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de cabezas
        NumCuerpos = fileBuff.getInteger()
    
        'Resize array
        ReDim BodyData(0 To NumCuerpos) As BodyData
        ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
        For i = 1 To NumCuerpos
            MisCuerpos(i).Body(1) = fileBuff.getLong()
            MisCuerpos(i).Body(2) = fileBuff.getLong()
            MisCuerpos(i).Body(3) = fileBuff.getLong()
            MisCuerpos(i).Body(4) = fileBuff.getLong()
            MisCuerpos(i).HeadOffsetX = fileBuff.getInteger()
            MisCuerpos(i).HeadOffsetY = fileBuff.getInteger()
            
            If MisCuerpos(i).Body(1) Then
                Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
                Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
                Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
                Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
                
                BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
                BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarAtaques()
'*************************************
'Autor: Lorwik
'Fecha:03/11/2020
'Descripción: Carga el index de Cuerpos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    Dim LaCabecera As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("Ataques.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Ataques.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de cabezas
        NumAtaques = fileBuff.getInteger()
    
        'Resize array
        ReDim MisAtaques(0 To NumAtaques) As tIndiceAtaque
        ReDim AtaqueData(0 To NumAtaques) As AtaqueAnimData
    
        For i = 1 To NumCuerpos
            MisAtaques(i).Body(1) = fileBuff.getLong()
            MisAtaques(i).Body(2) = fileBuff.getLong()
            MisAtaques(i).Body(3) = fileBuff.getLong()
            MisAtaques(i).Body(4) = fileBuff.getLong()
            MisAtaques(i).HeadOffsetX = fileBuff.getInteger()
            MisAtaques(i).HeadOffsetY = fileBuff.getInteger()
            
            If MisCuerpos(i).Body(1) Then
                Call InitGrh(AtaqueData(i).AtaqueWalk(1), MisAtaques(i).Body(1), 0)
                Call InitGrh(AtaqueData(i).AtaqueWalk(2), MisAtaques(i).Body(2), 0)
                Call InitGrh(AtaqueData(i).AtaqueWalk(3), MisAtaques(i).Body(3), 0)
                Call InitGrh(AtaqueData(i).AtaqueWalk(4), MisAtaques(i).Body(4), 0)
                
                AtaqueData(i).HeadOffset.X = MisAtaques(i).HeadOffsetX
                AtaqueData(i).HeadOffset.Y = MisAtaques(i).HeadOffsetY
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarFxs()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Fxs
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i           As Long
    Dim NumFxs      As Integer
    Dim LaCabecera  As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("FXs.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("FXs.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de Fxs
        NumFxs = fileBuff.getInteger()
        
        'Resize array
        ReDim FxData(1 To NumFxs) As tIndiceFx
        
        For i = 1 To NumFxs
            FxData(i).Animacion = fileBuff.getLong()
            FxData(i).OffsetX = fileBuff.getInteger()
            FxData(i).OffsetY = fileBuff.getInteger()
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Fxs.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If

End Sub

Public Sub CargarTips()
'************************************************************************************.
' Carga el JSON con los tips del juego en un objeto para su uso a lo largo del proyecto
'************************************************************************************
On Error GoTo errhandler:
    
    Dim TipFile As String
        TipFile = FileToString(Carga.Path(Lenguajes) & "tips_" & Language & ".json")
    
    Set JsonTips = JSON.parse(TipFile)

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo" & "tips_" & Language & ".json no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
End Sub

Sub CargarAnimArmas()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Armas
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim dLen        As Long
    Dim InfoHead    As INFOHEADER
    Dim i As Long
    Dim LaCabecera As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("Armas.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Armas.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de armas
        NumWeaponAnims = fileBuff.getInteger()
        
        'Resize array
        ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
        ReDim Weapons(1 To NumWeaponAnims) As tIndiceArmas
        
        For i = 1 To NumWeaponAnims
            Weapons(i).weapon(1) = fileBuff.getLong()
            Weapons(i).weapon(2) = fileBuff.getLong()
            Weapons(i).weapon(3) = fileBuff.getLong()
            Weapons(i).weapon(4) = fileBuff.getLong()
            
            If Weapons(i).weapon(1) Then
            
                Call InitGrh(WeaponAnimData(i).WeaponWalk(1), Weapons(i).weapon(1), 0)
                Call InitGrh(WeaponAnimData(i).WeaponWalk(2), Weapons(i).weapon(2), 0)
                Call InitGrh(WeaponAnimData(i).WeaponWalk(3), Weapons(i).weapon(3), 0)
                Call InitGrh(WeaponAnimData(i).WeaponWalk(4), Weapons(i).weapon(4), 0)
            
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Armas.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If

End Sub

Public Sub CargarColores()
'*************************************
'Autor: Lorwik
'Fecha: 30/08/2020
'Descripción: Carga los colores
'*************************************
On Error GoTo errhandler:
    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim LaCabecera  As tCabecera
    Dim fileBuff    As clsByteBuffer
    Dim i           As Long
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("Colores.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Colores.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
        
        For i = 0 To MAXCOLORES
        
            ColoresPJ(i) = fileBuff.getLong
        
        Next i
        
        Erase buffer
    End If
    
    Set fileBuff = Nothing

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Colores.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
End Sub

Sub CargarAnimEscudos()
'*************************************
'Autor: Lorwik
'Fecha: ???
'Descripción: Carga el index de Escudos
'*************************************
On Error GoTo errhandler:

    Dim buffer()    As Byte
    Dim InfoHead    As INFOHEADER
    Dim i As Long
    Dim LaCabecera As tCabecera
    Dim fileBuff  As clsByteBuffer
    
    InfoHead = File_Find(Carga.Path(ePath.recursos) & "\Scripts.WAO", LCase$("Escudos.ind"))
    
    If InfoHead.lngFileSize <> 0 Then
    
        Extract_File_Memory Scripts, LCase$("Escudos.ind"), buffer()
        
        Set fileBuff = New clsByteBuffer
        
        fileBuff.initializeReader buffer
        
        LaCabecera.Desc = fileBuff.getString(Len(LaCabecera.Desc))
        LaCabecera.CRC = fileBuff.getLong
        LaCabecera.MagicWord = fileBuff.getLong
    
        'num de escudos
        NumEscudosAnims = fileBuff.getInteger()
        
        'Resize array
        ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
        ReDim Shields(1 To NumEscudosAnims) As tIndiceEscudos
        
        For i = 1 To NumEscudosAnims
            Shields(i).shield(1) = fileBuff.getLong()
            Shields(i).shield(2) = fileBuff.getLong()
            Shields(i).shield(3) = fileBuff.getLong()
            Shields(i).shield(4) = fileBuff.getLong()
            
            If Shields(i).shield(1) Then
            
                Call InitGrh(ShieldAnimData(i).ShieldWalk(1), Shields(i).shield(1), 0)
                Call InitGrh(ShieldAnimData(i).ShieldWalk(2), Shields(i).shield(2), 0)
                Call InitGrh(ShieldAnimData(i).ShieldWalk(3), Shields(i).shield(3), 0)
                Call InitGrh(ShieldAnimData(i).ShieldWalk(4), Shields(i).shield(4), 0)
            
            End If
        Next i
    
        Erase buffer
    End If
    
    Set fileBuff = Nothing

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Escudos.ind no existe. Por favor, reinstale el juego.", , Form_Caption)
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarMapa(ByVal fileMap As String)

    On Error GoTo ErrorHandler

    Dim fh           As Integer
    
    Dim MH           As tMapHeader
    Dim Blqs()       As tDatosBloqueados

    Dim L1()         As Long
    Dim L2()         As tDatosGrh
    Dim L3()         As tDatosGrh
    Dim L4()         As tDatosGrh

    Dim Triggers()   As tDatosTrigger
    Dim Luces()      As tDatosLuces
    Dim Particulas() As tDatosParticulas
    Dim Objetos()    As tDatosObjs
    Dim NPCs()       As tDatosNPC
    Dim TEs()        As tDatosTE
    Dim Zonas()      As tDatosZonas
    
    Dim i            As Long
    Dim j            As Long

    Dim LaCabecera   As tCabecera
    
    Dim buffer()     As Byte
    Dim fileBuff     As clsByteBuffer
    
    DoEvents
    
    Extract_File_Memory srcFileType.Map, fileMap, buffer()
    
    Set fileBuff = New clsByteBuffer
        
    fileBuff.initializeReader buffer
    
    With LaCabecera
        .Desc = fileBuff.getString(Len(.Desc))
        .CRC = fileBuff.getLong
        .MagicWord = fileBuff.getLong
    End With
    
    With MH
        .NumeroBloqueados = fileBuff.getLong()
        .NumeroLayers(2) = fileBuff.getLong()
        .NumeroLayers(3) = fileBuff.getLong()
        .NumeroLayers(4) = fileBuff.getLong()
        .NumeroTriggers = fileBuff.getLong()
        .NumeroParticulas = fileBuff.getLong()
        .NumeroLuces = fileBuff.getLong()
        .NumeroZonas = fileBuff.getLong()
        .NumeroNPCs = fileBuff.getLong()
        .NumeroOBJs = fileBuff.getLong()
        .NumeroTE = fileBuff.getLong()
        .NumeroData = fileBuff.getInteger()
    End With
    
    With MapSize
        .XMax = fileBuff.getInteger()
        .XMin = fileBuff.getInteger()
        .YMax = fileBuff.getInteger()
        .YMin = fileBuff.getInteger()
    End With

    CantZonas = MH.NumeroData

    ReDim MapZonas(CantZonas) As tZonaInfo
    
    For i = 0 To CantZonas
        With MapZonas(i)
            
            .name = fileBuff.getString()
            Call fileBuff.getBoolean
            Call fileBuff.getBoolean
            Call fileBuff.getString
            .Music = fileBuff.getString()
            .Zona = fileBuff.getString()
            .Terreno = fileBuff.getString()
            .Ambient = fileBuff.getString()
            Call fileBuff.getString
            Call fileBuff.getBoolean
            Call fileBuff.getBoolean
            Call fileBuff.getBoolean
            Call fileBuff.getBoolean
            Call fileBuff.getBoolean
            Call fileBuff.getBoolean
            .LuzBase = fileBuff.getLong()
            Debug.Print .LuzBase
            Call fileBuff.getLong
            Call fileBuff.getBoolean
            
        End With
    Next i

    With MapSize
        ReDim L1(.XMin To .XMax, .YMin To .YMax)
    End With

    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax

            L1(i, j) = fileBuff.getLong()
            
            If L1(i, j) > 0 Then
                Call InitGrh(MapData(i, j).Graphic(1), L1(i, j))
            End If

        Next i
    Next j

    With MH

        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)

            For i = 1 To .NumeroBloqueados
                With Blqs(i)
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()
                    MapData(.X, .Y).Blocked = 1
                End With
            Next i

        End If
        
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            
            For i = 1 To .NumeroLayers(2)
            
                With L2(i)
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()
                    .GrhIndex = fileBuff.getLong()
                
                    Call InitGrh(MapData(.X, .Y).Graphic(2), .GrhIndex)
                End With
            Next i

        End If
        
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))

            For i = 1 To .NumeroLayers(3)
            
                With L3(i)
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()
                    .GrhIndex = fileBuff.getLong()
                
                    Call InitGrh(MapData(.X, .Y).Graphic(3), .GrhIndex)
                End With
            Next i

        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            
            For i = 1 To .NumeroLayers(4)
            
                With L4(i)
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()
                    .GrhIndex = fileBuff.getLong()
  
                    Call InitGrh(MapData(.X, .Y).Graphic(4), .GrhIndex)
                End With
            Next i

        End If
        
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            
            For i = 1 To .NumeroTriggers
                
                With Triggers(i)
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()
                    .Trigger = fileBuff.getInteger()
                
                    MapData(.X, .Y).Trigger = .Trigger
                End With
                
            Next i

        End If

        If .NumeroParticulas > 0 Then
        
            ReDim Particulas(1 To .NumeroParticulas)

            For i = 1 To .NumeroParticulas

                With Particulas(i)
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()
                    .Particula = fileBuff.getLong()

                    MapData(.X, .Y).Particle_Index = .Particula
                    Call General_Particle_Create(.Particula, .X, .Y)
                End With

            Next i

        End If

        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Dim p As Byte
            
            For i = 1 To .NumeroLuces
                With Luces(i)
                    .r = fileBuff.getInteger()
                    .G = fileBuff.getInteger()
                    .b = fileBuff.getInteger()
                    .range = fileBuff.getByte()
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()

                    Call Create_Light_To_Map(.X, .Y, .range, .r, .G, .b, False)
                End With
            Next i
            
        End If
        
        If .NumeroZonas > 0 Then
            ReDim Zonas(1 To .NumeroZonas)
            
            For i = 1 To .NumeroZonas
                
                With Zonas(i)
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()
                    .Zona = fileBuff.getInteger()
                
                    MapData(.X, .Y).ZonaIndex = .Zona
                End With
                
            Next i

        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            
            For i = 1 To .NumeroOBJs

                With Objetos(i)
                    .X = fileBuff.getInteger()
                    .Y = fileBuff.getInteger()
                    .ObjIndex = fileBuff.getInteger()
                    .ObjAmmount = fileBuff.getInteger()
                
                    'Erase OBJs
                    MapData(.X, .Y).ObjGrh.GrhIndex = 0
                End With
            Next i
            
        End If
        
    End With
    
    Erase buffer
    Set fileBuff = Nothing

ErrorHandler:
    
    If fh <> 0 Then Close fh
    
    If Err.number <> 0 Then
        'Call LogError(Err.number, Err.Description, "modCarga.CargarMapa")
        Call MsgBox("err: " & Err.number & " desc: " & Err.Description)
    End If

End Sub

Public Sub CargarPasos()

    ReDim Pasos(1 To NUM_PASOS) As tPaso

    Pasos(CONST_BOSQUE).CantPasos = 2
    ReDim Pasos(CONST_BOSQUE).Wav(1 To Pasos(CONST_BOSQUE).CantPasos) As Integer
    Pasos(CONST_BOSQUE).Wav(1) = 201
    Pasos(CONST_BOSQUE).Wav(2) = 202

    Pasos(CONST_NIEVE).CantPasos = 2
    ReDim Pasos(CONST_NIEVE).Wav(1 To Pasos(CONST_NIEVE).CantPasos) As Integer
    Pasos(CONST_NIEVE).Wav(1) = 199
    Pasos(CONST_NIEVE).Wav(2) = 200

    Pasos(CONST_CABALLO).CantPasos = 2
    ReDim Pasos(CONST_CABALLO).Wav(1 To Pasos(CONST_CABALLO).CantPasos) As Integer
    Pasos(CONST_CABALLO).Wav(1) = 23
    Pasos(CONST_CABALLO).Wav(2) = 24

    Pasos(CONST_DUNGEON).CantPasos = 2
    ReDim Pasos(CONST_DUNGEON).Wav(1 To Pasos(CONST_DUNGEON).CantPasos) As Integer
    Pasos(CONST_DUNGEON).Wav(1) = 23
    Pasos(CONST_DUNGEON).Wav(2) = 24

    Pasos(CONST_DESIERTO).CantPasos = 2
    ReDim Pasos(CONST_DESIERTO).Wav(1 To Pasos(CONST_DESIERTO).CantPasos) As Integer
    Pasos(CONST_DESIERTO).Wav(1) = 197
    Pasos(CONST_DESIERTO).Wav(2) = 198

    Pasos(CONST_PISO).CantPasos = 2
    ReDim Pasos(CONST_PISO).Wav(1 To Pasos(CONST_PISO).CantPasos) As Integer
    Pasos(CONST_PISO).Wav(1) = 23
    Pasos(CONST_PISO).Wav(2) = 24

    Pasos(CONST_PESADO).CantPasos = 3
    ReDim Pasos(CONST_PESADO).Wav(1 To Pasos(CONST_PESADO).CantPasos) As Integer
    Pasos(CONST_PESADO).Wav(1) = 220
    Pasos(CONST_PESADO).Wav(2) = 221
    Pasos(CONST_PESADO).Wav(3) = 222

End Sub

Public Sub CargarRecursos()
'*******************************
'Autor: Lorwik
'Fecha: 19/04/2021
'Descripción: Llamada a la carga de recursos
'*******************************

    Call frmCargando.ActualizarCarga(JsonLanguage.item("INICIA_CARGARECURSOS").item("TEXTO"), 62)
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarAtaques
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call LoadGraphics
    Call CargarParticulas
    Call frmCargando.ActualizarCarga(frmCargando.Caption = JsonLanguage.item("HECHO").item("TEXTO"), 63)
    
    'Inicializamos el conectar renderizado
    Call frmCargando.ActualizarCarga(JsonLanguage.item("INICIA_GUI").item("TEXTO"), 64)
    Call ModConectar.InicializarRndCNT
    Call frmCargando.ActualizarCarga(frmCargando.Caption = JsonLanguage.item("HECHO").item("TEXTO"), 65)
    
    '###################
    ' ANIMACIONES EXTRAS
    Call frmCargando.ActualizarCarga(JsonLanguage.item("INICIA_FXS").item("TEXTO"), 66)
    
    Call CargarTips
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    Call CargarPasos
    
End Sub
