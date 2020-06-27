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
    Skins
    Sounds
    Musica
    Mapas
    Lenguajes
    Fonts
    recursos
End Enum

Public Enum E_SISTEMA_MUSICA
    CONST_DESHABILITADA = 0
    CONST_MP3 = 1
End Enum

Public Type tSetupMods

    ' VIDEO
    Aceleracion As Byte
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
    
    ' OTHER
    MostrarTips As Byte
    MostrarBindKeysSelection As Byte
    
    'MOUSE
    MouseGeneral As Byte
    MouseBaston As Byte
    SkinSeleccionado As String
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
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
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

Private Type tDatosLuces
    X As Integer
    Y As Integer
    light_value(3) As Long
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
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
    objindex As Integer
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

Private Type tMapDat
    map_name As String
    battle_mode As Boolean
    backup_mode As Boolean
    restrict_mode As String
    music_number As String
    zone As String
    terrain As String
    Ambient As String
    lvlMinimo As String
    RoboNpcsPermitido As Boolean
    InvocarSinEfecto As Boolean
    OcultarSinEfecto As Boolean
    ResuSinEfecto As Boolean
    MagiaSinEfecto As Boolean
    InviSinEfecto As Boolean
    NoEncriptarMP As Boolean
    version As Long
End Type

Public MapSize As tMapSize
Public MapDat As tMapDat
'********************************
'END - Load Map with .CSM format
'********************************

'Conectar renderizado
Private Type tMapaConnect
    Map As Byte
    X As Byte
    Y As Byte
End Type
Public MapaConnect() As tMapaConnect
Public NumConnectMap As Byte 'Numero total de mapas cargados

Private FileManager As clsIniManager

Public NumHeads As Integer
Public NumCascos As Integer
Public NumEscudosAnims As Integer

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "WinterAO Resurrection mod Argentum Online by Noland Studios. http://winterao.com.ar"
        .CRC = Rnd * 245
        .MagicWord = Rnd * 92
    End With
    
End Sub

Public Function Path(ByVal PathType As ePath) As String

    Select Case PathType
        
        Case ePath.Script
            Path = App.Path & "\Recursos\INIT\"
            
        Case ePath.Init
            Path = App.Path & "\INIT\"
        
        Case ePath.Graficos
            Path = App.Path & "\Recursos\Graficos\"
        
        Case ePath.Skins
            Path = App.Path & "\Recursos\Graficos\Skins\"
            
        Case ePath.Interfaces
            Path = App.Path & "\Recursos\Graficos\Interfaces\"
            
        Case ePath.Lenguajes
            Path = App.Path & "\Recursos\Lenguajes\"
            
        Case ePath.Mapas
            Path = App.Path & "\Recursos\Mapas\"
            
        Case ePath.Musica
            Path = App.Path & "\Recursos\MP3\"
            
        Case ePath.Sounds
            Path = App.Path & "\Recursos\WAV\"
            
        Case ePath.recursos
            Path = App.Path & "\Recursos"
    
    End Select

End Function

Public Sub LeerConfiguracion()
    On Local Error GoTo fileErr:
    
    Call IniciarCabecera
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(Carga.Path(Init) & CLIENT_FILE)
    
    With ClientSetup
        ' VIDEO
        .Aceleracion = Lector.GetValue("VIDEO", "RenderMode")
        .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
        .bNoRes = CBool(Lector.GetValue("VIDEO", "DisableResolutionChange"))
        .ProyectileEngine = CBool(Lector.GetValue("VIDEO", "ProjectileEngine"))
        .PartyMembers = CBool(Lector.GetValue("VIDEO", "PartyMembers"))
        .UsarSombras = CBool(Lector.GetValue("VIDEO", "Sombras"))
        .UsarReflejos = CBool(Lector.GetValue("VIDEO", "Reflejos"))
        .UsarAuras = CBool(Lector.GetValue("VIDEO", "Auras"))
        .ParticleEngine = CBool(Lector.GetValue("VIDEO", "ParticleEngine"))
        .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LimitarFPS"))
        .HUD = CBool(Lector.GetValue("VIDEO", "HUD"))
        
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
        
        ' OTHER
        .MostrarTips = CBool(Lector.GetValue("OTHER", "MOSTRAR_TIPS"))
        .MostrarBindKeysSelection = CBool(Lector.GetValue("OTHER", "MOSTRAR_BIND_KEYS_SELECTION"))
        
        Debug.Print "Modo de Renderizado: " & IIf(.Aceleracion = 1, "Mixto (Hardware + Software)", "Hardware")
        Debug.Print "byMemory: " & .byMemory
        Debug.Print "bNoRes: " & .bNoRes
        Debug.Print "ProyectileEngine: " & .ProyectileEngine
        Debug.Print "PartyMembers: " & .PartyMembers
        Debug.Print "UsarSombras: " & .UsarSombras
        Debug.Print "UsarReflejos: " & .UsarReflejos
        Debug.Print "UsarAuras: " & .UsarAuras
        Debug.Print "ParticleEngine: " & .ParticleEngine
        Debug.Print "LimitarFPS: " & .LimiteFPS
        Debug.Print "bMusic: " & .bMusic
        Debug.Print "bSound: " & .bSound
        Debug.Print "MusicVolume: " & .MusicVolume
        Debug.Print "SoundVolume: " & .SoundVolume
        Debug.Print "bGuildNews: " & .bGuildNews
        Debug.Print "bGldMsgConsole: " & .bGldMsgConsole
        Debug.Print "bCantMsgs: " & .bCantMsgs
        Debug.Print "bDie: " & .bDie
        Debug.Print "bKill: " & .byMurderedLevel
        Debug.Print "bActive: " & .bActive
        Debug.Print "MostrarTips: " & .MostrarTips
        Debug.Print vbNullString
        
    End With
  
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
        Call Lector.ChangeValue("VIDEO", "RenderMode", .Aceleracion)
        Call Lector.ChangeValue("VIDEO", "DynamicMemory", .byMemory)
        Call Lector.ChangeValue("VIDEO", "DisableResolutionChange", IIf(.bNoRes, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "ParticleEngine", IIf(.ProyectileEngine, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "PartyMembers", IIf(.PartyMembers, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "Sombras", IIf(.UsarSombras, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "Reflejos", IIf(.UsarReflejos, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "Auras", IIf(.UsarAuras, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "ParticleEngine", IIf(.ParticleEngine, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "LimitarFPS", IIf(.LimiteFPS, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "HUD", IIf(.HUD, "True", "False"))
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "MUSICA", .bMusic)
        Call Lector.ChangeValue("AUDIO", "SONIDO", .bSound)
        Call Lector.ChangeValue("AUDIO", "AMBIENT", .bAmbient)
        Call Lector.ChangeValue("AUDIO", "VOLMUSICA", .MusicVolume)
        Call Lector.ChangeValue("AUDIO", "VOLAUDIO", .SoundVolume)
        Call Lector.ChangeValue("AUDIO", "VOLAMBIENT", .AmbientVol)
        
        ' GUILD
        Call Lector.ChangeValue("GUILD", "NEWS", IIf(.bGuildNews, "True", "False"))
        Call Lector.ChangeValue("GUILD", "MESSAGES", IIf(DialogosClanes.Activo, "True", "False"))
        Call Lector.ChangeValue("GUILD", "MAX_MESSAGES", CByte(DialogosClanes.CantidadDialogos))
        
        ' FRAGSHOOTER
        Call Lector.ChangeValue("FRAGSHOOTER", "DIE", IIf(.bDie, "True", "False"))
        Call Lector.ChangeValue("FRAGSHOOTER", "KILL", IIf(.bKill, "True", "False"))
        Call Lector.ChangeValue("FRAGSHOOTER", "MURDERED_LEVEL", CByte(.byMurderedLevel))
        Call Lector.ChangeValue("FRAGSHOOTER", "ACTIVE", IIf(.bActive, "True", "False"))
        
        ' OTHER
        ' Lo comento por que no tiene por que setearse aqui esto.
        ' Al menos no al hacer click en el boton Salir del formulario opciones (Recox)
        ' Call Lector.ChangeValue("OTHER", "MOSTRAR_TIPS", CBool(.MostrarTips))
    End With
    
    Call Lector.DumpFile(Carga.Path(Init) & CLIENT_FILE)
fileErr:

    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
    End If
End Sub

''
' Loads grh data using the new file format.
'

Public Sub LoadGrhData()
On Error GoTo ErrorHandler:

    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim Handle As Integer
    Dim fileVersion As Long
    Dim LaCabecera As tCabecera
    
    'Open files
    Handle = FreeFile()
    Open IniPath & "Graficos.ind" For Binary Access Read As Handle
    
        Get Handle, , LaCabecera
    
        Get Handle, , fileVersion
        
        Get Handle, , grhCount
        
        ReDim GrhData(0 To grhCount) As GrhData
        
        While Not EOF(Handle)
            Get Handle, , Grh
            
            With GrhData(Grh)
            
                '.active = True
                Get Handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ReDim .Frames(1 To .NumFrames)
                
                If .NumFrames > 1 Then
                
                    For Frame = 1 To .NumFrames
                        Get Handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                    Next Frame
                    
                    Get Handle, , .speed
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
                    
                    Get Handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                    
                End If
                
            End With
            
        Wend
    
    Close Handle
    
Exit Sub

ErrorHandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Graficos.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarCabezas()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Integer
    Dim LaCabecera As tCabecera
    
    N = FreeFile()
    Open Carga.Path(Script) & "Head.ind" For Binary Access Read As #N
    
        Get #N, , LaCabecera
    
        Get #N, , NumHeads   'cantidad de cabezas

        ReDim heads(0 To NumHeads) As tHead
            
        For i = 1 To NumHeads
            Get #N, , heads(i).Std
            Get #N, , heads(i).Texture
            Get #N, , heads(i).startX
            Get #N, , heads(i).startY
        Next i

    Close #N
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Head.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarCascos()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Integer
    Dim LaCabecera As tCabecera
    
    N = FreeFile()
    Open Carga.Path(Script) & "Helmet.ind" For Binary Access Read As #N
    
        Get #N, , LaCabecera
    
        Get #N, , NumCascos   'cantidad de cascos
             
        ReDim Cascos(0 To NumCascos) As tHead
             
        For i = 1 To NumCascos
            Get #N, , Cascos(i).Std
            Get #N, , Cascos(i).Texture
            Get #N, , Cascos(i).startX
            Get #N, , Cascos(i).startY
                
        Next i
         
    Close #N
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Helmet.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarCuerpos()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    Dim LaCabecera As tCabecera
    
    N = FreeFile()
    Open Carga.Path(Script) & "Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , LaCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            Call InitGrh(BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0)
            Call InitGrh(BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0)
            Call InitGrh(BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0)
            Call InitGrh(BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0)
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Personajes.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarFxs()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    Dim LaCabecera As tCabecera

    N = FreeFile
    Open Carga.Path(Script) & "FXs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , LaCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Fxs.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
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
        TipFile = FileToString(Carga.Path(Script) & "tips_" & Language & ".json")
    
    Set JsonTips = JSON.parse(TipFile)

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo" & "tips_" & Language & ".json no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            Call CloseClient
        End If
        
    End If
End Sub

Sub CargarAnimArmas()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim LaCabecera As tCabecera
    
    N = FreeFile
    Open Carga.Path(Script) & "Armas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , LaCabecera
    
    'num de cabezas
    Get #N, , NumWeaponAnims
    
    'Resize array
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    ReDim Weapons(1 To NumWeaponAnims) As tIndiceArmas
    
    For i = 1 To NumWeaponAnims
        Get #N, , Weapons(i)
        
        If Weapons(i).weapon(1) Then
        
            Call InitGrh(WeaponAnimData(i).WeaponWalk(1), Weapons(i).weapon(1), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(2), Weapons(i).weapon(2), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(3), Weapons(i).weapon(3), 0)
            Call InitGrh(WeaponAnimData(i).WeaponWalk(4), Weapons(i).weapon(4), 0)
        
        End If
    Next i
    
    Close #N

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Armas.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            Call CloseClient
        End If
        
    End If

End Sub


Public Sub CargarColores()
On Error GoTo errhandler:

    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Carga.Path(Script) & "colores.dat")
    
    Dim i As Long
    
    For i = 0 To 47 '48, 49 y 50 reservados para atacables, ciudadano y criminal
        ColoresPJ(i) = D3DColorXRGB(FileManager.GetValue(CStr(i), "R"), FileManager.GetValue(CStr(i), "G"), FileManager.GetValue(CStr(i), "B"))
    Next i
    
    '   Crimi
    ColoresPJ(50) = D3DColorXRGB(FileManager.GetValue("CR", "R"), FileManager.GetValue("CR", "G"), FileManager.GetValue("CR", "B"))

    '   Ciuda
    ColoresPJ(49) = D3DColorXRGB(FileManager.GetValue("CI", "R"), FileManager.GetValue("CI", "G"), FileManager.GetValue("CI", "B"))
    
    '   Atacable TODO: hay que implementar un color para los atacables y hacer que funcione.
    'ColoresPJ(48) = D3DColorXRGB(FileManager.GetValue("AT", "R"), FileManager.GetValue("AT", "G"), FileManager.GetValue("AT", "B"))
    
    For i = 51 To 56 'Colores reservados para la renderizacion de dano
        ColoresDano(i) = D3DColorXRGB(FileManager.GetValue(CStr(i), "R"), FileManager.GetValue(CStr(i), "G"), FileManager.GetValue(CStr(i), "B"))
    Next i
    
    Set FileManager = Nothing
    
errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo colores.dat no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            Call CloseClient
        End If
        
    End If
    
End Sub

Sub CargarAnimEscudos()
On Error GoTo errhandler:

    Dim N As Integer
    Dim i As Long
    Dim LaCabecera As tCabecera
    
    N = FreeFile
    Open Carga.Path(Script) & "Escudos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , LaCabecera
    
    'num de cabezas
    Get #N, , NumEscudosAnims
    
    'Resize array
    ReDim ShieldAnimData(1 To NumWeaponAnims) As ShieldAnimData
    ReDim Shields(1 To NumWeaponAnims) As tIndiceEscudos
    
    For i = 1 To NumEscudosAnims
        Get #N, , Shields(i)
        
        If Shields(i).shield(1) Then
        
            Call InitGrh(ShieldAnimData(i).ShieldWalk(1), Shields(i).shield(1), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(2), Shields(i).shield(2), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(3), Shields(i).shield(3), 0)
            Call InitGrh(ShieldAnimData(i).ShieldWalk(4), Shields(i).shield(4), 0)
        
        End If
    Next i
    
    Close #N

errhandler:
    
    If Err.number <> 0 Then
        
        If Err.number = 53 Then
            Call MsgBox("El archivo Escudos.ind no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
            Call CloseClient
        End If
        
    End If
    
End Sub

Public Sub CargarHechizos()
'********************************
'Author: Shak
'Last Modification:
'Cargamos los hechizos del juego. [Solo datos necesarios]
'********************************
On Error GoTo errorH

    Dim j As Long
    
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Carga.Path(Script) & "Hechizos.dat")

    NumHechizos = Val(FileManager.GetValue("INIT", "NumHechizos"))
 
    ReDim Hechizos(1 To NumHechizos) As tHechizos
    
    For j = 1 To NumHechizos
        
        With Hechizos(j)
            .Desc = FileManager.GetValue("HECHIZO" & j, "Desc")
            .PalabrasMagicas = FileManager.GetValue("HECHIZO" & j, "PalabrasMagicas")
            .Nombre = FileManager.GetValue("HECHIZO" & j, "Nombre")
            .SkillRequerido = Val(FileManager.GetValue("HECHIZO" & j, "MinSkill"))
         
            If j <> 38 And j <> 39 Then
                
                .EnergiaRequerida = Val(FileManager.GetValue("HECHIZO" & j, "StaRequerido"))
                 
                .HechiceroMsg = FileManager.GetValue("HECHIZO" & j, "HechizeroMsg")
                .ManaRequerida = Val(FileManager.GetValue("HECHIZO" & j, "ManaRequerido"))
             
                .PropioMsg = FileManager.GetValue("HECHIZO" & j, "PropioMsg")
                .TargetMsg = FileManager.GetValue("HECHIZO" & j, "TargetMsg")
                
            End If
            
        End With
        
    Next j
    
    Set FileManager = Nothing
    
Exit Sub
 
errorH:

    If Err.number <> 0 Then
        
        Select Case Err.number
            
            Case 9
                Call MsgBox("Error cargando el archivo Hechizos.dat (Hechizo " & j & "). Por favor, avise a los administradores enviandoles el archivo Errores.log que se encuentra en la carpeta del cliente.", , "Winter AO Resurrection")
                Call LogError(Err.number, Err.Description, "CargarHechizos")
            
            Case 53
                Call MsgBox("El archivo Hechizos.dat no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
        
        End Select
        
        Call CloseClient

    End If

End Sub

Sub CargarMapa(ByVal Map As Integer, ByVal Dir_Map As String)

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

    Dim i            As Long
    Dim j            As Long

    Dim LaCabecera   As tCabecera

    DoEvents
    
    fh = FreeFile
    Open Dir_Map For Binary Access Read As fh
    
    Get #fh, , LaCabecera
    
    Get #fh, , MH
    Get #fh, , MapSize
    
    Get #fh, , MapDat
    
    With MapSize
        ReDim MapData(.XMin To .XMax, .YMin To .YMax)
        ReDim L1(.XMin To .XMax, .YMin To .YMax)
    End With
    
    Get #fh, , L1
    
    With MH

        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs

            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i

        End If
        
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2

            For i = 1 To .NumeroLayers(2)
                Call InitGrh(MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).GrhIndex)
            Next i

        End If
        
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3

            For i = 1 To .NumeroLayers(3)
                Call InitGrh(MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).GrhIndex)
            Next i

        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4

            For i = 1 To .NumeroLayers(4)
                Call InitGrh(MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).GrhIndex)
            Next i

        End If
        
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers

            For i = 1 To .NumeroTriggers
                
                With Triggers(i)
                    MapData(.X, .Y).Trigger = .Trigger
                End With
                
            Next i

        End If
        
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
            
            For i = 1 To .NumeroParticulas

                With Particulas(i)
                    MapData(.X, .Y).Particle_Group_Index = General_Particle_Create(.Particula, .X, .Y)
                End With

            Next i

        End If
            
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
            
            'Aca metes la carga de las luces...
        End If
            
        If .NumeroOBJs > 0 Then
            ReDim Objetos(1 To .NumeroOBJs)
            Get #fh, , Objetos
            
            For i = 1 To .NumeroOBJs
                'Erase OBJs
                MapData(Objetos(i).X, Objetos(i).Y).ObjGrh.GrhIndex = 0
            Next i
            
        End If
            
        If .NumeroNPCs > 0 Then
            ReDim NPCs(1 To .NumeroNPCs)
            Get #fh, , NPCs
            
            For i = 1 To .NumeroNPCs
                MapData(NPCs(i).X, NPCs(i).Y).NPCIndex = NPCs(i).NPCIndex
            Next
                
        End If

        If .NumeroTE > 0 Then
            ReDim TEs(1 To .NumeroTE)
            Get #fh, , TEs

            For i = 1 To .NumeroTE
                
                With TEs(i)
                
                    MapData(.X, .Y).TileExit.Map = .DestM
                    MapData(.X, .Y).TileExit.X = .DestX
                    MapData(.X, .Y).TileExit.Y = .DestY
                
                End With
                
            Next i

        End If
        
    End With

    Close fh

    For j = MapSize.YMin To MapSize.YMax
        For i = MapSize.XMin To MapSize.XMax

            If L1(i, j) > 0 Then
                Call InitGrh(MapData(i, j).Graphic(1), L1(i, j))
            End If

        Next i
    Next j
    
    '*******************************
    'INFORMACION DEL MAPA
    '*******************************
    
    mapInfo.name = MapDat.map_name
    mapInfo.Music = MapDat.music_number
    mapInfo.Ambient = MapDat.Ambient
    mapInfo.Zona = MapDat.zone

    DeleteFile Dir_Map
ErrorHandler:
    
    If fh <> 0 Then Close fh
    
    If Err.number <> 0 Then
        'Call LogError(Err.number, Err.Description, "modCarga.CargarMapa")
        Call MsgBox("err: " & Err.number, "desc: " & Err.Description)
    End If

End Sub

Public Sub CargarConnectMaps()
'********************************
'Author: Lorwik
'Last Modification: 13/05/2020
'Cargamos los mapas del conectar renderizado
'********************************
On Error GoTo errorH
    Dim i As Byte
    
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(Carga.Path(Script) & "Maps.ini")
    
    NumConnectMap = Val(FileManager.GetValue("INIT", "NumMaps"))
    
    ReDim Preserve MapaConnect(NumConnectMap) As tMapaConnect
    
    For i = 1 To NumConnectMap
    
        MapaConnect(i).Map = Val(FileManager.GetValue("MAPA" & i, "Map"))
        MapaConnect(i).X = Val(FileManager.GetValue("MAPA" & i, "X"))
        MapaConnect(i).Y = Val(FileManager.GetValue("MAPA" & i, "Y"))
        
    Next i
    
    Set FileManager = Nothing
    
    Exit Sub
 
errorH:

    If Err.number <> 0 Then
        
        Select Case Err.number
            
            Case 9
                Call MsgBox("Error cargando el archivo de Mapas. Por favor, avise a los administradores enviandoles el archivo Errores.log que se encuentra en la carpeta del cliente.", , "Winter AO Resurrection")
                Call LogError(Err.number, Err.Description, "CargarHechizos")
            
            Case 53
                Call MsgBox("El archivo de configuracion de Mapas no existe. Por favor, reinstale el juego.", , "Winter AO Resurrection")
        
        End Select
        
        Call CloseClient

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
