Attribute VB_Name = "Game"
'Argentum Online 0.13.9.2

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
    Init
    Graficos
    Interfaces
    Skins
    sounds
    musica
    Mapas
    Lenguajes
    Fonts
End Enum

Public Type tSetupMods

    ' VIDEO
    Aceleracion As Byte
    byMemory    As Integer
    ProyectileEngine As Boolean
    PartyMembers As Boolean
    TonalidadPJ As Boolean
    UsarSombras As Boolean
    ParticleEngine As Boolean
    HUD As Boolean
    LimiteFPS As Boolean
    bNoRes      As Boolean
    FPSShow      As Boolean
    
    ' AUDIO
    bMusic    As E_SISTEMA_MUSICA
    bSound    As Byte
    bAmbient As Long
    Inverido As Byte
    bSoundEffects As Byte
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
Private Const CLIENT_FILE As String = "Config.ini"

Public Sub IniciarCabecera()

    With MiCabecera
        .Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
        .CRC = Rnd * 100
        .MagicWord = Rnd * 10
    End With
    
End Sub

Public Function Path(ByVal PathType As ePath) As String

    Select Case PathType
        
        Case ePath.Init
            Path = App.Path & "\Recursos\INIT\"
        
        Case ePath.Graficos
            Path = App.Path & "\Recursos\Graficos\"
        
        Case ePath.Skins
            Path = App.Path & "\Recursos\Graficos\Skins\"
            
        Case ePath.Interfaces
            Path = App.Path & "\Recursos\Graficos\Interfaces\"
            
        Case ePath.Fonts
            Path = App.Path & "\Recursos\Graficos\Fonts\"
            
        Case ePath.Lenguajes
            Path = App.Path & "\Recursos\Lenguajes\"
            
        Case ePath.Mapas
            Path = App.Path & "\Recursos\Mapas\"
            
        Case ePath.musica
            Path = App.Path & "\Recursos\MP3\"
            
        Case ePath.sounds
            Path = App.Path & "\Recursos\WAV\"
    
    End Select

End Function

Public Sub LeerConfiguracion()
    On Local Error GoTo fileErr:
    
    Call IniciarCabecera
    
    Set Lector = New clsIniManager
    Call Lector.Initialize(Game.Path(Init) & CLIENT_FILE)
    
    With ClientSetup
        ' VIDEO
        .Aceleracion = Lector.GetValue("VIDEO", "RenderMode")
        .byMemory = Lector.GetValue("VIDEO", "DynamicMemory")
        .bNoRes = CBool(Lector.GetValue("VIDEO", "DisableResolutionChange"))
        .ProyectileEngine = CBool(Lector.GetValue("VIDEO", "ProjectileEngine"))
        .PartyMembers = CBool(Lector.GetValue("VIDEO", "PartyMembers"))
        .TonalidadPJ = CBool(Lector.GetValue("VIDEO", "TonalidadPJ"))
        .UsarSombras = CBool(Lector.GetValue("VIDEO", "Sombras"))
        .ParticleEngine = CBool(Lector.GetValue("VIDEO", "ParticleEngine"))
        .LimiteFPS = CBool(Lector.GetValue("VIDEO", "LimitarFPS"))
        .HUD = CBool(Lector.GetValue("VIDEO", "HUD"))
        
        ' AUDIO
        .bMusic = CByte(Lector.GetValue("AUDIO", "MUSIC"))
        .bSound = CByte(Lector.GetValue("AUDIO", "SOUND"))
        .bAmbient = CByte(Lector.GetValue("AUDIO", "AMBIENT"))
        .AmbientVolume = CLng(Lector.GetValue("AUDIO", "AMBIENT"))
        .MusicVolume = CLng(Lector.GetValue("AUDIO", "MUSIC_VOLUME"))
        .SoundVolume = CLng(Lector.GetValue("AUDIO", "SOUND_VOLUME"))
        .Invertido = CByte(Lector.GetValue("AUDIO", "INVERTIDO"))
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
        Debug.Print "TonalidadPJ: " & .TonalidadPJ
        Debug.Print "UsarSombras: " & .UsarSombras
        Debug.Print "ParticleEngine: " & .ParticleEngine
        Debug.Print "LimitarFPS: " & .LimiteFPS
        Debug.Print "bMusic: " & .bMusic
        Debug.Print "bSound: " & .bSound
        Debug.Print "bSoundEffects: " & .bSoundEffects
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
    Call Lector.Initialize(Game.Path(Init) & CLIENT_FILE)

    With ClientSetup
        
        ' VIDEO
        Call Lector.ChangeValue("VIDEO", "RenderMode", .Aceleracion)
        Call Lector.ChangeValue("VIDEO", "DynamicMemory", .byMemory)
        Call Lector.ChangeValue("VIDEO", "DisableResolutionChange", IIf(.bNoRes, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "ParticleEngine", IIf(.ProyectileEngine, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "PartyMembers", IIf(.PartyMembers, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "TonalidadPJ", IIf(.TonalidadPJ, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "Sombras", IIf(.UsarSombras, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "ParticleEngine", IIf(.ParticleEngine, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "LimitarFPS", IIf(.LimiteFPS, "True", "False"))
        Call Lector.ChangeValue("VIDEO", "HUD", IIf(.HUD, "True", "False"))
        
        ' AUDIO
        Call Lector.ChangeValue("AUDIO", "MUSIC", IIf(.bMusic))
        Call Lector.ChangeValue("AUDIO", "SOUND", IIf(.bSound))
       Call Lector.ChangeValue("AUDIO", "AMBIENT", IIf(.bAmbient))
       Call Lector.ChangeValue("AUDIO", "MUSIC_VOLUME", Sound.AmbienteActual)
        Call Lector.ChangeValue("AUDIO", "MUSIC_VOLUME", Sound.MusicActual)
        Call Lector.ChangeValue("AUDIO", "SOUND_VOLUME", Sound.VolumenActual)
        
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
    
    Call Lector.DumpFile(Game.Path(Init) & CLIENT_FILE)
fileErr:

    If Err.number <> 0 Then
        MsgBox ("Ha ocurrido un error al guardar la configuracion del cliente. Error " & Err.number & " : " & Err.Description)
    End If
End Sub
