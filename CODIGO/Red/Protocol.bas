Attribute VB_Name = "Protocol_Handler"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
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

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martin Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Type tFont
    Red As Byte
    Green As Byte
    Blue As Byte
    bold As Boolean
    italic As Boolean
End Type

Private Enum ServerPacketID
    logged = 1                  ' LOGGED
    RemoveDialogs               ' QTDL
    RemoveCharDialog            ' QDL
    NavigateToggle              ' NAVEG
    Disconnect                  ' FINOK
    CommerceEnd                 ' FINCOMOK
    BankEnd                     ' FINBANOK
    CommerceInit                ' INITCOM
    BankInit                    ' INITBANCO
    CommerceChat
    UpdateSta                   ' ASS
    UpdateMana                  ' ASM
    UpdateHP                    ' ASH
    UpdateGold                  ' ASG
    UpdateBankGold
    UpdateExp                   ' ASE
    ChangeMap                   ' CM
    PosUpdate                   ' PU
    ChatOverHead                ' ||
    ConsoleMsg                  ' || - Beware!! its the same as above, but it was properly splitted
    ScreenMsg
    GuildChat                   ' |+
    ShowMessageBox              ' !!
    UserIndexInServer           ' IU
    UserCharIndexInServer       ' IP
    CharacterCreate             ' CC
    CharacterRemove             ' BP
    CharacterChangeNick
    CharacterMove               ' MP, +, * and _ '
    ForceCharMove
    CharacterChange             ' CP
    HeadingChange
    ObjectCreate                ' HO
    ObjectDelete                ' BO
    BlockPosition               ' BQ
    UserCommerceInit            ' INITCOMUSU
    UserCommerceEnd             ' FINCOMUSUOK
    UserOfferConfirm
    PlayMUSIC                    ' TM
    PlayWave                     ' TW
    guildList                    ' GL
    AreaChanged                  ' CA
    PauseToggle                  ' BKW
    ActualizarClima
    CreateFX                     ' CFX
    UpdateUserStats              ' EST
    ChangeInventorySlot          ' CSI
    ChangeBankSlot               ' SBO
    ChangeSpellSlot              ' SHS
    Atributes                    ' ATR
    InitTrabajo
    RestOK                       ' DOK
    ErrorMsg                     ' ERR
    Blind                        ' CEGU
    Dumb                         ' DUMB
    ShowSignal                   ' MCAR
    ChangeNPCInventorySlot       ' NPCI
    UpdateHungerAndThirst        ' EHYS
    Fame                         ' FAMA
    MiniStats                    ' MEST
    LevelUp                      ' SUNI
    AddForumMsg                  ' FMSG
    ShowForumForm                ' MFOR
    SetInvisible                 ' NOVER
    MeditateToggle               ' MEDOK
    BlindNoMore                  ' NSEGUE
    DumbNoMore                   ' NESTUP
    SendSkills                   ' SKILLS
    TrainerCreatureList          ' LSTCRI
    guildNews                    ' GUILDNE
    OfferDetails                 ' PEACEDE & ALLIEDE
    AlianceProposalsList         ' ALLIEPR
    PeaceProposalsList           ' PEACEPR
    CharacterInfo                ' CHRINFO
    GuildLeaderInfo              ' LEADERI
    GuildMemberInfo
    GuildDetails                 ' CLANDET
    ShowGuildFundationForm       ' SHOWFUN
    ParalizeOK                   ' PARADOK
    ShowUserRequest              ' PETICIO
    ChangeUserTradeSlot          ' COMUSUINV
    SendNight                    ' NOC
    Pong
    UpdateTagAndStatus
    BattleGs                     'Battlegrounds
    MostrarShop
    ActualizarGemasShop
    
    'GM =  messages
    SpawnList                    ' SPL
    ShowSOSForm                  ' MSOS
    ShowMOTDEditionForm          ' ZMOTD
    ShowGMPanelForm              ' ABPANEL
    UserNameList                 ' LISTUSU
    ShowDenounces
    RecordList
    RecordDetails
    
    ShowGuildAlign
    ShowPartyForm
    PeticionInvitarParty
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    StopWorking
    CancelOfferItem
    PlayAttackAnim
    FXtoMap
    EnviarPJUserAccount
    SearchList
    QuestDetails
    QuestListSend
    ActualizarNPCQuest
    CreateDamage                 ' CDMG
    UserInEvent
    DeletedChar
    EquitandoToggle
    InitCraftman
    EnviarListDeAmigos
    Proyectil
    CharParticle
    IniciarSubastaConsulta
    ConfirmarInstruccion
    SetSpeed
    AtaqueNPC
    MostrarPVP
    eBarFx
    ePrivilegios
End Enum

Public Enum FontTypeNames
    FONTTYPE_TALK = 0
    FONTTYPE_FIGHT = 1
    FONTTYPE_WARNING = 2
    FONTTYPE_INFO = 3
    FONTTYPE_INFOBOLD = 4
    FONTTYPE_EJECUCION = 5
    FONTTYPE_PARTY = 6
    FONTTYPE_VENENO = 7
    FONTTYPE_GUILD = 8
    FONTTYPE_SERVER = 9
    FONTTYPE_GUILDMSG = 10
    FONTTYPE_CONSEJO = 11
    FONTTYPE_CONSEJOCAOS = 12
    FONTTYPE_CONSEJOVesA = 13
    FONTTYPE_CONSEJOCAOSVesA = 14
    FONTTYPE_CENTINELA = 15
    FONTTYPE_GMMSG = 16
    FONTTYPE_GM = 17
    FONTTYPE_CITIZEN = 18
    FONTTYPE_CONSE = 19
    FONTTYPE_DIOS = 20
    FONTTYPE_CRIMINAL = 21
    FONTTYPE_EXP = 22
    FONTTYPE_PRIVADO = 23
    
End Enum

Public FontTypes(23) As tFont

Public Sub Connect(ByVal Modo As E_MODO)
    '*********************************************************************
    'Author: Jopi
    'Conexion al servidor mediante la API de Windows.
    '*********************************************************************
        
    'Evitamos enviar multiples peticiones de conexion al servidor.
    ModConectar.Conectando = False
        
    'Primero lo cerramos, para evitar errores.
    If frmMain.Client.State <> (sckClosed Or sckConnecting) Then
        frmMain.Client.CloseSck
        DoEvents
    End If
    
    EstadoLogin = Modo

    'Usamos la API de Windows
    Call frmMain.Client.Connect(CurServerIp, CurServerPort)

    'Vuelvo a activar el boton.
    ModConectar.Conectando = True
End Sub

''
' Initializes the fonts array

Public Sub InitFonts()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    With FontTypes(FontTypeNames.FONTTYPE_TALK)
        .Red = 204
        .Green = 255
        .Blue = 255
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_FIGHT)
        .Red = 255
        .Green = 102
        .Blue = 102
        .bold = 1
        .italic = 0
    End With

    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        .Red = 255
        .Green = 255
        .Blue = 102
        .bold = 1
        .italic = 0
    End With

    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        .Red = 255
        .Green = 204
        .Blue = 153
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
        .Red = 255
        .Green = 204
        .Blue = 153
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_EJECUCION)
        .Red = 255
        .Green = 0
        .Blue = 127
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PARTY)
        .Red = 252
        .Green = 203
        .Blue = 130
    End With

    With FontTypes(FontTypeNames.FONTTYPE_VENENO)
        .Red = 128
        .Green = 255
        .Blue = 0
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_GUILD)
        .Red = 205
        .Green = 101
        .Blue = 236
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_SERVER)
        .Red = 250
        .Green = 150
        .Blue = 237
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        .Red = 228
        .Green = 199
        .Blue = 27
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CONSEJO)
        .Red = 130
        .Green = 130
        .Blue = 255
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOS)
        .Red = 255
        .Green = 60
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOVesA)
        .Green = 200
        .Blue = 255
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
        .Red = 255
        .Green = 50
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CENTINELA)
        .Red = 240
        .Green = 230
        .Blue = 140
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_GMMSG)
        .Red = 255
        .Green = 255
        .Blue = 255
        .italic = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_GM)
        .Red = 30
        .Green = 255
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CITIZEN)
        .Red = 78
        .Green = 78
        .Blue = 252
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_CONSE)
        .Red = 30
        .Green = 150
        .Blue = 30
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_DIOS)
        .Red = 250
        .Green = 250
        .Blue = 150
        .bold = 1
    End With

    With FontTypes(FontTypeNames.FONTTYPE_CRIMINAL)
        .Red = 224
        .Green = 52
        .Blue = 17
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_EXP)
        .Red = 0
        .Green = 162
        .Blue = 232
        .bold = 1
    End With
    
    With FontTypes(FontTypeNames.FONTTYPE_PRIVADO)
        .Red = 182
        .Green = 226
        .Blue = 29
    End With
    
End Sub

''
' Handles incoming data.

Public Sub HandleIncomingData()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
On Error Resume Next

    Dim Packet As Long: Packet = CLng(incomingData.PeekByte())
    
    'Debug.Print Packet
    
    Select Case Packet
            
        Case ServerPacketID.logged                  ' LOGGED
            Call HandleLogged
        
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
        
        Case ServerPacketID.NavigateToggle          ' NAVEG
            Call HandleNavigateToggle
        
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             ' FINCOMOK
            Call HandleCommerceEnd
            
        Case ServerPacketID.CommerceChat
            Call HandleCommerceChat
        
        Case ServerPacketID.BankEnd                 ' FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            ' INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                ' INITBANCO
            Call HandleBankInit
        
        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              ' ASG
            Call HandleUpdateGold
            
        Case ServerPacketID.UpdateBankGold
            Call HandleUpdateBankGold

        Case ServerPacketID.UpdateExp               ' ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
            
        Case ServerPacketID.ScreenMsg
            Call HandleScreenMessage
        
        Case ServerPacketID.GuildChat               ' |+
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterChangeNick
            Call HandleCharacterChangeNick
            
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
            
        Case ServerPacketID.HeadingChange
            Call HandleHeadingChange
            
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
            
        Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
            Call HandleUserCommerceEnd
            
        Case ServerPacketID.UserOfferConfirm
            Call HandleUserOfferConfirm
        
        Case ServerPacketID.PlayMUSIC                ' TM
            Call HandlePlayMUSIC
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave
        
        Case ServerPacketID.guildList               ' GL
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.ActualizarClima
            Call HandleActualizarClima
        
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats

        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          ' SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes               ' ATR
            Call HandleAtributes
        
        Case ServerPacketID.InitTrabajo
            Call HandleInitTrabajo
            
        Case ServerPacketID.InitCraftman
            Call HandleInitCraftman
        
        Case ServerPacketID.RestOK                  ' DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind                   ' CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal              ' MCAR
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.Fame                    ' FAMA
            Call HandleFame
        
        Case ServerPacketID.MiniStats               ' MEST
            Call HandleMiniStats
            
        Case ServerPacketID.LevelUp                 ' SUNI
            Call HandleLevelUp
        
        Case ServerPacketID.AddForumMsg             ' FMSG
            Call HandleAddForumMessage
        
        Case ServerPacketID.ShowForumForm           ' MFOR
            Call HandleShowForumForm
        
        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible
        
        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             ' NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              ' NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              ' SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     ' LSTCRI
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.guildNews               ' GUILDNE
            Call HandleGuildNews
        
        Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList      ' PEACEPR
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharacterInfo           ' CHRINFO
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo         ' LEADERI
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails            ' CLANDET
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK
        
        Case ServerPacketID.ShowUserRequest         ' PETICIO
            Call HandleShowUserRequest

        Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
            Call HandleChangeUserTradeSlot
            
        Case ServerPacketID.SendNight               ' NOC
            Call HandleSendNight
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
        
        Case ServerPacketID.GuildMemberInfo
            Call HandleGuildMemberInfo
            
        Case ServerPacketID.PlayAttackAnim
            Call HandleAttackAnim
            
        Case ServerPacketID.FXtoMap
            Call HandleFXtoMap
        
        Case ServerPacketID.EnviarPJUserAccount
            Call HandleEnviarPJUserAccount
            
        Case ServerPacketID.SearchList              '/BUSCAR
            Call HandleSearchList

        Case ServerPacketID.QuestDetails
            Call HandleQuestDetails

        Case ServerPacketID.QuestListSend
            Call HandleQuestListSend
            
        Case ServerPacketID.ActualizarNPCQuest
            Call HandleActualizarNPCQuest

        Case ServerPacketID.ShowGuildAlign
            Call HandleShowGuildAlign
        
        Case ServerPacketID.ShowPartyForm
            Call HandleShowPartyForm
            
        Case ServerPacketID.PeticionInvitarParty
            Call HandlePeticionInvitarParty
        
        Case ServerPacketID.UpdateStrenghtAndDexterity
            Call HandleUpdateStrenghtAndDexterity
            
        Case ServerPacketID.UpdateStrenght
            Call HandleUpdateStrenght
            
        Case ServerPacketID.UpdateDexterity
            Call HandleUpdateDexterity
            
        Case ServerPacketID.AddSlots
            Call HandleAddSlots

        Case ServerPacketID.MultiMessage
            Call HandleMultiMessage
        
        Case ServerPacketID.StopWorking
            Call HandleStopWorking
            
        Case ServerPacketID.CancelOfferItem
            Call HandleCancelOfferItem
            
        Case ServerPacketID.CreateDamage            ' CDMG
            Call HandleCreateDamage
    
        Case ServerPacketID.UserInEvent
            Call HandleUserInEvent

        Case ServerPacketID.DeletedChar             ' BORRA USUARIO
            Call HandleDeletedChar

        Case ServerPacketID.EquitandoToggle         'Para las monturas
            Call HandleEquitandoToggle
        
        Case ServerPacketID.EnviarListDeAmigos
            Call HandleEnviarListDeAmigos
            
        Case ServerPacketID.Proyectil
            Call HandleProyectil
            
        Case ServerPacketID.CharParticle
            Call HandleCharParticle
            
        Case ServerPacketID.IniciarSubastaConsulta
            Call HandleIniciarSubasta
            
        Case ServerPacketID.ConfirmarInstruccion
            Call HandleConfirmarInstruccion

        Case ServerPacketID.AtaqueNPC
            Call HandleAtaqueNPC
            
        Case ServerPacketID.BattleGs
            Call HandleBattlegrounds
            
        Case ServerPacketID.MostrarShop
            Call HandleMostrarShop
            
        Case ServerPacketID.ActualizarGemasShop
            Call HandleActualizarGemasShop
            
        Case ServerPacketID.MostrarPVP
            Call HandleMostrarPVP
            
        Case ServerPacketID.eBarFx
            Call HandleBarFx
            
        Case ServerPacketID.ePrivilegios
            Call HandlePrivilegios

        '*******************
        'GM messages
        '*******************
        Case ServerPacketID.SpawnList               ' SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
            Call HandleShowSOSForm
            
        Case ServerPacketID.ShowDenounces
            Call HandleShowDenounces
            
        Case ServerPacketID.RecordDetails
            Call HandleRecordDetails
            
        Case ServerPacketID.RecordList
            Call HandleRecordList
            
        Case ServerPacketID.ShowMOTDEditionForm     ' ZMOTD
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            ' LISTUSU
            Call HandleUserNameList
            
        Case ServerPacketID.SetSpeed
            Call HandleSetSpeed
            
        Case Else
            'ERROR : Abort!
            Exit Sub
    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
        Err.Clear
        Call HandleIncomingData
    End If
End Sub

Public Sub HandleMultiMessage()

    '***************************************************
    'Author: Unknown
    'Last Modification: 11/16/2010
    ' 09/28/2010: C4b3z0n - Ahora se le saco la "," a los minutos de distancia del /hogar, ya que a veces quedaba "12,5 minutos y 30segundos"
    ' 09/21/2010: C4b3z0n - Now the fragshooter operates taking the screen after the change of killed charindex to ghost only if target charindex is visible to the client, else it will take screenshot like before.
    ' 11/16/2010: Amraphen - Recoded how the FragShooter works.
    ' 04/12/2019: jopiortiz - Carga de mensajes desde JSON.
    '***************************************************
    Dim BodyPart As Byte

    Dim Dano As Integer

    Dim SpellIndex As Integer

    Dim Nombre     As String
    
    With incomingData
        Call .ReadByte
    
        Select Case .ReadByte

            Case eMessages.NPCSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("TEXTO"), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(1), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(2), _
                        JsonLanguage.item("MENSAJE_CRIATURA_FALLA_GOLPE").item("COLOR").item(3), _
                        True, False, True)
        
            Case eMessages.NPCKillUser
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_CRIATURA_MATADO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.BlockedWithShieldUser
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.BlockedWithShieldOther
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.UserSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_FALLADO_GOLPE").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.SafeModeOn
                Call frmMain.ControlSM(eSMType.sSafemode, True)
        
            Case eMessages.SafeModeOff
                Call frmMain.ControlSM(eSMType.sSafemode, False)
        
            Case eMessages.CombatSafeOff
                Call frmMain.ControlSM(eSMType.sCombatMode, False)
         
            Case eMessages.CombatSafeOn
                Call frmMain.ControlSM(eSMType.sCombatMode, True)
        
            Case eMessages.NobilityLost
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_PIERDE_NOBLEZA").item("COLOR").item(3), _
                    False, False, True)
        
            Case eMessages.CantUseWhileMeditating
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_USAR_MEDITANDO").item("COLOR").item(3), _
                    False, False, True)
        
            Case eMessages.NPCHitUser

                Select Case incomingData.ReadByte()

                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("TEXTO") & CStr(incomingData.ReadInteger()) & "!!", _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("TEXTO") & CStr(incomingData.ReadInteger() & "!!"), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_GOLPE_TORSO").item("COLOR").item(3), _
                            True, False, True)

                End Select
        
            Case eMessages.UserHitNPC
                Dim MsgHitNpc As String
                    MsgHitNpc = JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("TEXTO")
                    MsgHitNpc = Replace$(MsgHitNpc, "VAR_DANO", CStr(incomingData.ReadLong()))
                    
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    MsgHitNpc, _
                    JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_DAMAGE_NPC").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.UserAttackedSwing
                Call AddtoRichTextBox(frmMain.RecTxt, _
                    charlist(incomingData.ReadInteger()).Nombre & JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("TEXTO"), _
                    JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_ATAQUE_FALLO").item("COLOR").item(3), _
                    True, False, True)
        
            Case eMessages.UserHittedByUser

                Dim AttackerName As String
            
                AttackerName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                BodyPart = incomingData.ReadByte()
                Dano = incomingData.ReadInteger()
            
                Select Case BodyPart

                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                        AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(1), _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(2), _
                        JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ").item("COLOR").item(3), _
                        True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            AttackerName & JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_RECIVE_IMPACTO_TORSO").item("COLOR").item(3), _
                            True, False, True)

                End Select
        
            Case eMessages.UserHittedUser

                Dim VictimName As String
            
                VictimName = GetRawName(charlist(incomingData.ReadInteger()).Nombre)
                BodyPart = incomingData.ReadByte()
                Dano = incomingData.ReadInteger()
            
                Select Case BodyPart

                    Case ePartesCuerpo.bCabeza
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_CABEZA").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoIzquierdo
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bBrazoDerecho
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaIzquierda
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bPiernaDerecha
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER").item("COLOR").item(3), _
                            True, False, True)
                
                    Case ePartesCuerpo.bTorso
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_1").item("TEXTO") & VictimName & JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("TEXTO") & Dano & MENSAJE_2, _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PRODUCE_IMPACTO_TORSO").item("COLOR").item(3), _
                            True, False, True)

                End Select
        
            Case eMessages.WorkRequestTarget
                UsingSkill = incomingData.ReadByte()
            
                frmMain.MousePointer = 2 'vbCrosshair
            
                Select Case UsingSkill

                    Case Magia
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MAGIA").item("COLOR").item(3))
                
                    Case pesca
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PESCA").item("COLOR").item(3))
                
                    Case Robar
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_ROBAR").item("COLOR").item(3))
                
                    Case Talar
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_TALAR").item("COLOR").item(3))
                
                    Case Mineria
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_MINERIA").item("COLOR").item(3))
                
                    Case FundirMetal
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_FUNDIRMETAL").item("COLOR").item(3))
                
                    Case Proyectiles
                        Call AddtoRichTextBox(frmMain.RecTxt, _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("TEXTO"), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_TRABAJO_PROYECTILES").item("COLOR").item(3))

                End Select

            Case eMessages.HaveKilledUser

                Dim KilledUser As Integer

                Dim EXP        As Long
                
                Dim MensajeExp As String
            
                KilledUser = .ReadInteger
                EXP = .ReadLong
            
                Call ShowConsoleMsg( _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("TEXTO") & charlist(KilledUser).Nombre & MENSAJE_22, _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(1), _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(2), _
                    JsonLanguage.item("MENSAJE_HAS_MATADO_A").item("COLOR").item(3), _
                    True, False)
                
                ' Para mejor lectura
                MensajeExp = JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("TEXTO") 'String original
                MensajeExp = Replace$(MensajeExp, "VAR_EXP_GANADA", EXP) 'Parte a reemplazar
                
                Call ShowConsoleMsg(MensajeExp, _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_HAS_GANADO_EXP").item("COLOR").item(3), _
                                    True, False)
            
                'Sacamos un screenshot si esta activado el FragShooter:
                If ClientSetup.bKill And ClientSetup.bActive Then
                    If EXP \ 2 > ClientSetup.byMurderedLevel Then
                        FragShooterNickname = charlist(KilledUser).Nombre
                        FragShooterKilledSomeone = True
                    
                        FragShooterCapturePending = True

                    End If

                End If
            
            Case eMessages.UserKill

                Dim KillerUser As Integer
            
                KillerUser = .ReadInteger
            
                Call ShowConsoleMsg(charlist(KillerUser).Nombre & JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("TEXTO"), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_TE_HA_MATADO").item("COLOR").item(3), _
                                    True, False)
            
                'Sacamos un screenshot si esta activado el FragShooter:
                If ClientSetup.bDie And ClientSetup.bActive Then
                    FragShooterNickname = charlist(KillerUser).Nombre
                    FragShooterKilledSomeone = False
                
                    FragShooterCapturePending = True

                End If
                
            Case eMessages.EarnExp
                'Dim MENSAJE_HAS_GANADO_EXP As String
                '    MENSAJE_HAS_GANADO_EXP = JsonLanguage.Item("MENSAJE_HAS_GANADO_EXP").Item("TEXTO")
                '    MENSAJE_HAS_GANADO_EXP = Replace$(MENSAJE_HAS_GANADO_EXP, "VAR_EXP_GANADA", .ReadLong)
                    
                'Call ShowConsoleMsg(MENSAJE_HAS_GANADO_EXP, _
                '                    JsonLanguage.Item("MENSAJE_HAS_GANADO_EXP").Item("COLOR").Item(1), _
                '                    JsonLanguage.Item("MENSAJE_HAS_GANADO_EXP").Item("COLOR").Item(2), _
                '                    JsonLanguage.Item("MENSAJE_HAS_GANADO_EXP").Item("COLOR").Item(3), _
                '                    True, False)
        
                   
            Case eMessages.FinishHome
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_HOGAR").item("TEXTO"), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_HOGAR").item("COLOR").item(3))
            
            Case eMessages.UserMuerto
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(2), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(1), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(2), _
                                    JsonLanguage.item("MENSAJE_USER_MUERTO").item("COLOR").item(3))
        
            Case eMessages.NpcInmune
                Call ShowConsoleMsg(JsonLanguage.item("NPC_INMUNE").item("TEXTO"), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(1), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(2), _
                                    JsonLanguage.item("NPC_INMUNE").item("COLOR").item(3))

        End Select

    End With

End Sub

''
' Handles the DeletedChar message.

Private Sub HandleDeletedChar()
'***************************************************
'Author: Lucas Recoaro (Recox)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    Call MostrarMensaje(JsonLanguage.item("BORRAR_PJ").item("TEXTO"))

    'Close connection
    'Call CloseConnectionAndResetAllInfo
End Sub

' Handles the EquitandoToggle message.
Private Sub HandleEquitandoToggle()
'***************************************************
'Author: Lorwik
'Last Modification: 06/04/2020
'06/04/2020: FrankoH298 - Recibimos el contador para volver a equiparnos la montura.
'23/10/2020: Lorwik - Ahora recibe la velocidad
'***************************************************

    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserEquitando = Not CurrentUser.UserEquitando
End Sub

Private Sub HandleLogged()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    ' Variable initialization
    CurrentUser.UserClase = incomingData.ReadByte

    EngineRun = True
    Nombres = True
    
    'Set connected state
    Call SetConnected
    
    If bShowTutorial Then
        Call frmTutorial.Show(vbModeless)
    End If
    
    'Show tip
    If ClientSetup.MostrarTips = True Then
        frmtip.Visible = True
    End If

    'Show Keyboard configuration
    If ClientSetup.MostrarBindKeysSelection = True Then
        Call frmKeysConfigurationSelect.Show(vbModeless, frmMain)
        Call frmKeysConfigurationSelect.SetFocus
    End If
    
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check if the packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserNavegando = Not CurrentUser.UserNavegando
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    
    'Remove packet ID
    Call incomingData.ReadByte

    Call ResetAllInfo(False)
    frmMain.Visible = False
    Call MostrarCuenta(True)
    
    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        If ClientSetup.bMusic <> CONST_DESHABILITADA Then
            Sound.NextMusic = MUS_VolverInicio
            Sound.Fading = 200
        End If
    End If
    
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvComNpc = Nothing
    
    'Hide form
    Unload frmComerciar
    
    'Reset vars
    Comerciando = False
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = Nothing
    Set InvBanco(1) = Nothing
    
    Unload frmBancoObj
    Comerciando = False
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/01/20
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = New clsGraphicalInventory
    Set InvComNpc = New clsGraphicalInventory
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciar.picInvUser, MAX_INVENTORY_SLOTS, , , , , , , , True)
    Call InvComNpc.Initialize(DirectD3D8, frmComerciar.picInvNpc, MAX_NPC_INVENTORY_SLOTS)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ObjIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .ObjIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .valor(i), .ItemName(i), .NoUsa(i))
            End With
        End If
    Next i
    
    ' Fill Npc inventory
    For i = 1 To MAX_NPC_INVENTORY_SLOTS
        If NPCInventory(i).ObjIndex <> 0 Then
            With NPCInventory(i)
                Call InvComNpc.SetItem(i, .ObjIndex, _
                .Amount, 0, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .valor, .name, .NoUsa)
            End With
        End If
    Next i
    
    'Set state and show form
    frmComerciar.Show , frmMain
    
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    Dim BankGold As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvBanco(0) = New clsGraphicalInventory
    Set InvBanco(1) = New clsGraphicalInventory
    
    BankGold = incomingData.ReadLong
    Call InvBanco(0).Initialize(DirectD3D8, frmBancoObj.PicBancoInv, MAX_BANCOINVENTORY_SLOTS)
    Call InvBanco(1).Initialize(DirectD3D8, frmBancoObj.PicInv, MAX_INVENTORY_SLOTS, , , , , , , , True)
    
    For i = 1 To MAX_INVENTORY_SLOTS
        With Inventario
            Call InvBanco(1).SetItem(i, .ObjIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .valor(i), .ItemName(i), .NoUsa(i))
        End With
    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        With UserBancoInventory(i)
            Call InvBanco(0).SetItem(i, .ObjIndex, _
                .Amount, .Equipped, .GrhIndex, _
                .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, _
                .valor, .name, .NoUsa)
        End With
    Next i
    
    'Set state and show form
    Comerciando = True
    
    frmBancoObj.lblUserGld.Caption = BankGold
    
    frmBancoObj.Show , frmMain

End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    TradingUserName = incomingData.ReadASCIIString
    
    Set InvComUsu = New clsGraphicalInventory
    Set InvOfferComUsu(0) = New clsGraphicalInventory
    Set InvOfferComUsu(1) = New clsGraphicalInventory
    Set InvOroComUsu(0) = New clsGraphicalInventory
    Set InvOroComUsu(1) = New clsGraphicalInventory
    Set InvOroComUsu(2) = New clsGraphicalInventory
    
    ' Initialize commerce inventories
    Call InvComUsu.Initialize(DirectD3D8, frmComerciarUsu.picInvComercio, MAX_INVENTORY_SLOTS)
    Call InvOfferComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaProp, INV_OFFER_SLOTS)
    Call InvOfferComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOfertaOtro, INV_OFFER_SLOTS)
    Call InvOroComUsu(0).Initialize(DirectD3D8, frmComerciarUsu.picInvOroProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(1).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaProp, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)
    Call InvOroComUsu(2).Initialize(DirectD3D8, frmComerciarUsu.picInvOroOfertaOtro, INV_GOLD_SLOTS, , TilePixelWidth * 2, TilePixelHeight, TilePixelWidth / 2)

    'Fill user inventory
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ObjIndex(i) <> 0 Then
            With Inventario
                Call InvComUsu.SetItem(i, .ObjIndex(i), _
                .Amount(i), .Equipped(i), .GrhIndex(i), _
                .OBJType(i), .MaxHit(i), .MinHit(i), .MaxDef(i), .MinDef(i), _
                .valor(i), .ItemName(i), .NoUsa(i))
            End With
        End If
    Next i

    ' Inventarios de oro
    Call InvOroComUsu(0).SetItem(1, ORO_INDEX, CurrentUser.UserGLD, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(1).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")
    Call InvOroComUsu(2).SetItem(1, ORO_INDEX, 0, 0, ORO_GRH, 0, 0, 0, 0, 0, 0, "Oro")


    'Set state and show form
    Comerciando = True
    Call frmComerciarUsu.Show(vbModeless, frmMain)
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Set InvComUsu = Nothing
    Set InvOroComUsu(0) = Nothing
    Set InvOroComUsu(1) = Nothing
    Set InvOroComUsu(2) = Nothing
    Set InvOfferComUsu(0) = Nothing
    Set InvOfferComUsu(1) = Nothing
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False
End Sub

''
' Handles the UserOfferConfirm message.
Private Sub HandleUserOfferConfirm()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    With frmComerciarUsu
        ' Now he can accept the offer or reject it
        .HabilitarAceptarRechazar True
        
        .PrintCommerceMsg TradingUserName & JsonLanguage.item("MENSAJE_COMM_OFERTA_ACEPTA").item("TEXTO"), FontTypeNames.FONTTYPE_CONSE
    End With
    
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    CurrentUser.UserMinSTA = incomingData.ReadInteger()
    
    frmMain.lblEnergia = CurrentUser.UserMinSTA & "/" & CurrentUser.UserMaxSTA
    
    frmMain.shpEnergia.Width = (((CurrentUser.UserMinSTA / 100) / (CurrentUser.UserMaxSTA / 100)) * 92)
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    CurrentUser.UserMinMAN = incomingData.ReadInteger()
    
    frmMain.lblMana = CurrentUser.UserMinMAN & "/" & CurrentUser.UserMaxMAN
    
    If CurrentUser.UserMaxMAN > 0 Then _
        frmMain.shpMana.Width = (((CurrentUser.UserMinMAN / 100) / (CurrentUser.UserMaxMAN / 100)) * 92)
        
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    CurrentUser.UserMinHP = incomingData.ReadInteger()
    
    frmMain.lblVida = CurrentUser.UserMinHP & "/" & CurrentUser.UserMaxHP

    frmMain.shpVida.Width = (((CurrentUser.UserMinHP / 100) / (CurrentUser.UserMaxHP / 100)) * 185)
    
    'Is the user alive??
    If CurrentUser.UserMinHP = 0 Then
        CurrentUser.UserEstado = 1
    
        CurrentUser.UserEquitando = 0
    Else
        CurrentUser.UserEstado = 0
    End If
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
'***************************************************
'Autor: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 09/21/10
'Last Modified By: C4b3z0n
'- 08/14/07: Tavo - Added GldLbl color variation depending on User Gold and Level
'- 09/21/10: C4b3z0n - Modified color change of gold ONLY if the player's level is greater than 12 (NOT newbie).
'***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    CurrentUser.UserGLD = incomingData.ReadLong()
    
    Call frmMain.SetGoldColor

    frmMain.GldLbl.Caption = CurrentUser.UserGLD
End Sub

''
' Handles the UpdateBankGold message.

Private Sub HandleUpdateBankGold()
'***************************************************
'Autor: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmBancoObj.lblUserGld.Caption = incomingData.ReadLong
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    CurrentUser.UserExp = incomingData.ReadLong()

    frmMain.UpdateProgressExperienceLevelBar (CurrentUser.UserExp)
End Sub

''
' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenghtAndDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************

    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    CurrentUser.UserFuerza = incomingData.ReadByte
    CurrentUser.UserAgilidad = incomingData.ReadByte
    
    frmMain.lblStrg.Caption = CurrentUser.UserFuerza
    frmMain.lblDext.Caption = CurrentUser.UserAgilidad
    
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateStrenght()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************

    'Check packet is complete
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    CurrentUser.UserFuerza = incomingData.ReadByte
    
    frmMain.lblStrg.Caption = CurrentUser.UserFuerza
    
End Sub

' Handles the UpdateStrenghtAndDexterity message.

Private Sub HandleUpdateDexterity()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'***************************************************
    'Check packet is complete
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    CurrentUser.UserAgilidad = incomingData.ReadByte
   
    frmMain.lblDext.Caption = CurrentUser.UserAgilidad

End Sub

''
' Handles the ChangeMap message.
Private Sub HandleChangeMap()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Call Err.Raise(incomingData.NotEnoughDataErrCode)
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserMap = incomingData.ReadInteger()
    
    'TODO: Once on-the-fly editor is implemented check for map version before loading....
    'For now we just drop it
    Call incomingData.ReadInteger
      
    Call SwitchMap(CurrentUser.UserMap)
    
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Map_RemoveOldUser
    
    '// Seteamos la Posicion en el Mapa
    Call Char_MapPosSet(incomingData.ReadInteger(), incomingData.ReadInteger())

    'Update pos label
    Call Char_UserPos
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim CharIndex As Integer
    Dim NoConsole As Boolean
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    chat = buffer.ReadASCIIString()
    CharIndex = buffer.ReadInteger()
    NoConsole = buffer.ReadBoolean()
    
    Red = buffer.ReadByte()
    Green = buffer.ReadByte()
    Blue = buffer.ReadByte()
    
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If Char_Check(CharIndex) Then
        Call Dialogos.CreateDialog(Trim$(chat), CharIndex, RGB(Red, Green, Blue))

        'Aqui escribimos el texto que aparece sobre la cabeza en la consola.
        If NoConsole = False Then Call Protocol_Write.WriteChatOverHeadInConsole(CharIndex, chat, Red, Green, Blue)
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    chat = buffer.ReadASCIIString()
    FontIndex = buffer.ReadByte()

    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
        
        ' Para no perder el foco cuando chatea por party
        If FontIndex = FontTypeNames.FONTTYPE_PARTY Then
            If MirandoParty Then frmParty.SendTxt.SetFocus
        End If
    End If
'    Call checkText(chat)
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 04/07/08 (NicoNZ)
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    chat = buffer.ReadASCIIString()
    
    If Not DialogosClanes.Activo Then
        If InStr(1, chat, "~") Then
            str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
            
            Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichTextBox(frmMain.RecTxt, chat, .Red, .Green, .Blue, .bold, .italic)
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(ReadField(1, chat, 126))
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleCommerceChat()
'***************************************************
'Author: ZaMa
'Last Modification: 03/12/2009
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String
    Dim FontIndex As Integer
    Dim str As String
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    
    chat = buffer.ReadASCIIString()
    FontIndex = buffer.ReadByte()
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
            If Val(str) > 255 Then
                Red = 255
            Else
                Red = Val(str)
            End If
            
            str = ReadField(3, chat, 126)
            If Val(str) > 255 Then
                Green = 255
            Else
                Green = Val(str)
            End If
            
            str = ReadField(4, chat, 126)
            If Val(str) > 255 Then
                Blue = 255
            Else
                Blue = Val(str)
            End If
            
        Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, Left$(chat, InStr(1, chat, "~") - 1), Red, Green, Blue, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmComerciarUsu.CommerceConsole, chat, .Red, .Green, .Blue, .bold, .italic)
        End With
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call MostrarMensaje(buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserIndex = incomingData.ReadInteger()
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Char_UserIndexSet(incomingData.ReadInteger())
                     
    'Update pos label
    Call Char_UserPos
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim Body As Integer
    Dim Head As Integer
    Dim Heading As E_Heading
    Dim x As Integer
    Dim y As Integer
    Dim weapon As Integer
    Dim shield As Integer
    Dim helmet As Integer
    Dim Ataque As Integer
    Dim privs As Integer
    Dim NickColor As Byte
    Dim AuraAnim As Long
    Dim AuraColor As Long
    
    CharIndex = buffer.ReadInteger()
    Body = buffer.ReadInteger()
    Head = buffer.ReadInteger()
    Heading = buffer.ReadByte()
    x = buffer.ReadInteger()
    y = buffer.ReadInteger()
    weapon = buffer.ReadInteger()
    shield = buffer.ReadInteger()
    helmet = buffer.ReadInteger()
    Ataque = buffer.ReadInteger()

    With charlist(CharIndex)
        Call Char_SetFx(CharIndex, buffer.ReadInteger(), buffer.ReadInteger())

        .Nombre = buffer.ReadASCIIString()
        .Clan = mid$(.Nombre, getTagPosition(.Nombre))
        NickColor = buffer.ReadByte()

        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        If NickColor = 8 Then .WorldBoss = True
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        privs = buffer.ReadByte()
        
        If privs <> 0 Then
            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil
            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil
            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster
            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0
        End If
        
        AuraAnim = buffer.ReadLong()
        AuraColor = buffer.ReadLong()
        
        .NoShadow = buffer.ReadByte()
        .EstadoQuest = buffer.ReadByte()
    End With
    
    Call Char_Make(CharIndex, Body, Head, Heading, x, y, weapon, shield, helmet, Ataque, AuraAnim, AuraColor)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleCharacterChangeNick()
'***************************************************
'Author: Budi
'Last Modification: 07/23/09
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
    Dim CharIndex As Integer
    CharIndex = incomingData.ReadInteger
    
    Call Char_SetName(CharIndex, incomingData.ReadASCIIString)
    
End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim Desaparece As Boolean
    
    CharIndex = incomingData.ReadInteger()

    Call Char_Erase(CharIndex)
    
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim x As Integer
    Dim y As Integer
    
    CharIndex = incomingData.ReadInteger()
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()

    With charlist(CharIndex)
        If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
            .FxIndex = 0
        End If
        
        ' Play steps sounds if the user is not an admin of any kind

        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)
        End If

    End With
    
    Call Char_MovebyPos(CharIndex, x, y)
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte
    
    Direccion = incomingData.ReadByte()

    Call Char_MovebyHead(UserCharIndex, Direccion)
    Call Char_MoveScreen(Direccion)
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 21/09/2010 - C4b3z0n
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'21/09/2010: C4b3z0n - Added code for FragShooter. If its waiting for the death of certain UserIndex, and it dies, then the capture of the screen will occur.
'***************************************************
    If incomingData.length < 18 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim Heading As Byte
    
    CharIndex = incomingData.ReadInteger()
    
    Heading = incomingData.ReadByte()
    charlist(CharIndex).Heading = Heading
    
    '// Char Body
    Call Char_SetBody(CharIndex, incomingData.ReadInteger())

    '// Char Head
    Call Char_SetHead(CharIndex, incomingData.ReadInteger)
    
    '// Char Weapon
    Call Char_SetWeapon(CharIndex, incomingData.ReadInteger())
        
    '// Char Shield
    Call Char_SetShield(CharIndex, incomingData.ReadInteger())
        
    '// Char Casco
    Call Char_SetCasco(CharIndex, incomingData.ReadInteger())
        
    '// Char Fx
    Call Char_SetFx(CharIndex, incomingData.ReadInteger(), incomingData.ReadInteger())
    
    '// Char Aura
    Call Char_SetAura(CharIndex, incomingData.ReadLong(), incomingData.ReadLong())
    
    'Quest
    charlist(CharIndex).EstadoQuest = incomingData.ReadByte()
    
End Sub

Private Sub HandleHeadingChange()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 21/09/2010 - C4b3z0n
'25/08/2009: ZaMa - Changed a variable used incorrectly.
'21/09/2010: C4b3z0n - Added code for FragShooter. If its waiting for the death of certain UserIndex, and it dies, then the capture of the screen will occur.
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte

    Dim CharIndex As Integer

    CharIndex = incomingData.ReadInteger()

    Call Char_SetHeading(CharIndex, incomingData.ReadByte())
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x               As Integer
    Dim y               As Integer
    Dim GrhIndex        As Long
    Dim ParticulaIndex  As Integer
    Dim Shadow          As Byte
    
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
    GrhIndex = incomingData.ReadLong()
    ParticulaIndex = incomingData.ReadInteger()
    Shadow = incomingData.ReadByte()
    
    Call Map_CreateObject(x, y, GrhIndex, ParticulaIndex, Shadow)
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x   As Integer
    Dim y   As Integer
    Dim obj As Long

    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
        
    obj = Map_PosExitsObject(x, y)

    Call Particle_Group_Remove(MapData(x, y).Particle_Group_Index)

    If (obj > 0) Then
        Call Map_DestroyObject(x, y)
    End If
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    Dim block As Boolean
    
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
    block = incomingData.ReadBoolean()
    
    If block Then
        Map_SetBlocked x, y, 1
    Else
        Map_SetBlocked x, y, 0
    End If
End Sub

''
' Handles the PlayMUSIC message.

Private Sub HandlePlayMUSIC()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 06/01/20
'Utilizamos PlayBackgroundMusic para el uso de  MUSIC, simplifique la funcion (Recox)
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim currentMusic As Integer
    Dim Loops As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    currentMusic = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    If ClientSetup.bMusic <> CONST_DESHABILITADA Then
        If ClientSetup.bMusic <> CONST_DESHABILITADA Then
            Sound.NextMusic = currentMusic
            Sound.Fading = 200
        End If
    End If
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
'***************************************************
'Autor: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Last Modified by: Rapsodius
'Added support for 3D Sounds.
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim wave As Integer
    Dim srcX As Integer
    Dim srcY As Integer
    
    wave = incomingData.ReadInteger()
    srcX = incomingData.ReadInteger()
    srcY = incomingData.ReadInteger()
        
    Call Sound.Sound_Play(wave, , Sound.Calculate_Volume(srcX, srcY), Sound.Calculate_Pan(srcX, srcY))
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildAdm
        'Clear guild's list
        .guildslist.Clear
        
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long
        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .guildslist.AddItem(GuildNames(i))
            End If
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(buffer)
        
        .Show vbModeless, frmMain
    End With
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x As Integer
    Dim y As Integer
    Dim Head As Byte
    
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
    Head = incomingData.ReadByte()
        
    Call CambioDeArea(x, y, Head)
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    pausa = Not pausa
End Sub

''
' Handles the ActualizarClima message.

Private Sub HandleActualizarClima()
'***************************************************
'Author: Lorwik
'Last Modification: 09/08/2020
'
'***************************************************
    Dim DayStatus As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Recibimos el estado del dia8
    DayStatus = incomingData.ReadByte
    
    Call Actualizar_Estado(DayStatus)
    
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim fX As Integer
    Dim Loops As Integer
    
    CharIndex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    Call Char_SetFx(CharIndex, fX, Loops)
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With CurrentUser
    
        .UserMaxHP = incomingData.ReadInteger()
        .UserMinHP = incomingData.ReadInteger()
        .UserMaxMAN = incomingData.ReadInteger()
        .UserMinMAN = incomingData.ReadInteger()
        .UserMaxSTA = incomingData.ReadInteger()
        .UserMinSTA = incomingData.ReadInteger()
        .UserGLD = incomingData.ReadLong()
        .UserLvl = incomingData.ReadByte()
        .UserPasarNivel = incomingData.ReadLong()
        .UserExp = incomingData.ReadLong()
        
        frmMain.UpdateProgressExperienceLevelBar (.UserExp)
        
        frmMain.GldLbl.Caption = .UserGLD
        frmMain.lblLvl.Caption = .UserLvl
        
        'Stats
        frmMain.lblMana = .UserMinMAN & "/" & .UserMaxMAN
        frmMain.lblVida = .UserMinHP & "/" & .UserMaxHP
        frmMain.lblEnergia = .UserMinSTA & "/" & .UserMaxSTA
        
        '***************************
        If .UserMaxMAN > 0 Then _
        frmMain.shpMana.Width = (((.UserMinMAN / 100) / (.UserMaxMAN / 100)) * 185)
        '***************************
        
        frmMain.shpVida.Width = (((.UserMinHP / 100) / (.UserMaxHP / 100)) * 185)
    
        '***************************
        
        frmMain.shpEnergia.Width = (((.UserMinSTA / 100) / (.UserMaxSTA / 100)) * 185)
        '***************************
        
        If .UserMinHP = 0 Then
            .UserEstado = 1
        Else
            .UserEstado = 0
        End If
    
        Call frmMain.SetGoldColor
    End With
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 24 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    Dim ObjIndex As Integer
    Dim name As String
    Dim Amount As Integer
    Dim Equipped As Boolean
    Dim GrhIndex As Long
    Dim OBJType As Byte
    Dim MaxHit As Integer
    Dim MinHit As Integer
    Dim MaxDef As Integer
    Dim MinDef As Integer
    Dim value As Single
    Dim NoUsa As Boolean
    
    slot = buffer.ReadByte()
    ObjIndex = buffer.ReadInteger()
    name = buffer.ReadASCIIString()
    Amount = buffer.ReadInteger()
    Equipped = buffer.ReadBoolean()
    GrhIndex = buffer.ReadLong()
    OBJType = buffer.ReadByte()
    MaxHit = buffer.ReadInteger()
    MinHit = buffer.ReadInteger()
    MaxDef = buffer.ReadInteger()
    MinDef = buffer.ReadInteger()
    value = buffer.ReadSingle()
    NoUsa = buffer.ReadBoolean()
    
    If Equipped Then
        Select Case OBJType
            Case eObjType.otWeapon
                lblWeapon = MinHit & "/" & MaxHit
                CurrentUser.UserWeaponEqpSlot = slot
            Case eObjType.otArmadura
                lblArmor = MinDef & "/" & MaxDef
                CurrentUser.UserArmourEqpSlot = slot
            Case eObjType.otescudo
                lblShielder = MinDef & "/" & MaxDef
                CurrentUser.UserHelmEqpSlot = slot
            Case eObjType.otcasco
                lblHelm = MinDef & "/" & MaxDef
                CurrentUser.UserShieldEqpSlot = slot
        End Select
    Else
        Select Case slot
            Case CurrentUser.UserWeaponEqpSlot
                lblWeapon = "0/0"
                CurrentUser.UserWeaponEqpSlot = 0
            Case CurrentUser.UserArmourEqpSlot
                lblArmor = "0/0"
                CurrentUser.UserArmourEqpSlot = 0
            Case CurrentUser.UserHelmEqpSlot
                lblShielder = "0/0"
                CurrentUser.UserHelmEqpSlot = 0
            Case CurrentUser.UserShieldEqpSlot
                lblHelm = "0/0"
                CurrentUser.UserShieldEqpSlot = 0
        End Select
    End If
    
    Call Inventario.SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, value, name, NoUsa)

    If frmComerciar.Visible Then
        Call InvComUsu.SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, value, name, NoUsa)
    End If

    If frmBancoObj.Visible Then
        Call InvBanco(1).SetItem(slot, ObjIndex, Amount, Equipped, GrhIndex, OBJType, MaxHit, MinHit, MaxDef, MinDef, value, name, NoUsa)
        frmBancoObj.NoPuedeMover = False
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

' Handles the AddSlots message.
Private Sub HandleAddSlots()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte
    
    MaxInventorySlots = incomingData.ReadByte
    Call Inventario.DrawInventory
End Sub

' Handles the StopWorking message.
Private Sub HandleStopWorking()
'***************************************************
'Author: Budi
'Last Modification: 12/01/09
'
'***************************************************

    Call incomingData.ReadByte

    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_WORK_FINISHED").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
    End With
End Sub

' Handles the CancelOfferItem message.

Private Sub HandleCancelOfferItem()
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 05/03/10
'
'***************************************************
    Dim slot As Byte
    Dim Amount As Long
    
    Call incomingData.ReadByte
    
    slot = incomingData.ReadByte
    
    With InvOfferComUsu(0)
        Amount = .Amount(slot)
        
        ' No tiene sentido que se quiten 0 unidades
        If Amount <> 0 Then
            ' Actualizo el inventario general
            Call frmComerciarUsu.UpdateInvCom(.ObjIndex(slot), Amount)
            
            ' Borro el item
            Call .SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
        End If
    End With
    
    ' Si era el unico item de la oferta, no puede confirmarla
    If Not frmComerciarUsu.HasAnyItem(InvOfferComUsu(0)) And Not frmComerciarUsu.HasAnyItem(InvOroComUsu(1)) Then
        Call frmComerciarUsu.HabilitarConfirmar(False)
    End If
    
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call frmComerciarUsu.PrintCommerceMsg(JsonLanguage.item("MENSAJE_NO_COMM_OBJETO").item("TEXTO"), FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 23 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    slot = buffer.ReadByte()
    
    With UserBancoInventory(slot)
        .ObjIndex = buffer.ReadInteger()
        .name = buffer.ReadASCIIString()
        .Amount = buffer.ReadInteger()
        .GrhIndex = buffer.ReadLong()
        .OBJType = buffer.ReadByte()
        .MaxHit = buffer.ReadInteger()
        .MinHit = buffer.ReadInteger()
        .MaxDef = buffer.ReadInteger()
        .MinDef = buffer.ReadInteger()
        .valor = buffer.ReadLong()
        .NoUsa = buffer.ReadBoolean()
        
        If frmBancoObj.Visible Then
            Call InvBanco(0).SetItem(slot, .ObjIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .valor, .name, .NoUsa)
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ChangeSpellSlot message.

Private Sub HandleChangeSpellSlot()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    slot = buffer.ReadByte()
    
    UserHechizos(slot) = buffer.ReadInteger()
    
    If slot <= frmMain.hlst.ListCount Then
        frmMain.hlst.List(slot - 1) = buffer.ReadASCIIString()
    Else
        Call frmMain.hlst.AddItem(buffer.ReadASCIIString())
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub


''
' Handles the Attributes message.

Private Sub HandleAtributes()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        CurrentUser.UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    LlegaronAtrib = True
End Sub

''
' Handles the InitTrabajo message.

Private Sub HandleInitTrabajo()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Count As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    MirandoTrabajo = buffer.ReadByte() 'Me sirve para saber que trabajo estoy mirado y para indicar que estoy trabajando
    
    Count = buffer.ReadInteger()
    
    ReDim ObjetoTrabajo(Count) As tItemsConstruibles
    
    For i = 1 To Count
        With ObjetoTrabajo(i)
            .name = buffer.ReadASCIIString()    'Get the object's name
            .GrhIndex = buffer.ReadLong()
            .PrecioConstruccion = buffer.ReadLong()
            
            For j = 1 To MAXMATERIALES
                .Materiales(j) = buffer.ReadLong()
                .CantMateriales(j) = buffer.ReadInteger()
                .NameMateriales(j) = buffer.ReadASCIIString()
            Next j
            
            .ObjIndex = buffer.ReadInteger()
        End With
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    'En otras palabras, a partir de ahora podes usar "Exit Sub" sin romper nada.
    Call incomingData.CopyBuffer(buffer)
    
    Call frmTrabajos.Show(vbModeless, frmMain)
    
    For i = 1 To MAX_LIST_ITEMS
        Set InvMaterialTrabajo(i) = New clsGraphicalInventory
    Next i
    
    With frmTrabajos
        ' Inicializo los inventarios
        Call InvMaterialTrabajo(1).Initialize(DirectD3D8, .picMaterial0, 4, , , , , , False)
        Call InvMaterialTrabajo(2).Initialize(DirectD3D8, .picMaterial1, 4, , , , , , False)
        Call InvMaterialTrabajo(3).Initialize(DirectD3D8, .picMaterial2, 4, , , , , , False)
        Call InvMaterialTrabajo(4).Initialize(DirectD3D8, .picMaterial3, 4, , , , , , False)
        
        Call .HideExtraControls(Count)
        Call .RenderList(1)
    End With

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleInitCraftman()
    '***************************************************
    'Author: WyroX
    'Last Modification: 27/01/2020
    '***************************************************

    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim CountObjs As Integer
    Dim CountCrafteo As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long

    frmArtesano.ArtesaniaCosto = buffer.ReadLong()

    CountObjs = buffer.ReadInteger()

    ReDim ObjArtesano(CountObjs) As tItemArtesano

    For i = 1 To CountObjs
        With ObjArtesano(i)
            .name = buffer.ReadASCIIString()
            .GrhIndex = buffer.ReadLong()
            .ObjIndex = buffer.ReadInteger()
            
            CountCrafteo = buffer.ReadByte()
            ReDim .ItemsCrafteo(CountCrafteo) As tItemCrafteo
            
            For j = 1 To CountCrafteo
                .ItemsCrafteo(j).name = buffer.ReadASCIIString()
                .ItemsCrafteo(j).GrhIndex = buffer.ReadLong()
                .ItemsCrafteo(j).ObjIndex = buffer.ReadInteger()
                .ItemsCrafteo(j).Amount = buffer.ReadInteger()
            Next j
        End With
    Next i

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

    Call frmArtesano.Show(vbModeless, frmMain)

    For i = 1 To MAX_LIST_ITEMS
        Set InvObjArtesano(i) = New clsGraphicalInventory
    Next i

    With frmArtesano
        ' Inicializo los inventarios
        Call InvObjArtesano(1).Initialize(DirectD3D8, .picObj0, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(2).Initialize(DirectD3D8, .picObj1, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(3).Initialize(DirectD3D8, .picObj2, MAX_ITEMS_CRAFTEO, , , , , , False)
        Call InvObjArtesano(4).Initialize(DirectD3D8, .picObj3, MAX_ITEMS_CRAFTEO, , , , , , False)

        Call .HideExtraControls(CountObjs)
        Call .RenderList(1)
    End With

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserDescansar = Not CurrentUser.UserDescansar
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call MostrarMensaje(buffer.ReadASCIIString())
    
    If frmConnect.Visible Then
        frmMain.Client.CloseSck
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserCiego = True
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserEstupido = True
End Sub

''
' Handles the ShowSignal message.

Private Sub HandleShowSignal()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim tmp As String
    tmp = buffer.ReadASCIIString()
    
    Call InitCartel(tmp, buffer.ReadLong())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 23 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim slot As Byte
    slot = buffer.ReadByte()
    
    With NPCInventory(slot)
        .name = buffer.ReadASCIIString()
        .Amount = buffer.ReadInteger()
        .valor = buffer.ReadSingle()
        .GrhIndex = buffer.ReadLong()
        .ObjIndex = buffer.ReadInteger()
        .OBJType = buffer.ReadByte()
        .MaxHit = buffer.ReadInteger()
        .MinHit = buffer.ReadInteger()
        .MaxDef = buffer.ReadInteger()
        .MinDef = buffer.ReadInteger()
        .NoUsa = buffer.ReadBoolean()
    
        If frmComerciar.Visible Then
            Call InvComNpc.SetItem(slot, .ObjIndex, .Amount, 0, .GrhIndex, .OBJType, .MaxHit, .MinHit, .MaxDef, .MinDef, .valor, .name, .NoUsa)
        End If
    End With
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With CurrentUser
        .UserMaxAGU = incomingData.ReadByte()
        .UserMinAGU = incomingData.ReadByte()
        .UserMaxHAM = incomingData.ReadByte()
        .UserMinHAM = incomingData.ReadByte()
        frmMain.lblHambre = .UserMinHAM & "%"
        frmMain.lblSed = .UserMinAGU & "%"
        
        frmMain.shpHambre.Height = (((.UserMinHAM / 100) / (.UserMaxHAM / 100)) * 20)
        '*********************************
        
        frmMain.shpSed.Visible = (((.UserMinAGU / 100) / (.UserMaxAGU / 100)) * 20)
    End With
End Sub

''
' Handles the Fame message.

Private Sub HandleFame()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 29 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With CurrentUser.UserReputacion
        .AsesinoRep = incomingData.ReadLong()
        .BandidoRep = incomingData.ReadLong()
        .BurguesRep = incomingData.ReadLong()
        .LadronesRep = incomingData.ReadLong()
        .NobleRep = incomingData.ReadLong()
        .PlebeRep = incomingData.ReadLong()
        .Promedio = incomingData.ReadLong()
    End With
    
    LlegoFama = True
End Sub

''
' Handles the MiniStats message.

Private Sub HandleMiniStats()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 20 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    With CurrentUser.UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .UsuariosMatados = incomingData.ReadLong()
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()
    End With
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = SkillPoints + incomingData.ReadInteger()

End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim ForumType As eForumMsgType
    Dim Title As String
    Dim Message As String
    Dim Author As String
    
    ForumType = buffer.ReadByte
    
    Title = buffer.ReadASCIIString()
    Author = buffer.ReadASCIIString()
    Message = buffer.ReadASCIIString()
    
    If Not frmForo.ForoLimpio Then
        clsForos.ClearForums
        frmForo.ForoLimpio = True
    End If

    Call clsForos.AddPost(ForumAlignment(ForumType), Title, Author, Message, EsAnuncio(ForumType))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmForo.Privilegios = incomingData.ReadByte
    frmForo.CanPostSticky = incomingData.ReadByte
    
    If Not MirandoForo Then
        frmForo.Show , frmMain
    End If
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim CharIndex As Integer
    Dim timeRemaining As Integer
    
    CharIndex = incomingData.ReadInteger()
    CurrentUser.UserInvisible = incomingData.ReadBoolean()
    Call Char_SetInvisible(CharIndex, CurrentUser.UserInvisible)
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    CurrentUser.UserMeditar = Not CurrentUser.UserMeditar
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserCiego = False
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CurrentUser.UserEstupido = False
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Now the server send the percentage of progress of the skills.
'***************************************************
    If incomingData.length < 2 + NUMSKILLS Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    CurrentUser.UserClase = incomingData.ReadByte
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        CurrentUser.UserSkills(i) = incomingData.ReadByte()
    Next i
    
    LlegaronSkills = True
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatures() As String
    Dim i As Long
    Dim Upper_creatures As Long
    
    creatures = Split(buffer.ReadASCIIString(), SEPARATOR)
    Upper_creatures = UBound(creatures())
    
    For i = 0 To Upper_creatures
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList() As String
    Dim Upper_guildList As Long
    Dim i As Long
    Dim sTemp As String
    
    'Get news' string
    frmGuildNews.news = buffer.ReadASCIIString()
    
    'Get Enemy guilds list
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    Upper_guildList = UBound(guildList)
    
    For i = 0 To Upper_guildList
        sTemp = frmGuildNews.txtClanesGuerra.Text
        frmGuildNews.txtClanesGuerra.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    'Get Allied guilds list
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To Upper_guildList
        sTemp = frmGuildNews.txtClanesAliados.Text
        frmGuildNews.txtClanesAliados.Text = sTemp & guildList(i) & vbCrLf
    Next i
    
    If ClientSetup.bGuildNews Or bShowGuildNews Then frmGuildNews.Show vbModeless, frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim vsGuildList() As String, Upper_vsGuildList As Long
    Dim i As Long
    
    vsGuildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    Upper_vsGuildList = UBound(vsGuildList())
    
    Call frmPeaceProp.lista.Clear
    For i = 0 To Upper_vsGuildList
        Call frmPeaceProp.lista.AddItem(vsGuildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList()     As String
    Dim Upper_guildList As Long
    Dim i               As Long
    
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    With frmPeaceProp
    
        .lista.Clear
    
        Upper_guildList = UBound(guildList())
    
        For i = 0 To Upper_guildList
            .lista.AddItem (guildList(i))
        Next i
    
        .ProposalType = TIPO_PROPUESTA.PAZ
        .Show vbModeless, frmMain
    
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 35 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .imgRechazar.Visible = False
            .imgAceptar.Visible = False
            .imgEchar.Visible = True
            .imgPeticion.Visible = False
        Else
            .imgRechazar.Visible = True
            .imgAceptar.Visible = True
            .imgEchar.Visible = False
            .imgPeticion.Visible = True
        End If
        
        .Nombre.Caption = buffer.ReadASCIIString()
        .Raza.Caption = ListaRazas(buffer.ReadByte())
        .Clase.Caption = ListaClases(buffer.ReadByte())
        
        If buffer.ReadByte() = 1 Then
            .Genero.Caption = "Hombre"
        Else
            .Genero.Caption = "Mujer"
        End If
        
        .Nivel.Caption = buffer.ReadByte()
        .Oro.Caption = buffer.ReadLong()
        .Banco.Caption = buffer.ReadLong()
        
        Dim reputation As Long
        reputation = buffer.ReadLong()
        
        .reputacion.Caption = reputation
        
        .txtPeticiones.Text = buffer.ReadASCIIString()
        .guildactual.Caption = buffer.ReadASCIIString()
        .txtMiembro.Text = buffer.ReadASCIIString()
        
        Dim armada As Boolean
        Dim caos As Boolean
        
        armada = buffer.ReadBoolean()
        caos = buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = JsonLanguage.item("ARMADA").item("TEXTO")
        ElseIf caos Then
            .ejercito.Caption = JsonLanguage.item("LEGION").item("TEXTO")
        End If
        
        .Ciudadanos.Caption = CStr(buffer.ReadLong())
        .criminales.Caption = CStr(buffer.ReadLong())
        
        If reputation > 0 Then
            .status.Caption = " " & JsonLanguage.item("CIUDADANO").item("TEXTO")
            .status.ForeColor = vbBlue
        Else
            .status.Caption = " " & JsonLanguage.item("CRIMINAL").item("TEXTO")
            .status.ForeColor = vbRed
        End If
        
        Call .Show(vbModeless, frmMain)
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim i As Long
    Dim List() As String

    With frmGuildLeader
        'Get list of existing guilds
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .guildslist.AddItem(GuildNames(i))
            End If
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(buffer.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .members.Clear

        For i = 0 To UBound(GuildMembers())
            Call .members.AddItem(GuildMembers(i))
        Next i
        
        .txtguildnews = buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear

        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        
        .Show , frmMain
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 26 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildBrief
        .imgDeclararGuerra.Visible = .EsLeader
        .imgOfrecerAlianza.Visible = .EsLeader
        .imgOfrecerPaz.Visible = .EsLeader
        
        .Nombre.Caption = buffer.ReadASCIIString()
        .fundador.Caption = buffer.ReadASCIIString()
        .creacion.Caption = buffer.ReadASCIIString()
        .lider.Caption = buffer.ReadASCIIString()
        .web.Caption = buffer.ReadASCIIString()
        .Miembros.Caption = buffer.ReadInteger()
        
        If buffer.ReadBoolean() Then
            .eleccion.Caption = UCase$(JsonLanguage.item("ABIERTA").item("TEXTO"))
        Else
            .eleccion.Caption = UCase$(JsonLanguage.item("CERRADA").item("TEXTO"))
        End If
        
        .lblAlineacion.Caption = buffer.ReadASCIIString()
        .Enemigos.Caption = buffer.ReadInteger()
        .Aliados.Caption = buffer.ReadInteger()
        .antifaccion.Caption = buffer.ReadASCIIString()
        
        Dim codexStr() As String
        Dim i As Long
        
        codexStr = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        For i = 0 To 7
            .Codex(i).Caption = codexStr(i)
        Next i
        
        .Desc.Text = buffer.ReadASCIIString()
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowGuildAlign message.

Private Sub HandleShowGuildAlign()
'***************************************************
'Author: ZaMa
'Last Modification: 14/12/2009
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    frmEligeAlineacion.Show vbModeless, frmMain
End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildFoundation.Show , frmMain
End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    CurrentUser.UserParalizado = Not CurrentUser.UserParalizado

End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************

    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    Dim OfferSlot As Byte
    Dim ObjIndex As Integer
    Dim Amount As Long
    Dim ObjGrhIndex As Long
    Dim OffObjType As Byte
    Dim MaximoHit As Integer
    Dim MinimoHit As Integer
    Dim MaximaDefensa As Integer
    Dim MinimaDefensa As Integer
    Dim PrecioValor As Long
    Dim NombreObjeto As String
    Dim NoUsa As Boolean
    
    With buffer
    
        'Remove packet ID
        Call .ReadByte
        OfferSlot = .ReadByte()
        ObjIndex = .ReadInteger()
        Amount = .ReadLong()
        
        ObjGrhIndex = .ReadLong()
        OffObjType = .ReadByte()
        MaximoHit = .ReadInteger()
        MinimoHit = .ReadInteger()
        MaximaDefensa = .ReadInteger()
        MinimaDefensa = .ReadInteger()
        PrecioValor = .ReadLong()
        NombreObjeto = .ReadASCIIString()
        NoUsa = .ReadBoolean()
        
        If OfferSlot = GOLD_OFFER_SLOT Then
            Call InvOroComUsu(2).SetItem(1, ObjIndex, Amount, 0, _
                                            ObjGrhIndex, OffObjType, MaximoHit, MinimoHit, _
                                            MaximaDefensa, MinimaDefensa, PrecioValor, NombreObjeto, NoUsa)

        Else
            Call InvOfferComUsu(1).SetItem(OfferSlot, ObjIndex, Amount, 0, _
                                            ObjGrhIndex, OffObjType, MaximoHit, MinimoHit, _
                                            MaximaDefensa, MinimaDefensa, PrecioValor, NombreObjeto, NoUsa)

        End If
    End With
    
    Call frmComerciarUsu.PrintCommerceMsg(TradingUserName & JsonLanguage.item("MENSAJE_COMM_OFERTA_CAMBIA").item("TEXTO"), FontTypeNames.FONTTYPE_VENENO)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the SendNight message.

Private Sub HandleSendNight()
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/08/07
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim tBool As Boolean 'CHECK, este handle no hace nada con lo que recibe.. porque, ehmm.. no hay noche?.. o si?
    tBool = incomingData.ReadBoolean()
End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatureList() As String
    Dim i As Long
    Dim Upper_creatureList As Long
    
    creatureList = Split(buffer.ReadASCIIString(), SEPARATOR)
    Upper_creatureList = UBound(creatureList())
    
    For i = 0 To Upper_creatureList
        Call frmSpawnList.lstCriaturas.AddItem(creatureList(i))
    Next i
    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim sosList() As String
    Dim i As Long
    Dim Upper_sosList As Long
    
    sosList = Split(buffer.ReadASCIIString(), SEPARATOR)
    Upper_sosList = UBound(sosList())
    
    For i = 0 To Upper_sosList
        Call frmMSG.List1.AddItem(sosList(i))
    Next i
    
    frmMSG.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowDenounces message.

Private Sub HandleShowDenounces()
'***************************************************
'Author: ZaMa
'Last Modification: 14/11/2010
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim DenounceList() As String
    Dim Upper_denounceList As Long
    Dim DenounceIndex As Long
    
    DenounceList = Split(buffer.ReadASCIIString(), SEPARATOR)
    Upper_denounceList = UBound(DenounceList())
    
    With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
        For DenounceIndex = 0 To Upper_denounceList
            Call AddtoRichTextBox(frmMain.RecTxt, DenounceList(DenounceIndex), .Red, .Green, .Blue, .bold, .italic)
        Next DenounceIndex
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowPartyForm()
'***************************************************
'Author: Budi
'Last Modification: 11/26/09
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim members() As String
    Dim Upper_members As Long
    Dim i As Long
    
    EsPartyLeader = CBool(buffer.ReadByte())
       
    members = Split(buffer.ReadASCIIString(), SEPARATOR)
    Upper_members = UBound(members())
    
    For i = 0 To Upper_members
        Call frmParty.lstMembers.AddItem(members(i))
    Next i
    
    frmParty.lblTotalExp.Caption = buffer.ReadLong
    frmParty.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandlePeticionInvitarParty()
'***************************************************
'Author: Lorwik
'Last Modification: 05/11/2020
'
'***************************************************

    Call incomingData.ReadByte
    
    frmMain.MousePointer = 2 'vbCrosshair
    
    InvitandoParty = True
    
    Call AddtoRichTextBox(frmMain.RecTxt, _
        JsonLanguage.item("MENSAJE_INVITAR_PARTY").item("TEXTO"), _
        JsonLanguage.item("MENSAJE_INVITAR_PARTY").item("COLOR").item(1), _
        JsonLanguage.item("MENSAJE_INVITAR_PARTY").item("COLOR").item(2), _
        JsonLanguage.item("MENSAJE_INVITAR_PARTY").item("COLOR").item(3))

End Sub


''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'*************************************Su**************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim id As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    id = incomingData.ReadByte
    
    If id = 0 Then
        frmPanelGm.Show vbModeless, frmMain
    ElseIf id = 1 Then
        frmBuscar.Show vbModeless, frmMain
    End If
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim userList() As String
    
    userList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear
        
        Dim i As Long
        Dim Upper_userlist As Long
            Upper_userlist = UBound(userList())
            
        For i = 0 To Upper_userlist
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleScreenMessage()

    Call incomingData.ReadByte
    
    renderMsgReset
    renderText = incomingData.ReadASCIIString
    renderTextPk = incomingData.ReadASCIIString
    'renderFont = incomingData.ReadInteger
    colorRender = 240
End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Call incomingData.ReadByte
    
    Dim MENSAJE_PING As String
        MENSAJE_PING = JsonLanguage.item("MENSAJE_PING").item("TEXTO")
        MENSAJE_PING = Replace$(MENSAJE_PING, "VAR_PING", (timeGetTime - pingTime))
        
    Call AddtoRichTextBox(frmMain.RecTxt, _
                            MENSAJE_PING, _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(1), _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(2), _
                            JsonLanguage.item("MENSAJE_PING").item("COLOR").item(3), _
                            True, False, True)
    
    pingTime = 0
End Sub

''
' Handles the Pong message.

Private Sub HandleGuildMemberInfo()
'***************************************************
'Author: ZaMa
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildMember
        'Clear guild's list
        .lstClanes.Clear
        
        GuildNames = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        Dim i As Long

        For i = 0 To UBound(GuildNames())
            If LenB(GuildNames(i)) <> 0 Then
                Call .lstClanes.AddItem(GuildNames(i))
            End If
        Next i
        
        'Get list of guild's members
        GuildMembers = Split(buffer.ReadASCIIString(), SEPARATOR)
        .lblCantMiembros.Caption = CStr(UBound(GuildMembers()) + 1)
        
        'Empty the list
        Call .lstMiembros.Clear

        For i = 0 To UBound(GuildMembers())
            Call .lstMiembros.AddItem(GuildMembers(i))
        Next i
        
        'If we got here then packet is complete, copy data back to original queue
        Call incomingData.CopyBuffer(buffer)
        
        .Show vbModeless, frmMain
    End With
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim CharIndex As Integer
    Dim NickColor As Byte
    Dim UserTag As String
    
    CharIndex = buffer.ReadInteger()
    NickColor = buffer.ReadByte()
    UserTag = buffer.ReadASCIIString()
    
    'Update char status adn tag!
    With charlist(CharIndex)
        If (NickColor And eNickColor.ieCriminal) <> 0 Then
            .Criminal = 1
        Else
            .Criminal = 0
        End If
        
        .Atacable = (NickColor And eNickColor.ieAtacable) <> 0
        
        .Nombre = UserTag
        .Clan = mid$(.Nombre, getTagPosition(.Nombre))
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the RecordList message.

Private Sub HandleRecordList()
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim NumRecords As Byte
    Dim i As Long
    
    NumRecords = buffer.ReadByte
    
    'Se limpia el ListBox y se agregan los usuarios
    frmPanelGm.lstUsers.Clear
    For i = 1 To NumRecords
        frmPanelGm.lstUsers.AddItem buffer.ReadASCIIString
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

''
' Handles the RecordDetails message.

Private Sub HandleRecordDetails()
'***************************************************
'Author: Amraphen
'Last Modification: 29/11/2010
'
'***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue: Set buffer = New clsByteQueue
    Dim tmpStr As String
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
       
    With frmPanelGm
        .txtCreador.Text = buffer.ReadASCIIString
        .txtDescrip.Text = buffer.ReadASCIIString
        
        'Status del pj
        If buffer.ReadBoolean Then
            .lblEstado.ForeColor = vbGreen
            .lblEstado.Caption = UCase$(JsonLanguage.item("EN_LINEA").item("TEXTO"))
        Else
            .lblEstado.ForeColor = vbRed
            .lblEstado.Caption = UCase$(JsonLanguage.item("DESCONECTADO").item("TEXTO"))
        End If
        
        'IP del personaje
        tmpStr = buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtIP.Text = tmpStr
        Else
            .txtIP.Text = JsonLanguage.item("USUARIO").item("TEXTO") & JsonLanguage.item("DESCONECTADO").item("TEXTO")
        End If
        
        'Tiempo online
        tmpStr = buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtTimeOn.Text = tmpStr
        Else
            .txtTimeOn.Text = JsonLanguage.item("USUARIO").item("TEXTO") & JsonLanguage.item("DESCONECTADO").item("TEXTO")
        End If
        
        'Observaciones
        tmpStr = buffer.ReadASCIIString
        If LenB(tmpStr) Then
            .txtObs.Text = tmpStr
        Else
            .txtObs.Text = JsonLanguage.item("MENSAJE_NO_NOVEDADES").item("TEXTO")
        End If
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleAttackAnim()
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim CharIndex As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    CharIndex = incomingData.ReadInteger
    'Set the animation trigger on true
    charlist(CharIndex).attacking = True 'should be done in separated sub?
End Sub

Private Sub HandleFXtoMap()

    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim x As Integer, y As Integer, FxIndex As Integer, Loops As Integer
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Loops = incomingData.ReadByte
    x = incomingData.ReadInteger
    y = incomingData.ReadInteger
    FxIndex = incomingData.ReadInteger
    
    'Comprobamos si las coordenadas estan dentro de lo esperado
    If Not Map_InBounds(x, y) Then Exit Sub

    'Set the fx on the map
    With MapData(x, y) 'TODO: hay que hacer una funcion separada que haga esto
        .FxIndex = FxIndex
    
        If .FxIndex > 0 Then
                        
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)
            .fX.Loops = Loops

        End If

    End With

End Sub

Private Sub HandleEnviarPJUserAccount()

    If incomingData.length < 13 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Security.Redundance = buffer.ReadByte
    CurrentUser.AccountName = buffer.ReadASCIIString
    CurrentUser.NumberOfCharacters = buffer.ReadByte
    
    CurrentUser.VIP = buffer.ReadASCIIString
    CurrentUser.esVIP = buffer.ReadBoolean

    'Cambiamos al modo cuenta
    Call ModConectar.MostrarCuenta(Not frmConnect.Visible)

    If CurrentUser.NumberOfCharacters > 0 Then
    
        ReDim cPJ(1 To CurrentUser.NumberOfCharacters) As PjCuenta
        
        Dim loopc As Long
        
        For loopc = 1 To CurrentUser.NumberOfCharacters
        
            With cPJ(loopc)
                .Nombre = buffer.ReadASCIIString
                .Body = buffer.ReadInteger
                .Head = buffer.ReadInteger
                .weapon = buffer.ReadInteger
                .shield = buffer.ReadInteger
                .helmet = buffer.ReadInteger
                .Class = buffer.ReadByte
                .Race = buffer.ReadByte
                .Map = buffer.ReadInteger
                .Level = buffer.ReadByte
                .Criminal = buffer.ReadBoolean
                .Dead = buffer.ReadBoolean
                
                If .Dead Then
                    .Head = eCabezas.CASPER_HEAD
                    .Body = iCuerpoMuerto
                    .weapon = 0
                    .helmet = 0
                    .shield = 0
                ElseIf (.Body = 397 Or .Body = 395 Or .Body = 399) Then
                    .Head = 0
                End If

                .GameMaster = buffer.ReadBoolean
            End With
            
        Next loopc
        
    End If

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleSearchList()
On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
 
    Dim num   As Integer
    Dim Datos As String
    Dim obj   As Boolean
        
    'Remove packet ID
    Call buffer.ReadByte
   
    num = buffer.ReadInteger()
    obj = buffer.ReadBoolean()
    Datos = buffer.ReadASCIIString()
 
    Call frmBuscar.AddItem(num, obj, Datos)

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error
 
End Sub
 
Private Sub HandleQuestDetails()
'*****************************************
'Recibe y maneja el paquete QuestDetails del servidor.
'Last modified: 31/01/2010 by Amraphen
'*****************************************
    If incomingData.length < 15 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    Dim tmpStr As String
    Dim tmpByte As Byte
    Dim QuestEmpezada As Boolean
    Dim i As Integer
    
    With buffer
        'Leemos el id del paquete
        Call .ReadByte
        
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)
        
        tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
        tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
        
        tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay NPCs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                If QuestEmpezada Then
                    tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                Else
                    tmpStr = tmpStr & vbCrLf
                End If
            Next i
        End If
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay NPCs para hablar
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Hablar con " & .ReadASCIIString & "."
                If QuestEmpezada Then
                    tmpStr = tmpStr & " (Has hablado con " & .ReadInteger & ")" & vbCrLf
                Else
                    tmpStr = tmpStr & vbCrLf
                End If
            Next i
        End If
        
        tmpByte = .ReadByte
        If tmpByte Then 'Hay OBJs
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) Conseguir " & .ReadInteger & " " & .ReadASCIIString & "." & vbCrLf
            Next i
        End If
 
        tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
        tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
        tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
        
        tmpByte = .ReadByte
        If tmpByte Then
            For i = 1 To tmpByte
                tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadASCIIString & vbCrLf
            Next i
        End If
    End With
    
    'Determinamos que formulario se muestra, segï¿½n si recibimos la informaciï¿½n y la quest estï¿½ empezada o no.
    If QuestEmpezada Then
        frmQuests.txtInfo.Text = tmpStr
        
    Else
        frmQuestInfo.txtInfo.Text = tmpStr
        frmQuestInfo.Show vbModeless, frmMain
        Comerciando = True
        
    End If
    
    Call incomingData.CopyBuffer(buffer)
    
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then _
        Err.Raise Error
End Sub
 
Public Sub HandleQuestListSend()
'*****************************************
'Recibe y maneja el paquete QuestListSend del servidor.
'Last modified: 31/01/2010 by Amraphen
'*****************************************
    If incomingData.length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errhandler
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    Dim i As Integer
    Dim tmpByte As Byte
    Dim tmpStr As String
    
    'Leemos el id del paquete
    Call buffer.ReadByte
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = buffer.ReadByte
    
    'Limpiamos el ListBox y el TextBox del formulario
    frmQuests.lstQuests.Clear
    frmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = buffer.ReadASCIIString
        
        'Agregamos los items
        For i = 1 To tmpByte
            frmQuests.lstQuests.AddItem ReadField(i, tmpStr, 45)
        Next i
    End If
    
    'Mostramos el formulario
    frmQuests.Show vbModeless, frmMain
    
    'Pedimos la informaciï¿½n de la primer quest (si la hay)
    If tmpByte Then Call Protocol_Write.WriteQuestDetailsRequest(1)
    
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(buffer)
 
errhandler:
    Dim Error As Long
    Error = Err.number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then _
        Err.Raise Error
End Sub
 
Public Sub HandleActualizarNPCQuest()

    '*****************************************
    'Autor: Lorwik
    'Fecha: 10/05/2021
    '*****************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler
    
    Dim CharIndex As Integer

    Dim buffer    As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Leemos el id del paquete
    Call buffer.ReadByte
    
    CharIndex = buffer.ReadInteger()
    charlist(CharIndex).EstadoQuest = buffer.ReadByte()
    
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(buffer)
 
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleCreateDamage()
 
    ' @ Crea dano en pos X e Y.
 
    With incomingData
        
        ' Leemos el ID del paquete.
        .ReadByte
     
        Call mDx8_Dibujado.Damage_Create(.ReadInteger(), .ReadInteger(), 0, .ReadLong(), .ReadByte())
     
    End With
 
End Sub

Private Sub HandleUserInEvent()
    Call incomingData.ReadByte
    
    CurrentUser.UserEvento = Not CurrentUser.UserEvento
End Sub

Private Sub HandleEnviarListDeAmigos()

    '***************************************************
    'Author: Abusivo#1215 (DISCORD)
    '***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler
    
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue
    Call buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim slot As Byte
    Dim i    As Integer

    slot = buffer.ReadByte()
    
    With frmMain.ListAmigos
    
        If slot <= .ListCount Then
            .List(slot - 1) = buffer.ReadASCIIString()
        Else
            Call .AddItem(buffer.ReadASCIIString())
        End If
    
    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    Dim Error As Long
        Error = Err.number
        
    On Error GoTo 0

    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Call Err.Raise(Error)
    
End Sub

Private Sub HandleProyectil()
'**************************
'Autor: Lorwik
'Fecha: 18/05/2020
'**************************

    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet id
    Call incomingData.ReadByte
        
    Dim CharSending      As Integer
    Dim CharRecieved     As Integer
    Dim GrhIndex         As Long
        
    CharSending = incomingData.ReadInteger()
    CharRecieved = incomingData.ReadInteger()
    GrhIndex = incomingData.ReadLong()
    
    Engine_Projectile_Create CharSending, CharRecieved, GrhIndex, 0
End Sub

Private Sub HandleCharParticle()
'***************************************************
'Author: Lorwik
'Last Modification: 20/07/2020
'***************************************************

    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim Create As Boolean
    Dim ParticulaID As Integer
    Dim char_index As Integer
    Dim Life As Long

    'Remove packet ID
    Call incomingData.ReadByte
    
    ParticulaID = incomingData.ReadInteger
    Create = incomingData.ReadBoolean
    char_index = incomingData.ReadInteger
    Life = incomingData.ReadLong
    
    'Si el create esta en true, creamos
    If Create Then
        Call General_Char_Particle_Create(ParticulaID, char_index, Life)
        
    Else 'si es False, destruimos
    
        Call Char_Particle_Group_Remove_All(char_index)
    End If
End Sub

Private Sub HandleIniciarSubasta()
    With incomingData
        Call .ReadByte
        frmSubastar.Show
    End With
End Sub

Private Sub HandleConfirmarInstruccion()
'***************************************************
'Author: Lorwik
'Last Modification: 19/08/2020
'***************************************************

    Dim Mensaje As String
    
    With incomingData
        Call .ReadByte
        
        Mensaje = .ReadASCIIString
        
        Call Sound.Sound_Play(SND_MSG)
        frmConfirmacion.msg.Caption = Mensaje
        frmConfirmacion.Show
    End With
End Sub

Private Sub HandleSetSpeed()
'***************************************************
'Author: Lorwik
'Last Modification: 23/10/2020
'Setea la nueva velocidad recibida por el server
'***************************************************

    Dim speed As Double
    
    With incomingData
        Call .ReadByte
        
        speed = .ReadDouble
        
        Call SetSpeedUsuario(speed)
        
    End With
End Sub

Private Sub HandleAtaqueNPC()

    Dim NPCAtaqueIndex As Integer

    'Remove packet ID
    Call incomingData.ReadByte

    NPCAtaqueIndex = incomingData.ReadInteger()

    With charlist(NPCAtaqueIndex)
            
        MapData(.Pos.x, .Pos.y).CharIndex = NPCAtaqueIndex
        .Ataque.AtaqueWalk(.Heading).Started = 1
        .NPCAttack = True
    End With
End Sub

Private Sub HandleBattlegrounds()
'*****************************
'Autor: Lorwik
'Fecha: 02/05/2022
'Descripción: Recibe la variable Battegrounds del server
'*****************************

    'Remove packet ID
    Call incomingData.ReadByte
    
    Battlegrounds = incomingData.ReadBoolean()

End Sub

Private Sub HandleMostrarShop()
'*****************************
'Autor: Lorwik
'Fecha: 016/05/2022
'Descripción: Recibe la variable Battegrounds del server
'*****************************

    Dim NUMSHOPS As Integer
    Dim i As Integer

    'Remove packet ID
    Call incomingData.ReadByte
    
    frmShop.lstItemsShop.Clear
    
    frmShop.lblCredits.Caption = incomingData.ReadInteger
    NUMSHOPS = incomingData.ReadInteger
    
    ReDim ShopObject(1 To NUMSHOPS) As ShopObj
    
    For i = 1 To NUMSHOPS
    
        ShopObject(i).ObjIndex = incomingData.ReadInteger
        ShopObject(i).Nombre = incomingData.ReadASCIIString
        ShopObject(i).Amount = incomingData.ReadInteger
        ShopObject(i).valor = incomingData.ReadInteger
    
    Next i
    
    For i = 1 To NUMSHOPS
    
        frmShop.lstItemsShop.AddItem ShopObject(i).Nombre
    
    Next i
    
    frmShop.lblNombre = vbNullString
    frmShop.Show

End Sub

Private Sub HandleActualizarGemasShop()
'***************************************
'Autor: Lorwik
'Fecha: 16/05/2022
'Descripcion: Actualiza las gemas Winter en la shop
'***************************************

    'Remove packet ID
    Call incomingData.ReadByte
    
    frmShop.lblCredits = incomingData.ReadLong
    
End Sub

Private Sub HandleMostrarPVP()
'***************************************************
'Author: Lorwik
'Last Modification: 22/05/2022
'
'***************************************************

    Call incomingData.ReadByte
    
    CurrentUser.UserNivelPVP = incomingData.ReadByte
    CurrentUser.UserEXPPVP = incomingData.ReadInteger
    CurrentUser.UserELVPVP = incomingData.ReadInteger
    CurrentUser.UserELO = incomingData.ReadLong

    Call frmPVP.IniciarLabels
    frmPVP.Show , frmMain
    
End Sub

Private Sub HandleBarFx()
'***************************************************
'Author: Lorwik
'Last Modification: 11/10/2023
'
'***************************************************

    On Error GoTo HandleBarFx_Err

    Dim CharIndex As Integer

    Dim BarTime   As Integer

    Dim BarAccion As Byte
    
    Call incomingData.ReadByte
    
    CharIndex = incomingData.ReadInteger
    BarTime = incomingData.ReadInteger
    BarAccion = incomingData.ReadInteger
    
    charlist(CharIndex).BarTime = 0
    charlist(CharIndex).BarAccion = BarAccion
    charlist(CharIndex).MaxBarTime = BarTime
    
    Exit Sub

HandleBarFx_Err:
    Call LogError(Err.number, Err.Description, "Protocol_Handler.HandleBarFx", Erl)
    
    
End Sub
 
Public Sub HandlePrivilegios()
'***************************************************
'Author: Lorwik
'Last Modification: 28/10/2023
'
'***************************************************
    On Error GoTo errhandler
    
    Call incomingData.ReadByte
    
    CurrentUser.esGM = incomingData.ReadBoolean

    If CurrentUser.esGM Then
        frmMain.lblPanelGM.Visible = True
        frmMain.lblBuscarNpc.Visible = True
        frmMain.lblInvisible.Visible = True
    Else
        frmMain.lblPanelGM.Visible = False
        frmMain.lblBuscarNpc.Visible = False
        frmMain.lblInvisible.Visible = False
    End If
    Exit Sub
errhandler:

    Call LogError(Err.number, Err.Description, "Protocol_Handler.HandlePrivilegios", Erl)
End Sub
