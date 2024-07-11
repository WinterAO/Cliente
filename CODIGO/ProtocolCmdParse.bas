Attribute VB_Name = "ProtocolCmdParse"
'Argentum Online
'
'Copyright (C) 2006 Juan Martin Sotuyo Dodero (Maraxus)
'Copyright (C) 2006 Alejandro Santos (AlejoLp)

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

Option Explicit

Public Enum eNumber_Types
    ent_byte
    ent_Integer
    ent_Long
    ent_Trigger
End Enum

Public Sub AuxWriteWhisper(ByVal UserName As String, ByVal Mensaje As String)
'***************************************************
'Author: Unknown
'Last Modification: 03/12/2010
'03/12/2010: Enanoh - Ahora se envia el nick en vez del index del usuario.
'***************************************************
    If LenB(UserName) = 0 Then Exit Sub
    
    
    If (InStrB(UserName, "+") <> 0) Then
        UserName = Replace$(UserName, "+", " ")
    End If
    
    UserName = UCase$(UserName)
    
    Call WriteWhisper(UserName, Mensaje)
    
End Sub

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.

Public Sub ParseUserCommand(ByVal RawCommand As String)
    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modification: 16/11/2009
    'Interpreta, valida y ejecuta el comando ingresado
    '26/03/2009: ZaMa - Flexibilizo la cantidad de parametros de /nene,  /onlinemap y /telep
    '16/11/2009: ZaMa - Ahora el /ct admite radio
    '18/09/2010: ZaMa - Agrego el comando /mod username vida xxx
    '***************************************************
    Dim TmpArgos()         As String
    
    Dim Comando            As String
    Dim ArgumentosAll()    As String
    Dim ArgumentosRaw      As String
    Dim Argumentos2()      As String
    Dim Argumentos3()      As String
    Dim Argumentos4()      As String
    Dim CantidadArgumentos As Long
    Dim notNullArguments   As Boolean
    
    Dim tmpArr()           As String
    Dim tmpInt             As Integer
    
    ' TmpArgs: Un array de a lo sumo dos elementos,
    ' el primero es el comando (hasta el primer espacio)
    ' y el segundo elemento es el resto. Si no hay argumentos
    ' devuelve un array de un solo elemento
    TmpArgos = Split(RawCommand, " ", 2)
    
    Comando = Trim$(UCase$(TmpArgos(0)))
    
    If UBound(TmpArgos) > 0 Then
        ' El string en crudo que este despues del primer espacio
        ArgumentosRaw = TmpArgos(1)
        
        'veo que los argumentos no sean nulos
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        
        ' Un array separado por blancos, con tantos elementos como
        ' se pueda
        ArgumentosAll = Split(TmpArgos(1), " ")
        
        ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        
        ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
        ' 2, 3 y 4 elementos respectivamente. Eso significa
        ' que pueden tener menos, por lo que es imperativo
        ' preguntar por CantidadArgumentos.
        
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0
    End If
    
    ' Sacar cartel APESTA!! (y es ilogico, estas diciendo una pausa/espacio  :rolleyes: )
    If LenB(Comando) = 0 Then Comando = " "
    
    If Left$(Comando, 1) = "/" Then
        ' Comando normal
        
        Select Case Comando

            Case "/PREPARADO"
                Call WritedueloSet(50)
                
            Case "/ONLINE"
                Call WriteOnline
                
            Case "/INVOCAR"
                Call WriteInvocar
                
            Case "/SUBASTA"
                Call WriteConsultaSubasta
           
            Case "/OFERTAR"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Long) Then
                        Call WriteOfertarSubasta(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /Ofertar oferta.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /Ofertar oferta.")
                End If
                
            Case "/FADD"

                If notNullArguments Then
                    Call WriteAddAmigo(ArgumentosRaw, 2)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("FRM_LISTAAMIGOS_PARAMETRO").item("TEXTO") & " " & JsonLanguage.item("FRM_LISTAAMIGOS_FADD").item("TEXTO"))

                End If
 
            Case "/FMSG"

                If notNullArguments Then
                    Call WriteMsgAmigo(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("FRM_LISTAAMIGOS_PARAMETRO").item("TEXTO") & " " & JsonLanguage.item("FRM_LISTAAMIGOS_FMSG").item("TEXTO"))
                End If

            Case "/FON"
                Call WriteOnAmigo
  
            Case "/DISCORD"

                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WriteDiscord(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                
            Case "/SALIR"

                If CurrentUser.UserParalizado Then 'Inmo

                    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NO_SALIR").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteQuit
                
            Case "/SALIRCLAN"
                Call WriteGuildLeave
                
            Case "/BALANCE"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteRequestAccountState
                
            Case "/QUIETO"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WritePetStand
                
            Case "/ACOMPANAR"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WritePetFollow
                
            Case "/LIBERAR"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteReleasePet
                
            Case "/ENTRENAR"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteTrainList
                
            Case "/DESCANSAR"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteRest
                
            Case "/MEDITAR"

                If CurrentUser.UserMinMAN = CurrentUser.UserMaxMAN Then Exit Sub
                
                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteMeditate
        
            Case "/CONSULTA"
                Call WriteConsultation
            
            Case "/RESUCITAR"
                Call WriteResucitate
                
            Case "/CURAR"
                Call WriteHeal
                              
            Case "/EST"
                Call WriteRequestStats
            
            Case "/AYUDA"
                Call WriteHelp
                
            Case "/COMERCIAR"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                
                ElseIf Comerciando Then 'Comerciando

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_COMERCIANDO").item("TEXTO"), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteCommerceStart
                
            Case "/BOVEDA"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                Call WriteBankStart
                
            Case "/ENLISTAR"
                Call WriteEnlist
                    
            Case "/INFORMACION"
                Call WriteInformation
                
            Case "/RECOMPENSA"
                Call WriteReward
                
            Case "/MOTD"
                Call WriteRequestMOTD
                
            Case "/UPTIME"
                Call WriteUpTime
                
            Case "/SALIRGRUPO"
                Call WritePartyLeave
            
            Case "/COMPARTIRNPC"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteShareNpc
                
            Case "/NOCOMPARTIRNPC"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteStopSharingNpc
                
            Case "/ENCUESTA"

                If CantidadArgumentos = 0 Then
                    ' Version sin argumentos: Inquiry
                    Call WriteInquiry
                Else

                    ' Version con argumentos: InquiryVote
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_byte) Then
                        Call WriteInquiryVote(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_ENCUESTA").item("TEXTO"))
                    End If
                End If
        
            Case "/CMSG"

                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WriteGuildMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
        
            Case "/PMSG"

                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WritePartyMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If

            Case "/CENTINELA"

                If notNullArguments Then
                    Call WriteCentinelReport(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CENTINELA").item("TEXTO"))
                End If
        
            Case "/ONLINECLAN"
                Call WriteGuildOnline
                
            Case "/ONLINEPARTY"
                Call WritePartyOnline
                
            Case "/BMSG"

                If notNullArguments Then
                    Call WriteCouncilMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                End If
                
            Case "/ROL"

                If notNullArguments Then
                    Call WriteRoleMasterRequest(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_ASK").item("TEXTO"))
                End If
                
            Case "/GM"
                frmGM.Show vbModeless, frmMain
                
            Case "/DESC"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteChangeDescription(ArgumentosRaw)
            
            Case "/VOTO"

                If notNullArguments Then
                    Call WriteGuildVote(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /voto NICKNAME.")
                End If
               
            Case "/PENAS"

                If notNullArguments Then
                    Call WritePunishments(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /penas NICKNAME.")
                End If

            Case "/APOSTAR"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteGamble(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CANTIDAD_INCORRECTA").item("TEXTO") & " /apostar CANTIDAD.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /apostar CANTIDAD.")
                End If
                
            Case "/RETIRARFACCION"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                Call WriteLeaveFaction
                
            Case "/RETIRAR"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If
                
                If notNullArguments Then

                    ' Version con argumentos: BankExtractGold
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankExtractGold(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CANTIDAD_INCORRECTA").item("TEXTO") & " /retirar CANTIDAD.")
                    End If
                End If

            Case "/DEPOSITAR"

                If CurrentUser.UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                    End With
                    Exit Sub
                End If

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankDepositGold(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_CANTIDAD_INCORRECTA").item("TEXTO") & " /depositar CANTIDAD.")
                    End If
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /depositar CANTIDAD.")
                End If
                
            Case "/DENUNCIAR"

                If notNullArguments Then
                    Call WriteDenounce(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Formule su denuncia.")
                End If
                
            Case "/FUNDARCLAN"

                If CurrentUser.UserLvl >= 25 Then
                    Call WriteGuildFundate
                Else
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FUNDAR_CLAN").item("TEXTO"))
                End If
            
            Case "/ECHARPARTY"

                If notNullArguments Then
                    Call WritePartyKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /echarparty NICKNAME.")
                End If
                
            Case "/PARTYLIDER"

                If notNullArguments Then
                    Call WritePartySetLeader(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /partylider NICKNAME.")
                End If
                
            Case "/ACCEPTPARTY"

                If notNullArguments Then
                    Call WritePartyAcceptMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /acceptparty NICKNAME.")
                End If
                
            Case "/MONITOR"
                frmMonitor.Show
                    
                '
                ' BEGIN GM COMMANDS
                '
            
            Case "/BUSCAR"
                If CurrentUser.esGM Then
                    Call WriteGMPanel(1)
                End If

            Case "/LIMPIARMUNDO"

                If CurrentUser.esGM Then
                    Call WriteLimpiarMundo
                End If

            Case "/GMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteGMMessage(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/SHOWNAME"

                If CurrentUser.esGM Then
                    Call WriteShowName
                End If

            Case "/ONLINEREAL"

                If CurrentUser.esGM Then
                    Call WriteOnlineRoyalArmy
                End If

            Case "/ONLINECAOS"

                If CurrentUser.esGM Then
                    Call WriteOnlineChaosLegion
                End If

            Case "/IRCERCA"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteGoNearby(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ircerca NICKNAME.")
                    End If
                End If

            Case "/REM"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteComment(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_COMENTARIO").item("TEXTO"))
                    End If
                End If

            Case "/HORA"

                If CurrentUser.esGM Then
                    Call Protocol_Write.WriteServerTime
                End If

            Case "/DONDE"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteWhere(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /donde NICKNAME.")
                    End If
                End If

            Case "/NENE"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                            Call WriteCreaturesInMap(ArgumentosRaw)
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_MAPA_INCORRECTO").item("TEXTO") & " /nene MAPA.")
                        End If
                    Else
                        ' Por defecto, toma el mapa en el que está
                        Call WriteCreaturesInMap(CurrentUser.UserMap)
                    End If
                End If

            Case "/TELEPLOC"

                If CurrentUser.esGM Then
                    Call WriteWarpMeToTarget
                End If

            Case "/TELEP"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 4 Then
                        If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Integer) Then
                            Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3), False)
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telep NICKNAME MAPA X Y.")
                        End If
                    ElseIf CantidadArgumentos = 3 Then

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) Then
                            ' Por defecto, si no se indica el nombre, se teletransporta el mismo usuario
                            Call WriteWarpChar("YO", ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), False)
                        ElseIf ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) Then
                            ' Por defecto, si no se indica el mapa, se teletransporta al mismo donde está el usuario
                            Call WriteWarpChar(ArgumentosAll(0), CurrentUser.UserMap, ArgumentosAll(1), ArgumentosAll(2), False)
                        Else
                            ' No uso ningún formato por defecto
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telep NICKNAME MAPA X Y.")
                        End If
                    ElseIf CantidadArgumentos = 2 Then

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            ' Por defecto, se considera que se quiere únicamente cambiar las coordenadas del usuario, en el mismo mapa
                            Call WriteWarpChar("YO", CurrentUser.UserMap, ArgumentosAll(0), ArgumentosAll(1), False)
                        Else
                            ' No uso ningún formato por defecto
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telep NICKNAME MAPA X Y.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /telep NICKNAME MAPA X Y.")
                    End If
                End If
                
            Case "/TELEPC"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 4 Then
                        If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Integer) Then
                            Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3), True)
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telepc NICKNAME CUADRANTE X Y.")
                        End If
                    ElseIf CantidadArgumentos = 3 Then

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) Then
                            ' Por defecto, si no se indica el nombre, se teletransporta el mismo usuario
                            Call WriteWarpChar("YO", ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), True)
                        ElseIf ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) Then
                            ' Por defecto, si no se indica el mapa, se teletransporta al mismo donde está el usuario
                            Call WriteWarpChar(ArgumentosAll(0), CurrentUser.UserMap, ArgumentosAll(1), ArgumentosAll(2), True)
                        Else
                            ' No uso ningún formato por defecto
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telepc NICKNAME CUADRANTE X Y.")
                        End If
                    ElseIf CantidadArgumentos = 2 Then

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            ' Por defecto, se considera que se quiere únicamente cambiar las coordenadas del usuario, en el mismo mapa
                            Call WriteWarpChar("YO", CurrentUser.UserMap, ArgumentosAll(0), ArgumentosAll(1), True)
                        Else
                            ' No uso ningún formato por defecto
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /telepc NICKNAME CUADRANTE X Y.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /telepc NICKNAME CUADRANTE X Y.")
                    End If
                End If

            Case "/SILENCIAR"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteSilence(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /silenciar NICKNAME.")
                    End If
                End If

            Case "/SHOW"

                If CurrentUser.esGM Then
                    If notNullArguments Then

                        Select Case UCase$(ArgumentosAll(0))

                            Case "SOS"
                                Call WriteSOSShowList

                            Case "INT"
                                Call WriteShowServerForm

                            Case "DENUNCIAS"
                                Call WriteShowDenouncesList
                        End Select
                    End If
                End If

            Case "/DENUNCIAS"

                If CurrentUser.esGM Then
                    Call WriteEnableDenounces
                End If

            Case "/IRA"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteGoToChar(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ira NICKNAME.")
                    End If
                End If

            Case "/INVISIBLE"

                If CurrentUser.esGM Then
                    Call WriteInvisible
                End If

            Case "/PANELGM"

                If CurrentUser.esGM Then
                    Call WriteGMPanel(0)
                End If

            Case "/TRABAJANDO"

                If CurrentUser.esGM Then
                    Call WriteWorking
                End If

            Case "/OCULTANDO"

                If CurrentUser.esGM Then
                    Call WriteHiding
                End If
                
            Case "/CARCEL"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        tmpArr = Split(ArgumentosRaw, "@")

                        If UBound(tmpArr) = 2 Then
                            If ValidNumber(tmpArr(2), eNumber_Types.ent_byte) Then
                                Call WriteJail(tmpArr(0), tmpArr(1), tmpArr(2))
                            Else
                                ' No es numérico
                                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_TIEMPO_INCORRECTO").item("TEXTO") & " /carcel NICKNAME@MOTIVO@TIEMPO.")
                            End If
                        Else
                            ' Faltan los parámetros con el formato propio
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /carcel NICKNAME@MOTIVO@TIEMPO.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /carcel NICKNAME@MOTIVO@TIEMPO.")
                    End If
                End If

            Case "/RMATA"

                If CurrentUser.esGM Then
                    Call WriteKillNPC
                End If

            Case "/ADVERTENCIA"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        tmpArr = Split(ArgumentosRaw, "@", 2)

                        If UBound(tmpArr) = 1 Then
                            Call WriteWarnUser(tmpArr(0), tmpArr(1))
                        Else
                            ' Faltan los parámetros con el formato propio
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /advertencia NICKNAME@MOTIVO.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /advertencia NICKNAME@MOTIVO.")
                    End If
                End If

            Case "/MOD"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 3 Then

                        Select Case UCase$(ArgumentosAll(1))

                            Case "BODY"
                                tmpInt = eEditOptions.eo_Body

                            Case "HEAD"
                                tmpInt = eEditOptions.eo_Head

                            Case "ORO"
                                tmpInt = eEditOptions.eo_Gold

                            Case "LEVEL"
                                tmpInt = eEditOptions.eo_Level

                            Case "SKILLS"
                                tmpInt = eEditOptions.eo_Skills

                            Case "CLASE"
                                tmpInt = eEditOptions.eo_Class

                            Case "EXP"
                                tmpInt = eEditOptions.eo_Experience

                            Case "CRI"
                                tmpInt = eEditOptions.eo_CriminalsKilled

                            Case "CIU"
                                tmpInt = eEditOptions.eo_CiticensKilled

                            Case "NOB"
                                tmpInt = eEditOptions.eo_Nobleza

                            Case "ASE"
                                tmpInt = eEditOptions.eo_Asesino

                            Case "SEX"
                                tmpInt = eEditOptions.eo_Sex

                            Case "RAZA"
                                tmpInt = eEditOptions.eo_Raza

                            Case "AGREGAR"
                                tmpInt = eEditOptions.eo_addGold

                            Case "VIDA"
                                tmpInt = eEditOptions.eo_Vida

                            Case "POSS"
                                tmpInt = eEditOptions.eo_Poss

                            Case "SPEED"
                                tmpInt = eEditOptions.eo_Speed

                            Case "EXPPVP"
                                tmpInt = eEditOptions.eo_ExperiencePVP

                            Case Else
                                tmpInt = -1
                        End Select

                        If tmpInt > 0 Then
                            If CantidadArgumentos = 3 Then
                                Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), vbNullString)
                            Else
                                Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), ArgumentosAll(3))
                            End If
                        Else
                            ' Avisar que no existe el comando
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_COMANDO_INCORRECTO").item("TEXTO"))
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO"))
                    End If
                End If

            Case "/INFO"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRequestCharInfo(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /info NICKNAME.")
                    End If
                End If
                
            Case "/STAT"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRequestCharStats(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /stat NICKNAME.")
                    End If
                End If

            Case "/BAL"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRequestCharGold(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /bal NICKNAME.")
                    End If
                End If

            Case "/INV"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRequestCharInventory(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /inv NICKNAME.")
                    End If
                End If

            Case "/BOV"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRequestCharBank(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /bov NICKNAME.")
                    End If
                End If

            Case "/SKILLS"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRequestCharSkills(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /skills NICKNAME.")
                    End If
                End If

            Case "/REVIVIR"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteReviveChar(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /revivir NICKNAME.")
                    End If
                End If

            Case "/ONLINEGM"

                If CurrentUser.esGM Then
                    Call WriteOnlineGM
                End If

            Case "/ONLINEMAP"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            Call WriteOnlineMap(ArgumentosAll(0))
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_MAPA_INCORRECTO").item("TEXTO") & " /ONLINEMAP")
                        End If
                    Else
                        Call WriteOnlineMap(CurrentUser.UserMap)
                    End If
                End If

            Case "/PERDON"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteForgive(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /perdon NICKNAME.")
                    End If
                End If

            Case "/ECHAR"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteKick(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /echar NICKNAME.")
                    End If
                End If

            Case "/EJECUTAR"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteExecute(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ejecutar NICKNAME.")
                    End If
                End If
                
            Case "/BAN"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        tmpArr = Split(ArgumentosRaw, "@", 2)

                        If UBound(tmpArr) = 1 Then
                            Call WriteBanChar(tmpArr(0), tmpArr(1))
                        Else
                            ' Faltan los parámetros con el formato propio
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /ban NICKNAME@MOTIVO.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ban NICKNAME@MOTIVO.")
                    End If
                End If

            Case "/UNBAN"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteUnbanChar(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /unban NICKNAME.")
                    End If
                End If

            Case "/SEGUIR"

                If CurrentUser.esGM Then
                    Call WriteNPCFollow
                End If

            Case "/SUM"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteSummonChar(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /sum NICKNAME.")
                    End If
                End If

            Case "/CC"

                If CurrentUser.esGM Then
                    Call WriteSpawnListRequest
                End If

            Case "/RESETINV"

                If CurrentUser.esGM Then
                    Call WriteResetNPCInventory
                End If

            Case "/RMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteServerMessage(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/MAPMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteMapMessage(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/NICK2IP"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteNickToIP(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /nick2ip NICKNAME.")
                    End If
                End If

            Case "/IP2NICK"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If validipv4str(ArgumentosRaw) Then
                            Call WriteIPToNick(str2ipv4l(ArgumentosRaw))
                        Else
                            ' No es una IP
                            Call ShowConsoleMsg("IP incorrecta. Utilice /ip2nick IP.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ip2nick IP.")
                    End If
                End If

            Case "/ONCLAN"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteGuildOnlineMembers(ArgumentosRaw)
                    Else
                        ' Avisar sintaxis incorrecta
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_ONCLAN").item("TEXTO"))
                    End If
                End If

            Case "/CT"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 3 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_byte) Then
                            If CantidadArgumentos = 3 Then
                                Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                            Else

                                If ValidNumber(ArgumentosAll(3), eNumber_Types.ent_byte) Then
                                    Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                                Else
                                    ' No es numérico
                                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /ct MAPA X Y RADIO(Opcional).")
                                End If
                            End If
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /ct MAPA X Y RADIO(Opcional).")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ct MAPA X Y RADIO(Opcional).")
                    End If
                End If

            Case "/DT"

                If CurrentUser.esGM Then
                    Call WriteTeleportDestroy
                End If

            Case "/DE"

                If CurrentUser.esGM Then
                    Call WriteExitDestroy
                End If

            Case "/METEO"

                If CurrentUser.esGM Then
                    If notNullArguments = False Then
                        Call WriteMeteoToggle
                    Else

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) Then
                            Call WriteMeteoToggle(ArgumentosAll(0))
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /METEO 0: Random, 1: Lluvia, 2: Niebla, 3: Niebla + Lluvia.")
                        End If
                    End If
                End If

            Case "/SETDESC"

                If CurrentUser.esGM Then
                    Call WriteSetCharDescription(ArgumentosRaw)
                End If

            Case "/FORCEMUSICMAP"

                If CurrentUser.esGM Then
                    If notNullArguments Then

                        ' Elegir el mapa es opcional
                        If CantidadArgumentos = 1 Then
                            If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) Then
                                ' Evitamos un mapa nulo para que tome el del usuario.
                                Call WriteForceMUSICToMap(ArgumentosAll(0), 0)
                            Else
                                ' No es numérico
                                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /forcemusicmap MUSIC MAPA")
                            End If
                        Else

                            If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                                Call WriteForceMUSICToMap(ArgumentosAll(0), ArgumentosAll(1))
                            Else
                                ' No es numérico
                                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /forcemusicmap MUSIC MAPA")
                            End If
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg("Utilice /forcemusicmap MUSIC MAPA")
                    End If
                End If

            Case "/FORCEWAVMAP"

                If CurrentUser.esGM Then
                    If notNullArguments Then

                        ' Elegir la posición es opcional
                        If CantidadArgumentos = 1 Then
                            If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) Then
                                ' Evitamos una posición nula para que tome la del usuario.
                                Call WriteForceWAVEToMap(ArgumentosAll(0), 0, 0, 0)
                            ElseIf CantidadArgumentos = 4 Then

                                If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_byte) Then
                                    Call WriteForceWAVEToMap(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                                Else
                                    ' No es numérico
                                    Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                                End If
                            Else
                                ' Avisar que falta el parámetro
                                Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                            End If
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")
                    End If
                End If

            Case "/REALMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRoyalArmyMessage(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/CAOSMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteChaosLegionMessage(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/CIUMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteCitizenMessage(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/CRIMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteCriminalMessage(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/TALKAS"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteTalkAsNPC(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/MASSDEST"

                If CurrentUser.esGM Then
                    Call WriteDestroyAllItemsInArea
                End If

            Case "/ACEPTCONSE"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteAcceptRoyalCouncilMember(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /aceptconse NICKNAME.")
                    End If
                End If

            Case "/ACEPTCONSECAOS"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteAcceptChaosCouncilMember(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /aceptconsecaos NICKNAME.")
                    End If
                End If

            Case "/PISO"

                If CurrentUser.esGM Then
                    Call WriteItemsInTheFloor
                End If

            Case "/ESTUPIDO"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteMakeDumb(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /estupido NICKNAME.")
                    End If
                End If

            Case "/NOESTUPIDO"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteMakeDumbNoMore(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /noestupido NICKNAME.")
                    End If
                End If

            Case "/DUMPSECURITY"

                If CurrentUser.esGM Then
                    Call WriteDumpIPTables
                End If

            Case "/KICKCONSE"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteCouncilKick(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /kickconse NICKNAME.")
                    End If
                End If

            Case "/TRIGGER"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Trigger) Then
                            Call WriteSetTrigger(ArgumentosRaw)
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /trigger NUMERO.")
                        End If
                    Else
                        ' Versión sin parámetro
                        Call WriteAskTrigger
                    End If
                End If

            Case "/BANIPLIST"

                If CurrentUser.esGM Then
                    Call WriteBannedIPList
                End If

            Case "/BANIPRELOAD"

                If CurrentUser.esGM Then
                    Call WriteBannedIPReload
                End If
                
            Case "/MIEMBROSCLAN"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteGuildMemberList(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /miembrosclan GUILDNAME.")
                    End If
                End If

            Case "/BANCLAN"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteGuildBan(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /banclan GUILDNAME.")
                    End If
                End If

            Case "/BANIP"

                If CurrentUser.esGM Then
                    If CantidadArgumentos >= 2 Then
                        If validipv4str(ArgumentosAll(0)) Then
                            Call WriteBanIP(True, str2ipv4l(ArgumentosAll(0)), vbNullString, Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                        Else
                            ' No es una IP, es un nick
                            Call WriteBanIP(False, str2ipv4l("0.0.0.0"), ArgumentosAll(0), Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /banip IP motivo o /banip nick motivo.")
                    End If
                End If

            Case "/UNBANIP"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If validipv4str(ArgumentosRaw) Then
                            Call WriteUnbanIP(str2ipv4l(ArgumentosRaw))
                        Else
                            ' No es una IP
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /unbanip IP.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /unbanip IP.")
                    End If
                End If

            Case "/CI"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos = 2 Then
                        If IsNumeric(ArgumentosAll(0)) And IsNumeric(ArgumentosAll(1)) Then
                            Call WriteCreateItem(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_OBJETO_INCORRECTO").item("TEXTO") & " /CI " & JsonLanguage.item("OBJETO").item("TEXTO") & " " & JsonLanguage.item("CANTIDAD").item("TEXTO"))
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ci OBJETO CANTIDAD.")
                    End If
                End If

            Case "/DEST"

                If CurrentUser.esGM Then
                    Call WriteDestroyItems
                End If

            Case "/NOCAOS"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteChaosLegionKick(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /nocaos NICKNAME.")
                    End If
                End If

            Case "/NOREAL"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRoyalArmyKick(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /noreal NICKNAME.")
                    End If
                End If

            Case "/FORCEMUSIC"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) Then
                            Call WriteForceMUSICAll(ArgumentosAll(0))
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_MIDI_INCORRECTO").item("TEXTO") & " /forceusic MUSIC.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /forcemusic MUSIC.")
                    End If
                End If

            Case "/FORCEWAV"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) Then
                            Call WriteForceWAVEAll(ArgumentosAll(0))
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_WAV_INCORRECTO").item("TEXTO") & " /forcewav WAV.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /forcewav WAV.")
                    End If
                End If
    
            Case "/MODIFICARPENA"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        tmpArr = Split(ArgumentosRaw, "@", 3)

                        If UBound(tmpArr) = 2 Then
                            Call WriteRemovePunishment(tmpArr(0), tmpArr(1), tmpArr(2))
                        Else
                            ' Faltan los parámetros con el formato propio
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /modificarpena NICK@PENA@NuevaPena.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /modificarpena NICK@PENA@NuevaPena.")
                    End If
                End If

            Case "/BLOQ"

                If CurrentUser.esGM Then
                    Call WriteTileBlockedToggle
                End If

            Case "/MATA"

                If CurrentUser.esGM Then
                    Call WriteKillNPCNoRespawn
                End If

            Case "/MASSKILL"

                If CurrentUser.esGM Then
                    Call WriteKillAllNearbyNPCs
                End If

            Case "/LASTIP"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteLastIP(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /lastip NICKNAME.")
                    End If
                End If

            Case "/MOTDCAMBIA"

                If CurrentUser.esGM Then
                    Call WriteChangeMOTD
                End If

            Case "/SMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteSystemMessage(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_INPUT_MSJ").item("TEXTO"))
                    End If
                End If

            Case "/ACC"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            Call WriteCreateNPC(ArgumentosAll(0), False)
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NPC_INCORRECTO").item("TEXTO") & " /acc NPC.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /acc NPC.")
                    End If
                End If

            Case "/RACC"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            Call WriteCreateNPC(ArgumentosAll(0), True)
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_NPC_INCORRECTO").item("TEXTO") & " /racc NPC.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /racc NPC.")
                    End If
                End If

            Case "/AI" ' 1 - 4

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 2 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteImperialArmour(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /ai ARMADURA OBJETO.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ai ARMADURA OBJETO.")
                    End If
                End If

            Case "/AC" ' 1 - 4

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 2 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteChaosArmour(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /ac ARMADURA OBJETO.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ac ARMADURA OBJETO.")
                    End If
                End If

            Case "/NAVE"

                If CurrentUser.esGM Then
                    Call WriteNavigateToggle
                End If

            Case "/HABILITAR"

                If CurrentUser.esGM Then
                    Call WriteServerOpenToUsersToggle
                End If

            Case "/APAGAR"

                If CurrentUser.esGM Then
                    Call WriteTurnOffServer
                End If

            Case "/CONDEN"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteTurnCriminal(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /conden NICKNAME.")
                    End If
                End If

            Case "/RAJAR"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteResetFactions(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /rajar NICKNAME.")
                    End If
                End If

            Case "/RAJARCLAN"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRemoveCharFromGuild(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /rajarclan NICKNAME.")
                    End If
                End If

            Case "/LASTEMAIL"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteRequestCharMail(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /lastemail NICKNAME.")
                    End If
                End If

            Case "/ANAME"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        tmpArr = Split(ArgumentosRaw, "@", 2)

                        If UBound(tmpArr) = 1 Then
                            Call WriteAlterName(tmpArr(0), tmpArr(1))
                        Else
                            ' Faltan los parámetros con el formato propio
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /aname ORIGEN@DESTINO.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /aname ORIGEN@DESTINO.")
                    End If
                End If

            Case "/SLOT"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        tmpArr = Split(ArgumentosRaw, "@", 2)

                        If UBound(tmpArr) = 1 Then
                            If ValidNumber(tmpArr(1), eNumber_Types.ent_byte) Then
                                Call WriteCheckSlot(tmpArr(0), tmpArr(1))
                            Else
                                ' Faltan o sobran los parámetros con el formato propio
                                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /slot NICK@SLOT.")
                            End If
                        Else
                            ' Faltan o sobran los parámetros con el formato propio
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /slot NICK@SLOT.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /slot NICK@SLOT.")
                    End If
                End If

            Case "/CENTINELAACTIVADO"

                If CurrentUser.esGM Then
                    Call WriteToggleCentinelActivated
                End If

            Case "/CREARPRETORIANOS"

                If CurrentUser.esGM Then
                    If CantidadArgumentos = 3 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) Then
                            Call WriteCreatePretorianClan(Val(ArgumentosAll(0)), Val(ArgumentosAll(1)), Val(ArgumentosAll(2)))
                        Else
                            ' Faltan o sobran los parámetros con el formato propio
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /CrearPretorianos MAPA X Y.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /CrearPretorianos MAPA X Y.")
                    End If
                End If

            Case "/ELIMINARPRETORIANOS"

                If CurrentUser.esGM Then
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            Call WriteDeletePretorianClan(Val(ArgumentosAll(0)))
                        Else
                            ' Faltan o sobran los parámetros con el formato propio
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /EliminarPretorianos MAPA.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /EliminarPretorianos MAPA.")
                    End If
                End If

            Case "/DOBACKUP"

                If CurrentUser.esGM Then
                    Call WriteDoBackup
                End If

            Case "/SHOWCMSG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteShowGuildMessages(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /showcmsg GUILDNAME.")
                    End If
                End If

            Case "/GUARDAMAPA"

                If CurrentUser.esGM Then
                    Call WriteSaveMap
                End If

            Case "/MODZONA" ' PK, BACKUP

                If CurrentUser.esGM Then
                    If CantidadArgumentos > 1 Then

                        Select Case UCase$(ArgumentosAll(0))

                            Case "PK" ' "/MODZona PK"
                                Call WriteChangeZonaPK(ArgumentosAll(1) = "1")

                            Case "BACKUP" ' "/MODZona BACKUP"
                                Call WriteChangeZonaBackup(ArgumentosAll(1) = "1")

                            Case "RESTRINGIR" '/MODZona RESTRINGIR
                                Call WriteChangeZonaRestricted(ArgumentosAll(1))

                            Case "MAGIASINEFECTO" '/MODZona MAGIASINEFECTO
                                Call WriteChangeZonaNoMagic(ArgumentosAll(1) = "1")

                            Case "INVISINEFECTO" '/MODZona INVISINEFECTO
                                Call WriteChangeZonaNoInvi(ArgumentosAll(1) = "1")

                            Case "RESUSINEFECTO" '/MODZona RESUSINEFECTO
                                Call WriteChangeZonaNoResu(ArgumentosAll(1) = "1")

                            Case "TERRENO" '/MODZona TERRENO
                                Call WriteChangeZonaLand(ArgumentosAll(1))

                            Case "ZONA" '/MODZona ZONA
                                Call WriteChangeZonaZone(ArgumentosAll(1))

                            Case "ROBONPC" '/MODZona ROBONPC
                                Call WriteChangeZonaStealNpc(ArgumentosAll(1) = "1")

                            Case "OCULTARSINEFECTO" '/MODZona OCULTARSINEFECTO
                                Call WriteChangeZonaNoOcultar(ArgumentosAll(1) = "1")

                            Case "INVOCARSINEFECTO" '/MODZona INVOCARSINEFECTO
                                Call WriteChangeZonaNoInvocar(ArgumentosAll(1) = "1")
                        End Select
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " : PK, BACKUP, RESTRINGIR, MAGIASINEFECTO, INVISINEFECTO, RESUSINEFECTO, TERRENO, ZONA")
                    End If
                End If

            Case "/GRABAR"

                If CurrentUser.esGM Then
                    Call WriteSaveChars
                End If

            Case "/BORRAR"

                If CurrentUser.esGM Then
                    If notNullArguments Then

                        Select Case UCase$(ArgumentosAll(0))

                            Case "SOS" ' "/BORRAR SOS"
                                Call WriteCleanSOS
                        End Select
                    End If
                End If

            Case "/NOCHE"

                If CurrentUser.esGM Then
                    Call WriteNight
                End If

            Case "/ECHARTODOSPJS"

                If CurrentUser.esGM Then
                    Call WriteKickAllChars
                End If

            Case "/RELOADNPCS"

                If CurrentUser.esGM Then
                    Call WriteReloadNPCs
                End If

            Case "/RELOADSINI"

                If CurrentUser.esGM Then
                    Call WriteReloadServerIni
                End If

            Case "/RELOADHECHIZOS"

                If CurrentUser.esGM Then
                    Call WriteReloadSpells
                End If

            Case "/RELOADOBJ"

                If CurrentUser.esGM Then
                    Call WriteReloadObjects
                End If

            Case "/REINICIAR"

                If CurrentUser.esGM Then
                    Call WriteRestart
                End If

            Case "/AUTOUPDATE"

                If CurrentUser.esGM Then
                    Call WriteResetAutoUpdate
                End If

            Case "/CHATCOLOR"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 3 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_byte) Then
                            Call WriteChatColor(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                        Else
                            ' No es numérico
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /chatcolor R G B.")
                        End If
                    ElseIf Not notNullArguments Then ' Volver al valor predeterminado
                        Call WriteChatColor(0, 255, 0)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /chatcolor R G B.")
                    End If
                End If
            
            Case "/IGNORADO"

                If CurrentUser.esGM Then
                    Call WriteIgnored
                End If

            Case "/PING"

                If CurrentUser.esGM Then
                    Call WritePing
                End If

            Case "/RETOS"

                If CurrentUser.esGM Then
                    Call FrmRetos.Show(vbModeless, frmMain)
                End If

            Case "/CERRARCLAN"

                If CurrentUser.esGM Then
                    Call WriteCloseGuild
                End If

            Case "/ACEPTAR"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteFightAccept(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /ACEPTAR NICKNAME.")
                    End If
                End If

            Case "/QUEST"

                If CurrentUser.esGM Then
                    Call WriteQuest
                End If

            Case "/SETINIVAR"

                If CurrentUser.esGM Then
                    If CantidadArgumentos = 3 Then
                        ArgumentosAll(2) = Replace(ArgumentosAll(2), "+", " ")
                        Call WriteSetIniVar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FORMATO_INCORRECTO").item("TEXTO") & " /SETINIVAR LLAVE CLAVE VALOR")
                    End If
                End If

            Case "/CVC"

                If CurrentUser.esGM Then
                    Call WriteEnviaCvc
                End If

            Case "/ACVC"

                If CurrentUser.esGM Then
                    Call WriteAceptarCvc
                End If

            Case "/IRCVC"

                If CurrentUser.esGM Then
                    Call WriteIrCvc
                End If

            Case "/HOGAR"

                If CurrentUser.esGM Then
                    Call WriteHome
                End If

            Case "/SETDIALOG"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteSetDialog(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /SETDIALOG DIALOGO.")
                    End If
                End If

            Case "/IMPERSONAR"

                If CurrentUser.esGM Then
                    Call WriteImpersonate
                End If

            Case "/SILENCIARGLOBAL"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteSilenciarGlobal(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg("Faltan parámetros. Utilice /SilenciarGlobal NICKNAME.")
                    End If
                End If

            Case "/TOGGLEGLOBAL"

                If CurrentUser.esGM Then
                    Call WriteToggleGlobal
                End If

            Case "/MIMETIZAR"

                If CurrentUser.esGM Then
                    Call WriteImitate
                End If

            Case "/EDITGEMS"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 2 Then
                        If Not IsNumeric(ArgumentosAll(0)) And IsNumeric(ArgumentosAll(1)) Then
                            Call WriteEditGems(ArgumentosAll(0), ArgumentosAll(1), 0)
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /EDITGEMS NICKNAME CANTIDAD.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /EDITGEMS NICKNAME CANTIDAD.")
                    End If
                End If

            Case "/SUMARGEMS"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 2 Then
                        If Not IsNumeric(ArgumentosAll(0)) And IsNumeric(ArgumentosAll(1)) Then
                            Call WriteEditGems(ArgumentosAll(0), ArgumentosAll(1), 1)
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /EDITGEMS NICKNAME CANTIDAD.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /EDITGEMS NICKNAME CANTIDAD.")
                    End If
                End If

            Case "/RESTARGEMS"

                If CurrentUser.esGM Then
                    If notNullArguments And CantidadArgumentos >= 2 Then
                        If Not IsNumeric(ArgumentosAll(0)) And IsNumeric(ArgumentosAll(1)) Then
                            Call WriteEditGems(ArgumentosAll(0), ArgumentosAll(1), 2)
                        Else
                            Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_VALOR_INCORRECTO").item("TEXTO") & " /EDITGEMS NICKNAME CANTIDAD.")
                        End If
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /EDITGEMS NICKNAME CANTIDAD.")
                    End If
                End If

            Case "/CONSULTARGEMS"

                If CurrentUser.esGM Then
                    If notNullArguments Then
                        Call WriteConsultarGems(ArgumentosRaw)
                    Else
                        ' Avisar que falta el parámetro
                        Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_FALTAN_PARAMETROS").item("TEXTO") & " /CONSULTARGEMS NICKNAME.")
                    End If
                End If
        
        End Select
        
    ElseIf Left$(Comando, 1) = ";" Then

        If notNullArguments Then
            If CurrentUser.UserEstado = 1 Then 'Muerto

                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            Call WriteGlobalChat(ArgumentosRaw)
        Else
            'Avisar que falta el parametro
            Call ShowConsoleMsg("Escriba un mensaje.")
        End If
        
    ElseIf Left$(Comando, 1) = "\" Then

        If CurrentUser.UserEstado = 1 Then 'Muerto

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
            End With
            Exit Sub
        End If
        ' Mensaje Privado
        Call AuxWriteWhisper(mid$(Comando, 2), ArgumentosRaw)
        
    ElseIf Left$(Comando, 1) = "-" Then

        If CurrentUser.UserEstado = 1 Then 'Muerto

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg(JsonLanguage.item("MENSAJE_USER_MUERTO").item("TEXTO").item(1), .Red, .Green, .Blue, .bold, .italic)
            End With
            Exit Sub
        End If
        ' Gritar
        Call WriteYell(mid$(RawCommand, 2))
        
    Else
        ' Hablar
        Call WriteTalk(RawCommand)
    End If
End Sub

''
' Show a console message.
'
' @param    Message The message to be written.
' @param    red Sets the font red color.
' @param    green Sets the font green color.
' @param    blue Sets the font blue color.
' @param    bold Sets the font bold style.
' @param    italic Sets the font italic style.

Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal Red As Integer = 255, Optional ByVal Green As Integer = 255, Optional ByVal Blue As Integer = 255, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/03/07
'
'***************************************************
    Call AddtoRichTextBox(frmMain.RecTxt, Message, Red, Green, Blue, bold, italic)
End Sub

''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, ByVal tipo As eNumber_Types) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim Minimo As Long
    Dim Maximo As Long
    
    If Not IsNumeric(Numero) Then _
        Exit Function
    
    Select Case tipo
        Case eNumber_Types.ent_byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 6
    End Select
    
    If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then _
        ValidNumber = True
End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal Ip As String) As Boolean
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 01/06/07
'
'***************************************************
    Dim tmpArr() As String
    
    tmpArr = Split(Ip, ".")
    
    If UBound(tmpArr) <> 3 Then _
        Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_byte) Or _
      Not ValidNumber(tmpArr(1), eNumber_Types.ent_byte) Or _
      Not ValidNumber(tmpArr(2), eNumber_Types.ent_byte) Or _
      Not ValidNumber(tmpArr(3), eNumber_Types.ent_byte) Then _
        Exit Function
    
    validipv4str = True
End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Private Function str2ipv4l(ByVal Ip As String) As Byte()
'***************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Bytes
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr() As String
    Dim bArr(3) As Byte
    
    tmpArr = Split(Ip, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr
End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()
'***************************************************
'Author: Lucas Tavolaro Ortuz (Tavo)
'Useful for AEMAIL BUG FIX
'Last Modification: 07/26/07
'Last Modified By: Rapsodius
'Specify Return Type as Array of Strings
'Otherwise, the default is a Variant or Array of Variants, that slows down
'the function
'***************************************************
    Dim tmpArr(0 To 1) As String
    Dim Pos As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString
    End If
    
    AEMAILSplit = tmpArr
End Function
