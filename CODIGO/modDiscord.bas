Attribute VB_Name = "modDiscord"
Option Explicit

Private Declare Sub ab1_discord_initialize Lib "vb6-discord-rich-presence.dll" (ByVal clientID As String)
Private Declare Sub ab1_discord_release Lib "vb6-discord-rich-presence.dll" ()
Private Declare Sub ab1_discord_presence_set Lib "vb6-discord-rich-presence.dll" (ByVal state As String, ByVal details As String)
Private Declare Sub ab1_discord_presence_clear Lib "vb6-discord-rich-presence.dll" ()

Public Sub Discord_Presence_Start()
    On Error GoTo ErrorHandler

    Call ab1_discord_initialize("1315686608581820527")
    Call ab1_discord_presence_set("5.0.0 - Alpha", "https://winterao.com")
    
ErrorHandler:
    Call RegistrarError(Err.number, Err.Description, "ModDiscord.Discord_Presence_Start", Erl)
End Sub

Public Sub Discord_Presence_End()
    On Error GoTo ErrorHandler

    Call ab1_discord_presence_clear
    Call ab1_discord_release
    
ErrorHandler:
    Call RegistrarError(Err.number, Err.Description, "MModDiscord.Discord_Presence_End", Erl)
End Sub
