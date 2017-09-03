VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCPS"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCpsLock"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtText"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtChat"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDatabase"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraServer"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraPrinter"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraGlobal"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraXP"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Daily Reward"
      TabPicture(3)   =   "frmServer.frx":170DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmDailyRewards"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraXP 
         Caption         =   "Bonus XP"
         Height          =   3135
         Left            =   -69960
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         Begin VB.OptionButton optGoDown 
            Caption         =   "Go Down"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   2280
            Width           =   1215
         End
         Begin VB.OptionButton optGoUp 
            Caption         =   "Go Up"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   2040
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.Timer tmrXPDec 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   960
            Top             =   2520
         End
         Begin VB.HScrollBar scrlStopAt 
            Height          =   255
            Left            =   120
            Max             =   100
            Min             =   1
            TabIndex        =   47
            Top             =   1680
            Value           =   10
            Width           =   1215
         End
         Begin VB.HScrollBar scrlXPDec 
            Height          =   255
            Left            =   120
            Max             =   100
            Min             =   1
            TabIndex        =   46
            Top             =   1080
            Value           =   1
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   2640
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdXPSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   2640
            Width           =   1215
         End
         Begin VB.HScrollBar scrlDur 
            Height          =   255
            Left            =   120
            Max             =   1000
            Min             =   1
            TabIndex        =   43
            Top             =   480
            Value           =   1
            Width           =   1215
         End
         Begin VB.Label lblExit 
            Alignment       =   2  'Center
            Caption         =   "x"
            Height          =   255
            Left            =   1080
            TabIndex        =   53
            Top             =   120
            Width           =   255
         End
         Begin VB.Label lblStopAt 
            Caption         =   "Stop At: 1.0"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblXPDec 
            Caption         =   "XP: 0.1"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblDur 
            Caption         =   "Dur: 1"
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraGlobal 
         Height          =   1455
         Left            =   -69960
         TabIndex        =   37
         Top             =   2040
         Width           =   1455
         Begin VB.HScrollBar scrlXPRate 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   41
            Top             =   1080
            Value           =   10
            Width           =   1215
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   4
            TabIndex        =   39
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblXPRate 
            Caption         =   "XP Rate: 1.0"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblAccessLevel 
            Caption         =   "Access: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraPrinter 
         Caption         =   "Printer"
         Height          =   1695
         Left            =   -69960
         TabIndex        =   33
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdRe_SaveData 
            Caption         =   "Re-Save Data"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdShowPrinter 
            Caption         =   "Show Printer"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdRe_RunData 
            Caption         =   "Re-Run Data"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame frmDailyRewards 
         Caption         =   "Daily Rewards"
         Height          =   2895
         Left            =   -74760
         TabIndex        =   21
         Top             =   360
         Width           =   6015
         Begin VB.Frame Frame1 
            Height          =   975
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   3015
            Begin VB.CommandButton cmdSave 
               Caption         =   "Save"
               Height          =   255
               Left            =   1920
               TabIndex        =   31
               Top             =   600
               Width           =   975
            End
            Begin VB.HScrollBar scrlItemReward 
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   600
               Value           =   1
               Width           =   1695
            End
            Begin VB.Label lblItemReward 
               Caption         =   "Item Reward:"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.CheckBox chkAllowDaily 
            Caption         =   "Daily Allowed?: False"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   2520
            Width           =   2535
         End
         Begin VB.CommandButton cmdResetDaily 
            Caption         =   "Set New Daily / Reset Daily"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label lblLastDate 
            Caption         =   "Unknown"
            Height          =   375
            Left            =   1440
            TabIndex        =   25
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label3 
            Caption         =   "Last Daily Set:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblDailyValue 
            Caption         =   "0"
            Height          =   255
            Left            =   1440
            TabIndex        =   23
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label Label1 
            Caption         =   "Current Value:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         Height          =   3135
         Left            =   -71880
         TabIndex        =   1
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton cmdEditPlayers 
            Caption         =   "Edit Online Players"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   2520
            Width           =   1575
         End
         Begin VB.CommandButton cmdSavePlayers 
            Caption         =   "Save Online Players"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton cmdFailsafe 
            Caption         =   "Failsafe"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CheckBox chkServerLog 
            Caption         =   "Server Log"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fraDatabase 
         Caption         =   "Reload"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   2895
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animations"
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Resources"
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Items"
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Npcs"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Shops"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Spells"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Maps"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Classes"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   6255
      End
      Begin VB.TextBox txtText 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   6255
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5318
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[Unlock]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblCPS 
         Caption         =   "CPS: 0"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
      End
      Begin VB.Menu mnuMute 
         Caption         =   "Mute"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Kill"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAllowDaily_Click()
Dim i As String

With chkAllowDaily

If .Value = 1 Then
    .Caption = "Daily Allowed: True"
    i = "True"
    Open App.Path & "\data\daily\AllowDaily.txt" For Output As #1
    Print #1, i
    Close #1
Else
    .Caption = "Daily Allowed: False"
    i = "False"
    Open App.Path & "\data\daily\AllowDaily.txt" For Output As #1
    Print #1, i
    Close #1
End If

End With

End Sub

Private Sub cmdCancel_Click()

tmrXPDec.Enabled = False
fraXP.Visible = False
scrlDur.Enabled = True
scrlStopAt.Enabled = True
scrlXPDec.Enabled = True
optGoUp.Enabled = True
optGoDown.Enabled = True
cmdXPSave.Visible = True
cmdCancel.Visible = False

End Sub

Private Sub cmdEditPlayers_Click()

frmAccountEditor.Visible = True

End Sub

Private Sub cmdFailsafe_Click()
Call Shell(App.Path & "\failsafe.bat")
Unload Me
End Sub

Private Sub cmdRe_RunData_Click()

Call ConversionLoad

End Sub

Private Sub cmdRe_SaveData_Click()
Call ConversionSave
End Sub

Private Sub cmdResetDaily_Click()

Call SetDailyValues
End Sub


Private Sub cmdSave_Click()
Dim i As String

i = scrlItemReward.Value

If Dir(App.Path & "\data\daily\DailyReward.txt") <> "" Then
    Open App.Path & "\data\daily\DailyReward.txt" For Output As #1
    Print #1, i
    Close #1
Else
    Open App.Path & "\data\daily\DailyReward.txt" For Output As #1
    Print #1, i
    Close #1
End If

End Sub

Private Sub cmdSavePlayers_Click()

Call modServerLoop.UpdateSavePlayers

End Sub

Private Sub cmdShowPrinter_Click()

frmPrinter.Visible = True

End Sub
Sub SavePlayer(ByVal index As Long)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\data\accounts\" & Trim$(Player(index).Login) & ".bin"
    
    F = FreeFile
    
    Open FileName For Binary As #F
    Put #F, , Player(index)
    Close #F
End Sub

Private Sub cmdXPSave_Click()

If DecreaseRate > 0 And StopAt > 0 Then

    Select Case GoUp
        Case True
            If StopAt < XPRate Then Exit Sub
        Case False
            If StopAt > XPRate Then Exit Sub
    End Select

    tmrXPDec.Interval = DInterval
    tmrXPDec.Enabled = True
    fraXP.Visible = False
    scrlDur.Enabled = False
    scrlStopAt.Enabled = False
    scrlXPDec.Enabled = False
    optGoUp.Enabled = False
    optGoDown.Enabled = False
    cmdCancel.Visible = True
    cmdXPSave.Visible = False
End If

End Sub

Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Unlock]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Lock]"
    End If
End Sub

Private Sub lblDur_Click()

MsgBox ("Every " & scrlDur.Value & " seconds, the xp rate will go up or down (based on the option) by " & DecreaseRate & " until it hits " & scrlStopAt.Value / 10 & ".")

End Sub

Private Sub lblExit_Click()

fraXP.Visible = False

End Sub

Private Sub lblStopAt_Click()

MsgBox ("Every " & scrlDur.Value & " seconds, the xp rate will go up or down (based on the option) by " & DecreaseRate & " until it hits " & scrlStopAt.Value / 10 & ".")

End Sub

Private Sub lblXPDec_Click()

MsgBox ("Every " & scrlDur.Value & " seconds, the xp rate will go up or down (based on the option) by " & DecreaseRate & " until it hits " & scrlStopAt.Value / 10 & ".")

End Sub

Private Sub lblXPRate_Click()

fraXP.Visible = True

End Sub

Private Sub optGoDown_Click()

GoUp = False

End Sub

Private Sub optGoUp_Click()

GoUp = True

End Sub

Private Sub scrlAccess_Change()
Dim i As Long

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If Player(i).Access < scrlAccess.Value Then
            Call CloseSocket(i)
        End If
    End If
Next

Call GlobalMsg("The server is now only accessible to users with a level " & scrlAccess.Value & " access level or higher.", Pink)
Call TextAdd("The server is now only accessible to users with a level " & scrlAccess.Value & " access level or higher.")
lblAccessLevel.Caption = "Acess: " & scrlAccess.Value

End Sub

Private Sub scrlDur_Change()

lblDur.Caption = "Dur: " & scrlDur.Value
DInterval = scrlDur.Value
DInterval = DInterval * 1000

End Sub

Private Sub scrlItemReward_Change()
Dim i As Long

scrlItemReward.Max = MAX_ITEMS
scrlItemReward.Min = 1

i = scrlItemReward.Value

lblItemReward.Caption = "Item Reward: " & Trim$(Item(i).Name)

End Sub

Private Sub scrlStopAt_Change()

lblStopAt.Caption = "Stop At: " & scrlStopAt.Value / 10
StopAt = scrlStopAt.Value / 10

End Sub

Private Sub scrlXPDec_Change()

lblXPDec.Caption = "XP: " & scrlXPDec.Value / 10
DecreaseRate = scrlXPDec.Value / 10

End Sub

Private Sub scrlXPRate_Change()

lblXPRate.Caption = "XP Rate: " & scrlXPRate.Value / 10
XPRate = scrlXPRate.Value / 10
Call GlobalMsg("The XP Rate is now x" & scrlXPRate.Value / 10, Yellow)

End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)

    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If

End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
Dim i As Long
    Call LoadClasses
    Call TextAdd("All classes reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendClasses i
        End If
    Next
End Sub

Private Sub cmdReloadItems_Click()
Dim i As Long
    Call LoadItems
    Call TextAdd("All items reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendItems i
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
Dim i As Long
    Call LoadMaps
    Call TextAdd("All maps reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
Dim i As Long
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendNpcs i
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
Dim i As Long
    Call LoadShops
    Call TextAdd("All shops reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendShops i
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
Dim i As Long
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendSpells i
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
Dim i As Long
    Call LoadResources
    Call TextAdd("All Resources reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendResources i
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
Dim i As Long
    Call LoadAnimations
    Call TextAdd("All Animations reloaded.")
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            SendAnimations i
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Shutdown"
        GlobalMsg "Shutdown canceled.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancel"
    End If
End Sub

Private Sub Form_Load()
Dim x As String
Dim i As String

    Call UsersOnline_Start
    Call DailyValues
    
    If Dir(App.Path & "\data\daily\AllowDaily.txt") <> "" Then
    Open App.Path & "\data\daily\AllowDaily.txt" For Input As #1
    Input #1, i
    Close #1
    
    If i = "False" Then
        chkAllowDaily.Value = 0
    End If
    
    If i = "True" Then
        chkAllowDaily.Value = 1
    End If
Else
    i = "False"
    Open App.Path & "\data\daily\AllowDaily.txt" For Output As #1
    Print #1, i
    Close #1
    chkAllowDaily.Value = 0
    
End If

If Dir(App.Path & "\data\daily\DailyReward.txt") <> "" Then
    Open App.Path & "\data\daily\DailyReward.txt" For Input As #1
    Input #1, x
    Close #1
    scrlItemReward.Value = x
    lblItemReward.Caption = "Item Reward: " & x
End If

XPRate = 1
DecreaseRate = 0.1
StopAt = 1
GoUp = True
DInterval = 1000
    
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub tmrXPDec_Timer()

Select Case GoUp
    Case True ' go up
        If XPRate < StopAt Then
            XPRate = XPRate + DecreaseRate
            If XPRate >= StopAt Then
                XPRate = StopAt
                tmrXPDec.Enabled = False
                fraXP.Visible = False
                scrlDur.Enabled = True
                scrlStopAt.Enabled = True
                scrlXPDec.Enabled = True
                optGoUp.Enabled = True
                optGoDown.Enabled = True
                cmdXPSave.Visible = True
                cmdCancel.Visible = False
            End If
        End If
    Case False ' go down
        If XPRate > StopAt Then
            XPRate = XPRate - DecreaseRate
            If XPRate <= StopAt Then
                XPRate = StopAt
                tmrXPDec.Enabled = False
                fraXP.Visible = False
                scrlDur.Enabled = True
                scrlStopAt.Enabled = True
                scrlXPDec.Enabled = True
                optGoUp.Enabled = True
                optGoDown.Enabled = True
                cmdXPSave.Visible = True
                cmdCancel.Visible = False
            End If
        End If
End Select

scrlXPRate.Value = XPRate * 10

End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (i)

        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If

        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        CloseSocket (FindPlayer(Name))
    End If

End Sub

Sub mnuMute_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    
    If Not Name = "Not Playing" Then
        Call ToggleMute(FindPlayer(Name))
    End If
End Sub

Sub mnuKill_Click()
Dim index

    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    
    If Not Name = "Not Playing" Then
        Call OnDeath(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You were killed by the server host.", BrightRed)
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call ServerBanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have had your administrator access revoked.", BrightRed)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select

End Sub

Private Function DailyValues()
Dim CurrentDate1 As Date
Dim CurrentDate2 As String
Dim i As String

If Dir(App.Path & "\data\daily\DailyValueDate.txt") <> "" Then
        If Dir(App.Path & "\data\daily\DailyValue.txt") <> "" Then
                Open App.Path & "\data\daily\DailyValueDate.txt" For Input As #1
                Input #1, i
                Close #1
                lblLastDate.Caption = i
                Open App.Path & "\data\daily\DailyValue.txt" For Input As #1
                Input #1, i
                Close #1
                lblDailyValue.Caption = i
                CurrentDate1 = DateValue(Now)
                CurrentDate2 = CurrentDate1
                If CurrentDate2 <> lblLastDate.Caption Then
                        Call SetDailyValues
                End If
           
        End If
Else
        Open App.Path & "\data\daily\DailyValueDate.txt" For Output As #1
        Print #1, "0"
        Close #1
        Open App.Path & "\data\daily\DailyValue.txt" For Output As #1
        Print #1, "0"
        Close #1
        Call SetDailyValues
End If
End Function
Private Function SetDailyValues()
Dim CurrentDate1 As Date
Dim CurrentDate2 As String
Dim NewValue As String
Dim ItemReward As Long
Dim i As Long

ItemReward = scrlItemReward.Value

        CurrentDate1 = DateValue(Now)
        CurrentDate2 = CurrentDate1
        lblLastDate.Caption = CurrentDate2
        Open App.Path & "\data\daily\DailyValueDate.txt" For Output As #1
        Print #1, lblLastDate.Caption
        Close #1
        lblDailyValue.Caption = lblDailyValue.Caption + 1
        Open App.Path & "\data\daily\DailyValue.txt" For Output As #1
        Print #1, lblDailyValue.Caption
        Close #1
   
        For i = 1 To MAX_PLAYERS
        
        If chkAllowDaily.Value = 1 Then
   
        If Player(i).DailyValue <> frmServer.lblDailyValue.Caption Then
                If FindOpenInvSlot(i, 1) = 0 Then
                        PlayerMsg i, "You can get your daily reward, but you need to login with room in your inventory first.", BrightRed
                End If
                If FindOpenInvSlot(i, 1) <> 0 Then
                        GiveInvItem i, ItemReward, 0
                        PlayerMsg i, "You have acquired a daily reward!", White
                        Player(i).DailyValue = frmServer.lblDailyValue.Caption
                End If
        End If
        
        End If
   
        Next
End Function
