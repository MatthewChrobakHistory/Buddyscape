VERSION 5.00
Begin VB.Form frmEditor_Conv 
   Caption         =   "Form1"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   13005
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Conversation List"
      Height          =   8655
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   2655
      Begin VB.ListBox lstIndex 
         Height          =   8055
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   1695
      Left            =   2760
      TabIndex        =   22
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   5415
      End
      Begin VB.HScrollBar scrlChatCount 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   23
         Top             =   1200
         Value           =   1
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label lblChatCount 
         Caption         =   "Chat count:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   5415
      End
   End
   Begin VB.Frame fraConv 
      Caption         =   "Conversation (1/50)"
      Height          =   6495
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   10215
      Begin VB.HScrollBar scrlCustomScript 
         Height          =   255
         Left            =   1680
         TabIndex        =   63
         Top             =   6120
         Width           =   1935
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   4320
         TabIndex        =   61
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   600
         Width           =   2415
      End
      Begin VB.Frame fraRequirements 
         Caption         =   "Reply Requirements"
         Height          =   5895
         Left            =   5640
         TabIndex        =   29
         Top             =   120
         Width           =   4455
         Begin VB.Frame Frame3 
            Caption         =   "4"
            Height          =   1335
            Index           =   3
            Left            =   120
            TabIndex        =   51
            Top             =   4440
            Width           =   2775
            Begin VB.HScrollBar scrlQuestIndex 
               Height          =   255
               Index           =   4
               Left            =   1200
               TabIndex        =   54
               Top             =   240
               Value           =   1
               Width           =   1455
            End
            Begin VB.HScrollBar scrlTaskIndex 
               Height          =   255
               Index           =   4
               Left            =   1200
               TabIndex        =   53
               Top             =   600
               Width           =   1455
            End
            Begin VB.HScrollBar scrlCondi 
               Height          =   255
               Index           =   4
               Left            =   1200
               Max             =   3
               Min             =   1
               TabIndex        =   52
               Top             =   960
               Value           =   1
               Width           =   1455
            End
            Begin VB.Label lblQuestIndex 
               Caption         =   "Quest: 000"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblTaskIndex 
               Caption         =   "Task: 0"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   56
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblCondi 
               Caption         =   "Is equal to"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   55
               Top             =   960
               Width           =   1575
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "3"
            Height          =   1335
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   3120
            Width           =   2775
            Begin VB.HScrollBar scrlQuestIndex 
               Height          =   255
               Index           =   3
               Left            =   1200
               TabIndex        =   47
               Top             =   240
               Value           =   1
               Width           =   1455
            End
            Begin VB.HScrollBar scrlTaskIndex 
               Height          =   255
               Index           =   3
               Left            =   1200
               TabIndex        =   46
               Top             =   600
               Width           =   1455
            End
            Begin VB.HScrollBar scrlCondi 
               Height          =   255
               Index           =   3
               Left            =   1200
               Max             =   3
               Min             =   1
               TabIndex        =   45
               Top             =   960
               Value           =   1
               Width           =   1455
            End
            Begin VB.Label lblQuestIndex 
               Caption         =   "Quest: 000"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblTaskIndex 
               Caption         =   "Task: 0"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   49
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblCondi 
               Caption         =   "Is equal to"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   48
               Top             =   960
               Width           =   1575
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "2"
            Height          =   1335
            Index           =   1
            Left            =   120
            TabIndex        =   37
            Top             =   1680
            Width           =   2775
            Begin VB.HScrollBar scrlQuestIndex 
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   40
               Top             =   240
               Value           =   1
               Width           =   1455
            End
            Begin VB.HScrollBar scrlTaskIndex 
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   39
               Top             =   600
               Width           =   1455
            End
            Begin VB.HScrollBar scrlCondi 
               Height          =   255
               Index           =   2
               Left            =   1200
               Max             =   3
               Min             =   1
               TabIndex        =   38
               Top             =   960
               Value           =   1
               Width           =   1455
            End
            Begin VB.Label lblQuestIndex 
               Caption         =   "Quest: 000"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblTaskIndex 
               Caption         =   "Task: 0"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   42
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblCondi 
               Caption         =   "Is equal to"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   41
               Top             =   960
               Width           =   1575
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "1"
            Height          =   1335
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2775
            Begin VB.HScrollBar scrlCondi 
               Height          =   255
               Index           =   1
               Left            =   1200
               Max             =   3
               Min             =   1
               TabIndex        =   36
               Top             =   960
               Value           =   1
               Width           =   1455
            End
            Begin VB.HScrollBar scrlTaskIndex 
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   35
               Top             =   600
               Width           =   1455
            End
            Begin VB.HScrollBar scrlQuestIndex 
               Height          =   255
               Index           =   1
               Left            =   1200
               TabIndex        =   34
               Top             =   240
               Value           =   1
               Width           =   1455
            End
            Begin VB.Label lblCondi 
               Caption         =   "Is equal to"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   33
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label lblTaskIndex 
               Caption         =   "Task: 0"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   32
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label lblQuestIndex 
               Caption         =   "Quest: 000"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   1335
            End
         End
      End
      Begin VB.HScrollBar scrlCurChat 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   15
         Top             =   240
         Value           =   1
         Width           =   5295
      End
      Begin VB.TextBox txtConvText 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   960
         Width           =   5415
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   3855
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   3600
         Width           =   3855
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   3960
         Width           =   3855
      End
      Begin VB.ComboBox cmbToConv 
         Height          =   315
         Index           =   1
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox cmbToConv 
         Height          =   315
         Index           =   2
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
      End
      Begin VB.ComboBox cmbToConv 
         Height          =   315
         Index           =   3
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox cmbToConv 
         Height          =   315
         Index           =   4
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3960
         Width           =   1455
      End
      Begin VB.ComboBox cmbEvent 
         Height          =   315
         ItemData        =   "frmEditor_Conv.frx":0000
         Left            =   120
         List            =   "frmEditor_Conv.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4560
         Width           =   5415
      End
      Begin VB.HScrollBar scrlData1 
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   4920
         Width           =   3350
      End
      Begin VB.HScrollBar scrlData2 
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   5280
         Width           =   3350
      End
      Begin VB.HScrollBar scrlData3 
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   5640
         Width           =   3350
      End
      Begin VB.Label lblCustomScript 
         Caption         =   "Custom Script: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "length:"
         Height          =   375
         Left            =   3720
         TabIndex        =   60
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   600
         TabIndex        =   59
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Replies:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Event:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label lblData1 
         Caption         =   "Data1: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label lblData2 
         Caption         =   "Data2: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label lblData3 
         Caption         =   "Data3: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   5640
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   8280
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditor_Conv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbEvent_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
    
    Conv(EditorIndex).Chat(scrlCurChat.Value).Event = cmbEvent.ListIndex
    If ConvEditorLoaded Then InitEventData

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbEvent_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
    
    If cmbSound.ListIndex >= 0 Then
        Conv(EditorIndex).Chat(scrlCurChat.Value).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Conv(EditorIndex).Chat(scrlCurChat.Value).Sound = "None."
    End If
End Sub

Private Sub cmbToConv_Click(Index As Integer)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
    
    Conv(EditorIndex).Chat(scrlCurChat.Value).ReplyConvTo(Index) = cmbToConv(Index).ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbToConv_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ConvEditorOK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Dim i As Byte

For i = 1 To 4
    scrlQuestIndex(i).Max = MAX_QUESTS
    scrlTaskIndex(i).Max = MAX_QUEST_TASKS
Next


End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Re-init the editor
    If ConvEditorLoaded = True Then ConvEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlChatCount_Change()
Dim curIndex(1 To 4) As Byte, i As Long, j As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Conv(EditorIndex).ChatCount = scrlChatCount.Value
    scrlCurChat.Max = Conv(EditorIndex).ChatCount
    lblChatCount.Caption = "Chat count: " & scrlChatCount.Value
    fraConv.Caption = "Conversation: (" & scrlCurChat.Value & "/" & scrlChatCount.Value & ")"
    
    ' Reset the conv-to boxes
    For i = 1 To 4
        If cmbToConv(i).ListIndex > 0 Then
            curIndex(i) = cmbToConv(i).ListIndex
        Else
            curIndex(i) = 0
        End If
        
        cmbToConv(i).Clear
        cmbToConv(i).AddItem "None", 0
        
        For j = 1 To scrlChatCount.Value
            cmbToConv(i).AddItem CStr(j), j
        Next
        
        If Conv(EditorIndex).Chat(scrlCurChat.Value).ReplyConvTo(i) > scrlChatCount.Value Then
            Me.cmbToConv(i).ListIndex = 0
        Else
            cmbToConv(i).ListIndex = Conv(EditorIndex).Chat(scrlCurChat.Value).ReplyConvTo(i)
        End If
        
        ' reset the list index
        'If curIndex(i) <= cmbToConv(i).ListIndex Then cmbToConv(i).ListIndex = curIndex(i)
    Next
    
    ' Reset the data
    If Conv(EditorIndex).Chat(scrlChatCount.Value).Data1 = 0 Then Conv(EditorIndex).Chat(scrlChatCount.Value).Data1 = 1
    If Conv(EditorIndex).Chat(scrlChatCount.Value).Data2 = 0 Then Conv(EditorIndex).Chat(scrlChatCount.Value).Data2 = 1
    If Conv(EditorIndex).Chat(scrlChatCount.Value).Data3 = 0 Then Conv(EditorIndex).Chat(scrlChatCount.Value).Data3 = 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlChatCount_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCondi_Change(Index As Integer)

Select Case scrlCondi(Index).Value
    Case 1
        lblCondi(Index).Caption = "Is equal to"
    Case 2
        lblCondi(Index).Caption = "Is less than"
    Case 3
        lblCondi(Index).Caption = "Is greater than"
End Select

Conv(EditorIndex).Chat(scrlCurChat.Value).QuestRequirement(Index).Condition = scrlCondi(Index).Value

End Sub

Private Sub scrlCurChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    fraConv.Caption = "Conversation: (" & scrlCurChat.Value & "/" & Conv(EditorIndex).ChatCount & ")"
    ConvEditorInit True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlData1_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
   
    ' Change it based on what it is
    Select Case cmbEvent.ListIndex
        Case 2
            Conv(EditorIndex).Chat(scrlCurChat.Value).Data1 = scrlData1.Value
            lblData1.Caption = "Shop Index: " & scrlData1.Value
        Case 3, 4
            Conv(EditorIndex).Chat(scrlCurChat.Value).Data1 = scrlData1.Value
            lblData1.Caption = "Item Index: " & scrlData1.Value
        Case 5
            Conv(EditorIndex).Chat(scrlCurChat.Value).Data1 = scrlData1.Value
            lblData1.Caption = "Map Index: " & scrlData1.Value
        Case Is > 5
            Conv(EditorIndex).Chat(scrlCurChat.Value).Data1 = scrlData1.Value
            lblData1.Caption = "Quest Index: " & scrlData1.Value
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlData1_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlData2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
   
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
   
    ' Change it based on what it is
    Select Case cmbEvent.ListIndex
        Case 3, 4
            Conv(EditorIndex).Chat(scrlCurChat.Value).Data2 = scrlData2.Value
            lblData2.Caption = "Value: " & scrlData2.Value
        Case 5
            Conv(EditorIndex).Chat(scrlCurChat.Value).Data2 = scrlData2.Value
            lblData2.Caption = "X index: " & scrlData2.Value
        Case Is > 5
            Conv(EditorIndex).Chat(scrlCurChat.Value).Data2 = scrlData2.Value
            lblData2.Caption = "Finish Quest Task: " & scrlData2.Value
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlData2_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlData3_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Conv(EditorIndex).Chat(scrlCurChat.Value).Data3 = scrlData3.Value
    
    Select Case cmbEvent.ListIndex
        Case 5
            lblData3.Caption = "Y index: " & scrlData3.Value
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlData3_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestIndex_Change(Index As Integer)

lblQuestIndex(Index).Caption = "Quest: " & scrlQuestIndex(Index).Value
Conv(EditorIndex).Chat(scrlCurChat.Value).QuestRequirement(Index).Index = scrlQuestIndex(Index).Value

End Sub

Private Sub scrlTaskIndex_Change(Index As Integer)

lblTaskIndex(Index).Caption = "Task: " & scrlTaskIndex(Index).Value
Conv(EditorIndex).Chat(scrlCurChat.Value).QuestRequirement(Index).Task = scrlTaskIndex(Index).Value

End Sub

Private Sub txtConvText_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Conv(EditorIndex).Chat(scrlCurChat.Value).Text = Trim$(txtConvText.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtConvText_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtInterval_Change()

If IsNumeric(txtInterval.Text) = False Then txtInterval.Text = "1000"
Conv(EditorIndex).Chat(scrlCurChat.Value).SoundLength = txtInterval.Text

End Sub

Private Sub txtName_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Conv(EditorIndex).Name = txtName.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtReply_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Conv(EditorIndex).Chat(scrlCurChat.Value).ReplyText(Index) = txtReply(Index).Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

