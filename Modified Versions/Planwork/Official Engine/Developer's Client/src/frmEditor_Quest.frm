VERSION 5.00
Begin VB.Form frmEditor_Quest 
   Caption         =   "Form1"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tasks"
      Height          =   4215
      Left            =   2640
      TabIndex        =   6
      Top             =   1440
      Width           =   5175
      Begin VB.Frame fraTaskData 
         Caption         =   "Task 1/50"
         Height          =   3255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   4935
         Begin VB.HScrollBar scrlReward 
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   2520
            Value           =   1
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAmount 
            Height          =   255
            Left            =   840
            Min             =   1
            TabIndex        =   19
            Top             =   2880
            Value           =   1
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.HScrollBar scrlXP 
            Height          =   255
            Left            =   2880
            TabIndex        =   18
            Top             =   2520
            Width           =   1935
         End
         Begin VB.HScrollBar scrlDataValue 
            Height          =   255
            Left            =   2640
            TabIndex        =   17
            Top             =   1800
            Width           =   2175
         End
         Begin VB.HScrollBar scrlDataIndex 
            Height          =   255
            Left            =   2640
            TabIndex        =   16
            Top             =   1440
            Width           =   2175
         End
         Begin VB.ComboBox cmbTaskType 
            Height          =   315
            ItemData        =   "frmEditor_Quest.frx":0000
            Left            =   120
            List            =   "frmEditor_Quest.frx":0013
            TabIndex        =   13
            Text            =   "None"
            Top             =   840
            Width           =   4695
         End
         Begin VB.HScrollBar scrlTask 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   10
            Top             =   240
            Value           =   1
            Width           =   4695
         End
         Begin VB.Label lblReward 
            Caption         =   "Reward: None"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label lblAmount 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2880
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblXP 
            Caption         =   "XP Reward: None"
            Height          =   255
            Left            =   2880
            TabIndex        =   21
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label lblDataValue 
            Caption         =   "Data Value: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1800
            Width           =   2415
         End
         Begin VB.Label lblDataIndex 
            Caption         =   "Data Index: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "TaskType"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   4695
         End
      End
      Begin VB.HScrollBar scrlTaskCount 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   8
         Top             =   480
         Value           =   1
         Width           =   4935
      End
      Begin VB.Label lblTaskCount 
         Caption         =   "Task Count: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "General Data"
      Height          =   1215
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.ListBox lstIndex 
         Height          =   5910
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   5760
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTaskType_Click()

    Quest(lstIndex.ListIndex + 1).Task(scrlTask.Value).TaskType = cmbTaskType.ListIndex
    
    scrlDataIndex.Visible = False
    scrlDataValue.Visible = False
    lblDataIndex.Visible = False
    lblDataValue.Visible = False
    
    Select Case cmbTaskType.ListIndex
        Case 2 ' kill
            scrlDataIndex.Visible = True
            scrlDataValue.Visible = True
            lblDataIndex.Visible = True
            lblDataValue.Visible = True
            lblDataIndex.Caption = "NPC Index: " & scrlDataIndex.Value
            lblDataValue.Caption = "NPC Value: " & scrlDataValue.Value
        Case 4 ' resource
            scrlDataIndex.Visible = True
            scrlDataValue.Visible = True
            lblDataIndex.Visible = True
            lblDataValue.Visible = True
            lblDataIndex.Caption = "Resource Index: " & scrlDataIndex.Value
            lblDataValue.Caption = "Resource Value: " & scrlDataValue.Value
    End Select

End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call QuestEditorOK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()

scrlTaskCount.Max = MAX_QUEST_TASKS

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    'If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Re-init the editor
    If QuestEditorLoaded = True Then QuestEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAmount_Change()

    lblReward.Caption = "Reward: " & scrlAmount.Value & "x " & Trim$(Item(scrlReward.Value).Name)
    Quest(lstIndex.ListIndex + 1).RewardAmount = scrlAmount.Value

End Sub

Private Sub scrlDataIndex_Change()

    lblDataIndex.Caption = "Data Index: " & scrlDataIndex.Value
    Quest(lstIndex.ListIndex + 1).Task(scrlTask.Value).DataIndex = scrlDataIndex.Value

    Select Case cmbTaskType.ListIndex
        Case 2 ' kill
            lblDataIndex.Caption = "NPC Index: " & scrlDataIndex.Value
            lblDataValue.Caption = "NPC Value: " & scrlDataValue.Value
        Case 4 ' resource
            lblDataIndex.Caption = "Resource Index: " & scrlDataIndex.Value
            lblDataValue.Caption = "Resource Value: " & scrlDataValue.Value
    End Select

End Sub

Private Sub scrlDataValue_Change()

    lblDataValue.Caption = "Data Value: " & scrlDataValue.Value
    Quest(lstIndex.ListIndex + 1).Task(scrlTask.Value).DataAmount = scrlDataValue.Value
    
    Select Case cmbTaskType.ListIndex
        Case 2 ' kill
            lblDataIndex.Caption = "NPC Index: " & scrlDataIndex.Value
            lblDataValue.Caption = "NPC Value: " & scrlDataValue.Value
        Case 4 ' resource
            lblDataIndex.Caption = "Resource Index: " & scrlDataIndex.Value
            lblDataValue.Caption = "Resource Value: " & scrlDataValue.Value
    End Select

End Sub

Private Sub scrlReward_Change()

    If scrlReward.Value <> 0 Then
        If Item(scrlReward.Value).Type = ITEM_TYPE_CURRENCY Then
            scrlAmount.Visible = True
            lblAmount.Visible = True
            lblReward.Caption = "Reward: " & scrlAmount.Value & "x " & Trim$(Item(scrlReward.Value).Name)
        Else
            lblReward.Caption = "Reward: " & Trim$(Item(scrlReward.Value).Name)
            Quest(lstIndex.ListIndex + 1).RewardAmount = 1
            scrlAmount.Visible = False
            lblAmount.Visible = False
        End If
    Else
        lblReward.Caption = "Reward: None"
    End If
    Quest(lstIndex.ListIndex + 1).Reward = scrlReward.Value
    
End Sub

Private Sub scrlTask_Change()

If QuestEditorLoaded = False Then Exit Sub

fraTaskData.Caption = "Task " & scrlTask.Value & "/" & scrlTaskCount.Value

scrlDataValue.Value = Quest(lstIndex.ListIndex + 1).Task(scrlTask.Value).DataAmount
scrlDataIndex.Value = Quest(lstIndex.ListIndex + 1).Task(scrlTask.Value).DataIndex

lblDataValue.Caption = "Data Value: " & scrlDataValue.Value
lblDataIndex.Caption = "Data Index: " & scrlDataIndex.Value
cmbTaskType.ListIndex = Quest(lstIndex.ListIndex + 1).Task(scrlTask.Value).TaskType

End Sub

Private Sub scrlTaskCount_Change()

lblTaskCount.Caption = "Task Count: " & scrlTaskCount.Value
Quest(EditorIndex).TaskCount = scrlTaskCount.Value

If scrlTaskCount.Value < scrlTask.Value Then scrlTask.Value = scrlTaskCount.Value
scrlTask.Max = scrlTaskCount.Value
fraTaskData.Caption = "Task " & scrlTask.Value & "/" & scrlTaskCount.Value

End Sub

Private Sub scrlXP_Change()

    If scrlXP.Value > 0 Then
        lblXP.Caption = "XP Reward: " & scrlXP.Value
    Else
        lblXP.Caption = "XP Reward: None"
    End If
    
    Quest(EditorIndex).XPReward = scrlXP.Value
    
End Sub

Private Sub txtName_Change()
Quest(EditorIndex).Name = txtName.text
End Sub
