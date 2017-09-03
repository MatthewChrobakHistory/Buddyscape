VERSION 5.00
Begin VB.Form frmAccountEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   15390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   2880
      TabIndex        =   75
      Top             =   4320
      Width           =   2175
      Begin VB.HScrollBar scrlCharStat 
         Height          =   255
         Index           =   1
         Left            =   600
         Max             =   99
         Min             =   1
         TabIndex        =   79
         Top             =   360
         Value           =   1
         Width           =   1335
      End
      Begin VB.HScrollBar scrlCharStat 
         Height          =   255
         Index           =   2
         Left            =   600
         Max             =   99
         Min             =   1
         TabIndex        =   78
         Top             =   840
         Value           =   1
         Width           =   1335
      End
      Begin VB.HScrollBar scrlCharStat 
         Height          =   255
         Index           =   3
         Left            =   600
         Max             =   99
         Min             =   1
         TabIndex        =   77
         Top             =   1320
         Value           =   1
         Width           =   1335
      End
      Begin VB.HScrollBar scrlCharStat 
         Height          =   255
         Index           =   4
         Left            =   600
         Max             =   99
         Min             =   1
         TabIndex        =   76
         Top             =   1800
         Value           =   1
         Width           =   1335
      End
      Begin VB.Label lblCharStat 
         Caption         =   "Att: 0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCharStat 
         Caption         =   "Str: 0"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCharStat 
         Caption         =   "Def: 0"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   81
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblCharStat 
         Caption         =   "Will: 0"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   80
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame frameQuests 
      Caption         =   "Quests"
      Height          =   7215
      Left            =   12000
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.HScrollBar scrlRemaining 
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   74
         Top             =   5640
         Width           =   1455
      End
      Begin VB.HScrollBar scrlTaskOn 
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   5640
         Width           =   1455
      End
      Begin VB.CommandButton cmdResetQuests 
         Caption         =   "Reset Quest"
         Height          =   375
         Left            =   840
         TabIndex        =   70
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Quest"
         Height          =   375
         Left            =   840
         TabIndex        =   69
         Top             =   6240
         Width           =   1455
      End
      Begin VB.ListBox lstQuests 
         Height          =   4935
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblRemaining 
         Caption         =   "Remaining: 0"
         Height          =   375
         Left            =   1680
         TabIndex        =   73
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label lblTaskOn 
         Caption         =   "Task On: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   5280
         Width           =   1455
      End
   End
   Begin VB.TextBox txtUserNameLoad 
      Height          =   285
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmdFindPlayer 
      Caption         =   "Find Player"
      Height          =   255
      Left            =   0
      TabIndex        =   64
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdSavePlayer 
      Caption         =   "Save Player"
      Height          =   255
      Left            =   0
      TabIndex        =   63
      Top             =   840
      Width           =   1335
   End
   Begin VB.Frame FrameAccountDetails 
      Caption         =   "Account Details"
      Height          =   4215
      Left            =   2520
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   960
         TabIndex        =   52
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   960
         TabIndex        =   51
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtLogin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   50
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtAccess 
         Height          =   285
         Left            =   960
         TabIndex        =   49
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   960
         TabIndex        =   48
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   47
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtSprite 
         Height          =   285
         Left            =   960
         TabIndex        =   46
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtXP 
         Height          =   285
         Left            =   960
         TabIndex        =   45
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtPoints 
         Height          =   285
         Left            =   960
         TabIndex        =   44
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtDaily 
         Height          =   285
         Left            =   960
         TabIndex        =   43
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Login:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Access:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Class:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Level:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   57
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Sprite: "
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "XP: "
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Points: "
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Daily Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3840
         Width           =   855
      End
   End
   Begin VB.Frame frameBank 
      Caption         =   "Bank"
      Height          =   6735
      Left            =   5160
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ListBox lstBank 
         Height          =   4935
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   3135
      End
      Begin VB.HScrollBar scrlBankItem 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   5520
         Width           =   3015
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   960
         TabIndex        =   37
         Text            =   "0"
         Top             =   5880
         Width           =   2175
      End
      Begin VB.CommandButton cmdSaveBank 
         Caption         =   "Save Slot"
         Height          =   375
         Left            =   840
         TabIndex        =   36
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label lblBankItem 
         Caption         =   "Bank item: None"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   5280
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Ammount:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   5880
         Width           =   3135
      End
   End
   Begin VB.Frame frameSkills 
      Caption         =   "Skills"
      Height          =   3255
      Left            =   0
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtPotionBrewing 
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtWoodcutting 
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtMining 
         Height          =   285
         Left            =   1200
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtFishing 
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSmithing 
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCooking 
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtFletching 
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtCrafting 
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Woodcutting: "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Mining: "
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Fishing: "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Smithing: "
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Cooking: "
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Fletching: "
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Crafting: "
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Potion Brewing: "
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   1215
      End
   End
   Begin VB.Frame frameInventory 
      Caption         =   "Inventory"
      Height          =   6735
      Left            =   8640
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ListBox lstInventory 
         Height          =   4935
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtAmountInv 
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Text            =   "0"
         Top             =   5880
         Width           =   2175
      End
      Begin VB.HScrollBar scrlInvItem 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   5520
         Width           =   3015
      End
      Begin VB.CommandButton cmdSaveInventory 
         Caption         =   "Save Slot"
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label lblInvItem 
         Caption         =   "Inv item: None"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Ammount:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   5880
         Width           =   3015
      End
   End
   Begin VB.Frame frameEquipiment 
      Caption         =   "Equipment"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   2775
      Begin VB.CommandButton cmdUnequipHelm 
         Caption         =   "unequip"
         Height          =   195
         Left            =   1920
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdUnequipBody 
         Caption         =   "unequip"
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdUnequipLegs 
         Caption         =   "unequip"
         Height          =   195
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdUnequipShield 
         Caption         =   "unequip"
         Height          =   195
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdUnequipWeapon 
         Caption         =   "unequip"
         Height          =   195
         Left            =   1920
         TabIndex        =   1
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblBody 
         Caption         =   "Body:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblHelm 
         Caption         =   "Helm:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblLegs 
         Caption         =   "Legs:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblShield 
         Caption         =   "Shield:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblWeapon 
         Caption         =   "Weapon:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   0
      TabIndex        =   66
      Top             =   6960
      Width           =   8415
   End
End
Attribute VB_Name = "frmAccountEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFindPlayer_Click()
Dim Username As String
Dim i As Byte

Username = txtUserNameLoad.Text
lstBank.Clear
lstInventory.Clear
frameBank.Visible = False
FrameAccountDetails.Visible = False
frameInventory.Visible = False

For i = 1 To Player_HighIndex
    If IsPlaying(i) = True Then
        If Trim$(Player(i).Name) = Username Then
            EditUserIndex = i
            Call AccountEditorInit(i)
        Else
            AddInfo ("Player not online, or username did not match!")
        End If
    End If
Next

End Sub

Private Sub cmdResetQuests_Click()
Dim i As Long

For i = 1 To MAX_QUESTS

    With Player(EditUserIndex).PlayerQuest(i)
        .DataAmountLeft = 0
        .QuestStatus = 0
        .TaskOn = 0
    End With

Next

End Sub

Private Sub cmdSave_Click()

Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).TaskOn = scrlTaskOn.Value
If scrlTaskOn.Value > 0 Then
    Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).QuestStatus = 1
End If
If scrlTaskOn.Value > Quest(lstQuests.ListIndex + 1).TaskCount Then
    Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).QuestStatus = 2
End If
If scrlTaskOn.Value = 0 Then
    Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).QuestStatus = 0
End If

If Quest(lstQuests.ListIndex + 1).Task(Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).TaskOn).TaskType = 4 Or Quest(lstQuests.ListIndex + 1).Task(Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).TaskOn).TaskType = 2 Then
    Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).DataAmountLeft = scrlRemaining.Value
End If

Call SendPlayerData(EditUserIndex)

End Sub

Private Sub cmdSaveBank_Click()

If IsPlaying(EditUserIndex) = False Then
    Call AddInfo("Player not online!")
    Exit Sub
End If

Bank(EditUserIndex).Item(lstBank.ListIndex + 1).Num = scrlBankItem.Value
Bank(EditUserIndex).Item(lstBank.ListIndex + 1).Value = txtAmount.Text

Call SaveBank(EditUserIndex)
Call BankEditorInit

End Sub

Private Sub cmdSaveInventory_Click()

If IsPlaying(EditUserIndex) = False Then
    Call AddInfo("Player not online!")
    Exit Sub
End If

Player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Num = scrlInvItem.Value
Player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Value = txtAmountInv.Text

Call SendInventoryUpdate(EditUserIndex, lstInventory.ListIndex + 1)

Call InvEditorInit

End Sub

Private Sub cmdSavePlayer_Click()

If IsPlaying(EditUserIndex) = False Then
    AddInfo ("User no longer online!")
    Exit Sub
End If

Call SaveEditPlayer(EditUserIndex)

End Sub

Private Sub cmdUnequipBody_Click()
Call UnequipEquipment(EditUserIndex, Equipment.Armor)
End Sub

Private Sub cmdUnequipHelm_Click()
Call UnequipEquipment(EditUserIndex, Equipment.Helmet)
End Sub

Private Sub cmdUnequipLegs_Click()
Call UnequipEquipment(EditUserIndex, Equipment.Legs)
End Sub

Private Sub cmdUnequipShield_Click()
Call UnequipEquipment(EditUserIndex, Equipment.Shield)
End Sub

Private Sub cmdUnequipWeapon_Click()
Call UnequipEquipment(EditUserIndex, Equipment.weapon)
End Sub

Private Sub Form_Load()
Dim i As Byte

scrlBankItem.Max = MAX_ITEMS
scrlTaskOn.Max = MAX_QUEST_TASKS

cmbClass.Text = Trim$(Class(1).Name)
For i = 1 To Max_Classes
    cmbClass.AddItem Trim$(Class(i).Name)
Next
End Sub

Private Sub lstInventory_Click()
Dim ItemName As String

If Player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Num = 0 Then
    ItemName = "None"
Else
    ItemName = Trim$(Item(Player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Num).Name)
End If

lblInvItem.Caption = "Inv item: " & ItemName
txtAmountInv.Text = Player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Value
scrlInvItem.Value = Player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Num

End Sub

Private Sub lstBank_Click()
Dim ItemName As String

If Bank(EditUserIndex).Item(lstBank.ListIndex + 1).Num = 0 Then
    ItemName = "None"
Else
    ItemName = Trim$(Item(Bank(EditUserIndex).Item(lstBank.ListIndex + 1).Num).Name)
End If

lblBankItem.Caption = "Bank item: " & ItemName
txtAmount.Text = Bank(EditUserIndex).Item(lstBank.ListIndex + 1).Value
scrlBankItem.Value = Bank(EditUserIndex).Item(lstBank.ListIndex + 1).Num

End Sub

Private Sub lstQuests_Click()

scrlTaskOn.Value = Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).TaskOn
scrlRemaining.Value = Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).DataAmountLeft

If Quest(lstQuests.ListIndex + 1).Task(Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).TaskOn).TaskType = 4 Or Quest(lstQuests.ListIndex + 1).Task(Player(EditUserIndex).PlayerQuest(lstQuests.ListIndex + 1).TaskOn).TaskType = 2 Then
    scrlRemaining.Enabled = True
Else
    scrlRemaining.Enabled = False
End If

End Sub

Private Sub scrlBankItem_Change()

If scrlBankItem.Value = 0 Then
    lblBankItem.Caption = "Bank item: None"
Else
    lblBankItem.Caption = "Bank item: " & Item(scrlBankItem.Value).Name
End If

End Sub

Private Sub scrlCharStat_Change(Index As Integer)

Select Case Index
    Case 1
        lblCharStat(Index).Caption = "Att: " & scrlCharStat(Index).Value
    Case 2
        lblCharStat(Index).Caption = "Str: " & scrlCharStat(Index).Value
    Case 3
        lblCharStat(Index).Caption = "Def: " & scrlCharStat(Index).Value
    Case 4
        lblCharStat(Index).Caption = "Agil: " & scrlCharStat(Index).Value
End Select

End Sub

Private Sub scrlInvItem_Change()

If scrlInvItem.Value = 0 Then
    lblInvItem.Caption = "Inv item: None"
Else
    lblInvItem.Caption = "Inv item: " & Item(scrlInvItem.Value).Name
End If

End Sub

Private Sub scrlRemaining_Change()

lblRemaining.Caption = "Remaining: " & scrlRemaining.Value

End Sub

Private Sub scrlTaskOn_Change()

lblTaskOn.Caption = "Task On: " & scrlTaskOn.Value

If Quest(lstQuests.ListIndex + 1).Task(scrlTaskOn.Value).TaskType = 4 Or Quest(lstQuests.ListIndex + 1).Task(scrlTaskOn.Value).TaskType = 2 Then
    scrlRemaining.Enabled = True
    scrlRemaining.Value = Quest(lstQuests.ListIndex + 1).Task(scrlTaskOn).DataAmount
Else
    scrlRemaining.Enabled = False
    lblRemaining.Caption = "Remaining: 0"
End If

End Sub

Private Sub txtAccess_Change()

If IsNumeric(txtAccess.Text) = False Then txtAccess.Text = Player(EditUserIndex).Access

End Sub

Private Sub txtAmountInv_Change()

If IsNumeric(txtAmountInv.Text) = False Then txtAmountInv.Text = Player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Value
If txtAmountInv.Text > 2000000000 Then txtAmountInv.Text = 2000000000

End Sub

Private Sub txtCooking_Change()

If IsNumeric(txtCooking.Text) = False Then txtCooking.Text = Player(EditUserIndex).CookingXP

End Sub

Private Sub txtCrafting_Change()

If IsNumeric(txtCrafting.Text) = False Then txtCrafting.Text = Player(EditUserIndex).CraftingXP

End Sub

Private Sub txtDaily_Change()

If IsNumeric(txtDaily.Text) = False Then txtDaily.Text = Player(EditUserIndex).DailyValue

End Sub

Private Sub txtFishing_Change()

If IsNumeric(txtFishing.Text) = False Then txtFishing.Text = Player(EditUserIndex).FishingXP

End Sub

Private Sub txtFletching_Change()

If IsNumeric(txtFletching.Text) = False Then txtFletching.Text = Player(EditUserIndex).FletchingXP

End Sub

Private Sub txtLevel_Change()

If IsNumeric(txtLevel.Text) = False Then txtLevel.Text = Player(EditUserIndex).Level
If txtLevel.Text > MAX_LEVELS Then txtLevel.Text = Player(EditUserIndex).Level

End Sub

Private Sub txtLogin_Change()

If txtLogin.Text = vbNullString Then txtLogin.Text = Player(EditUserIndex).Login

End Sub

Private Sub txtMining_Change()

If IsNumeric(txtMining.Text) = False Then txtMining.Text = Player(EditUserIndex).MiningXP

End Sub

Private Sub txtPassword_Change()

If txtPassword.Text = vbNullString Then txtPassword.Text = Player(EditUserIndex).Password

End Sub

Private Sub txtPoints_Change()

If IsNumeric(txtPoints.Text) = False Then txtPoints.Text = Player(EditUserIndex).POINTS

End Sub

Private Sub txtPotionBrewing_Change()

If IsNumeric(txtPotionBrewing.Text) = False Then txtPotionBrewing.Text = Player(EditUserIndex).PotionBrewingXP

End Sub

Private Sub txtSmithing_Change()

If IsNumeric(txtSmithing.Text) = False Then txtSmithing.Text = Player(EditUserIndex).SmithingXP

End Sub

Private Sub txtSprite_Change()

If IsNumeric(txtSprite.Text) = False Then txtSprite.Text = Player(EditUserIndex).Sprite

End Sub

Private Sub txtUserName_Change()

If txtUserName.Text = vbNullString Then txtUserName.Text = Player(EditUserIndex).Name

End Sub

Private Sub txtAmount_Change()

If IsNumeric(txtAmount.Text) = False Then txtAmount.Text = Bank(EditUserIndex).Item(lstBank.ListIndex + 1).Value
If txtAmount.Text > 2000000000 Then txtAmount.Text = 2000000000

End Sub

Private Sub txtWoodcutting_Change()

If IsNumeric(txtWoodcutting.Text) = False Then txtWoodcutting.Text = Player(EditUserIndex).WoodcuttingXP

End Sub

Private Sub txtXP_Change()

If IsNumeric(txtXP.Text) = False Then txtXP.Text = Player(EditUserIndex).exp
If txtXP.Text > 2000000000 Then txtXP.Text = 2000000000

End Sub

