Attribute VB_Name = "modAccountEditor"
Option Explicit

Public EditUserIndex As Byte

Public Sub AddInfo(ByVal Text As String)

frmAccountEditor.lblInfo.Caption = Text

End Sub

Public Sub AccountEditorInit(ByVal Index As Byte)
Dim i As Byte
Dim ItemName As String
Dim QState As String

With frmAccountEditor
    .FrameAccountDetails.Visible = True
    .txtUserName.Text = Trim$(Player(Index).Name)
    .txtPassword.Text = Trim$(Player(Index).Password)
    .txtLogin.Text = Trim$(Player(Index).Login)
    .txtAccess.Text = Trim$(Player(Index).Access)
    .cmbClass.ListIndex = Player(Index).Class - 1
    .txtLevel.Text = Player(Index).Level
    .txtSprite.Text = Player(Index).Sprite
    .txtPoints.Text = Player(Index).POINTS
    .txtXP.Text = Player(Index).exp
    'skills
    .frameSkills.Visible = True
    .txtWoodcutting.Text = Player(Index).WoodcuttingXP
    .txtMining.Text = Player(Index).MiningXP
    .txtFishing.Text = Player(Index).FishingXP
    .txtSmithing.Text = Player(Index).SmithingXP
    .txtCooking.Text = Player(Index).CookingXP
    .txtCrafting.Text = Player(Index).CraftingXP
    .txtFletching.Text = Player(Index).FletchingXP
    .txtPotionBrewing.Text = Player(Index).PotionBrewingXP
    
    .txtDaily.Text = Player(Index).DailyValue
    
    'bank
    .frameBank.Visible = True
    For i = 1 To 99
        If Bank(EditUserIndex).Item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Bank(EditUserIndex).Item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).Item(i).Value)
    Next
    .lstBank.ListIndex = 0
    
    'inventory
    .frameInventory.Visible = True
    For i = 1 To 35
        If Player(Index).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Player(Index).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & Player(Index).Inv(i).Value)
    Next
    .lstInventory.ListIndex = 0
    
    'quests
    .lstQuests.Clear
    For i = 1 To MAX_QUESTS
        .frameQuests.Visible = True
        Select Case Player(Index).PlayerQuest(i).DataAmountLeft
            Case 0 'not started
                QState = " (Not Started) "
            Case 1 'started
                QState = " (Started) "
            Case 2 'finished
                QState = " (Finished) "
        End Select
        .lstQuests.AddItem (i & ": " & QState & Trim$(Quest(i).Name))
        If i = MAX_QUESTS Then Exit For
    Next
    .lstQuests.ListIndex = 0
    
    'equipment
    .cmdUnequipHelm.Visible = False
    .cmdUnequipBody.Visible = False
    .cmdUnequipLegs.Visible = False
    .cmdUnequipShield.Visible = False
    .cmdUnequipWeapon.Visible = False
    .lblHelm.Caption = "Helm: "
    .lblBody.Caption = "Body: "
    .lblLegs.Caption = "Legs: "
    .lblShield.Caption = "Shield: "
    .lblWeapon.Caption = "Weapon: "
    If GetPlayerEquipment(Index, Helmet) > 0 Then
        .lblHelm.Caption = "Helm: " & Trim$(Item(GetPlayerEquipment(Index, Helmet)).Name)
        .cmdUnequipHelm.Visible = True
    End If
    If GetPlayerEquipment(Index, Armor) > 0 Then
        .lblBody.Caption = "Armor: " & Trim$(Item(GetPlayerEquipment(Index, Armor)).Name)
        .cmdUnequipBody.Visible = True
    End If
    If GetPlayerEquipment(Index, Legs) > 0 Then
        .lblLegs.Caption = "Legs: " & Trim$(Item(GetPlayerEquipment(Index, Legs)).Name)
        .cmdUnequipLegs.Visible = True
    End If
    If GetPlayerEquipment(Index, Shield) > 0 Then
        .lblShield.Caption = "Shield: " & Trim$(Item(GetPlayerEquipment(Index, Shield)).Name)
        .cmdUnequipShield.Visible = True
    End If
    If GetPlayerEquipment(Index, weapon) > 0 Then
        .lblWeapon.Caption = "Weapon: " & Trim$(Item(GetPlayerEquipment(Index, weapon)).Name)
        .cmdUnequipWeapon.Visible = True
    End If
    
    For i = 1 To Stats.Stat_Count - 1
        Select Case i
            Case 1
                .lblCharStat(i).Caption = "Att: " & GetPlayerStat(Index, i)
            Case 2
                .lblCharStat(i).Caption = "Str: " & GetPlayerStat(Index, i)
            Case 3
                .lblCharStat(i).Caption = "Def: " & GetPlayerStat(Index, i)
            Case 4
                .lblCharStat(i).Caption = "Agil: " & GetPlayerStat(Index, i)
        End Select
    Next
End With

End Sub

Public Sub BankEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    .lstBank.Clear
    For i = 1 To 99 '99 bank space
        If Bank(EditUserIndex).Item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Bank(EditUserIndex).Item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).Item(i).Value)
    Next
    .lstBank.ListIndex = 0
End With

End Sub

Public Sub SaveEditPlayer(ByVal Index As Byte)
Dim i As Long

With Player(Index)
    .Name = frmAccountEditor.txtUserName.Text
    .Password = frmAccountEditor.txtPassword.Text
    .Login = frmAccountEditor.txtLogin.Text
    .Access = frmAccountEditor.txtAccess.Text
    .Class = frmAccountEditor.cmbClass.ListIndex + 1
    .Level = frmAccountEditor.txtLevel.Text
    .Sprite = frmAccountEditor.txtSprite.Text
    '.exp = frmAccountEditor.txtXP.Text
    .POINTS = frmAccountEditor.txtPoints.Text
    'skills
    .WoodcuttingXP = frmAccountEditor.txtWoodcutting.Text
    .MiningXP = frmAccountEditor.txtMining.Text
    .FishingXP = frmAccountEditor.txtFishing.Text
    .SmithingXP = frmAccountEditor.txtSmithing.Text
    .CookingXP = frmAccountEditor.txtCooking.Text
    .CraftingXP = frmAccountEditor.txtCrafting.Text
    .FletchingXP = frmAccountEditor.txtFletching.Text
    .PotionBrewingXP = frmAccountEditor.txtPotionBrewing.Text
    
    .DailyValue = frmAccountEditor.txtDaily.Text
    
    For i = 1 To Stats.Stat_Count - 1
        Player(Index).Stat(i) = frmAccountEditor.scrlCharStat(i).Value
    Next
End With

Call CheckPlayerLevelUp(EditUserIndex)
Call SendPlayerData(Index)

Call PlayerMsg(Index, "Your account was edited by an admin!", Pink)

End Sub

Public Sub InvEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    'inventory
    .lstInventory.Clear
    .frameInventory.Visible = True
    For i = 1 To 35
        If Player(EditUserIndex).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(Item(Player(EditUserIndex).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & Player(EditUserIndex).Inv(i).Value)
    Next
    .lstInventory.ListIndex = 0
End With

End Sub

Public Sub UnequipEquipment(ByVal Index As Long, ByVal Equip As Byte)

Call PlayerUnequipItem(Index, Equip)
frmAccountEditor.lstBank.Clear
frmAccountEditor.lstInventory.Clear
Call AccountEditorInit(Index)

End Sub
