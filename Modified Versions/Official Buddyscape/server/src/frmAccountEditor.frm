VERSION 5.00
Begin VB.Form frmAccountEditor 
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameInventory 
      Caption         =   "Inventory"
      Height          =   6735
      Left            =   6120
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdSaveInventory 
         Caption         =   "Save"
         Height          =   375
         Left            =   840
         TabIndex        =   25
         Top             =   6240
         Width           =   1455
      End
      Begin VB.HScrollBar scrlInvItem 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   5520
         Width           =   3015
      End
      Begin VB.TextBox txtAmountInv 
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Text            =   "0"
         Top             =   5880
         Width           =   2175
      End
      Begin VB.ListBox lstInventory 
         Height          =   4935
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label13 
         Caption         =   "Ammount:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   5880
         Width           =   3015
      End
      Begin VB.Label lblInvItem 
         Caption         =   "Inv item: None"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   5280
         Width           =   3015
      End
   End
   Begin VB.Frame frameBank 
      Caption         =   "Bank"
      Height          =   6735
      Left            =   2640
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdSaveBank 
         Caption         =   "Save"
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Text            =   "0"
         Top             =   5880
         Width           =   2175
      End
      Begin VB.HScrollBar scrlBankItem 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   5520
         Width           =   3015
      End
      Begin VB.ListBox lstBank 
         Height          =   4935
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Ammount:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   5880
         Width           =   3135
      End
      Begin VB.Label lblBankItem 
         Caption         =   "Bank item: None"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5280
         Width           =   3015
      End
   End
   Begin VB.Frame FrameAccountDetails 
      Caption         =   "Account Details"
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtSprite 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtAccess 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Sprite: "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Class:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Access:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSavePlayer 
      Caption         =   "Save Player"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindPlayer 
      Caption         =   "Find Player"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtUserNameLoad 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   6840
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

Private Sub Form_Load()
Dim i As Byte

scrlBankItem.Max = MAX_ITEMS

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

Private Sub scrlBankItem_Change()

If scrlBankItem.Value = 0 Then
    lblBankItem.Caption = "Bank item: None"
Else
    lblBankItem.Caption = "Bank item: " & Item(scrlBankItem.Value).Name
End If

End Sub

Private Sub scrlInvItem_Change()

If scrlInvItem.Value = 0 Then
    lblInvItem.Caption = "Inv item: None"
Else
    lblInvItem.Caption = "Inv item: " & Item(scrlInvItem.Value).Name
End If

End Sub

Private Sub txtAccess_Change()

If IsNumeric(txtAccess.Text) = False Then txtAccess.Text = Player(EditUserIndex).Access

End Sub

Private Sub txtAmountInv_Change()

If IsNumeric(txtAmountInv.Text) = False Then txtAmountInv.Text = Player(EditUserIndex).Inv(lstInventory.ListIndex + 1).Value
If txtAmountInv.Text > 2000000000 Then txtAmountInv.Text = 2000000000

End Sub

Private Sub txtPassword_Change()

If txtPassword.Text = vbNullString Then txtPassword.Text = Player(EditUserIndex).Password

End Sub

Private Sub txtSprite_Change()

If IsNumeric(txtSprite.Text) = False Then txtSprite.Text = Player(edituseindex).Sprite

End Sub

Private Sub txtUserName_Change()

If txtUserName.Text = vbNullString Then txtUserName.Text = Player(EditUserIndex).Name

End Sub

Private Sub txtAmount_Change()

If IsNumeric(txtAmount.Text) = False Then txtAmount.Text = Bank(EditUserIndex).Item(lstBank.ListIndex + 1).Value
If txtAmount.Text > 2000000000 Then txtAmount.Text = 2000000000

End Sub

