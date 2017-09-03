VERSION 5.00
Begin VB.Form frmPrinter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPrinter 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3480
      Top             =   1920
   End
   Begin VB.TextBox txtPrinter 
      Height          =   7575
      Left            =   120
      MaxLength       =   65535
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblCurOn 
      Caption         =   "Currently On: 0"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   7800
      Width           =   3735
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrPrinter_Timer()
Dim i As Long

If CurConvOn = MAX_ITEMS + 1 Then
    tmrPrinter.Enabled = False
    CurConvOn = 1
    Exit Sub
End If

If CurConvOn = 30 Then MsgBox "It is recommended that you cut and paste this information into a notepad document to optimize the speed that the printer works at."

PrintC ("Item(" & (CurConvOn) & ").MagicDefense = " & Item(CurConvOn).MagicDefense)
    PrintC ("Item(" & (CurConvOn) & ").MagicOffense = " & Item(CurConvOn).MagicOffense)
    PrintC ("Item(" & (CurConvOn) & ").Mastery = " & Item(CurConvOn).Mastery)
    PrintC ("Item(" & (CurConvOn) & ").MeleeDefense = " & Item(CurConvOn).MeleeDefense)
    PrintC ("Item(" & (CurConvOn) & ").MeleeOffense = " & Item(CurConvOn).MeleeOffense)
        PrintC ("Item(" & (CurConvOn) & ").Projectile.damage = " & Item(CurConvOn).ProjecTile.damage)
    PrintC ("Item(" & (CurConvOn) & ").Projectile.Direction = " & Item(CurConvOn).ProjecTile.Direction)
    PrintC ("Item(" & (CurConvOn) & ").Projectile.Pic = " & Item(CurConvOn).ProjecTile.Pic)
    PrintC ("Item(" & (CurConvOn) & ").Projectile.Range = " & Item(CurConvOn).ProjecTile.Range)
    PrintC ("Item(" & (CurConvOn) & ").Projectile.Speed = " & Item(CurConvOn).ProjecTile.Speed)
    PrintC ("Item(" & (CurConvOn) & ").Projectile.TravelTime = " & Item(CurConvOn).ProjecTile.TravelTime)
    PrintC ("Item(" & (CurConvOn) & ").Projectile.y = " & Item(CurConvOn).ProjecTile.y)
    PrintC ("Item(" & (CurConvOn) & ").Projectile.x = " & Item(CurConvOn).ProjecTile.x)
    PrintC ("Item(" & (CurConvOn) & ").RangedDefense = " & Item(CurConvOn).RangedDefense)
    PrintC ("Item(" & (CurConvOn) & ").RangedOffense = " & Item(CurConvOn).RangedOffense)
    PrintC ("Item(" & (CurConvOn) & ").Name = " & Trim$(Item(CurConvOn).Name))
    
    lblCurOn.Caption = "Currently On: " & CurConvOn
    CurConvOn = CurConvOn + 1

    With frmPrinter
        .txtPrinter.SelStart = 0
        .txtPrinter.SelLength = Len(.txtPrinter.Text)
    End With

End Sub
