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

    PrintC ("Item(" & CurConvOn & ").AccessReq = " & Item(CurConvOn).AccessReq)
    PrintC ("Item(" & CurConvOn & ").Add_Stat(1) = " & Item(CurConvOn).Add_Stat(1))
    PrintC ("Item(" & CurConvOn & ").Add_Stat(2) = " & Item(CurConvOn).Add_Stat(2))
    PrintC ("Item(" & CurConvOn & ").Add_Stat(3) = " & Item(CurConvOn).Add_Stat(3))
    PrintC ("Item(" & CurConvOn & ").Add_Stat(4) = " & Item(CurConvOn).Add_Stat(4))
    PrintC ("Item(" & CurConvOn & ").Add_Stat(5) = " & Item(CurConvOn).Add_Stat(5))
    PrintC ("Item(" & CurConvOn & ").AddEXP = " & Item(CurConvOn).AddEXP)
    PrintC ("Item(" & CurConvOn & ").AddHP = " & Item(CurConvOn).AddHP)
    PrintC ("Item(" & CurConvOn & ").AddMP = " & Item(CurConvOn).AddMP)
    PrintC ("Item(" & CurConvOn & ").Animation = " & Item(CurConvOn).Animation)
    PrintC ("Item(" & CurConvOn & ").BindType = " & Item(CurConvOn).BindType)
    PrintC ("Item(" & CurConvOn & ").CastSpell = " & Item(CurConvOn).CastSpell)
    PrintC ("Item(" & CurConvOn & ").ClassReq = " & Item(CurConvOn).ClassReq)
    PrintC ("Item(" & CurConvOn & ").ConsumeItem = " & Item(CurConvOn).ConsumeItem)
    PrintC ("Item(" & CurConvOn & ").CoRew = " & Item(CurConvOn).CoRew)
    PrintC ("Item(" & CurConvOn & ").CoXP = " & Item(CurConvOn).CoXP)
    PrintC ("Item(" & CurConvOn & ").CrRew = " & Item(CurConvOn).CrRew)
    PrintC ("Item(" & (CurConvOn) & ").CrXP = " & Item(CurConvOn).CrXP)
    PrintC ("Item(" & (CurConvOn) & ").Data1 = " & Item(CurConvOn).Data1)
    PrintC ("Item(" & (CurConvOn) & ").Data2 = " & Item(CurConvOn).Data2)
    PrintC ("Item(" & (CurConvOn) & ").Data3 = " & Item(CurConvOn).Data3)
    PrintC ("Item(" & (CurConvOn) & ").Desc = "" & item(CurConvOn).Desc & """)
    PrintC ("Item(" & (CurConvOn) & ").EqCoXP = " & Item(CurConvOn).EqCoXP)
    PrintC ("Item(" & (CurConvOn) & ").EqCrXP = " & Item(CurConvOn).EqCrXP)
    PrintC ("Item(" & (CurConvOn) & ").EqFlXP = " & Item(CurConvOn).EqFlXP)
    PrintC ("Item(" & (CurConvOn) & ").EqPbXP = " & Item(CurConvOn).EqPBXP)
    PrintC ("Item(" & (CurConvOn) & ").EqSmXP = " & Item(CurConvOn).EqSmXP)
    PrintC ("Item(" & (CurConvOn) & ").FlRew = " & Item(CurConvOn).FlRew)
    PrintC ("Item(" & (CurConvOn) & ").FlXP = " & Item(CurConvOn).FlXP)
    PrintC ("Item(" & (CurConvOn) & ").FXP = " & Item(CurConvOn).FXP)
    PrintC ("Item(" & (CurConvOn) & ").Handed = " & Item(CurConvOn).Handed)
    PrintC ("Item(" & (CurConvOn) & ").instaCast = " & Item(CurConvOn).instaCast)
    PrintC ("Item(" & (CurConvOn) & ").istwohander = " & Item(CurConvOn).istwohander)
    PrintC ("Item(" & (CurConvOn) & ").LevelReq = " & Item(CurConvOn).LevelReq)
    lblCurOn.Caption = "Currently On: " & CurConvOn
    CurConvOn = CurConvOn + 1

    With frmPrinter
        .txtPrinter.SelStart = 0
        .txtPrinter.SelLength = Len(.txtPrinter.Text)
    End With

End Sub
