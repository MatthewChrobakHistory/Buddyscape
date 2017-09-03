VERSION 5.00
Begin VB.Form frmPrinter 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrinter 
      Height          =   7575
      Left            =   0
      MaxLength       =   65535
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Timer tmrPrinter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   1800
   End
   Begin VB.Label lblCurOn 
      Caption         =   "Currently On: 0"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7680
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
Dim x As Long
Dim y As Long

Me.Visible = True

tmrPrinter.Enabled = False

    For x = 0 To Map(3).MaxX
        For y = 0 To Map(3).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                PrintC ("Map(3).Tile(" & x & "," & y & ").Layer(i).x = " & Map(3).Tile(x, y).Layer(i).x)
                PrintC ("Map(3).Tile(" & x & "," & y & ").Layer(i).y = " & Map(3).Tile(x, y).Layer(i).y)
                PrintC ("Map(3).Tile(" & x & "," & y & ").Layer(i).Tileset = " & Map(3).Tile(x, y).Layer(i).Tileset)
            Next
            PrintC ("Map(3).Tile(" & x & "," & y & ").Type = " & Map(3).Tile(x, y).Type)
            PrintC ("Map(3).Tile(" & x & "," & y & ").Data1 = " & Map(3).Tile(x, y).Data1)
            PrintC ("Map(3).Tile(" & x & "," & y & ").Data2 = " & Map(3).Tile(x, y).Data2)
            PrintC ("Map(3).Tile(" & x & "," & y & ").Data3 = " & Map(3).Tile(x, y).Data3)
            PrintC ("Map(3).Tile(" & x & "," & y & ").DirBlock = " & Map(3).Tile(x, y).DirBlock)
        Next
    Next
    

    
End Sub

