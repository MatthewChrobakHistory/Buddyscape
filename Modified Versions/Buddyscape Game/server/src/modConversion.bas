Attribute VB_Name = "modConversion"
Option Explicit

Public CurTextOn As Byte
Public CanPrint As Boolean
Public CurConvOn As Long

' Here's how this works; ConversionLoad is called first. From here you can decide you want to print info, or save info.
' Copy and paste data as needed.

Public Sub ConversionLoad()

Call ConversionSave
'Call ConversionPrint
CurConvOn = 1
frmPrinter.Visible = True
frmPrinter.tmrPrinter.Enabled = True

End Sub

Public Sub PrintC(ByVal Text As String)

frmPrinter.txtPrinter.Text = frmPrinter.txtPrinter.Text + Text & vbCrLf

End Sub

Public Sub ConversionSave()

End Sub

