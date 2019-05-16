VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  If Target.Column <> 2 Then Exit Sub
  Dim y As Integer: y = Target.Row
  Dim x As Integer: x = Target.Column
  If Cells(y, x).Value <> "" Then
    If InStr(Cells(y, x).Value, "jpg") <> 0 Then
      Call Image1_Load(Cells(y, x).Value)
    End If
  End If
End Sub

Sub Image1_Load(imgFile As String)
  Image1.BorderStyle = fmBorderStyleNone 'ògñ≥Çµ
  Image1.PictureSizeMode = fmPictureSizeModeZoom 'âÊëúÇÃècâ°î‰ÇÕï€Ç¡Çƒï\é¶
  
  Dim MyPath As String
  If Cells(4, 2).Value = "" Then
    MyPath = ThisWorkbook.Path & "\" & imgFile
  Else
    MyPath = Cells(4, 2).Value & "\" & imgFile
  End If
  
  If Dir(MyPath) = "" Then
    Image1.Picture = LoadPicture("")
  Else
    Image1.Picture = LoadPicture(MyPath)
    Image1.Width = 400
    Image1.Height = 225
  End If
End Sub

