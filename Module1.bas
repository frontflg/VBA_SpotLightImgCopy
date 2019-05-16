Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function GdiplusStartup Lib "gdiplus" ( _
                                        ByRef token As Long, _
                                        ByRef inputBuf As GdiplusStartupInput, _
                                        ByVal outputBuf As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" ( _
                                        ByVal token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" ( _
                                        ByVal FileName As Long, _
                                        ByRef image As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" ( _
                                        ByVal image As Long, _
                                        ByRef Width As Single, _
                                        ByRef Height As Single) As Long
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Sub GetSpotLightImg()
  Dim buf As String
  Dim cnt As Long
  Dim newName As String
  
' 参照設定が必要（Windows Script Host Object Model）
  Dim oNetwork As New IWshRuntimeLibrary.WshNetwork
  Dim UsrId As String
  UsrId = oNetwork.UserName
  
  Dim Leng As Long
  If Cells(2, 2).Value = "" Then
    Leng = 450
  Else
    Leng = Cells(2, 2).Value
  End If
  
  Dim FromYmd As Date
  If IsDate(Cells(3, 2).Value) Then
    FromYmd = Cells(3, 2).Value
  Else
    FromYmd = "2016/01/01"
  End If
  
  Dim Path As String
  Path = "C:\Users\" & UsrId & "\AppData\Local\" _
  & "Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets\"
  
  Dim MyPath As String
  If Cells(4, 2).Value = "" Then
    MyPath = ThisWorkbook.Path & "\"
  Else
    MyPath = Cells(4, 2).Value & "\"
  End If
  
  If Cells(6, 1).Value <> "" Then
    Cells(6, 1).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Cells(5, 1).Select
  End If
  
  buf = Dir(Path & "*.*")
  cnt = 5
  Do While buf <> ""
    If FileLen(Path & buf) >= Leng * 1000 Then
      If FileDateTime(Path & buf) > FromYmd Then
        If isYoko(Path & buf) Then
          cnt = cnt + 1
          newName = Right(buf, 10) + ".jpg"
          FileCopy Path & buf, MyPath & newName
          Cells(cnt, 1) = cnt - 5
          Cells(cnt, 2) = newName
          Cells(cnt, 3) = FileDateTime(Path & buf)
          Cells(cnt, 4) = FileLen(Path & buf)
        End If
      End If
    End If
    buf = Dir()
  Loop
  Shell "C:\Windows\Explorer.exe " & MyPath, vbNormalFocus
End Sub

Function isYoko(ByVal sImageFilePath As String) As Boolean
  Dim uGdiStartupInput As GdiplusStartupInput
  Dim nGdiToken As Long
  Dim nStatus As Long
  Dim hImage As Long
  isYoko = False
  Dim x, y As Single
  x = 0: y = 0
  uGdiStartupInput.GdiplusVersion = 1
  nStatus = GdiplusStartup(nGdiToken, uGdiStartupInput, 0&)
  If nStatus = 0 Then
    nStatus = GdipLoadImageFromFile(ByVal StrPtr(sImageFilePath), hImage)
    If nStatus = 0 Then
      nStatus = GdipGetImageDimension(hImage, x, y)
      If nStatus = 0 And x > y Then
        isYoko = True
      End If
    End If
    Call GdiplusShutdown(nGdiToken)
  End If
End Function
