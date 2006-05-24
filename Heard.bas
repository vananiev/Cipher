Attribute VB_Name = "Heard"
Option Explicit
Dim Pass As New Executer
Dim intCount As Long
Dim BytText() As Byte
Dim lngText As Long
Dim strText As String
Sub Main()
Pass.Show
strText = Pass.Password
BytText = StrConv(strText, vbFromUnicode)
For intCount = 0 To LenB(strText)
   lngText = BytText(intCount) * 2 ^ intCount + lngText
Next intCount
lngText = lngText + Now
End Sub

