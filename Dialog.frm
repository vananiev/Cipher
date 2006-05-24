VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cipher"
   ClientHeight    =   2175
   ClientLeft      =   4560
   ClientTop       =   4800
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4905
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox txtOrg 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblCode 
      Caption         =   "Ваш код"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Введите информацию о себе для получения кода."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Организация"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Ваше имя"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intCount As Long
Dim BytText() As Byte
Dim lngText As Long
Dim strText As String
Private Sub txtName_Change()
lngText = Now
txtCode = ""
strText = txtName & txtOrg
BytText() = StrConv(strText, vbFromUnicode)
For intCount = LBound(BytText) To UBound(BytText)
   txtCode = Right(Str(BytText(intCount) * 2 ^ intCount), Len(Str(BytText(intCount))) - 1) & txtCode
Next intCount
txtCode = txtCode & Right(Str(lngText ^ 2), Len(Str(lngText ^ 2)) - 1)
End Sub
Private Sub txtOrg_Change()
lngText = Now
txtCode = ""
strText = txtName & txtOrg
BytText() = StrConv(strText, vbFromUnicode)
For intCount = LBound(BytText) To UBound(BytText)
   txtCode = Right(Str(BytText(intCount) * 2 ^ intCount), Len(Str(BytText(intCount))) - 1) & txtCode
Next intCount
txtCode = txtCode & Right(Str(lngText ^ 2), Len(Str(lngText ^ 2)) - 1)
End Sub
