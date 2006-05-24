VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Executer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Sub Class_Terminate()
' выгружаем форму при уничтожении объекта
Unload Dialog
End Sub
Private Sub Class_Initialize()
' создаем форму сразу после создания объекта, вызывая метод Show
Show
End Sub
Property Get Password() As String
Password = Dialog.Text1 & Dialog.Text2
End Property
Sub Show()
Dialog.Show vbModal
End Sub

