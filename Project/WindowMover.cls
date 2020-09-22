VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim OldX As Single, OldY As Single, Moving As Boolean

Function StartMove(X As Single, Y As Single)
    OldX = X
    OldY = Y
    Moving = True
End Function

Function Move(Window As Form, X As Single, Y As Single)
    If Moving = False Then Exit Function
    Window.Left = Window.Left + X - OldX
    Window.Top = Window.Top + Y - OldY
End Function

Function EndMove()
    Moving = False
End Function
