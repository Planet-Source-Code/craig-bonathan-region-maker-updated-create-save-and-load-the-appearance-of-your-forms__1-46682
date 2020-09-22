VERSION 5.00
Begin VB.Form RegionTestForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Region Test"
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   LinkTopic       =   "Form1"
   ScaleHeight     =   79
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "RegionTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mover As New WindowMover

Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mover.StartMove X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mover.Move Me, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mover.EndMove
End Sub
