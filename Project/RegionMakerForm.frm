VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form RegionMakerForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Region Maker"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   409
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Load Region"
      Height          =   255
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Include"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      MousePointer    =   2  'Cross
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   480
      Width           =   5655
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Preview"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Picture"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "RegionMakerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileName As String

Private Sub Command1_Click()
    On Error GoTo SkipOpen::
    FileDialog.DialogTitle = "Load Bitmap"
    FileDialog.Filter = "Windows Bitmap (*.bmp)|*.bmp"
    FileDialog.ShowOpen
    On Error GoTo 0
    Picture1.Picture = LoadPicture(FileDialog.FileName)
    FileName = FileDialog.FileName
    If Picture1.Left * 2 + Picture1.Width > 409 Then
        Me.Width = (Picture1.Left * 2 + Picture1.Width) * Screen.TwipsPerPixelX
    Else
        Me.Width = 409 * Screen.TwipsPerPixelX
    End If
    If Picture1.Top * 2 + Picture1.Height + Picture1.Left > 80 Then
        Me.Height = (Picture1.Top * 2 + Picture1.Height + Picture1.Left) * Screen.TwipsPerPixelY
    Else
        Me.Height = 80 * Screen.TwipsPerPixelY
    End If
    
    Command3.Enabled = True
    Command4.Enabled = True
    Picture1.Enabled = True
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    FileDialog.FileName = ""
    Exit Sub
SkipOpen::
End Sub

Private Sub Command2_Click()
    Dim RegionData As RegionData_Type, PixelCount As Long
    On Error GoTo SkipOpen::
    FileDialog.DialogTitle = "Load Region Data"
    FileDialog.Filter = "Region Data (*.rgd)|*.rgd"
    FileDialog.ShowOpen
    On Error GoTo 0
    RegionData = LoadRegion(FileDialog.FileName)
    RegionTestForm.Show
    ProjectRegion RegionTestForm, RegionData
    If RegionData.Valid = False Then
        Unload RegionTestForm
        MsgBox ("Invalid region data")
        Exit Sub
    End If
    RegionTestForm.Width = RegionData.Width * Screen.TwipsPerPixelX
    RegionTestForm.Height = RegionData.Height * Screen.TwipsPerPixelY
    RegionTestForm.Left = (Screen.Width - RegionTestForm.Width) / 2
    RegionTestForm.Top = (Screen.Height - RegionTestForm.Height) / 2
    FileDialog.FileName = ""
    Exit Sub
SkipOpen::
End Sub

Private Sub Command3_Click()
    Dim RegionData As RegionData_Type, PixelCount As Long
    If Check1.Value = 0 Then
        PixelCount = CreateRegion(RegionMakerForm, FileName, Label2.BackColor, False, RegionData)
    Else
        PixelCount = CreateRegion(RegionMakerForm, FileName, Label2.BackColor, True, RegionData)
    End If
    If PixelCount < 100 Then
        MsgBox ("There would only be " & CStr(PixelCount) & " visible pixels in this region," & vbCrLf & _
                "which is not suitable for a window. Please choose" & vbCrLf & _
                "a different colour, mode or picture.")
        Exit Sub
    End If
    RegionTestForm.Show
    RegionTestForm.Width = RegionData.Width * Screen.TwipsPerPixelX
    RegionTestForm.Height = RegionData.Height * Screen.TwipsPerPixelY
    RegionTestForm.Picture = Picture1.Picture
    ProjectRegion RegionTestForm, RegionData
    RegionTestForm.Left = (Screen.Width - RegionTestForm.Width) / 2
    RegionTestForm.Top = (Screen.Height - RegionTestForm.Height) / 2
End Sub

Private Sub Command4_Click()
    Dim RegionData As RegionData_Type, PixelCount As Long
    On Error GoTo SkipSave::
    FileDialog.DialogTitle = "Save Region Data"
    FileDialog.Filter = "Region Data (*.rgd)|*.rgd"
    FileDialog.ShowSave
    On Error GoTo 0
    If Check1.Value = 0 Then
        PixelCount = CreateRegion(RegionMakerForm, FileName, Label2.BackColor, False, RegionData)
    Else
        PixelCount = CreateRegion(RegionMakerForm, FileName, Label2.BackColor, True, RegionData)
    End If
    If PixelCount < 100 Then
        MsgBox ("There would only be " & CStr(PixelCount) & " visible pixels in this region," & vbCrLf & _
                "which is not suitable for a window. Please choose" & vbCrLf & _
                "a different colour, mode or picture.")
        Exit Sub
    End If
    SaveRegion FileDialog.FileName, RegionData
    MsgBox ("Region data saved")
    FileDialog.FileName = ""
    Exit Sub
SkipSave::
End Sub

Private Sub Form_Activate()
    Unload RegionTestForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload RegionTestForm
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.BackColor = Picture1.Point(X, Y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.BackColor = Picture1.Point(X, Y)
End Sub
