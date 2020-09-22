Attribute VB_Name = "RegionManagement"
Option Explicit

Private Const Header As String = "RegionData"

' Region functions
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

' Bitmap functions
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

Type RegionData_Type
    Width As Long
    Height As Long
    Data() As Byte
    Valid As Boolean
End Type

Function ProjectRegion(Window As Form, RegionData As RegionData_Type)
    Dim X As Long, Y As Long, Temp As Byte
    Dim RegionHandle1 As Long, RegionHandle2 As Long
    If Window.BorderStyle <> 0 Then
        MsgBox ("Error applying region: Window must not have a border")
        Exit Function
    End If
    For X = 0 To RegionData.Width - 1
        For Y = 0 To RegionData.Height - 1
            Temp = RegionData.Data(X * RegionData.Height + Y)
            If Temp = 1 Then
                If RegionHandle1 = 0 Then
                    RegionHandle1 = CreateRectRgn(X, Y, X + 1, Y + 1)
                Else
                    RegionHandle2 = CreateRectRgn(X, Y, X + 1, Y + 1)
                    CombineRgn RegionHandle1, RegionHandle1, RegionHandle2, 2
                    DeleteObject RegionHandle2
                End If
            End If
        Next
    Next
    SetWindowRgn Window.hWnd, RegionHandle1, True
End Function

Function CreateRegion(Window As Form, BitmapFile As String, Colour As Long, Include As Boolean, _
        RegionData As RegionData_Type) As Long
    Dim Picture As IPictureDisp
    Dim Width As Long, Height As Long, X As Long, Y As Long
    Dim NewHandle As Long, OldHandle As Long, DeviceContextHandle As Long
    Dim ShownPixelCount As Long
    
    If BitmapFile <> "" Then
        Set Picture = LoadPicture(BitmapFile)
    Else
        Set Picture = Window.Picture
    End If
    Width = Window.ScaleX(Picture.Width, vbHimetric, vbPixels)
    Height = Window.ScaleY(Picture.Height, vbHimetric, vbPixels)
    NewHandle = Picture.Handle
    DeviceContextHandle = CreateCompatibleDC(0)
    OldHandle = SelectObject(DeviceContextHandle, NewHandle)
    
    ReDim RegionData.Data(Width * Height - 1)
    RegionData.Width = Width
    RegionData.Height = Height
    For X = 0 To Width - 1
        For Y = 0 To Height - 1
            If Include = True Then
                If GetPixel(DeviceContextHandle, X, Y) = Colour Then
                    RegionData.Data(X * Height + Y) = 1
                    ShownPixelCount = ShownPixelCount + 1
                Else
                    RegionData.Data(X * Height + Y) = 0
                End If
            Else
                If GetPixel(DeviceContextHandle, X, Y) = Colour Then
                    RegionData.Data(X * Height + Y) = 0
                Else
                    RegionData.Data(X * Height + Y) = 1
                    ShownPixelCount = ShownPixelCount + 1
                End If
            End If
        Next
    Next
    
    NewHandle = SelectObject(DeviceContextHandle, OldHandle)
    DeleteDC DeviceContextHandle
    Set Picture = Nothing
    
    CreateRegion = ShownPixelCount
End Function

Function SaveRegion(FileName As String, RegionData As RegionData_Type)
    Dim FileNum As Long
    On Error Resume Next
    Kill FileName
    On Error GoTo 0
    FileNum = FreeFile
    Open FileName For Binary As #FileNum
        Put #FileNum, 1, Header
        Put #FileNum, 1 + LenB(Header), RegionData.Width
        Put #FileNum, 5 + LenB(Header), RegionData.Height
        Put #FileNum, 9 + LenB(Header), RegionData.Data()
    Close #FileNum
End Function

Function LoadRegion(FileName As String) As RegionData_Type
    Dim FileNum As Long, Temp As String
    FileNum = FreeFile
    LoadRegion.Valid = True
    Open FileName For Binary As #FileNum
        Temp = String(Len(Header), Chr(0))
        Get #FileNum, 1, Temp
        If Temp <> "RegionData" Then
            LoadRegion.Valid = False
            Close #FileNum
            Exit Function
        End If
        Get #FileNum, 1 + LenB(Header), LoadRegion.Width
        Get #FileNum, 5 + LenB(Header), LoadRegion.Height
        ReDim LoadRegion.Data(LoadRegion.Width * LoadRegion.Height - 1)
        Get #FileNum, 9 + LenB(Header), LoadRegion.Data()
    Close #FileNum
End Function
