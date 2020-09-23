VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB RAW Viewer! - By Simon Price - visit www.VBgames.co.uk for more!"
   ClientHeight    =   5052
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   7932
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar CameraZ 
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   1812
   End
   Begin VB.PictureBox Backbuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4800
      Left            =   6720
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.PictureBox Primary 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4800
      Left            =   2040
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   5760
   End
   Begin VB.FileListBox File1 
      Height          =   3720
      Left            =   120
      Pattern         =   "*.raw"
      TabIndex        =   0
      Top             =   360
      Width           =   1812
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label lblZoom 
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1812
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' program to view .raw files, by Simon Price 15/11/00
' part of research for a 3D modelling program
' visit www.vbgames.co.uk for more


Option Explicit ' that's just in case wgreen7 is reading this code

Private Type POINTAPI
    x As Integer
    y As Integer
End Type

Private Type t3Dvector
    x As Single
    y As Single
    z As Single
End Type

Private Type tPolygon
    v(0 To 2) As t3Dvector
    p(0 To 2) As POINTAPI
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Polygon() As tPolygon
Private Camera As t3Dvector

Private PrimaryHDC As Long
Private BackbufferHDC As Long

Private EndNow As Boolean

Private Const CAMERALENS = 256
Private Const SPEEDX = 0.05
Private Const SPEEDY = 0.02
Private Const PI = 3.1415


Function LoadRAW(RawFilename As String) As Boolean
On Error GoTo FileError
MousePointer = vbHourglass
ReDim Polygon(0)
Open RawFilename For Input As #1
    On Error Resume Next
    Dim N(0 To 8) As Single
    Dim temp As String
    Dim i As Long
    Dim i2 As Long
    Dim i3 As Long
    Dim i4 As Long
    Do
        DoEvents
        i4 = 0
        Input #1, temp
        If Len(temp) > 16 Then
            For i = 1 To Len(temp)
                If Mid(temp, i, 1) = " " Then
                    N(i4) = Val(Mid(temp, i2, i3 - i2 + 1))
                    i4 = i4 + 1
                    Do
                        If Mid(temp, i + 1, 1) = " " Then
                            i = i + 1
                        Else
                            i2 = i
                            Exit Do
                        End If
                    Loop
                Else
                    i3 = i
                End If
            Next
            N(0) = Val(temp)
            N(8) = Val(Mid(temp, i2, i3 - i2 + 1))
            ReDim Preserve Polygon(0 To UBound(Polygon) + 1)
            With Polygon(UBound(Polygon))
                .v(0).x = N(0)
                .v(0).y = N(1)
                .v(0).z = N(2)
                .v(1).x = N(3)
                .v(1).y = N(4)
                .v(1).z = N(5)
                .v(2).x = N(6)
                .v(2).y = N(7)
                .v(2).z = N(8)
                For i = 0 To 2
                    If .v(i2).z > Camera.z / 2 Then Camera.z = .v(i2).z * 2
                Next
            End With
            End If
    Loop Until EOF(1)
Close #1
MousePointer = vbDefault
LoadRAW = True
Exit Function
FileError:
On Error Resume Next
Close #1
MousePointer = vbDefault
LoadRAW = False
End Function

Private Sub CameraZ_Change()
Camera.z = CameraZ
End Sub

Private Sub File1_Click()
On Error Resume Next
If LoadRAW(File1.Path & "/" & File1.FileName) Then
    CameraZ.Max = Camera.z * 5
    CameraZ = Camera.z
    MainBit
Else
    MsgBox "FILE ERROR: Could not load " & File1.FileName, vbExclamation, "FILE LOAD ERROR!"
End If
End Sub

Sub MainBit()
Dim i As Long
Dim i2 As Byte
Dim LensDivDistance As Single
Dim NumPolys As Long
Dim lpPoint As POINTAPI
Dim vec As t3Dvector
Dim COSofSPEEDX As Double
Dim COSofSPEEDY As Double
Dim SINofSPEEDX As Double
Dim SINofSPEEDY As Double
On Error Resume Next
NumPolys = UBound(Polygon)
If NumPolys < 1 Then Exit Sub
COSofSPEEDX = Cos(SPEEDX)
COSofSPEEDY = Cos(SPEEDY)
SINofSPEEDX = Sin(SPEEDX)
SINofSPEEDY = Sin(SPEEDY)
Do
    i = 0
    For i = 1 To NumPolys
        With Polygon(i)
            For i2 = 0 To 2
                vec.z = (.v(i2).z * COSofSPEEDX) - (.v(i2).x * SINofSPEEDX)
                .v(i2).x = (.v(i2).x * COSofSPEEDX) + (.v(i2).z * SINofSPEEDX)
                .v(i2).z = vec.z
                vec.y = (.v(i2).y * COSofSPEEDY) - (.v(i2).z * SINofSPEEDY)
                .v(i2).z = (.v(i2).z * COSofSPEEDY) + (.v(i2).y * SINofSPEEDY)
                .v(i2).y = vec.y
                LensDivDistance = CAMERALENS / (.v(i2).z - Camera.z)
                .p(i2).x = .v(i2).x * LensDivDistance + 240
                .p(i2).y = .v(i2).y * LensDivDistance + 200
            Next
            MoveToEx BackbufferHDC, .p(2).x, .p(2).y, lpPoint
            LineTo BackbufferHDC, .p(0).x, .p(0).y
            LineTo BackbufferHDC, .p(1).x, .p(1).y
            LineTo BackbufferHDC, .p(2).x, .p(2).y
        End With
    Next
    DoEvents
    BitBlt PrimaryHDC, 0, 0, 480, 400, BackbufferHDC, 0, 0, vbSrcCopy
    'Backbuffer.Cls
    Backbuffer.Line (0, 0)-(480, 400), vbBlack, BF
    If EndNow = True Then Exit Do
Loop
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
File1.Path = App.Path
PrimaryHDC = Primary.hdc
BackbufferHDC = Backbuffer.hdc
Camera.z = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
EndNow = True
End Sub
