VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Curvez"
   ClientHeight    =   5850
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton opn 
      Caption         =   "Open 8xi.jpg"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   240
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4200
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton reser 
      Caption         =   "Reset"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   0
      Width           =   1695
   End
   Begin VB.HScrollBar ba 
      Height          =   135
      Left            =   0
      Max             =   512
      TabIndex        =   11
      Top             =   5640
      Value           =   256
      Width           =   3855
   End
   Begin VB.HScrollBar ga 
      Height          =   135
      Left            =   0
      Max             =   512
      TabIndex        =   10
      Top             =   5160
      Value           =   256
      Width           =   3855
   End
   Begin VB.HScrollBar ra 
      Height          =   135
      Left            =   0
      Max             =   512
      TabIndex        =   9
      Top             =   4680
      Value           =   256
      Width           =   3855
   End
   Begin VB.HScrollBar bs 
      Height          =   255
      Left            =   0
      Max             =   512
      TabIndex        =   8
      Top             =   5400
      Value           =   256
      Width           =   3855
   End
   Begin VB.HScrollBar gs 
      Height          =   255
      Left            =   0
      Max             =   512
      TabIndex        =   7
      Top             =   4920
      Value           =   256
      Width           =   3855
   End
   Begin VB.HScrollBar rs 
      Height          =   255
      Left            =   0
      Max             =   512
      TabIndex        =   6
      Top             =   4440
      Value           =   256
      Width           =   3855
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      Left            =   3840
      Max             =   256
      TabIndex        =   5
      Top             =   0
      Value           =   128
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   4080
      Picture         =   "frm.frx":0000
      ScaleHeight     =   152
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   960
      Width           =   1920
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   360
      TabIndex        =   1
      Top             =   3840
      Value           =   180
      Width           =   3855
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   0
      Max             =   1000
      TabIndex        =   2
      Top             =   4080
      Value           =   1
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   0
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   3825
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.PictureBox shad 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   3960
      Picture         =   "frm.frx":255C
      ScaleHeight     =   152
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Menu Open 
      Caption         =   "Open"
   End
   Begin VB.Menu Save 
      Caption         =   "Save"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Dim Lum(255) As Integer
Dim rP(255) As Integer
Dim gP(255) As Integer
Dim bP(255) As Integer
Dim Small As Boolean
Dim mode As Integer
Dim mY As Integer
Dim mX As Integer

Sub doit()
Picture1.Refresh
Dim i As Integer
Dim h As Integer
Dim a As Double
Dim s As Integer
Dim m As Double
Dim df As Integer
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    Dim ra2 As Integer
    Dim ga2 As Integer
    Dim ba2 As Integer
    r = rs - 256
    g = gs - 256
    b = bs - 256
    ra2 = ra - 256
    ga2 = ga - 256
    ba2 = ba - 256


m = HScroll2 / 2
df = VScroll1 - 128

Set Picture1.Picture = Nothing


For i = 0 To 255
a = ((i / 255) * 360) * m
h = Sin(a * 3.1415 / 180) * r
Picture1.PSet (i, 255 - i + h + ra2), vbRed
rP(i) = i + h + ra2
Next
For i = 0 To 255
a = ((i / 255) * 360) * m
h = Sin(a * 3.1415 / 180) * g
Picture1.PSet (i, 255 - i + h + ga2), vbGreen
gP(i) = i + h + ga2
Next
For i = 0 To 255
a = ((i / 255) * 360) * m
h = Sin(a * 3.1415 / 180) * b
Picture1.PSet (i, 255 - i + h + ba2), vbBlue
bP(i) = i + h + ba2
Next

For i = 0 To 255
a = ((i / 255) * 360) * m
s = 180 - HScroll1
h = Sin(a * 3.1415 / 180) * s
Picture1.PSet (i, 255 - i + h + df), vbBlack
Lum(i) = h + df
Next

    Picture1.Picture = Picture1.Image
    Picture1.Refresh
End Sub

Private Sub Command1_Click()
   On Error Resume Next
    Dim x As Long, y As Long
    Dim iArray() As Byte
    Dim fDraw As New FastDrawing
    Dim stableR(255) As Integer
    Dim stableG(255) As Integer
    Dim stableB(255) As Integer

    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    r = rs - 128
    g = gs - 128
    b = bs - 128
    
    For x = 0 To 255
    stableR(x) = stabalize(rP(x) + Lum(x))
    stableG(x) = stabalize(gP(x) + Lum(x))
    stableB(x) = stabalize(bP(x) + Lum(x))
    Next
        
    fDraw.GetImageData shad, iArray()
    Dim TempWidth As Long, TempHeight As Long
    TempWidth = fDraw.GetImageWidth(Picture2) - 1
    TempHeight = fDraw.GetImageHeight(Picture2) - 1
    For x = 0 To TempWidth
    For y = 0 To TempHeight
        iArray(2, x, y) = stableR(iArray(2, x, y))
        iArray(1, x, y) = stableG(iArray(1, x, y))
        iArray(0, x, y) = stableB(iArray(0, x, y))
    Next y
    Next x
    
    fDraw.SetImageData Picture2, fDraw.GetImageWidth(Picture2), fDraw.GetImageHeight(Picture2), iArray()
    Erase iArray
    
    Picture2.Picture = Picture2.Image
    Picture2.Refresh
End Sub

Function stabalize(i)
stabalize = i
    If i > 255 Then stabalize = 255: Exit Function
    If i < 0 Then stabalize = 0: Exit Function
End Function

Private Sub Form_Load()
If Command = " " Then
MsgBox "Compile this program for Fast and Best Results.", vbInformation, "Compile"
End If
Small = True
mode = 1
End Sub

Private Sub HScroll1_Change()
doit
Command1_Click
mode = 1
End Sub

Private Sub HScroll1_Scroll()
doit
Command1_Click
End Sub

Private Sub Open_Click()
On Error Resume Next
cd.FileName = ""
cd.ShowOpen
If cd.FileName <> "" Then
reser_Click
Set shad.Picture = LoadPicture(cd.FileName)
Picture2.Width = shad.Width
Picture2.Height = shad.Height

'Exit Sub
'Width = (shad.Width + Picture1.Width + 17) * 15
'Height = (shad.Top + shad.Height + 5) * 15
'If Height < 6510 Then Height = 6510
'If shad.Height > 256 And shad.Width > 256 Then
'Small = False
'Else
'Small = True
'End If
doit
Command1_Click
End If
End Sub

Private Sub opn_Click()
reser_Click
Set shad.Picture = LoadPicture(App.Path & "\8Xi.jpg")
Picture2.Width = shad.Width
Picture2.Height = shad.Height
doit
Command1_Click

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
mY = y
mX = x
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then

If GetKeyState(82) < 0 And Shift <> 2 Then
        rs = -(mY - y) + 256
ElseIf GetKeyState(71) < 0 And Shift <> 2 Then
        gs = -(mY - y) + 256
ElseIf GetKeyState(66) < 0 And Shift <> 2 Then
        bs = -(mY - y) + 256
ElseIf GetKeyState(82) < 0 Then
        ra = -(mY - y) + 256
ElseIf GetKeyState(71) < 0 Then
        ga = -(mY - y) + 256
ElseIf GetKeyState(66) < 0 Then
        ba = -(mY - y) + 256
ElseIf Shift = 2 Then
        VScroll1.Value = -(mY - y) + 128
Else
        HScroll1.Value = (mY - y) + 180
End If

End If
End Sub

Private Sub ra_Change()
doit
Command1_Click
mode = 8
End Sub

Private Sub ra_Scroll()
If Small = True Then
doit
Command1_Click
End If
End Sub

Private Sub ga_Change()
doit
Command1_Click
mode = 5
End Sub

Private Sub ga_Scroll()
If Small = True Then
doit
Command1_Click
End If
End Sub

Private Sub ba_Change()
doit
Command1_Click
mode = 7
End Sub

Private Sub ba_Scroll()
If Small = True Then
doit
Command1_Click
End If
End Sub

Private Sub reser_Click()
VScroll1 = 128
DoEvents
HScroll1 = 180
DoEvents
HScroll2 = 1

rs = 256
DoEvents
gs = 256
DoEvents
bs = 256
DoEvents

ra = 256
DoEvents
ga = 256
DoEvents
ba = 256
DoEvents

End Sub

Private Sub rs_Change()
doit
Command1_Click
mode = 3
End Sub

Private Sub rs_Scroll()
If Small = True Then
doit
Command1_Click
End If
End Sub

Private Sub gs_Change()
doit
Command1_Click
mode = 4
End Sub

Private Sub gs_Scroll()
If Small = True Then
doit
Command1_Click
End If
End Sub

Private Sub bs_Change()
doit
Command1_Click
mode = 6
End Sub

Private Sub bs_Scroll()
If Small = True Then
doit
Command1_Click
End If
End Sub

Private Sub Save_Click()
On Error Resume Next
cd.FileName = ""
cd.ShowSave
If cd.FileName <> "" Then
SavePicture Picture2.Picture, cd.FileName
End If
End Sub

Private Sub vScroll1_Change()
doit
Command1_Click
mode = 2
End Sub

Private Sub vScroll1_Scroll()
If Small = True Then
doit
Command1_Click
End If
End Sub

Private Sub HScroll2_Change()
doit
Command1_Click
End Sub

Private Sub HScroll2_Scroll()
If Small = True Then
doit
Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
doit
Timer1 = False
End Sub

Private Sub Timer2_Timer()
Command1_Click
End Sub
