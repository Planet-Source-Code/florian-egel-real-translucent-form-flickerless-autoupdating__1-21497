VERSION 5.00
Begin VB.Form frmTransForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'Kein
   Caption         =   "TransForm - by Florian Egel"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   Icon            =   "TransForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   4230
      Left            =   4200
      ScaleHeight     =   4230
      ScaleWidth      =   3030
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   4800
      Left            =   0
      ScaleHeight     =   4800
      ScaleWidth      =   3885
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      Begin VB.Image Image3 
         Height          =   420
         Left            =   1215
         MouseIcon       =   "TransForm.frx":1FF2
         MousePointer    =   99  'Benutzerdefiniert
         Top             =   4380
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   210
         Left            =   3555
         MouseIcon       =   "TransForm.frx":2144
         MousePointer    =   99  'Benutzerdefiniert
         Top             =   345
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   210
         Left            =   3285
         MouseIcon       =   "TransForm.frx":2296
         MousePointer    =   99  'Benutzerdefiniert
         Top             =   345
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1215
      Left            =   1680
      ScaleHeight     =   1215
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmTransForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The code will not work good so good in the IDE,
'it probably leaves "ghosts" on the screen,
'so compile the code or rename the TransForm.exe_
'to Transform.exe.

'The code makes use of a function of FoxCBmp3.dll, which I also
'posted on Planet-Source-Code, along with many sample codes
'showing how to use it. go to
'http://www.planet-source-code.com/xq/ASP/txtCodeId.21470/lngWId.1/qx/vb/scripts/ShowCode.htm

'Have fun, Florian Egel

Private Declare Function FoxAlphaBlend Lib "FoxCBmp3.dll" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hScrDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal alpha As Byte, Optional ByVal MaskColor As Long, Optional ByVal Flags As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim CurX As Single, CurY As Single
Dim WH As Long, WD As Long

Private Sub Form_Load()
    Picture3.Picture = LoadPicture("back.bmp")
    Width = Picture3.Width
    Height = Picture3.Height
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image1_Click()
    WindowState = 1
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub Image3_Click()
    MsgBox "Test"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CurX = X
        CurY = Y
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DeltaX As Long, DeltaY As Long
    Dim WH As Long, WD As Long
    If Button = 1 Then
        WH = GetDesktopWindow
        WD = GetDC(WH)
        DeltaX = X - CurX
        DeltaY = Y - CurY
        BitBlt Picture2.HDC, 0, 0, Width \ 15, Height \ 15, Picture2.HDC, DeltaX \ 15, DeltaY \ 15, vbSrcCopy
        If DeltaX > 0 Then
            BitBlt Picture2.HDC, (Width - DeltaX) \ 15, 0, DeltaX \ 15, ScaleHeight \ 15, WD, (Left + Width) \ 15, (Top + DeltaY) \ 15, vbSrcCopy
        ElseIf DeltaX < 0 Then
            BitBlt Picture2.HDC, 0, 0, -DeltaX \ 15, Height \ 15, WD, (Left + DeltaX) \ 15, (Top + DeltaY) \ 15, vbSrcCopy
        End If
        If DeltaY > 0 Then
            BitBlt Picture2.HDC, 0, (ScaleHeight - DeltaY) \ 15, ScaleWidth \ 15, DeltaY \ 15, WD, (Left + DeltaX) \ 15, (Top + Height) \ 15, vbSrcCopy
        ElseIf DeltaY < 0 Then
            BitBlt Picture2.HDC, 0, 0, ScaleWidth \ 15, -DeltaY \ 15, WD, (Left + DeltaX) \ 15, (Top + DeltaY) \ 15, vbSrcCopy
        End If
        Picture2.Refresh
        BitBlt Picture1.HDC, 0, 0, Width \ 15, Height \ 15, Picture2.HDC, 0, 0, vbSrcCopy
        FoxAlphaBlend Picture1.HDC, 0, 0, Width \ 15, Height \ 15, Picture3.HDC, 0, 0, 128, &HFF00FF, 1
        Move Left + DeltaX, Top + DeltaY
        Picture1.Refresh
        BitBlt Me.HDC, 0, 0, Width \ 15, Height \ 15, Picture1.HDC, 0, 0, vbSrcCopy
        Sleep 10
        ReleaseDC WH, WD
    End If
End Sub

Private Sub Form_Resize()
    Picture1.Move 0, 0, Width, Height
    Picture2.Move 0, 0, Width, Height
    WH = GetDesktopWindow
    WD = GetDC(WH)
    BitBlt Picture2.HDC, 0, 0, Width \ 15, Height \ 15, WD, Left \ 15, Top \ 15, vbSrcCopy
    BitBlt Picture1.HDC, 0, 0, Width \ 15, Height \ 15, Picture2.HDC, 0, 0, vbSrcCopy
    FoxAlphaBlend Picture1.HDC, 0, 0, Width \ 15, Height \ 15, Picture3.HDC, 0, 0, 128, &HFF00FF, 1
    ReleaseDC WH, WD
    Picture2.Refresh
End Sub

Private Sub Picture1_Resize()
    'cmdExit.Move Picture1.Width - cmdExit.Width - 45, 45
    'cmdMinimize.Move cmdExit.Left - cmdMinimize.Width - 30, 45
End Sub
