VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Timer tmrInt 
      Left            =   4080
      Top             =   2640
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API-функция для получения копии всего экрана
Private Declare Function BitBlt _
Lib "gdi32" ( _
ByVal hDestDC As Long, _
ByVal IngX As Long, ByVal IngY As Long, _
ByVal nWidth As Long, _
ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal IngXSrc As Long, ByVal IngYSrc As Long, _
ByVal dwRop As Long _
) As Long
'API-функция для получения описателя дисплейного изображения
Private Declare Function GetDesktopWindow _
Lib "user32" () As Long
'API-функция для получения контекста устройства по описателю
Private Declare Function GetDC _
Lib "user32" ( _
ByVal hwnd As Long _
) As Long
'API-функция,освобождающая контекст устройства
Private Declare Function ReleaseDC _
Lib "user32" ( _
ByVal hwnd As Long, _
ByVal hdc As Long _
) As Long

'константы для некоторых API-функций
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Declare Function SystemParametersInfo _
Lib "user32" Alias "SystemParametersInfoA" ( _
ByVal uAction As Long, _
ByVal uParam As Long, _
ByRef IpvParam As Any, _
ByVal fuWinlni As Long _
) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndlnsertAfter As Long, ByVal x As Long, _
ByVal у As Long, _
ByVal ex As Long, _
ByVal cy As Long, _
ByVal wFlags As Long _
) As Long
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPHOST = -1
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPHOST = -2

    Dim dwRop As Long
    Dim hwndSrc As Long
    Dim hSrcDC As Long
    Dim lngRes As Long
    Dim strDir As String
    Dim intInter As Integer

Private Sub Form_Load()
    Open App.Path & "\Setting" For Input As #1
        Input #1, strDir
        strDir = strDir & "\"
        Input #1, intInter
        tmrInt.Interval = (Abs(intInter) And 7) * 1000
    Close #1
End Sub

Private Sub tmrInt_Timer()
Dim strFile As String
Dim i, x, y As Integer
    'копируем весь экран в окно рисунка
    ScaleMode = vbPixels
    Move 0, 0, Screen.Width + 1, Screen.Height + 1
    dwRop = &HCC0020
    hwndSrc = GetDesktopWindow()
    hSrcDC = GetDC(hwndSrc)
    lngRes = BitBlt(frmMain.hdc, 0, 0, ScaleWidth, ScaleHeight, hSrcDC, 0, 0, dwRop)
    i = 0
    pic.Move 0, 0, ScaleWidth \ 2, ScaleHeight \ 2
    For x = 0 To ScaleWidth Step 2
        BitBlt pic.hdc, i, 0, 1, ScaleHeight, frmMain.hdc, x, 0, vbSrcAnd
        i = i + 1
    Next x
    i = 0
    For y = 0 To ScaleHeight Step 2
        BitBlt pic.hdc, 0, i, ScaleWidth \ 2, 1, pic.hdc, 0, y, vbSrcCopy
        i = i + 1
    Next y
    pic.Move 0, 0, ScaleWidth \ 2, ScaleHeight \ 2
    pic.Refresh
    'Save
    strFile = strDir & Replace(Now, ":", ".") & ".bmp"
    SavePicture pic.image, strFile
    Set pic.Picture = Nothing
End Sub
