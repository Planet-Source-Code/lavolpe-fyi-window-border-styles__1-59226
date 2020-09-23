VERSION 5.00
Begin VB.Form frmBorders 
   Caption         =   "Click PicBox & Then Styles"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   15
      Left            =   7140
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   15
      Top             =   90
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   14
      Left            =   7140
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   14
      Top             =   1290
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   13
      Left            =   7140
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   13
      Top             =   2490
      Width           =   2250
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "frmBorders.frx":0000
      Left            =   7185
      List            =   "frmBorders.frx":0010
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   3840
      Width           =   2145
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   11
      Left            =   4815
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   11
      Top             =   3690
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   10
      Left            =   2505
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   10
      Top             =   3690
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   9
      Left            =   180
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   9
      Top             =   3690
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   8
      Left            =   4815
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   8
      Top             =   2490
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   7
      Left            =   2505
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   7
      Top             =   2490
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   6
      Left            =   180
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   6
      Top             =   2490
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   5
      Left            =   4815
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   5
      Top             =   1290
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   4
      Left            =   2505
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   4
      Top             =   1290
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   3
      Left            =   180
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   3
      Top             =   1290
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   2
      Left            =   4815
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   2
      Top             =   90
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   1
      Left            =   2505
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   1
      Top             =   90
      Width           =   2250
   End
   Begin VB.PictureBox Picbox 
      AutoRedraw      =   -1  'True
      Height          =   1155
      Index           =   0
      Left            =   180
      ScaleHeight     =   1095
      ScaleWidth      =   2190
      TabIndex        =   0
      Top             =   90
      Width           =   2250
   End
End
Attribute VB_Name = "frmBorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_BORDER As Long = &H800000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_EX_DLGMODALFRAME As Long = &H1&
Private Const WS_DLGFRAME As Long = &H400000
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE As Long = -16
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private picID As Integer


Private Sub List1_Click()
Dim lStyles(0 To 4) As Long
lStyles(1) = WS_BORDER
lStyles(2) = WS_THICKFRAME
lStyles(3) = WS_EX_CLIENTEDGE
lStyles(4) = WS_EX_STATICEDGE

Dim I As Integer, sType As String
Dim pStyle As Long, pExStyle As Long
With Picbox(picID)
    ' get current styles
    pStyle = GetWindowLong(.hwnd, GWL_STYLE)
    pExStyle = GetWindowLong(.hwnd, GWL_EXSTYLE)
    ' remove border styles
    For I = 1 To 2
        pStyle = pStyle And Not lStyles(I)
    Next
    For I = 3 To 4
        pExStyle = pExStyle And Not lStyles(I)
    Next
    ' calculate new border styles
    For I = 0 To 1
        If List1.Selected(I) Then
            pStyle = pStyle Or lStyles(I + 1)
            sType = sType & Choose(I + 1, "ws_border,", "ws_thickframe,")
        End If
    Next
    For I = 2 To 3
        If List1.Selected(I) Then
            pExStyle = pExStyle Or lStyles(I + 1)
            sType = sType & Choose(I - 1, "ws_ex_clientedge,", "ws_ex_staticedge,")
        End If
    Next
    ' apply new styles
    SetWindowLong .hwnd, GWL_STYLE, pStyle
    SetWindowLong .hwnd, GWL_EXSTYLE, pExStyle
    ' force styles to be drawn
    SetWindowPos .hwnd, 0, 0, 0, 0, 0, 55
    
    ' calculate border width
    Dim cRect As RECT, wRect As RECT, cPt As POINTAPI
    GetWindowRect .hwnd, wRect
    GetClientRect .hwnd, cRect
    ClientToScreen .hwnd, cPt
    OffsetRect cRect, cPt.x, cPt.y
    
    .Cls
End With

    Picbox(picID).Print Replace$(sType, ",", vbCrLf); "Border cx/cy: "; wRect.Right - cRect.Right
End Sub

Private Sub Picbox_Click(Index As Integer)
picID = Index
End Sub
