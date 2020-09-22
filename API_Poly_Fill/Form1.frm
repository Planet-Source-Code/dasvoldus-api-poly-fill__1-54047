VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WTF"
   ClientHeight    =   375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleWidth      =   1695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   0
      Top             =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type COORD
    x As Long
    y As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Const ALTERNATE = 1 ' ALTERNATE and WINDING are
Const WINDING = 2 ' constants for FillMode.
Const BLACKBRUSH = 4 ' Constant for brush type.





Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()

    Dim Position As POINTAPI
    
    GetCursorPos Position 'Get the current cursor position
       
    Dim poly(1 To 3) As COORD, NumCoords As Long, hBrush As Long, hRgn As Long
    
    ' Number of vertices in polygon.
    NumCoords = 3
    ' Set scalemode to pixels to set up points of triangle.
  
    
    poly(1).x = Position.x - Int(Rnd * 20)
    poly(1).y = Position.y - Int(Rnd * 20)
    poly(2).x = Position.x + Int(Rnd * 20)
    poly(2).y = Position.y + Int(Rnd * 20) '
    poly(3).x = Position.x - Int(Rnd * 20)
    poly(3).y = Position.y + Int(Rnd * 20)
    
    
    ' Polygon function creates unfilled polygon on screen.
    ' Remark FillRgn statement to see results.
    Polygon GetWindowDC(0), poly(1), NumCoords
    ' Gets stock black brush.
    hBrush = CreateSolidBrush(RGB(Rnd * 127, Rnd * 127, Rnd * 127)) 'GetStockObject(BLACKBRUSH)
    'RGB(Rnd * 127, Rnd * 127, Rnd * 127)

    ' Creates region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    ' If the creation of the region was successful then color.
    If hRgn Then FillRgn GetWindowDC(0), hRgn, hBrush
    DeleteObject hRgn
    DeleteObject hBrush
    
    
End Sub
