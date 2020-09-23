Attribute VB_Name = "Module1"
' Module1 (Module1.bas)
Option Explicit

' APIs
' For extracting
Public Declare Function BitBlt Lib "gdi32" _
   (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

'--------------------------------------------------------------------------
'  API to make application stay on top

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Public Const hWndInsertAfter = -1
Public Const wFlags = &H40 Or &H20

'------------------------------------------------------------------------------

' Twips/pixel
Public STX As Long
Public STY As Long

Public SurroundColor As Long

Public Const pi# = 3.14159265

Public Sub FixScrollbars(picC As PictureBox, picP As PictureBox, HS As HScrollBar, VS As VScrollBar)
   ' picC = Container = picFrame
   ' picP = Picture   = picDisplay
      HS.Max = picP.Width - picC.Width + 12   ' +4 to allow for border
      VS.Max = picP.Height - picC.Height + 12 ' +4 to allow for border
      HS.LargeChange = picC.Width \ 10
      HS.SmallChange = 1
      VS.LargeChange = picC.Height \ 10
      VS.SmallChange = 1
      HS.Top = picC.Top + picC.Height + 1
      HS.Left = picC.Left
      HS.Width = picC.Width
      If picP.Width < picC.Width Then
         HS.Visible = False
         'HS.Enabled = False
      Else
         HS.Visible = True
         'HS.Enabled = True
      End If
      VS.Top = picC.Top
      VS.Left = picC.Left - VS.Width - 1
      VS.Height = picC.Height
      If picP.Height < picC.Height Then
         VS.Visible = False
         'VS.Enabled = False
      Else
         VS.Visible = True
         'VS.Enabled = True
      End If
End Sub

