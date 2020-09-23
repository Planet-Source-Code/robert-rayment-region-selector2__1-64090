Attribute VB_Name = "Selector"
' Selector.bas

' Drawing selected shapes using Shape controls

' Requirements:
' On main Form: Shape Rectangle(SR), Shape Line(SL(0)) (Xor DrawMode) on a display Picture box,
'               an invisible picture box (P, pixels) for mask.
'               Label for X,Y coords
Option Explicit

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
' Used:-
'ar = SetRect(IR, 0, 0, picWidth - 1, picHeight - 1)
'ar = InvertRect(PIC(1).hdc, IR)

' To invert picture box
Public Declare Function SetRect Lib "user32" (lpRect As RECT, _
   ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function InvertRect Lib "user32" _
   (ByVal hdc As Long, lpRect As RECT) As Long
' To fill mask
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1


' Publics
Public picWidth As Long, picHeight As Long
Public aSelect As Boolean
Public aSelectDone As Boolean
Public SelectType As Long
Public ExtractType As Long

' Rectangle selection coords
Public XS1 As Single
Public YS1 As Single
Public XS2 As Single
Public YS2 As Single
Public SW As Long    ' Selection width
Public SH As Long    ' Selection height

' Final Rectangle selection coords
' at MouseUp
Public SXS1 As Single
Public SYS1 As Single
Public SXS2 As Single
Public SYS2 As Single
' For Square, Circle & Rounded Square
Public zDiagRad As Single
'Mover start
Public XM As Single
Public YM As Single
' Mover increment
Public XD As Single
Public YD As Single
Public aMover As Boolean
Public aMouseDown As Boolean

' For lasso & polys
Public NumLassoLines As Long
' For polys
Public NumLeftClicks As Long

Public Sub StartAll(frm As Form, P As PictureBox, Button As Integer, X As Single, Y As Single)
' From PIC_MouseDown,  P is picBack ( invisible mask )
Dim i As Long
Dim IR As RECT
Dim ar As Long
   If aSelect Then
   If Button = vbLeftButton Then
         aSelectDone = False
         Select Case SelectType
         Case 0 To 5: Start_Select frm, frm.SR, X, Y
         Case 6: Start_Lasso frm, X, Y       ' Lasso
            frm.LabWH = ""
         Case 7                          ' Poly
            frm.LabWH = ""
            NumLeftClicks = NumLeftClicks + 1
            If NumLeftClicks = 1 Then
               Start_Poly frm, X, Y
            End If
         End Select
   ElseIf Button = vbRightButton Then
      If SelectType = 7 Then ' Poly
      If NumLeftClicks > 1 Then
         If Not aMover Then
            ' Close shape
            NumLassoLines = NumLassoLines + 1
            Load frm.SL(NumLassoLines - 1)
            With frm.SL(NumLassoLines - 1)
               .X1 = X
               .Y1 = Y
               .X2 = frm.SL(0).X1
               .Y2 = frm.SL(0).Y1
               .Visible = True
            End With
         End If
         aSelectDone = True
      
         ' Transfer to picBack
         For i = 0 To NumLassoLines - 1
            With frm.SL(i)
               P.Line (.X1, .Y1)-(.X2, .Y2), vbWhite
            End With
         Next i
         Fill Form1.picBack, 1, 1   ' [W(B)]
         ' Invert
         ' Have mask White surrounded by Black [B(W)]
         ar = SetRect(IR, 0, 0, picWidth, picHeight)
         'ar = InvertRect(frm.picBack.hdc, IR)
         ar = InvertRect(P.hdc, IR)
         'frm.picBack.Refresh
         P.Refresh
         NumLeftClicks = 0
         FindMaxMins frm
         frm.LabWH = Str$(SW + 1) & " ," & Str$(SH + 1) & " "

      End If
      End If
   End If
   End If
End Sub

Public Sub MoveAll(frm As Form, Button As Integer, X As Single, Y As Single)
' From PIC_MouseMove
   If aSelect Then
      If Not aSelectDone Then
         If Button = vbLeftButton Then
            Select Case SelectType
            Case 0 To 5: Draw_Select frm.SR, X, Y
               frm.LabWH = Str$(SW + 1) & " ," & Str$(SH + 1) & " "
            Case 6: Draw_Lasso frm, X, Y      ' Lasso
            Case 7: Draw_Poly frm, X, Y       ' Poly
            End Select
         ElseIf NumLeftClicks > 0 And SelectType = 7 Then
            Draw_Poly frm, X, Y        ' Poly
         End If
      End If
   End If

End Sub

Public Sub MouseUpAll(frm As Form, P As PictureBox, Button As Integer, X As Single, Y As Single)
' From PIC_MouseUp,  P is picBack ( invisible mask )
Dim i As Long
Dim zrad As Single
Dim zaspect As Single
Dim IR As RECT
Dim ar As Long
   
   If aSelect Then
      If Not aSelectDone Then
         With P
            .BackColor = 0
            .Cls
            .DrawWidth = 2
            .FillColor = vbWhite
            .FillStyle = vbFSSolid
         End With
      End If
      Select Case SelectType
      Case 0, 1: ' Rectangle & Square
         aSelectDone = True
         P.Line (SXS1, SYS1)-Step(SW, SH), vbWhite, B
      
      Case 4, 5: ' Rounded rectangle & Rounded Square
         ' Corner circ radius = 1/6 of shortest dimension
         aSelectDone = True
         If frm.SR.Width > frm.SR.Height Then
            zrad = frm.SR.Height / 6
         Else
            zrad = frm.SR.Width / 6
         End If
         With frm.SR
            SXS1 = .Left
            If SXS1 < 0 Then SXS1 = 1
            SXS2 = .Left + SW
            If SXS2 >= picWidth Then
               SXS2 = picWidth - 1
               SW = SXS2 - SXS1
            End If
            SYS1 = .Top
            If SYS1 < 0 Then SYS1 = 1
            SYS2 = .Top + SH
            If SYS2 >= picHeight Then
               SYS2 = picHeight - 1
               SH = SYS2 - SYS1
            End If
         End With
         P.Line (SXS1 + zrad, SYS1)-(SXS2 - zrad, SYS1), vbWhite ' Top
         P.Line (SXS1 + zrad, SYS2)-(SXS2 - zrad, SYS2), vbWhite ' Bottom
         P.Line (SXS1, SYS1 + zrad)-(SXS1, SYS2 - zrad), vbWhite ' Left
         P.Line (SXS2, SYS1 + zrad)-(SXS2, SYS2 - zrad), vbWhite ' Right
         P.Circle (SXS1 + zrad, SYS1 + zrad), zrad, vbWhite, pi# / 2, pi#     ' TL
         P.Circle (SXS2 - zrad - 1, SYS1 + zrad), zrad, vbWhite, 0, pi# / 2     ' TR
         P.Circle (SXS1 + zrad + 1, SYS2 - zrad - 1), zrad, vbWhite, pi#, 3 * pi# / 2 ' BL
         P.Circle (SXS2 - zrad - 1, SYS2 - zrad - 2), zrad, vbWhite, 3 * pi# / 2, 0 ' BR
         
         Fill P, (SXS1 + SXS2) / 2, (SYS1 + SYS2) / 2
         
      Case 2 ' Oval
         aSelectDone = True
         ' zaspect ' <1 horz, >1 vert
         zaspect = frm.SR.Height / frm.SR.Width
         If zaspect >= 1 Then
            'zrad = Abs(((SXS1 + SXS2) / 2 - XS1)) * zaspect
            zrad = Abs(SW \ 2) * zaspect
         Else
            If zaspect = 0 Then zaspect = 4
            'zrad = Abs(((SYS1 + SYS2) / 2 - SYS1)) / zaspect
            zrad = Abs(SH \ 2) / zaspect
         End If
         If zrad = 0 Then zrad = 1
         P.Circle ((SXS1 + SXS2) / 2, (SYS1 + SYS2) / 2), zrad, vbWhite, , , zaspect
      
      Case 3   ' Circle
         aSelectDone = True
         With frm.SR
            SXS1 = .Left
            SXS2 = .Left + SW
            SYS1 = .Top
            SYS2 = .Top + SH
         End With
         zrad = Abs(SW \ 2)
         P.Circle ((SXS1 + SXS2) / 2, (SYS1 + SYS2) / 2), zrad, vbWhite
      
      Case 6  ' Lasso
         aSelectDone = True
         If Not aMover Then
            If X < 1 Then X = 1
            If Y < 1 Then Y = 1
            If X > picWidth - 2 Then X = picWidth - 2
            If Y > picHeight - 2 Then Y = picHeight - 2
            ' Close shape
            NumLassoLines = NumLassoLines + 1
            Load frm.SL(NumLassoLines - 1)
            With frm.SL(NumLassoLines - 1)
               .X1 = X
               .Y1 = Y
               .X2 = frm.SL(0).X1
               .Y2 = frm.SL(0).Y1
               .Visible = True
            End With
         End If
         ' Transfer to picBack
         For i = 0 To NumLassoLines - 1
            With frm.SL(i)
               P.Line (.X1, .Y1)-(.X2, .Y2), vbWhite
            End With
         Next i
         Fill P, 1, 1   ' [W(B)]
         ' Invert
         ' Have mask White surrounded by Black [B(W)]
         ar = SetRect(IR, 0, 0, picWidth, picHeight)
         ar = InvertRect(P.hdc, IR)
         P.Refresh
         
         FindMaxMins frm
         frm.LabWH = Str$(SW + 1) & " ," & Str$(SH + 1) & " "

      
      Case 7  ' Poly - finished with RightClick at MouseDown
         If Not aMover Then
            If X < 1 Then X = 1
            If Y < 1 Then Y = 1
            If X > picWidth - 2 Then X = picWidth - 2
            If Y > picHeight - 2 Then Y = picHeight - 2
            ' Close shape
            NumLassoLines = NumLassoLines + 1
            Load frm.SL(NumLassoLines - 1)
            With frm.SL(NumLassoLines - 1)
               .X1 = X
               .Y1 = Y
               .X2 = X
               .Y2 = Y
               .Visible = True
            End With
         End If
      End Select
      
      
      If aMover Then
         NumLeftClicks = 2    ' So it filters through
         Button = 2           ' Simulate Right click
         StartAll frm, P, Button, X, Y
      End If
      
      If aSelectDone Then
         With P
            .DrawWidth = 1
            .FillStyle = vbFSTransparent
         End With
      
         P.Refresh
      End If
   End If

End Sub


Public Sub Start_Select(frm As Form, SR As Shape, X As Single, Y As Single)
   With frm.SR
      .Left = X
      .Top = Y
      .Width = 4
      .Height = 4
      .Visible = True
   End With
   XS1 = X: YS1 = Y
   XS2 = X + 4: YS2 = Y + 4
End Sub

Public Sub Start_Lasso(frm As Form, X As Single, Y As Single)
   CheckLassoLines frm
   If X < 3 Then X = 3
   If Y < 3 Then Y = 3
   If X > picWidth - 3 Then X = picWidth - 3
   If Y > picHeight - 3 Then Y = picHeight - 3
   XS1 = X
   YS1 = Y
   XS2 = X
   YS2 = Y
   With frm.SL(0)
      .X1 = XS1
      .Y1 = YS1
      .X2 = XS2
      .Y2 = YS2
   End With
   frm.SL(0).Visible = True
End Sub

Public Sub Start_Poly(frm As Form, X As Single, Y As Single)
   CheckLassoLines frm
   NumLeftClicks = 1
   If X < 3 Then X = 3
   If Y < 3 Then Y = 3
   If X > picWidth - 3 Then X = picWidth - 3
   If Y > picHeight - 3 Then Y = picHeight - 3
   XS1 = X
   YS1 = Y
   XS2 = X
   YS2 = Y
   With frm.SL(0)
      .X1 = XS1
      .Y1 = YS1
      .X2 = XS2
      .Y2 = YS2
   End With
   Form1.SL(0).Visible = True
End Sub

Public Sub Draw_Select(SR As Shape, X As Single, Y As Single)
   If X < 1 Then X = 1
   If Y < 1 Then Y = 1
   If X > picWidth - 2 Then X = picWidth - 2
   If Y > picHeight - 2 Then Y = picHeight - 2
   
   If SelectType = 3 Then ' Circle  XS1,YS1, zDiagRad
      SXS1 = XS1 - zDiagRad
      SYS1 = YS1 - zDiagRad
      SXS2 = XS1 + zDiagRad
      SYS2 = YS1 + zDiagRad
      SW = 2 * zDiagRad
      SH = SW
      With SR
         .Left = SXS1
         .Top = SYS1
      End With
   Else
      SW = Abs(X - XS1)
      SH = Abs(Y - YS1)
      
      If X > XS1 Then
         SR.Left = XS1
         SXS1 = XS1
         SXS2 = X
      Else  ' X <= XS1 Then
         SR.Left = X
         SXS1 = X
         SXS2 = X + SW
      End If
      If Y > YS1 Then
         SR.Top = YS1
         SYS1 = YS1
         SYS2 = Y
      Else   'Y <= YS1 Then
         SR.Top = Y
         SYS1 = Y
         SYS2 = Y + SH
      End If
   End If
   
   With SR     ' Rectangle, Oval, Rounded rectangle
      .Width = SW
      .Height = SH
   End With
End Sub

Public Sub Draw_Lasso(frm As Form, X As Single, Y As Single)
   If X < 3 Then X = 3
   If Y < 3 Then Y = 3
   If X > picWidth - 3 Then X = picWidth - 3
   If Y > picHeight - 3 Then Y = picHeight - 3
   XS1 = XS2
   YS1 = YS2
   XS2 = X
   YS2 = Y
   NumLassoLines = NumLassoLines + 1
   Load frm.SL(NumLassoLines - 1)
   With frm.SL(NumLassoLines - 1)
      .X1 = XS1
      .Y1 = YS1
      .X2 = XS2 + 1
      .Y2 = YS2 + 1
      .Visible = True
   End With
End Sub

Public Sub Draw_Poly(frm As Form, X As Single, Y As Single)
   If X < 3 Then X = 3
   If Y < 3 Then Y = 3
   If X > picWidth - 3 Then X = picWidth - 3
   If Y > picHeight - 3 Then Y = picHeight - 3
   XS2 = X
   YS2 = Y
   With frm.SL(NumLassoLines - 1)
      .X2 = XS2
      .Y2 = YS2
      .Visible = True
   End With
End Sub


Public Sub AdjustRegularShapes(X As Single, Y As Single)
   XD = (X - XS1)
   YD = (Y - YS1)
   zDiagRad = Sqr(XD * XD + YD * YD)   ' Public
   Select Case SelectType
   Case 1, 5   ' Square & Rounded square
      ' Can use Sgn function of XD & YD here but this is clearer
      ' .7071 = Sin & Cos (45 deg)
      If Y >= YS1 Then
         If X > XS1 Then  ' BR quadrant
            Y = YS1 + zDiagRad * 0.7071
            X = XS1 + zDiagRad * 0.7071
         Else  ' X<=XS1   ' BL quadrant
            Y = YS1 + zDiagRad * 0.7071
            X = XS1 - zDiagRad * 0.7071
         End If
      Else ' Y < YS1
         If X > XS1 Then  ' TR quadrant
            Y = YS1 - zDiagRad * 0.7071
            X = XS1 + zDiagRad * 0.7071
         Else  ' X<=XS1   ' TL quadrant
            Y = YS1 - zDiagRad * 0.7071
            X = XS1 - zDiagRad * 0.7071
         End If
      End If
   Case 3   ' Circle
      ' Dealt with in MoveAll using zDiagRad
   End Select
End Sub

Public Sub StartMover(X As Single, Y As Single)
   XM = X
   YM = Y
   aSelectDone = False  ' so will redo selection
End Sub

Public Sub MoveMover(frm As Form, X As Single, Y As Single)
Dim k As Long
   XD = X - XM
   YD = Y - YM
   'Move all Shapes XD,YD
   Select Case SelectType
   Case 0 To 5 ' SR
      With frm.SR
         .Left = .Left + XD
         .Top = .Top + YD
      End With
   ' NumLassoLines
   Case 6, 7 ' SL()
      For k = 0 To NumLassoLines - 1
         With frm.SL(k)
            .X1 = .X1 + XD
            .Y1 = .Y1 + YD
            .X2 = .X2 + XD
            .Y2 = .Y2 + YD
         End With
      Next k
   End Select
   XM = X
   YM = Y
End Sub

Public Sub EndMover(frm As Form, X As Single, Y As Single)
   Select Case SelectType
   Case 0 To 5 ' SR
      SXS1 = frm.SR.Left
      SYS1 = frm.SR.Top
      If SelectType = 2 Then  ' Oval
         SXS2 = SXS1 + frm.SR.Width
         SYS2 = SYS1 + frm.SR.Height
      End If
   Case 6, 7
      ' No change here
   End Select
   aSelectDone = False
End Sub


Public Sub CheckLassoLines(frm As Form)
Dim i As Long
   If NumLassoLines > 1 Then ' Clear extra lasso lines SL(1)-SL(NumLassoLines-1)
      For i = 1 To NumLassoLines - 1
         Unload frm.SL(i)
      Next i
      NumLassoLines = 1
   End If
   frm.SL(0).Visible = False
   NumLeftClicks = 0
End Sub

Public Sub FindMaxMins(frm As Form)
Dim i As Long
' SXS1, SYS1    Mins
' SXS2, SYS2    Maxs
' SL(0) to SL(NumLassoLines - 1)
   SXS1 = 10000
   SYS1 = 10000
   SXS2 = -10000
   SYS2 = -10000
   For i = 0 To NumLassoLines - 1
      If frm.SL(i).X1 < SXS1 Then SXS1 = frm.SL(i).X1
      If frm.SL(i).Y1 < SYS1 Then SYS1 = frm.SL(i).Y1
      If frm.SL(i).X2 > SXS2 Then SXS2 = frm.SL(i).X2
      If frm.SL(i).Y2 > SYS2 Then SYS2 = frm.SL(i).Y2
   Next i
   SW = SXS2 - SXS1
   SH = SYS2 - SYS1
'   LabWH = Str$(SW + 1) & " ," & Str$(SH + 1) & " "
End Sub

Public Sub Fill(APIC As PictureBox, X As Single, Y As Single)
   ' Fill with FillColor = DrawColor at X,Y
   APIC.DrawStyle = vbSolid
   APIC.FillColor = vbWhite
   APIC.FillStyle = vbFSSolid
   
   ' FLOODFILLSURFACE = 1
   ' Fills with FillColor so long as point surrounded by
   ' color = APIC.Point(X, Y)
   
   ExtFloodFill APIC.hdc, X, Y, APIC.Point(X, Y), FLOODFILLSURFACE
   
   APIC.Refresh
End Sub

