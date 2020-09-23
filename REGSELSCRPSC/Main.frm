VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Region selectors2  by Robert Rayment"
   ClientHeight    =   5670
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9765
   DrawWidth       =   2
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   651
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInstructions 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      Height          =   1155
      Left            =   6570
      ScaleHeight     =   1095
      ScaleWidth      =   3000
      TabIndex        =   32
      Top             =   15
      Width           =   3060
   End
   Begin VB.CommandButton cmdSColor 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   4815
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   " Choose surround color "
      Top             =   -15
      Width           =   330
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   9
      Left            =   3240
      Picture         =   "Main.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   " Move selection "
      Top             =   15
      Width           =   270
   End
   Begin VB.CommandButton cmdExtract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   4530
      Picture         =   "Main.frx":0514
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   " Extract mask - black background "
      Top             =   15
      Width           =   285
   End
   Begin VB.CommandButton cmdExtract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   4215
      Picture         =   "Main.frx":05E6
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   " Extract mask - white background "
      Top             =   15
      Width           =   285
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H80000018&
      Caption         =   "?"
      Height          =   270
      Left            =   6285
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   30
      Width           =   225
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   5
      Left            =   2100
      Picture         =   "Main.frx":06B8
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   " Select rounded square"
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   3
      Left            =   1530
      Picture         =   "Main.frx":078A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   " Select circle "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   960
      Picture         =   "Main.frx":085C
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   " Select square "
      Top             =   15
      Width           =   270
   End
   Begin VB.PictureBox picSelect 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   9990
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   22
      Top             =   3195
      Width           =   645
   End
   Begin VB.CommandButton cmdUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   5220
      Picture         =   "Main.frx":092E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   " Undo "
      Top             =   15
      Width           =   330
   End
   Begin VB.CommandButton cmdExtract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   3915
      Picture         =   "Main.frx":0A78
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   " Extract selection "
      Top             =   15
      Width           =   285
   End
   Begin VB.CommandButton cmdExtract 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   3585
      Picture         =   "Main.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   " Show selection "
      Top             =   15
      Width           =   300
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   2670
      Picture         =   "Main.frx":158C
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   " Select polygon "
      Top             =   0
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   2
      Left            =   1245
      Picture         =   "Main.frx":165E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   " Select oval "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   4
      Left            =   1815
      Picture         =   "Main.frx":1730
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   " Select rounded rectangle "
      Top             =   15
      Width           =   270
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   10005
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   15
      Top             =   1020
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   300
      Picture         =   "Main.frx":1802
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   " Save resulting picture "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   8
      Left            =   2955
      Picture         =   "Main.frx":1D8C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   " Deselect "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   6
      Left            =   2385
      Picture         =   "Main.frx":1ED6
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " Select lasso "
      Top             =   15
      Width           =   270
   End
   Begin VB.OptionButton optSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   675
      Picture         =   "Main.frx":1FA8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   " Select rectangle "
      Top             =   15
      Width           =   270
   End
   Begin VB.CommandButton cmdLoadPic 
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   15
      Picture         =   "Main.frx":207A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Load picture "
      Top             =   15
      Width           =   270
   End
   Begin VB.HScrollBar HS 
      Height          =   210
      Index           =   1
      Left            =   5265
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4920
      Width           =   2490
   End
   Begin VB.HScrollBar HS 
      Height          =   210
      Index           =   0
      Left            =   525
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4980
      Width           =   2490
   End
   Begin VB.VScrollBar VS 
      Height          =   2130
      Index           =   1
      Left            =   5025
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   855
      Width           =   210
   End
   Begin VB.VScrollBar VS 
      Height          =   2130
      Index           =   0
      Left            =   330
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   870
      Width           =   195
   End
   Begin VB.PictureBox picC 
      AutoRedraw      =   -1  'True
      Height          =   4155
      Index           =   0
      Left            =   540
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   0
      Top             =   855
      Width           =   4320
      Begin VB.PictureBox PIC 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   795
         Index           =   0
         Left            =   15
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   2
         Top             =   15
         Width           =   930
         Begin VB.Shape SR 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Height          =   240
            Left            =   495
            Shape           =   1  'Square
            Top             =   135
            Width           =   240
         End
         Begin VB.Line SL 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            Index           =   0
            X1              =   21
            X2              =   10
            Y1              =   9
            Y2              =   16
         End
      End
   End
   Begin VB.PictureBox picC 
      AutoRedraw      =   -1  'True
      Height          =   4095
      Index           =   1
      Left            =   5310
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   1
      Top             =   855
      Width           =   3810
      Begin VB.PictureBox PIC 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   1
         Left            =   0
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   3
         Top             =   0
         Width           =   930
      End
   End
   Begin VB.Label LabWH 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   555
      TabIndex        =   26
      ToolTipText     =   " Selection size "
      Top             =   570
      Width           =   105
   End
   Begin VB.Label LabInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Picture to Save"
      Height          =   195
      Index           =   1
      Left            =   5340
      TabIndex        =   10
      Top             =   585
      Width           =   1095
   End
   Begin VB.Label LabInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   555
      TabIndex        =   9
      Top             =   315
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&FILE"
      Begin VB.Menu mnuFileOPS 
         Caption         =   "&Load picture"
         Index           =   0
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "&Save resulting picture"
         Index           =   1
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "L&oad from Clipboard"
         Index           =   3
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "&Copy to Clipboard"
         Index           =   4
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "&Get screen"
         Index           =   6
      End
      Begin VB.Menu mnuFileOPS 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu Brk0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "&SELECT"
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Rectangle"
         Index           =   0
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Square"
         Index           =   1
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Oval"
         Index           =   2
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Circle"
         Index           =   3
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "Ro&unded rectangle"
         Index           =   4
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "Rounded s&quare"
         Index           =   5
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Lasso"
         Index           =   6
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Polygon"
         Index           =   7
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Deselect"
         Checked         =   -1  'True
         Index           =   8
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "&Move selection"
         Index           =   9
      End
      Begin VB.Menu mnuSelectOPS 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuSColor 
         Caption         =   "D&efault surround color (black)"
         Index           =   0
      End
      Begin VB.Menu mnuSColor 
         Caption         =   "C&hoose surround color"
         Index           =   1
      End
   End
   Begin VB.Menu mnuExt 
      Caption         =   "&EXTRACT"
      Begin VB.Menu mnuExtract 
         Caption         =   "&Show selection"
         Index           =   0
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "E&xtract selection"
         Index           =   1
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "Show mask - &Black background"
         Index           =   2
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "Show mask - &White background"
         Index           =   3
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "&Undo"
         Index           =   4
      End
   End
   Begin VB.Menu mnuFileSpec 
      Caption         =   "FileSpec"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Eight Region Selectors by Robert Rayment  2/6/05

' 21/1/06
' Fixed minor PIC(1) scroll bars

' 15/1/06
' Added screen capture & selectable surround color of
' extracted shape.

' 5/6/05
' Correction  to extraction for 1 pixel error
' Pixel mask swapped over

' 4/6/05
' Update 4  Option to invert mask and move the selection.
'           Maintain fixed starting point for regular
'           selections ie Square, Circle & Rounded square
'           Circle drawn out from center.

' 3/5/05
' Update 3  Load/Copy from/to Clipboard

'  2/6/05
' Update 2  Added extract mask
' Update 1  Moved selector code to separate module
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' With INI recent files list, Drag-Drop onto input picbox
' or EXE or EXE shortcut.

' Shapes:-
'  SR   Rect,Square,Oval,Circle,Round Rect,Round Square
'  SL() Lasso & Polygon

' Pic boxes:-
'  PIC(0)    loaded picture
'  PIC(1)    resulting picture (to save)
'  picC(0/1) PIC() containers
'  picBack   mask hidden
'  picSelect extracted image hidden

' Scrollbars:_
'  HS(0/1), VS(0/1) also arrow key oprated

' Select op buttons:-
'  optSelect(0/8) -> mnuSelectOPS(0/8)

' Show resultant image:-
'  mnuExtract(0/1/2) -> cmdExtract(0/1/2/3)

Private i As Long
Private aClipBoard As Boolean
Private aClipBoardUsed As Boolean
Private aScreenShot As Boolean
Private FileLength0 As Long
Const ShortLen = 30
Dim CommonDialog1 As OSDialog

Dim CF As CFDialog
Dim SurroundColor As Long

Private Sub cmdSColor_Click()
   mnuSColor_Click 1
End Sub

Private Sub Form_Load()
Dim Infile$
Dim a$
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   ' Locate & Make application stay on top
   SetWindowPos Form1.hWnd, hWndInsertAfter, _
   200, 200, Form1.Width / STX, Form1.Height / STY, wFlags
   
   On Error GoTo NotPic
   
   picInstructions.Visible = False
   a$ = ""
   a$ = a$ & " All selectors are drawn with mouse down," & vbCrLf
   a$ = a$ & " apart from polygons where left click" & vbCrLf
   a$ = a$ & " starts a new line & right click completes" & vbCrLf
   a$ = a$ & " the shape.  Arrow keys can be used to" & vbCrLf
   a$ = a$ & " operate any scrollbars."
   picInstructions.Print a$
   
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   
   LoadPath$ = PathSpec$
   SavePath$ = PathSpec$
   IniTitle$ = "RegSel"
   ReDim RecentFilesList$(1)
   Get_Ini_Info
   
   Me.KeyPreview = True
   
   With picC(1)
      .Width = picC(0).Width
      .Height = picC(0).Height
      .Top = picC(0).Top
   End With
   With PIC(0)
      .Left = 0
      .Top = 0
   End With
   With PIC(1)
      .Left = 0
      .Top = 0
   End With
   HS(0).TabStop = False
   HS(1).TabStop = False
   VS(0).TabStop = False
   VS(1).TabStop = False
   FixScrollbars picC(0), PIC(0), HS(0), VS(0)
   FixScrollbars picC(1), PIC(1), HS(1), VS(1)
   
   cmdSave.Enabled = False
   mnuFileOPS(1).Enabled = False
   mnuSelect.Enabled = False
   mnuExt.Enabled = False
   For i = 0 To 9
      optSelect(i).Enabled = False
   Next i
   For i = 0 To 3
      cmdExtract(i).Enabled = False
   Next i
   cmdUndo.Enabled = False
   cmdSColor.Enabled = False
   aSelect = False
   SL(0).Visible = False
   SR.Visible = False
   NumLassoLines = 1 ' SL(0)
   NumLeftClicks = 0 ' SL(0) for paralines
   ExtractType = 0   ' PIC(1) size= PIC(0) size
   picBack.Visible = False
   picSelect.Visible = False
   aClipBoard = False
   aClipBoardUsed = False
   aMover = False
   aMouseDown = False
   SL(0).DrawMode = vbXorPen
   Show
   
   PIC(0).OLEDropMode = 1
   
   If Command$ <> "" Then   ' Loading pic on to exe
      If Left$(Command$, 1) = Chr(34) Then ' Strip off quotes
          Infile$ = Mid$(Command$, 2, Len(Command$) - 2)
      Else
          Infile$ = Command$
      End If
      If Not CheckLoadImage(Infile$) Then Exit Sub
   End If
   
   
   aScreenShot = False
   SurroundColor = 0
   cmdSColor.BackColor = SurroundColor
   CheckerpicC
   On Error GoTo 0
   Exit Sub
'=========
NotPic:
  MsgBox Err.Description & vbCrLf & vbCrLf & "Can't use this file", vbExclamation, "ERR# " & Err
End Sub

Private Sub mnuSColor_Click(Index As Integer)
Dim svSurroundColor As Long
   svSurroundColor = SurroundColor
   If Index = 0 Then
      SurroundColor = 0
   Else
      Set CF = New CFDialog
      If Not CF.VBChooseColor(SurroundColor, , , , Me.hWnd) Then
         SurroundColor = svSurroundColor
      End If
      Set CF = Nothing
   End If
   cmdSColor.BackColor = SurroundColor
End Sub

Private Sub PIC_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Infile$
   On Error GoTo NoPic
   Infile$ = Data.Files(1)
   
   If Not CheckLoadImage(Infile$) Then Exit Sub
   
   On Error GoTo 0
   Exit Sub
'=========
NoPic:
  MsgBox Err.Description & vbCrLf & vbCrLf & "Drag drop error", vbExclamation, "ERR# " & Err
End Sub

Private Function CheckLoadImage(FSpec$) As Boolean
Dim k As Long
Dim Ext$
   CheckLoadImage = False
   Ext$ = LCase$(FindExtension$(FSpec$))
   Select Case Ext$
   Case "bmp", "jpg", "gif"
   Case Else
      MsgBox "Not bmp, jpg or gif", vbCritical
      Exit Function
   End Select
   CheckLoadImage = True
   ' Check if InSpec$ already there
   For k = 0 To UBound(RecentFilesList$)
      If RecentFilesList$(k) = FSpec$ Then Exit For
   Next k
   If k > UBound(RecentFilesList$) Then ' File not listed
      LoadSpec$ = FSpec$
      FixMenuItems   ' Move items down
      RecentFilesList$(0) = LoadSpec$  ' Fill item 0
      mnuRecentFiles(0).Caption = ShortenFileSpec$(LoadSpec$, ShortLen)  ' Put most recent at top
   End If
   LoadSpec$ = FSpec$
   NewPicture
End Function

Private Sub mnuFileOPS_Click(Index As Integer)
   Select Case Index
   Case 0: cmdLoadPic_Click   ' Load picture
   Case 1: cmdSave_Click      ' Show PIC(1) as bmp
   Case 2   ' -
   Case 3   ' Load from Clipboard
      aClipBoard = True
      NewPicture
      aClipBoard = False
   Case 4   ' Copy to Clipboard
      Clipboard.Clear
      Clipboard.SetData PIC(1).Image, vbCFBitmap
      aClipBoardUsed = True
      DoEvents
   Case 5   ' -
   Case 6   ' Get screen
      ScreenShot
   Case 7   ' -
   End Select
End Sub

Private Sub ScreenShot()
Dim WinState As Long
Dim rDC As Long
    WinState = Form1.WindowState
    PIC(0).Width = Screen.Width \ STX
    PIC(0).Height = Screen.Height \ STY
    PIC(0).AutoRedraw = True
    Form1.WindowState = vbMinimized
    rDC = GetDC(0&)
    BitBlt PIC(0).hdc, 0, 0, Screen.Width \ STX, Screen.Height \ STY, rDC, 0, 0, vbSrcCopy
    PIC(0).Picture = PIC(0).Image
    
    Form1.WindowState = WinState
    PIC(0).Picture = PIC(0).Image
    
    ReleaseDC 0&, rDC
    
    aScreenShot = True
    NewPicture
End Sub

Private Sub cmdLoadPic_Click()
Dim k As Long
Dim Title$, Filt$, InDir$

   ' LOAD STANDARD VB PICTURES
   
   MousePointer = vbDefault
   
   If CommonDialog1 Is Nothing Then Set CommonDialog1 = New OSDialog
   
   Title$ = "Load a picture file"
   Filt$ = "Pics bmp,jpg,gif|*.bmp;*.jpg;*.gif"
   InDir$ = LoadPath$
   LoadSpec$ = ""
   
   CommonDialog1.ShowOpen LoadSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
   If Len(LoadSpec$) = 0 Then
      Close
      Set CommonDialog1 = Nothing
      Exit Sub
   End If
   Set CommonDialog1 = Nothing
   
   ' Check if LoadSpec$ already there
   For k = 0 To UBound(RecentFilesList$)
      If RecentFilesList$(k) = LoadSpec$ Then Exit For
   Next k
   If k > UBound(RecentFilesList$) Then ' File not listed
      FixMenuItems   ' Move items down
      RecentFilesList$(0) = LoadSpec$  ' Fill item 0
      mnuRecentFiles(0).Caption = ShortenFileSpec$(LoadSpec$, ShortLen)  ' Put most recent at top
   End If
   
   NewPicture
End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)
Dim k As Long, kk As Long
   If mnuRecentFiles(Index).Caption = "" Then Exit Sub
   
   If Not FileExists(RecentFilesList$(Index)) Then
      ' Re-write ini skipping Index
      On Error Resume Next
      Kill IniSpec$
      
      If mnuRecentFiles.Count = 1 Then
         ReDim RecentFilesList$(1)
         mnuRecentFiles(0).Caption = ""
         MsgBox " File not there "
         Exit Sub
      End If
      kk = 1
      For k = 0 To UBound(RecentFilesList$)
         LoadSpec$ = RecentFilesList$(k)
         If k <> Index Then
            If LoadPath$ <> "" Then
               WriteINI "RecentFiles", Str$(kk) & ".", LoadSpec$, IniSpec$
               kk = kk + 1
            End If
         End If
      Next k
      SavePath$ = FindPath$(SaveSpec$)
      If SavePath$ <> "" Then
         WriteINI "LastSavePath", "SavePath", SavePath$, IniSpec$
      End If
      ReDim RecentFilesList$(1)
      For k = 1 To mnuRecentFiles.Count - 1
         Unload mnuRecentFiles(k)
      Next k
      mnuRecentFiles(0).Caption = ""
      Get_Ini_Info
      
      MsgBox " File not there "
      Exit Sub
   Else
      LoadSpec$ = RecentFilesList$(Index)
      NewPicture
   End If
End Sub

Private Sub NewPicture()
   On Error GoTo NewPicError
            
   If aClipBoard Then
      PIC(0).Picture = Clipboard.GetData(vbCFBitmap)
      PIC(1).Picture = Clipboard.GetData(vbCFBitmap)
      FixScrollbars picC(0), PIC(0), HS(0), VS(0)
      FixScrollbars picC(1), PIC(1), HS(1), VS(1)
      
      picWidth = PIC(0).Width
      picHeight = PIC(0).Height
      
      LabInfo(0) = "Clipboard" & "  WxH =" & Str$(picWidth) & " x" & Str$(picHeight)
   
   ElseIf aScreenShot Then
      
      picWidth = PIC(0).Width
      picHeight = PIC(0).Height
      PIC(1).Width = picWidth
      PIC(1).Height = picHeight
      
      FixScrollbars picC(0), PIC(0), HS(0), VS(0)
      FixScrollbars picC(1), PIC(1), HS(1), VS(1)
      
      LoadSpec$ = "  Captured Screen"
      LabInfo(0) = "Size = screen size  WxH =" & Str$(picWidth) & " x" & Str$(picHeight)
      aScreenShot = False
   
   Else
      PIC(0).Picture = LoadPicture(LoadSpec$)  ' fails here if not valid file
      PIC(1).Picture = LoadPicture
      PIC(1).Picture = LoadPicture(LoadSpec$)
   
      LoadPath$ = FindPath$(LoadSpec$)
      FileLength0 = FileLen(LoadSpec$)
      
      FixScrollbars picC(0), PIC(0), HS(0), VS(0)
      FixScrollbars picC(1), PIC(1), HS(1), VS(1)
      
      picWidth = PIC(0).Width
      picHeight = PIC(0).Height
      
      LabInfo(0) = "Size =" & Str$(FileLength0) & "B  WxH =" & Str$(picWidth) & " x" & Str$(picHeight)
   End If
   
   cmdSave.Enabled = True
   
   mnuFileOPS(1).Enabled = True
   mnuSelect.Enabled = True
   mnuExt.Enabled = True
   For i = 0 To 9
      optSelect(i).Enabled = True
   Next i
   For i = 0 To 3
      cmdExtract(i).Enabled = True
   Next i
   cmdSColor.Enabled = True
   ' For masking
   picBack.Width = picWidth
   picBack.Height = picHeight
   
   mnuFileSpec.Caption = LoadSpec$
   
   ' Cancel any selection
   aSelect = False
   SR.Visible = False
   CheckLassoLines Form1
   PIC(0).MousePointer = vbDefault
   optSelect(5).Value = True
   optSelect(5).Value = False
   aMover = False
   aMouseDown = False
   DoEvents
   PIC(0).SetFocus
   On Error GoTo 0
   Exit Sub
'==========
NewPicError:
MsgBox Err.Description & vbCrLf & vbCrLf & "Could not load file", vbExclamation, "ERR# " & Err
On Error GoTo 0
End Sub

Private Sub cmdSave_Click()
Dim Title$, Filt$, InDir$
   
   If CommonDialog1 Is Nothing Then Set CommonDialog1 = New OSDialog
   
   Title$ = "Save BMP"
   Filt$ = "Pics bmp|*.bmp"
   InDir$ = SavePath$
   
   CommonDialog1.ShowSave SaveSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
   If Len(SaveSpec$) = 0 Then
      Close
      Set CommonDialog1 = Nothing
      Exit Sub
   End If
   
   FixExtension SaveSpec$, "bmp"
   SavePath$ = FindPath$(SaveSpec$)
   Set CommonDialog1 = Nothing
   
   SavePicture PIC(1).Image, SaveSpec$
   
'   If aSelect Or NumLassoLines > 1 Then
'   Else
'   End If
End Sub

Private Sub mnuExtract_Click(Index As Integer)
   If Index <> 4 Then  ' ie not Undo
      cmdExtract_Click Index
   Else    ' Undo
      cmdUndo_Click
   End If
End Sub

Private Sub cmdExtract_Click(Index As Integer)
' picBack, picWidth & picHeight set up in NewPicture
Dim IR As RECT
Dim ar As Long
   If aSelectDone Then
      cmdUndo.Enabled = True
      ExtractType = Index
      'Black pic(1) with colored select - Same Size as original
      PIC(1).Picture = LoadPicture
      PIC(1).Width = picWidth
      PIC(1).Height = picHeight
      BitBlt PIC(1).hdc, 0, 0, picWidth, picHeight, PIC(0).hdc, 0, 0, vbSrcCopy
      BitBlt PIC(1).hdc, 0, 0, picWidth, picHeight, picBack.hdc, 0, 0, vbSrcAnd
      ' Set surround color, filling from coords 0,0 on PIC(1) ie a black point 0&
      If SelectType > 1 Then   ' ie not rectangle or square
         PIC(1).DrawStyle = vbSolid
         PIC(1).FillColor = SurroundColor
         PIC(1).FillStyle = vbFSSolid
         'ExtFloodFill PIC(1).hdc, 0, 0, PIC(1).Point(0, 0), FLOODFILLSURFACE
         ExtFloodFill PIC(1).hdc, 0, 0, 0&, FLOODFILLSURFACE
      End If
   
      PIC(1).Refresh
      Select Case Index
      Case 1
         ' Ditto BUT reduced to rectangle (0,0)-(SW,SH) found
         ' when selected shape done - in Selector.bas
         picSelect.Width = SW
         picSelect.Height = SH
         picSelect.Picture = LoadPicture
         BitBlt picSelect.hdc, 0, 0, SW, SH, PIC(1).hdc, SXS1, SYS1, vbSrcCopy
         PIC(1).Picture = LoadPicture
         PIC(1).Width = SW
         PIC(1).Height = SH
         BitBlt PIC(1).hdc, 0, 0, SW, SH, picSelect.hdc, 0, 0, vbSrcCopy
         PIC(1).Refresh
      
      Case 2, 3  ' Show Mask - white or black background
         picSelect.Width = SW
         picSelect.Height = SH
         picSelect.Picture = LoadPicture
         BitBlt picSelect.hdc, 0, 0, SW, SH, picBack.hdc, SXS1, SYS1, vbSrcCopy
         PIC(1).Picture = LoadPicture
         PIC(1).Width = SW
         PIC(1).Height = SH
         BitBlt PIC(1).hdc, 0, 0, SW, SH, picSelect.hdc, 0, 0, vbSrcCopy
         PIC(1).Refresh
         ar = SetRect(IR, 0, 0, SW, SH)
         ar = InvertRect(PIC(1).hdc, IR)
         PIC(1).Refresh
         If Index = 3 Then   ' Show Mask - black background
            ar = SetRect(IR, 0, 0, SW, SH)
            ar = InvertRect(PIC(1).hdc, IR)
            PIC(1).Refresh
         End If
      End Select
      ' Reduce picbox memory
      picSelect.Picture = LoadPicture
      picSelect.Width = 4
      picSelect.Height = 4
      FixScrollbars picC(1), PIC(1), HS(1), VS(1)
      HS_Scroll (0)
      VS_Scroll (0)
   End If
   PIC(0).SetFocus
End Sub


Private Sub cmdUndo_Click()
Dim ii As Integer
   PIC(1).Width = picWidth
   PIC(1).Height = picHeight
   BitBlt PIC(1).hdc, 0, 0, picWidth, picHeight, PIC(0).hdc, 0, 0, vbSrcCopy
   PIC(1).Refresh
   CheckLassoLines Form1
   aSelectDone = False
   SR.Visible = False
   If aMover Then ' Mover
      ii = Int(SelectType)
      optSelect(ii).Value = True
      Call optSelect_MouseUp(ii, 0, 0, 0, 0)
      aMover = False
   End If
   FixScrollbars picC(1), PIC(1), HS(1), VS(1)
   LabWH = ""
   PIC(0).SetFocus
End Sub


'#### SELECTING #############################################################

Private Sub optSelect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
   Case 0 To 7   ' Select Rectangle, Rounded rectangle, Oval, Lasso or Poly
       mnuSelectOPS_Click Index
   Case 8   ' Deselect
      mnuSelectOPS_Click 8
   Case 9   ' Mover
      mnuSelectOPS_Click 9
   End Select
End Sub

Private Sub mnuSelectOPS_Click(Index As Integer)
Dim ii As Integer
   For i = 0 To 9
      mnuSelectOPS(i).Checked = False
   Next i
   If Index <> 9 Then   ' ie Mover
      SR.Visible = False
      CheckLassoLines Form1
   End If

' vbShapeRectangle         '0 (Default) Rectangle
' vbShapeSquare            '1 Square
' vbShapeOval              '2 Oval
' vbShapeCircle            '3 Circle
' vbShapeRoundedRectangle  '4 Rounded Rectangle
' vbShapeRoundedSquare     '5 Rounded Square
   
   Select Case Index
   Case 0 To 5
      aMover = False
      aSelect = True
      aSelectDone = False
      SelectType = Index
      PIC(0).MousePointer = 2    ' Cross
      SR.Shape = Index
   Case 6   ' Lasso
      aMover = False
      aSelect = True
      aSelectDone = False
      SelectType = 6
      PIC(0).MousePointer = 10   ' Up arrow
   Case 7   ' Poly
      aMover = False
      aSelect = True
      aSelectDone = False
      SelectType = 7
      PIC(0).MousePointer = 10   ' Up arrow
   Case 8   ' Deselect
      aMover = False
      cmdUndo_Click
      aSelect = False
      aSelectDone = False
      PIC(0).MousePointer = vbDefault
   Case 9   ' Move SelectType
      If aSelect And aSelectDone Then
         aMover = True
         PIC(0).MousePointer = 5    ' Mover Size cross
      Else  ' Cancel Mover pick up last SelectType
         ii = Int(SelectType)
         optSelect(ii).Value = True
         Call optSelect_MouseUp(ii, 0, 0, 0, 0)
         aMover = False
      End If
   End Select
   
   mnuSelectOPS(Index).Checked = True
   PIC(1).SetFocus
End Sub

Private Sub PIC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   aMouseDown = True
   If aMover Then   ' aSelect And aSelectDone = True
      StartMover X, Y
   Else
      StartAll Form1, picBack, Button, X, Y
   End If
End Sub

Private Sub PIC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aMover And aMouseDown Then
      MoveMover Form1, X, Y
   Else
      AdjustRegularShapes X, Y
      MoveAll Form1, Button, X, Y
   End If
End Sub

Private Sub PIC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aMover And aMouseDown Then
      EndMover Form1, X, Y
   End If
   MouseUpAll Form1, picBack, Button, X, Y
   aMouseDown = False
End Sub
'#### END SELECTING #############################################################



'#### mnuRecentFiles(), ini & RecentFilesList$() #################

Private Sub Get_Ini_Info()
' Public RecentFilesList$()
Dim aBool As Boolean
Dim k As Long, kup As Long
Dim Ret$
   IniSpec$ = PathSpec$ & IniTitle$ & ".ini"
   k = 0
   aBool = False
   Do
      If GetINI("RecentFiles", Str$(k + 1) & ".", Ret$, IniSpec$) Then
                '    [      ],      key,          info, ini file full spec
         If k = 0 Then aBool = True ' ie at least one entry
         RecentFilesList$(k) = Ret$
         k = k + 1
         ReDim Preserve RecentFilesList$(k)
      Else
         Exit Do
      End If
   Loop
   If aBool Then
      kup = k - 1
      ReDim Preserve RecentFilesList$(kup)
      mnuRecentFiles(0).Caption = ShortenFileSpec$(RecentFilesList$(0), ShortLen)    ' Put most recent at top
      If k > 0 Then
         For k = 1 To kup
            Load mnuRecentFiles(k)
            mnuRecentFiles(k).Caption = ShortenFileSpec$(RecentFilesList$(k), ShortLen)
         Next k
      End If
   End If
   If GetINI("LastSavePath", "SavePath", Ret$, IniSpec$) Then
      If Len(Ret$) > 1 Then SavePath$ = Ret$
   End If
End Sub

Private Sub FixMenuItems()
Dim k As Long
   If mnuRecentFiles.Count < MaxRecentFiles Then  ' extend
      k = mnuRecentFiles.UBound + 1
      Load mnuRecentFiles(k)
      ReDim Preserve RecentFilesList$(k)
   End If
   ' Move items down so most recent is in top (ie index 0)
   For k = mnuRecentFiles.UBound To 1 Step -1
      mnuRecentFiles(k).Caption = mnuRecentFiles(k - 1).Caption
      RecentFilesList$(k) = RecentFilesList$(k - 1)
   Next k
End Sub

'################################################################

Private Sub CheckerpicC()    ' Checker picC()
Dim k As Long, j As Long
   For k = 0 To 1
      picC(k).BackColor = vbWhite
      For j = 0 To picC(k).Height Step 32
      For i = 0 To picC(k).Width Step 32
         picC(k).Line (i, j)-(i + 16, j + 16), &HD0E0D0, BF
         picC(k).Line (i + 16, j + 16)-(i + 32, j + 32), &HD0E0D0, BF
      Next i
      Next j
      picC(k).Refresh
   Next k
End Sub

'#### PIC(0/1) SCROLL BARS #########################################

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   For i = 0 To 1
      If HS(i).Visible Then
         Select Case KeyCode
         Case vbKeyLeft
            If HS(i).Value > HS(i).Min Then
               HS(i).Value = HS(i).Value - 1
            End If
         Case vbKeyRight
            If HS(i).Value < HS(i).Max Then
               HS(i).Value = HS(i).Value + 1
            End If
         End Select
      End If
      
      If VS(i).Visible Then
         VS(i).TabStop = False
         Select Case KeyCode
         Case vbKeyUp
            If VS(i).Value > VS(i).Min Then
               VS(i).Value = VS(i).Value - 1
            End If
         Case vbKeyDown
            If VS(i).Value < VS(i).Max Then
               VS(i).Value = VS(i).Value + 1
            End If
         End Select
      End If
   Next i
End Sub

Private Sub HS_Change(Index As Integer)
   Call HS_Scroll(Index)
End Sub

Private Sub HS_Scroll(Index As Integer)
   If ExtractType = 0 Then      ' PIC(1) size= PIC(0) size
      PIC(Index).Left = -HS(Index).Value
      If Index = 0 Then
         PIC(1).Left = -HS(0).Value
         HS(1).Value = HS(0).Value
      Else
         PIC(0).Left = -HS(1).Value
         HS(0).Value = HS(1).Value
      End If
   Else
      PIC(Index).Left = -HS(Index).Value
   End If
End Sub

Private Sub VS_Change(Index As Integer)
   Call VS_Scroll(Index)
End Sub

Private Sub VS_Scroll(Index As Integer)
   If ExtractType = 0 Then      ' PIC(1) size= PIC(0) size
      PIC(Index).Top = -VS(Index).Value
      If Index = 0 Then
         PIC(1).Top = -VS(0).Value
         VS(1).Value = VS(0).Value
      Else
         PIC(0).Top = -VS(1).Value
         VS(0).Value = VS(1).Value
      End If
   Else
      PIC(Index).Top = -VS(Index).Value
   End If
End Sub
'#### END PIC(0/1) SCROLL BARS #########################################

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   picInstructions.Visible = True
End Sub
Private Sub cmdHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   picInstructions.Visible = False
   PIC(0).SetFocus
End Sub

Private Sub Form_Resize()
'Resize pic containers
Dim FW As Long
Dim PICCW As Long, PICCH As Long

On Error Resume Next
   
'   ' Cancel any selection
'   aSelect = False
'   SR.Visible = False
'   For i = 0 To NumLassoLines - 1
'      'Unload NOT PERMITTED IN RESIZE EVENT
'      SL(i).Visible = False
'   Next i

   FW = Me.Width \ STX
   PICCW = (FW - 150) \ 2
   PICCH = Me.Height \ STY - 150
   
   picC(0).Left = 50
   picC(0).Width = PICCW
   picC(0).Height = PICCH
   picC(1).Left = picC(0).Left + picC(0).Width + 50
   picC(1).Width = PICCW
   picC(1).Height = PICCH
   LabInfo(0).Left = picC(0).Left
   LabInfo(1).Left = picC(1).Left
   LabWH.Left = picC(0).Left
  
   FixScrollbars picC(0), PIC(0), HS(0), VS(0)
   FixScrollbars picC(1), PIC(1), HS(1), VS(1)
   
'   PIC(0).MousePointer = vbDefault
'   optSelect(5).Value = True
'   optSelect(5).Value = False
   CheckerpicC
End Sub


Private Sub mnuExit_Click()
   Form_QueryUnload 1, 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Form As Form
Dim k As Long
   If UnloadMode = 0 Then    'Close on Form1 pressed
'      k = MsgBox("", vbQuestion + vbYesNo + vbSystemModal, "Save result ?")
'      If k = vbYes Then
'         Cancel = True
'      Else
         Cancel = False
         Screen.MousePointer = vbDefault
         
         If aClipBoardUsed Then
            k = MsgBox("Clipboard used: Clear", vbYesNo + vbQuestion, "Exitting")
            If k = vbYes Then
               Clipboard.Clear
            End If
         End If
         
         If mnuRecentFiles(0).Caption <> "" Then  ' Check if any loaded files
            On Error Resume Next
            Kill IniSpec$
            For k = 0 To UBound(RecentFilesList$)
               LoadSpec$ = RecentFilesList$(k)
               If LoadPath$ <> "" Then
                  WriteINI "RecentFiles", Str$(k + 1) & ".", LoadSpec$, IniSpec$
               End If
            Next k
            SavePath$ = FindPath$(SaveSpec$)
            If SavePath$ <> "" Then
               WriteINI "LastSavePath", "SavePath", SavePath$, IniSpec$
            End If
         End If
         
         ' Make sure all forms cleared
         For Each Form In Forms
            Unload Form
            Set Form = Nothing
         Next Form
         End
      End If
'   End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   End
End Sub



