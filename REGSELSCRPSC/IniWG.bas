Attribute VB_Name = "INIWG"
' INIWG.bas

Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
 ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
 ByVal lpDefault As String, ByVal lpReturnedString As String, _
 ByVal nSize As Long, ByVal lpFileName As String) As Long
 ' lpDefault is the string return if no ini file found.
'------------------------------------------------------

'Dim PathSpec$, CurrentPath$, FileSpec$, FileSpecPath$

' File  stuff
Public PathSpec$, LoadSpec$, SaveSpec$
Public LoadPath$, SavePath$
Public IniTitle$, IniSpec$
Public Const MaxRecentFiles = 4 'Max of files to show in list
Public RecentFilesList$()
 
Public Function WriteINI(Title$, TheKey$, Info$, FileSpec$) As Boolean
   WritePrivateProfileString Title$, TheKey$, Info$, FileSpec$
End Function

Public Function GetINI(Title$, TheKey$, Ret$, FileSpec$) As Boolean
Dim N As Long
   On Error GoTo NoINI
   Ret$ = String(255, 0)
   N = GetPrivateProfileString(Title$, TheKey$, "", Ret$, 255, FileSpec$)
   'N is the number of characters copied to Ret$
   If N <> 0 Then
     GetINI = True
     Ret$ = Left$(Ret$, N)
   Else
     GetINI = False
     Ret$ = ""
   End If
   On Error GoTo 0
   Exit Function
'==========
NoINI:
GetINI = False
Ret$ = ""
End Function

Public Function FileExists(FSpec) As Boolean
  On Error Resume Next
  FileExists = FileLen(FSpec)
End Function

Public Sub FixExtension(FSpec$, Ext$)
' Enter Ext$ as jpg, bmp etc  no dot
Dim P As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   P = FindLastCharPos(FSpec$, ".")
   
   If P = 0 Then
      FSpec$ = FSpec$ & "." & Ext$
   Else
      If LCase$(Mid$(FSpec$, P + 1)) <> Ext$ Then FSpec$ = Mid$(FSpec$, 1, P) & Ext$
   End If

End Sub

Public Function FindPath$(FSpec$)
Dim P As Long
   FindPath$ = ""
   If Len(FSpec$) = 0 Then Exit Function
   
   P = FindLastCharPos(FSpec$, "\")
   If P = 0 Then Exit Function
   FindPath$ = Left$(FSpec$, P)
End Function

Public Function FindName$(FSpec$)
Dim P As Long
   FindName$ = ""
   If Len(FSpec$) = 0 Then Exit Function
   P = FindLastCharPos(FSpec$, "\")
   If P = 0 Then Exit Function
   FindName$ = Mid$(FSpec$, P + 1)
End Function

Public Function FindExtension$(FSpec$)
Dim P As Long
   P = FindLastCharPos(FSpec$, ".")
   If P = 0 Then
      FindExtension$ = ""
   Else
      FindExtension$ = Mid$(FSpec$, P + 1)
   End If
End Function

Public Function ShortenFileSpec$(FSpec$, L As Long)
' API for this ?
Dim P$, N$
Dim LPath As Long, LName As Long, LLeft As Long
Dim NDots As Long
   If Len(FSpec$) <= L Then
      ShortenFileSpec$ = FSpec$
      Exit Function
   End If
   P$ = FindPath$(FSpec$)
   N$ = FindName$(FSpec$)
   LPath = Len(P$)
   LName = Len(N$)
   LLeft = L - LName
   If LLeft = 0 Then       ' LName = L
      ShortenFileSpec$ = N$
   ElseIf LLeft < 0 Then   ' Very long name
      ShortenFileSpec$ = Left$(N$, L \ 2 - 2) & ".." & Right$(N$, L \ 2)
   Else
      NDots = LLeft \ 2
      ShortenFileSpec$ = Left$(P$, NDots) & String$(NDots, ".") & N$
   End If
End Function

Public Function FindLastCharPos(InString$, SerChar$) As Long
' Also VB5
Dim P As Long
    For P = Len(InString$) To 1 Step -1
      If Mid$(InString$, P, 1) = SerChar$ Then Exit For
    Next P
    If P < 1 Then P = 0
    FindLastCharPos = P
End Function



