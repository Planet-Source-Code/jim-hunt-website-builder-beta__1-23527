Attribute VB_Name = "modHTMLTagColors"
Option Explicit

'modGeneral (modGeneral.mod)
'----------------------------
'
'Purpose: Used to hold the general variables,
'subs and functions related to tag coloring.
'


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const EM_CHARFROMPOS& = &HD7
Public Const EM_GETFIRSTVISIBLELINE = &HCE '(0&,pt)
Public Const EM_FMTLINES = &HC8
Public Const EM_GETLINE = &HC4 '(line num,pt)=len
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB '(char num,0&)
Public Const EM_LINEFROMCHAR = &HC9 '(char num,0&)


'Public subs.

Public Function GetFirstLinePos(Line As Integer, Start As Integer, rtf As RichTextBox) As Integer
On Error Resume Next
Dim i As Integer, c As Integer

For i = Start To 0 Step -1
    c = SendMessage(rtf.hWnd, EM_LINEFROMCHAR, _
    i, 0&)
    
    If c < Line Then
        GetFirstLinePos = i + 1
    
        Exit For
    ElseIf i = 0 Then
        GetFirstLinePos = 0
    End If
Next i

End Function

Public Function GetLastLinePos(Line As Integer, Start As Integer, rtf As RichTextBox) As Integer
On Error Resume Next
Dim i As Integer, c As Integer

For i = Start To Len(rtf.Text)
    c = SendMessage(rtf.hWnd, EM_LINEFROMCHAR, _
    i, 0&)
    
    If c > Line Then
        GetLastLinePos = i - 1
    
        Exit For
    ElseIf i = Len(rtf.Text) Then
        GetLastLinePos = Len(rtf.Text)
    End If
Next i

End Function

Public Sub ColorTags(iStart As Integer, iEnd As Integer, rtf As RichTextBox, Optional Color As Long = vbBlue, Optional ErrorColor As Long = vbRed, Optional CommentColor As Long = &H8000&, Optional ParamColor As Long = &H800080, Optional IncludeColor As Long = &H40C0&)
On Error Resume Next
Dim iFirst As Integer
Dim iLast As Integer
Dim tmp As Long
Dim i As Integer, c As Integer, t As Integer
Dim OldStart As Integer

'Turn refreshing off.
tmp = LockWindowUpdate(rtf.hWnd)

OldStart = iStart

rtf.SelStart = iStart
rtf.SelLength = iEnd - iStart
rtf.SelColor = vbBlack
rtf.SelLength = 0

iStart = InStr(iStart + 1, rtf.Text, "<")

If iStart > 0 Then _
    rtf.SelStart = iStart - 1

iFirst = iStart - 1
'iFirst = InStr(iFirst + 1, rtf.Text, "<") - 1

c = OldStart + 1
i = InStr(OldStart + 1, rtf.Text, "<")

Do
    c = InStr(c + 1, rtf.Text, ">")
    
    If c < i Then
        rtf.SelStart = c - 1
        rtf.SelLength = 1
        
        rtf.SelColor = ErrorColor
        
        If iStart > 0 Then
            rtf.SelStart = iStart - 1
        Else
            rtf.SelStart = 0
        End If
        rtf.SelLength = 0
        
    Else
        Exit Do
    End If
Loop


iLast = InStr(iFirst + 1, rtf.Text, ">")

i = iLast
c = iLast
Do
    'A ">" without "<".
    i = InStr(i + 1, rtf.Text, "<")
    c = InStr(c + 1, rtf.Text, ">")

    If (c < i And c > 0) And (c < iEnd) Or _
    (c > 0 And i = 0) And (c < iEnd) Then
        rtf.SelStart = c - 1
        rtf.SelLength = 1
        rtf.SelColor = ErrorColor
        rtf.SelLength = 0

        rtf.SelStart = iFirst
    End If

    If i = 0 Then
        i = iLast
    End If
Loop Until (c = 0)


i = 0
c = 0

Do Until iFirst = -1
    iLast = InStr(iFirst + 1, rtf.Text, ">")
    rtf.SelStart = iFirst
        
    'A "<" without ">"
    tmp = InStr(iFirst + 2, rtf.Text, "<")
    
    If tmp < iLast And tmp > 0 Or iLast = 0 Then
        If tmp = 0 Then
            rtf.SelLength = Len(rtf.Text)
        Else
            rtf.SelLength = tmp - iFirst - 1
        End If
        
        rtf.SelColor = ErrorColor
    Else
        rtf.SelLength = iLast - iFirst
        
        If Mid(rtf.Text, iFirst + 1, 4) = "<!--" And _
        Mid(rtf.Text, iLast - 2, 3) = "-->" Then
        
            tmp = InStr(iFirst, rtf.Text, "#include", vbTextCompare)
            
            If tmp > 0 Then
                If Trim(Mid(rtf.Text, iFirst + 5, tmp - iFirst - 5)) = "" Then
                    rtf.SelColor = IncludeColor
                Else
                    rtf.SelColor = CommentColor
                End If
            Else
                rtf.SelColor = CommentColor
            End If
        Else
            rtf.SelColor = Color
        End If
        
        'Color the parameters.
        t = iFirst
        
        Do
            tmp = InStr(t + 1, rtf.Text, "=")
            
            If tmp > 0 And tmp < iLast Then
                For i = tmp + 1 To iLast
                    If Mid(rtf.Text, i, 1) <> " " And _
                    Mid(rtf.Text, i, 1) <> vbCr And _
                    Mid(rtf.Text, i, 1) <> vbLf Then
                        Exit For
                    End If
                Next i
                
                If i >= iLast Then
                    'A '=' without a parameter.
                    rtf.SelStart = tmp - 1
                    rtf.SelLength = 1
                    rtf.SelColor = ErrorColor
                    Exit Do
                End If
                
                For c = i + 1 To iLast
                    If Mid(rtf.Text, c, 1) = """" And _
                    Mid(rtf.Text, i, 1) = """" Then
                        Exit For
                    ElseIf Mid(rtf.Text, c, 1) = " " And _
                    Mid(rtf.Text, i, 1) <> """" Or _
                    Mid(rtf.Text, c, 1) = vbCr And _
                    Mid(rtf.Text, i, 1) <> """" Or _
                    Mid(rtf.Text, c, 1) = vbLf And _
                    Mid(rtf.Text, i, 1) <> """" Then
                        Exit For
                    End If
                Next c
                
                If c >= iLast And _
                Mid(rtf.Text, i, 1) = """" Then
                    'A parameter starting with
                    ''"' and doesn't end with one.
                    
                    rtf.SelStart = i - 1
                    rtf.SelLength = iLast - i
                    rtf.SelColor = ErrorColor
                    
                    Exit Do
                End If
                
                'Color the parameter.
                rtf.SelStart = i - 1
                rtf.SelLength = c - i + 1
                
                If rtf.SelColor = CommentColor Then
                    Exit Do
                End If
                
                rtf.SelColor = ParamColor
                
                t = tmp + 1
            Else
                Exit Do
            End If
        Loop
    End If
    
    iFirst = rtf.Find("<", iFirst + 1, , rtfNoHighlight)
    
    If iFirst > iEnd Then Exit Do
Loop
rtf.SelStart = iStart


'Allow repainting (Refreshing).
tmp = LockWindowUpdate(0)

rtf.SelStart = OldStart
End Sub


Public Function InStrRev(Start As Integer, SearchIn As String, SearchFor As String, Optional MatchCase As Boolean = False) As Integer
On Error Resume Next
Dim i As Integer

For i = Start To 1 Step -1
    If (UCase(Mid(SearchIn, i, _
    Len(SearchFor))) = UCase(SearchFor)) And _
    MatchCase = False Then
        'Found match.
        InStrRev = i
        
        Exit For
    ElseIf (Mid(SearchIn, i, _
    Len(SearchFor)) = SearchFor) And _
    MatchCase = True Then
        'Found match.
        InStrRev = i
        
        Exit For
    End If
Next i

End Function


