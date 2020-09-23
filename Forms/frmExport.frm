VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Website"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMkDir 
      Caption         =   "<-- &New Folder"
      Height          =   390
      Left            =   3000
      TabIndex        =   5
      Top             =   3525
      Width           =   1740
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3000
      TabIndex        =   3
      Top             =   750
      Width           =   1740
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Drive and Folder"
      Height          =   3840
      Left            =   150
      TabIndex        =   4
      Top             =   75
      Width           =   2640
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   2340
      End
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   150
         TabIndex        =   1
         Top             =   675
         Width           =   2340
      End
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "&Begin"
      Default         =   -1  'True
      Height          =   390
      Left            =   3000
      TabIndex        =   2
      Top             =   225
      Width           =   1740
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBegin_Click()
On Error GoTo ErrHandler
    Dim tmpFolder As String
    Dim tmpFilename As String
    
    Dim htmlHeaderStart As String
    Dim htmlMetaTags As String
    Dim htmlScript As String
    Dim htmlPageTitle As String
    Dim htmlHeaderEnd As String
    Dim htmlBodyTag As String
    Dim htmlBodyHeader As String
    Dim htmlBody As String
    Dim htmlFooter As String
    Dim HTML As String
    Dim db As Database
    Dim rsPages As Recordset
    Dim rsProjects As Recordset
    
    tmpFolder = Dir1.Path & "\"
    
    Set db = Workspaces(0).OpenDatabase(frmMain.datRecords.DatabaseName)
    Set rsProjects = db.OpenRecordset("Template", dbOpenSnapshot)
    
    htmlHeaderStart = "<HTML><HEAD>" & vbCrLf
    htmlMetaTags = rsProjects.Fields("htmlMetaTags").Value & vbCrLf
    htmlScript = rsProjects.Fields("htmlScript").Value & vbCrLf
    htmlHeaderEnd = "</HEAD>" & vbCrLf
    htmlBodyTag = "<BODY"
    htmlBodyTag = htmlBodyTag & " bgcolor=" & Chr(34) & rsProjects.Fields("htmlBGColor").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & " text=" & Chr(34) & rsProjects.Fields("htmlText").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & " link=" & Chr(34) & rsProjects.Fields("htmlLink").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & " alink=" & Chr(34) & rsProjects.Fields("htmlALink").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & " vlink=" & Chr(34) & rsProjects.Fields("htmlVLink").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & " leftmargin=" & Chr(34) & rsProjects.Fields("htmlLeftMargin").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & " marginwidth=" & Chr(34) & rsProjects.Fields("htmlLeftMargin").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & " topmargin=" & Chr(34) & rsProjects.Fields("htmlTopMargin").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & " marginheight=" & Chr(34) & rsProjects.Fields("htmlTopMargin").Value & Chr(34)
    htmlBodyTag = htmlBodyTag & ">" & vbCrLf
    htmlBodyHeader = rsProjects.Fields("htmlBodyHeader").Value & vbCrLf
    htmlFooter = rsProjects.Fields("htmlFooter").Value & vbCrLf & "</BODY></HTML>"
    rsProjects.Close

    Set rsPages = db.OpenRecordset("Pages", dbOpenSnapshot)
    rsPages.MoveLast
    rsPages.MoveFirst
    
    Do Until rsPages.EOF
        htmlPageTitle = "<TITLE>" & rsPages.Fields("PageTitle").Value & "</TITLE>"
        htmlBody = "" & rsPages.Fields("PageHTML").Value & vbCrLf
        tmpFilename = tmpFolder & rsPages.Fields("PageFilename").Value
        HTML = htmlHeaderStart & htmlPageTitle & htmlMetaTags & htmlScript & htmlHeaderEnd & htmlBodyTag & htmlBodyHeader & htmlBody & htmlFooter
        Open tmpFilename For Output As 1
        Print #1, HTML
        Close #1
        rsPages.MoveNext
    Loop
    
    rsPages.Close
    db.Close
    MsgBox "Website was built in folder" & vbCrLf & tmpFolder
    Unload Me
    
    Exit Sub

ErrHandler:
    MsgBox Err.Number & vbCrLf & Err.Description
    If Err.Number = 76 Then
        MkDir (App.Path & "\HTML")
        Resume Next
    Else
        db.Close
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMkDir_Click()
On Error GoTo ErrHandler
    Dim Response As String
    Response = InputBox("Enter name of new folder", "New Folder", , 0, 0)
    MkDir Dir1.Path & "\" & Response
    Dir1.Refresh
    Exit Sub
ErrHandler:
    MsgBox "Error Number " & Err.Number & " occured while creating a new folder" & vbCrLf & Err.Description
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Me.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Dir1.Path = App.Path & "\HTML"
    'Check for records in Pages table
    Dim db As Database
    Dim rsPages As Recordset
    Set db = Workspaces(0).OpenDatabase(frmMain.datRecords.DatabaseName)
    Set rsPages = db.OpenRecordset("Pages", dbOpenSnapshot)
    If rsPages.BOF And rsPages.EOF Then
        MsgBox "There isn't any data to work with!"
        rsPages.Close
        db.Close
        cmdBegin.Enabled = False
    Else
        rsPages.Close
        db.Close
    End If
    
    Exit Sub
ErrHandler:
    If Err.Number = 76 Then 'If HTML folder does not exist
        MkDir App.Path & "\HTML"
        Resume
        Exit Sub
    Else
        MsgBox "Error Number " & Err.Number & " has occurred." & vbCrLf & Err.Description
    End If
End Sub
