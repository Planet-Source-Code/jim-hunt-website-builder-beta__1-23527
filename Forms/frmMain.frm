VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Website Builder Beta"
   ClientHeight    =   5910
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5265
      Left            =   150
      TabIndex        =   4
      Top             =   525
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   9287
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Style"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Meta"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMetaTags"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Script"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtScript"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Body Start"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtBodyHeader"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Body End"
      TabPicture(4)   =   "frmMain.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtFooter"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Notes"
      TabPicture(5)   =   "frmMain.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtDescription"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Margin"
         Height          =   2565
         Left            =   4650
         TabIndex        =   23
         Top             =   225
         Width           =   3915
         Begin VB.TextBox txtBody 
            DataField       =   "htmlLeftMargin"
            DataSource      =   "datRecords"
            Height          =   315
            Index           =   5
            Left            =   1725
            TabIndex        =   27
            Top             =   300
            Width           =   1140
         End
         Begin VB.TextBox txtBody 
            DataField       =   "htmlTopMargin"
            DataSource      =   "datRecords"
            Height          =   315
            Index           =   6
            Left            =   1725
            TabIndex        =   26
            Top             =   750
            Width           =   1140
         End
         Begin VB.VScrollBar scrlLeftMargin 
            Height          =   315
            Left            =   2850
            Max             =   0
            Min             =   -50
            TabIndex        =   25
            Top             =   300
            Width           =   165
         End
         Begin VB.VScrollBar scrlTopMargin 
            Height          =   315
            Left            =   2850
            Max             =   0
            Min             =   -50
            TabIndex        =   24
            Top             =   750
            Width           =   165
         End
         Begin VB.Label Label7 
            Caption         =   "topmargin"
            Height          =   240
            Left            =   225
            TabIndex        =   29
            Top             =   825
            Width           =   1440
         End
         Begin VB.Label Label6 
            Caption         =   "leftmargin"
            Height          =   240
            Left            =   225
            TabIndex        =   28
            Top             =   375
            Width           =   1440
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Colors"
         Height          =   2565
         Left            =   225
         TabIndex        =   12
         Top             =   225
         Width           =   4065
         Begin VB.TextBox txtBody 
            DataField       =   "htmlBGColor"
            DataSource      =   "datRecords"
            Height          =   315
            Index           =   0
            Left            =   1725
            TabIndex        =   17
            Top             =   300
            Width           =   1440
         End
         Begin VB.TextBox txtBody 
            DataField       =   "htmlAlink"
            DataSource      =   "datRecords"
            Height          =   315
            Index           =   3
            Left            =   1725
            TabIndex        =   16
            Top             =   1650
            Width           =   1440
         End
         Begin VB.TextBox txtBody 
            DataField       =   "htmlText"
            DataSource      =   "datRecords"
            Height          =   315
            Index           =   1
            Left            =   1725
            TabIndex        =   15
            Top             =   750
            Width           =   1440
         End
         Begin VB.TextBox txtBody 
            DataField       =   "htmlVlink"
            DataSource      =   "datRecords"
            Height          =   315
            Index           =   4
            Left            =   1725
            TabIndex        =   14
            Top             =   2100
            Width           =   1440
         End
         Begin VB.TextBox txtBody 
            DataField       =   "htmlLink"
            DataSource      =   "datRecords"
            Height          =   315
            Index           =   2
            Left            =   1725
            TabIndex        =   13
            Top             =   1200
            Width           =   1440
         End
         Begin VB.Label Label5 
            Caption         =   "vlink"
            Height          =   240
            Left            =   225
            TabIndex        =   22
            Top             =   2175
            Width           =   1440
         End
         Begin VB.Label Label4 
            Caption         =   "alink"
            Height          =   240
            Left            =   225
            TabIndex        =   21
            Top             =   1725
            Width           =   1440
         End
         Begin VB.Label Label3 
            Caption         =   "link"
            Height          =   240
            Left            =   225
            TabIndex        =   20
            Top             =   1275
            Width           =   1440
         End
         Begin VB.Label Label2 
            Caption         =   "text"
            Height          =   240
            Left            =   225
            TabIndex        =   19
            Top             =   825
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "bgcolor"
            Height          =   240
            Left            =   225
            TabIndex        =   18
            Top             =   375
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Body Tag Preview"
         Height          =   1665
         Left            =   225
         TabIndex        =   10
         Top             =   3000
         Width           =   8340
         Begin VB.Label lblTagPreview 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   1215
            Left            =   150
            TabIndex        =   11
            Top             =   300
            Width           =   8040
         End
      End
      Begin VB.TextBox txtDescription 
         DataField       =   "Notes"
         DataSource      =   "datRecords"
         Height          =   4665
         Left            =   -74850
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   150
         Width           =   8640
      End
      Begin VB.TextBox txtFooter 
         DataField       =   "htmlFooter"
         DataSource      =   "datRecords"
         Height          =   4665
         Left            =   -74850
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   150
         Width           =   8640
      End
      Begin VB.TextBox txtBodyHeader 
         DataField       =   "htmlBodyHeader"
         DataSource      =   "datRecords"
         Height          =   4665
         Left            =   -74850
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   150
         Width           =   8640
      End
      Begin VB.TextBox txtScript 
         DataField       =   "htmlScript"
         DataSource      =   "datRecords"
         Height          =   4665
         Left            =   -74850
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   150
         Width           =   8640
      End
      Begin VB.TextBox txtMetaTags 
         DataField       =   "htmlMetaTags"
         DataSource      =   "datRecords"
         Height          =   4665
         Left            =   -74850
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   150
         Width           =   8640
      End
   End
   Begin VB.FileListBox lstProjects 
      Height          =   870
      Left            =   75
      TabIndex        =   3
      Top             =   4950
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Frame Frame4 
      Height          =   30
      Left            =   -30
      TabIndex        =   2
      Top             =   360
      Width           =   9300
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   75
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":050E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":07C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0922
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EFE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Data datRecords 
      Caption         =   "datRecords"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\WINDOWS\Desktop\WebBuilder\Other Suite Projects\WebWizard\wwdata.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5175
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Projects"
      Top             =   5475
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.TextBox txtProjectID 
      DataField       =   "PrjID"
      DataSource      =   "datRecords"
      Height          =   315
      Left            =   6900
      TabIndex        =   0
      Text            =   "PrjID"
      Top             =   5475
      Visible         =   0   'False
      Width           =   2190
   End
   Begin MSComctlLib.Toolbar tlbProject 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   635
      ButtonWidth     =   1508
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "Create New Project"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Object.ToolTipText     =   "Open Existing Project"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Object.ToolTipText     =   "Save Current Project"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pages"
            Object.ToolTipText     =   "Manage Webpages For Current Project"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export"
            Object.ToolTipText     =   "Build HTML Files For Current Project"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Object.ToolTipText     =   "Show WebWizard Help"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Object.ToolTipText     =   "About WebWizard"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Exit WebWizard"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Project"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileWebpages 
         Caption         =   "Webpages"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer 'Used in For loops
Dim ProjectFilename As String

Private Sub datRecords_Reposition()
On Error Resume Next 'If margin is greater than 50, continue anyway
    scrlLeftMargin.Value = txtBody(5).Text * -1
    scrlTopMargin.Value = txtBody(6).Text * -1
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Me.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    tlbProject.ImageList = ImageList1
    tlbProject.Buttons(2).Image = 1
    tlbProject.Buttons(4).Image = 2
    tlbProject.Buttons(6).Image = 3
    tlbProject.Buttons(8).Image = 4
    tlbProject.Buttons(10).Image = 5
    tlbProject.Buttons(12).Image = 6
    tlbProject.Buttons(14).Image = 7
    tlbProject.Buttons(16).Image = 8
    ' Clear the datacontrol
    datRecords.DatabaseName = ""
    datRecords.RecordSource = ""
    Me.Show
'    datRecords.Refresh
    MkDir App.Path & "\HTML"
    MkDir App.Path & "\Projects"
    lstProjects.Path = App.Path & "\Projects"
    'Check for any projects (If none, create one!)
    If lstProjects.ListCount = 0 Then
        'There are no projects available, would you like to create a new project?
        Dim Response As Integer
        Response = MsgBox("There are no projects available." & vbCrLf & "Would you like to start a new project?", vbYesNo + vbInformation, "WebWizard")
        If Response = 6 Then mnuFileNew_Click 'Show create new inputbox
    Else
        mnuFileOpen_Click  'Show open dialog
    End If
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 75 Then 'If folder exists
        Resume Next
    Else
        MsgBox "Error Number " & Err.Number & " has occurred." & vbCrLf & Err.Description
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    datRecords.UpdateRecord
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
    End
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExport_Click()
    frmExport.Show
End Sub

Private Sub mnuFileNew_Click()
    Dim Response As Boolean
    ProjectFilename = SaveDialog(Me, "WebWizard Files(*.mdb)|*.mdb", "Select Webwizard Filename", App.Path & "\Projects")
    If ProjectFilename = "" Then
        Exit Sub 'User cancelled
    Else
        ProjectFilename = ProjectFilename & ".mdb"
    End If
    Response = Database_Create(ProjectFilename)  'Create the new project in the Projects folder
    If Response = True Then
        CreateFirstRecord ProjectFilename, "Template"
        datRecords.DatabaseName = ProjectFilename
        datRecords.RecordSource = "Template"
        datRecords.Refresh
    Else
        MsgBox "There was an error while creating the new project." & vbCrLf & "Please check your settings and try again." & vbCrLf & vbCrLf & "If you are trying to replace a previous project," & vbCrLf & "delete the old project first (Right-click and select Delete)"
    End If
End Sub

Private Sub mnuFileOpen_Click()
    ProjectFilename = OpenDialog(Me, "WebWizard Files(*.mdb)|*.mdb", "Select Webwizard Project", App.Path & "\Projects")
    If ProjectFilename <> "" Then
        datRecords.DatabaseName = ProjectFilename
        datRecords.RecordSource = "Template"
        datRecords.Refresh
    End If
End Sub

Private Sub mnuFileSave_Click()
    datRecords.UpdateRecord
    datRecords.Refresh
End Sub

Private Sub mnuFileWebpages_Click()
    frmPages.Show
End Sub

Private Sub mnuHelpAbout_Click()
    Dim Msg As String
    Msg = Msg & "This software is Freeware." & vbCrLf
    Msg = Msg & "You may redistribute this software provided" & vbCrLf
    Msg = Msg & "the original installation archive is intact." & vbCrLf & vbCrLf
    Msg = Msg & "Website Builder Beta is distributed 'as is'.           " & vbCrLf
    Msg = Msg & "No warranty of any kind is expressed or implied." & vbCrLf
    Msg = Msg & "You use at your own risk.  The author will not " & vbCrLf
    Msg = Msg & "be held liable for any data loss, damages, loss " & vbCrLf
    Msg = Msg & "of profits or any other kind of loss resulting  " & vbCrLf
    Msg = Msg & "from the use or misuse of this software." & vbCrLf
    MsgBox "Website Builder Beta v0.9" & vbCrLf & vbCrLf & Msg & vbCrLf & "Copyright © 2001 Hunt Computer Services", vbOKOnly + vbInformation, "About WebWizard"
End Sub

Private Sub mnuHelpContents_Click()
    frmHelp.Show
End Sub

Private Sub scrlLeftMargin_Change()
    txtBody(5).Text = Abs(scrlLeftMargin.Value)
End Sub

Private Sub scrlTopMargin_Change()
    txtBody(6).Text = Abs(scrlTopMargin.Value)
End Sub

Private Sub tlbProject_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Pages"
            mnuFileWebpages_Click
        Case "Export"
            mnuFileExport_Click
        Case "Help"
            mnuHelpContents_Click
        Case "About"
            mnuHelpAbout_Click
        Case "Exit"
            mnuFileExit_Click
    End Select
End Sub

Private Sub txtBody_Change(Index As Integer)
    Dim counter As Integer
    Dim NewCaption As String
    NewCaption = "<BODY "
    NewCaption = NewCaption & "bgcolor=" & Chr(34) & txtBody(0).Text & Chr(34) & " "
    NewCaption = NewCaption & "text=" & Chr(34) & txtBody(1).Text & Chr(34) & " "
    NewCaption = NewCaption & "link=" & Chr(34) & txtBody(2).Text & Chr(34) & " "
    NewCaption = NewCaption & "alink=" & Chr(34) & txtBody(3).Text & Chr(34) & " "
    NewCaption = NewCaption & "vlink=" & Chr(34) & txtBody(4).Text & Chr(34) & " "
    NewCaption = NewCaption & "leftmargin=" & Chr(34) & txtBody(5).Text & Chr(34) & " "
    NewCaption = NewCaption & "marginwidth=" & Chr(34) & txtBody(5).Text & Chr(34) & " "
    NewCaption = NewCaption & "topmargin=" & Chr(34) & txtBody(6).Text & Chr(34) & " "
    NewCaption = NewCaption & "marginheight=" & Chr(34) & txtBody(6).Text & Chr(34) & ">"
    lblTagPreview.Caption = NewCaption
End Sub

Private Sub CreateFirstRecord(fname As String, tbl As String)
    Dim db As Database
    Dim rs As Recordset
    Set db = Workspaces(0).OpenDatabase(fname)
    Set rs = db.OpenRecordset(tbl)
    With rs
        .AddNew
        !htmlBGColor = "#FFFFFF"
        !htmlText = "#000000"
        !htmlLink = "#0000FF"
        !htmlAlink = "#FF0000"
        !htmlVlink = "#800080"
        !htmlLeftMargin = "10"
        !htmlTopMargin = "10"
        !htmlMetaTags = "<META name='author' content='Your name or company name'>" & vbCrLf & "<META name='description' content='Type your description here'>" & vbCrLf & "<META name='keywords' content='keywords, describing, your, company'>"
        !htmlScript = "<SCRIPT>" & vbCrLf & "<!-- Insert your client side script here -->" & vbCrLf & "</SCRIPT>"
        !htmlBodyHeader = "<!-- HTML to include on every page above the contents should go here. -->" & vbCrLf & "<!-- Don't include the BODY tag - that's built from options selected in the Body Style tab -->"
        !htmlFooter = "<!-- HTML that will follow any page content should go here. -->" & vbCrLf & "<!-- (Don't add the ending </BODY> and </HTML> tags - they will automatically be added -->"
        .Update
    End With
    rs.Close
    db.Close
End Sub

