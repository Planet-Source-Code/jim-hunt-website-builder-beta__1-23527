VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmPages 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Website Builder HTML Pages"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmPages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   390
      Left            =   8175
      TabIndex        =   14
      ToolTipText     =   "Save any changes to this page"
      Top             =   600
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   390
      Left            =   8175
      TabIndex        =   13
      ToolTipText     =   "Add a new page to this project"
      Top             =   150
      Width           =   915
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   390
      Left            =   8175
      TabIndex        =   12
      ToolTipText     =   "Delete this page from the project"
      Top             =   1500
      Width           =   915
   End
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "&Back"
      Height          =   390
      Left            =   8175
      TabIndex        =   11
      ToolTipText     =   "Return to the previous form"
      Top             =   3300
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   8175
      TabIndex        =   10
      ToolTipText     =   "Cancel any changes"
      Top             =   1050
      Width           =   915
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   390
      Left            =   8175
      TabIndex        =   9
      Top             =   2175
      Width           =   915
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   390
      Left            =   8175
      TabIndex        =   8
      Top             =   2625
      Width           =   915
   End
   Begin VB.Data datPageList 
      Caption         =   "datPageList"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\WINDOWS\Desktop\WebBuilder\Other Suite Projects\WebWizard\wwdata.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1125
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pages"
      Top             =   5550
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Page from List"
      Height          =   6015
      Left            =   75
      TabIndex        =   7
      Top             =   75
      Width           =   2790
      Begin MSDBCtls.DBList dblPageList 
         Bindings        =   "frmPages.frx":030A
         Height          =   5520
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   9737
         _Version        =   393216
         ListField       =   "PageTitle"
         BoundColumn     =   "PageID"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Page Details"
      Height          =   6015
      Left            =   3000
      TabIndex        =   6
      Top             =   75
      Width           =   5040
      Begin VB.Data datPages 
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\WINDOWS\Desktop\WebBuilder\Other Suite Projects\WebsiteWizard\wwdata.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   150
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Pages"
         Top             =   5475
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.CommandButton cmdExpandedView 
         Caption         =   "Expand HTML View"
         Height          =   390
         Left            =   3075
         TabIndex        =   4
         ToolTipText     =   "Opens a full-size window to work with your HTML code"
         Top             =   5475
         Width           =   1815
      End
      Begin VB.TextBox txtPageTitle 
         DataField       =   "PageTitle"
         DataSource      =   "datPages"
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   300
         Width           =   3240
      End
      Begin VB.TextBox txtPageHTML 
         DataField       =   "PageHTML"
         DataSource      =   "datPages"
         Height          =   3990
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Text            =   "frmPages.frx":0331
         Top             =   1425
         Width           =   4740
      End
      Begin VB.TextBox txtFilename 
         DataField       =   "PageFilename"
         DataSource      =   "datPages"
         Height          =   315
         Left            =   1650
         TabIndex        =   2
         Text            =   "Text3"
         Top             =   675
         Width           =   3240
      End
      Begin VB.Label Label3 
         Caption         =   "Page HTML"
         Height          =   240
         Left            =   150
         TabIndex        =   17
         Top             =   1125
         Width           =   1440
      End
      Begin VB.Label Label2 
         Caption         =   "Filename"
         Height          =   240
         Left            =   150
         TabIndex        =   16
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Page Title"
         Height          =   240
         Left            =   150
         TabIndex        =   15
         Top             =   375
         Width           =   1440
      End
   End
   Begin VB.TextBox txtPageID 
      DataField       =   "PageID"
      DataSource      =   "datPages"
      Height          =   315
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      ToolTipText     =   "This is an automatically generated number that cannot be changed"
      Top             =   5550
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "frmPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentProject As Long
Dim NewRecord As Boolean

Private Sub cmdAdd_Click()
    NewRecord = True
    datPages.RecordSource = "Pages"
    datPages.Refresh
    datPages.Recordset.AddNew
    txtPageTitle.SetFocus
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrHandler
    datPages.Recordset.CancelUpdate
    datPages.RecordSource = "Pages"
    datPages.Refresh
    datPageList.Refresh
    datPageList.UpdateControls
    Exit Sub
ErrHandler:
    If Err.Number = 3020 Then Exit Sub 'Record wasn't changed
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    If NewRecord Then Exit Sub
    datPages.Recordset.Delete
    datPages.Recordset.MoveFirst
    datPageList.Refresh
    datPageList.UpdateControls
End Sub

Private Sub cmdExpandedView_Click()
    frmExpanded.Show 1
End Sub

Private Sub cmdExport_Click()
    If NewRecord Then Exit Sub
    frmExport.Show
End Sub

Private Sub cmdHelp_Click()
    frmHelp.Show
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    If NewRecord Then
        datPages.UpdateRecord
        datPages.Recordset.Bookmark = datPages.Recordset.LastModified
        NewRecord = False
        datPageList.Refresh
        datPageList.UpdateControls
    Else
        datPages.UpdateRecord
    End If
End Sub

Private Sub dblPageList_Click()
On Error GoTo ErrHandler
    If NewRecord Then Exit Sub
    datPages.RecordSource = "SELECT * FROM Pages WHERE PageID = " & dblPageList.BoundText
    datPages.Refresh
    Exit Sub
ErrHandler:
    If Err.Number = 3075 Then Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    Me.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    datPages.DatabaseName = frmMain.datRecords.DatabaseName
    datPages.RecordSource = "Pages"
    datPages.Refresh
    datPageList.DatabaseName = frmMain.datRecords.DatabaseName
    datPageList.RecordSource = "Pages"
    datPageList.Refresh
    NewRecord = False
    Exit Sub
ErrHandler:
    If Err.Number = 3021 Then
        MsgBox "There are currently no pages in this project." & vbCrLf & "Click the 'Add' button to insert a new page."
    Else
        MsgBox "Error Number " & Err.Number & " has occurred while loading pages." & vbCrLf & Err.Description
    End If
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    datPages.UpdateRecord
End Sub
