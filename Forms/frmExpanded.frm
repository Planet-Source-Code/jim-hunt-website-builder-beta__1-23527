VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmExpanded 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expanded HTML View"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmExpanded.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh Colours"
      Height          =   390
      Left            =   75
      TabIndex        =   3
      Top             =   5775
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   8250
      TabIndex        =   2
      Top             =   5775
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   7275
      TabIndex        =   1
      Top             =   5775
      Width           =   915
   End
   Begin RichTextLib.RichTextBox rtfHTML 
      BeginProperty DataFormat 
         Type            =   4
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   8
      EndProperty
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   10081
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmExpanded.frx":000C
   End
End
Attribute VB_Name = "frmExpanded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    frmPages.txtPageHTML.Text = rtfHTML.Text
    DoEvents
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Call ColorTags(0, Len(rtfHTML.Text), rtfHTML)
    cmdRefresh.Enabled = False
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    rtfHTML.Text = frmPages.txtPageHTML.Text
    Me.Refresh
    'Colour the existing tags
    Call ColorTags(0, Len(rtfHTML.Text), rtfHTML)
    cmdRefresh.Enabled = False
End Sub

Private Sub rtfHTML_Change()
    cmdRefresh.Enabled = True
End Sub
