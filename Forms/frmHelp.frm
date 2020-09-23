VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Website Builder Help"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close (Hidden!)"
      Default         =   -1  'True
      Height          =   390
      Left            =   7275
      TabIndex        =   0
      Top             =   5700
      Width           =   1590
   End
   Begin RichTextLib.RichTextBox rtfHelp 
      Height          =   6045
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   10663
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      FileName        =   "C:\WINDOWS\Desktop\WebsiteBuilderBeta\Notes\Documentation.rtf"
      TextRTF         =   $"frmHelp.frx":030A
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    cmdClose.Top = ScaleHeight + 60 'Hide close button
End Sub
