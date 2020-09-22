VERSION 5.00
Begin VB.Form frmIntro 
   Caption         =   "OE Email Xtract"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.PictureBox Picture2 
         Height          =   1575
         Left            =   240
         Picture         =   "frmIntro.frx":0442
         ScaleHeight     =   1515
         ScaleWidth      =   7275
         TabIndex        =   7
         Top             =   1440
         Width           =   7335
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   915
         Left            =   240
         Picture         =   "frmIntro.frx":1404
         ScaleHeight     =   855
         ScaleWidth      =   795
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "End"
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next>>"
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label lblIntro3 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   5280
         Width           =   7335
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   6135
      End
      Begin VB.Label lblTitle 
         Caption         =   "OE Email Xtract (Beta version)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   6375
      End
      Begin VB.Label lblIntro2 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   4200
         Width           =   7215
      End
      Begin VB.Label lblIntro1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   3240
         Width           =   7575
      End
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  EndProgram
End Sub

Private Sub cmdNext_Click()
  Me.Hide
  frmStep1.Show
End Sub

Private Sub Form_Load()

  lblIntro1.Caption = "OE Email XTract is a tool to extract email addresses from " & _
        "Microsoft Outlook Express dbx files. It allows manual/auto selection of " & _
        "DBX files from your computer and then scans them for emails. The emails can then be exported into a file."
        
  lblIntro2.Caption = "The process is broken down into 3 simple steps. Press the " & _
     "Next button below to start the process"
  
  lblIntro3.Caption = "It is preferable to close Outlook Express before running this application."
  
  lblTitle.Caption = "OE Email Xtract version " & App.Major & "." & App.Minor & "." & App.Revision & "(Beta)"
  lblCopyright.Caption = "(C) 2004, Amit Sengupta, amit@logical-magic.com"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
  frmFinish.Show
End Sub
