VERSION 5.00
Begin VB.Form frmFinish 
   Caption         =   "OE Email Xtract"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmFinish.frx":0000
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
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   915
         Left            =   120
         Picture         =   "frmFinish.frx":0442
         ScaleHeight     =   855
         ScaleWidth      =   795
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Again"
         Height          =   375
         Left            =   5280
         TabIndex        =   2
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "Finish"
         Height          =   375
         Left            =   6600
         TabIndex        =   1
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblEnd 
         Alignment       =   2  'Center
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
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   7215
      End
      Begin VB.Label lblTitle 
         Caption         =   "Label1"
         Height          =   735
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmFinish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
  EndProgram
  
End Sub

Private Sub cmdStart_Click()
  EndProgram
  frmIntro.Show
End Sub

Private Sub Form_Load()
   lblTitle.Caption = frmIntro.lblTitle.Caption
   lblTitle.Font = frmIntro.lblTitle.Font
   lblTitle.FontSize = frmIntro.lblTitle.FontSize
   lblTitle.FontBold = frmIntro.lblTitle.FontBold
   
   lblEnd.Caption = "Thank you for using OE Email Xtract"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
  EndProgram
End Sub
