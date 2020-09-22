VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStep3 
   Caption         =   "OE Email XTract (Step 3 of 3)"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmStep3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlgSave 
      Left            =   3000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdEnd 
         Caption         =   "End"
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<<Prev"
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next>>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6480
         TabIndex        =   14
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   4575
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   4575
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save.."
            Height          =   375
            Left            =   240
            TabIndex        =   13
            Top             =   4080
            Width           =   975
         End
         Begin VB.OptionButton optRow 
            Caption         =   "Save as row delimited values"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   2880
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.OptionButton optTAB 
            Caption         =   "Save as tab delimited values"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   1800
            Width           =   3135
         End
         Begin VB.OptionButton optCSV 
            Caption         =   "Save as comma separated values (CSV)"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label6 
            Caption         =   "email3"
            Height          =   255
            Left            =   720
            TabIndex        =   12
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "email2"
            Height          =   255
            Left            =   720
            TabIndex        =   11
            Top             =   3360
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "email1"
            Height          =   255
            Left            =   720
            TabIndex        =   10
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "email1<tab space>email2<tab space>email3"
            Height          =   495
            Left            =   720
            TabIndex        =   9
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "email1,email2, email3,email4"
            Height          =   495
            Left            =   720
            TabIndex        =   8
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label lblEmails 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   255
            Left            =   2040
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Total Emails"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.ListBox lstEmails 
         Height          =   5910
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Step 3 of 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmStep3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private collEmails As Collection
Private bFileSaved As Boolean

Private Sub cmdEnd_Click()

    If Not bFileSaved Then
      If MsgBox("You have not saved the emails yet.Sure you want to stop now?", vbQuestion + vbYesNo, _
            App.Title) = vbYes Then
        Me.Hide
        frmFinish.Show
     End If
    End If
End Sub

Private Sub cmdNext_Click()
   frmFinish.Show
   Me.Hide
End Sub

Private Sub cmdPrev_Click()
  Unload Me
  frmStep2.Show
End Sub

'*****
' fill listbox with emails
' Parameters: None
' Returns   : None
'*******
Private Sub FillList()
   Dim nLoop As Integer
   Dim Item As Variant
   
    lstEmails.Clear
    
    For Each Item In collEmails
        lstEmails.AddItem Item
    Next
    
    lblEmails.Caption = collEmails.Count
    

End Sub

Private Sub cmdSave_Click()
  Dim objExport As DBXExport
  Dim nError As Integer
  Dim strfile As String
  
  cdlgSave.ShowSave
  strfile = cdlgSave.FileName
  
  Screen.MousePointer = vbHourglass
  
  If strfile <> "" Then
    If optCSV.Value Then
      Set objExport = New DBXExport
      objExport.Init collEmails, strfile
      nError = objExport.ExportAsCSV()
      If nError <> 0 Then
          MsgBox "Error in exporting as CSV File", vbExclamation + vbOKOnly, _
                  App.Title
      End If
    ElseIf optTAB.Value Then
      Set objExport = New DBXExport
      objExport.Init collEmails, strfile
      nError = objExport.ExportAsTabDelimited()
      If nError <> 0 Then
          MsgBox "Error in exporting as Tab delimited File", vbExclamation + vbOKOnly, _
                  App.Title
      End If
    ElseIf optRow.Value Then
      Set objExport = New DBXExport
      objExport.Init collEmails, strfile
      nError = objExport.ExportAsRowDelimited()
      If nError <> 0 Then
          MsgBox "Error in exporting as Row delimited File", vbExclamation + vbOKOnly, _
                  App.Title
      End If
    End If
    
    If nError = 0 Then
          MsgBox "Emails successfully saved", vbExclamation + vbOKOnly, _
                  App.Title
    End If
    bFileSaved = nError = 0
    cmdNext.Enabled = True
   End If
   Screen.MousePointer = vbNormal
   
End Sub

Private Sub Form_Load()

  Set collEmails = frmStep2.GetEmails
  FillList
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  frmFinish.Show
End Sub
