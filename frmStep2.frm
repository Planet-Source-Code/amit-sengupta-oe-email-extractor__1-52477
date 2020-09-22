VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmStep2 
   Caption         =   "OE Email XTract (Step 2 of 3)"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmStep2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdEnd 
         Caption         =   "End"
         Height          =   375
         Left            =   3240
         TabIndex        =   21
         Top             =   6600
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<<Prev"
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   6600
         Width           =   1455
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next>>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   19
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Search Emails"
         ForeColor       =   &H00FF0000&
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   7695
         Begin VB.Frame Frame4 
            Caption         =   "Search Statistics"
            ForeColor       =   &H00FF0000&
            Height          =   2295
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   4935
            Begin ComctlLib.ProgressBar prgFile 
               Height          =   255
               Left            =   1800
               TabIndex        =   15
               Top             =   1560
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   1
            End
            Begin ComctlLib.ProgressBar prgTotal 
               Height          =   255
               Left            =   1800
               TabIndex        =   16
               Top             =   1920
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   1
            End
            Begin VB.Label Label5 
               Caption         =   "Total Progress"
               Height          =   255
               Left            =   240
               TabIndex        =   18
               Top             =   1920
               Width           =   1335
            End
            Begin VB.Label Label4 
               Caption         =   "Current File"
               Height          =   255
               Left            =   240
               TabIndex        =   17
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Label lblDuplicates 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Height          =   255
               Left            =   1800
               TabIndex        =   14
               Top             =   1080
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Duplicates ignored"
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label lblEmailsAdded 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Height          =   255
               Left            =   1800
               TabIndex        =   12
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Emails added"
               Height          =   255
               Left            =   240
               TabIndex        =   11
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label lblFilesProcessed 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Height          =   255
               Left            =   1800
               TabIndex        =   10
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Files Processed"
               Height          =   255
               Left            =   240
               TabIndex        =   9
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop Search"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1920
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "Start Search"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.ListBox lstEmails 
            Height          =   2790
            ItemData        =   "frmStep2.frx":0442
            Left            =   5160
            List            =   "frmStep2.frx":0444
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Selected DBX Files"
         ForeColor       =   &H00FF0000&
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   7695
         Begin MSFlexGridLib.MSFlexGrid grdFiles1 
            Height          =   1815
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   3201
            _Version        =   393216
         End
         Begin VB.Label lblCount 
            Caption         =   "0 files"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Step 2 of 3"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmStep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrSelectedDbxFiles() As String
Private bInterrupted As Boolean
Private nEmailsAdded As Integer
Private nDuplicates As Integer
Private nFileHandle As Integer
Private collEmails As Collection
Private collExclude As Collection

Private Sub cmdEnd_Click()
 If MsgBox("Sure you want to stop now?", vbQuestion + vbYesNo, _
        App.Title) = vbYes Then
    Me.Hide
    frmFinish.Show
 End If

End Sub

Private Sub cmdNext_Click()
   frmStep3.Show
   Me.Hide
End Sub

Private Sub cmdPrev_Click()
  frmStep1.Show
  Me.Hide
End Sub

'***
' Sub to get list of selected dbx files
' Parameters: None
' Returns   : None
'*****
Private Sub GetSelectedDbxFiles()
   Dim bInitFlag As Boolean
   Dim nSize, nLoop As Integer
   Dim strfile As String
                    
                    ' get it from prev form
   arrSelectedDbxFiles = frmStep1.GetSelectedDbxFiles
                    
          With grdFiles1
                        ' init grid
            .Clear
            .Rows = 2
            .Cols = 2
            .FixedRows = 1
            
            .Row = 0
            .ColWidth(0) = DBX_GRID_COLWIDTH_0
            .ColWidth(1) = DBX_GRID_COLWIDTH_1
            .Col = 1
            .Text = "Selected DBX Files"
          End With
          bInitFlag = True
                        ' add entries
    With grdFiles1
    For nLoop = 0 To UBound(arrSelectedDbxFiles) - 1
        If nLoop > 0 Then
             .Rows = .Rows + 1
        End If
        .Row = .Rows - 1
        strfile = arrSelectedDbxFiles(nLoop)
        If Not bInitFlag Then
            .Rows = .Rows + 1
        End If
        .Row = .Rows - 1
        .Col = 1
        If Len(strfile) > GRID_CHARS_VISIBLE Then
            .Text = "..." & Right(strfile, GRID_CHARS_VISIBLE)
        Else
            .Text = strfile
        End If
    Next
    
    End With
                        ' show count
    lblCount.Caption = UBound(arrSelectedDbxFiles) - 1 & " files"
                    

End Sub

Private Sub cmdStart_Click()
  
  cmdStart.Enabled = False
  cmdStop.Enabled = True
  cmdPrev.Enabled = False
  cmdNext.Enabled = False
  
  StartSearch
  If collEmails.Count > 0 Then
    cmdNext.Enabled = True
  End If
  
  cmdStart.Enabled = True
  cmdStop.Enabled = False
  bInterrupted = False
  cmdPrev.Enabled = True
  cmdNext.Enabled = True
  
End Sub

Private Sub cmdStop_Click()
   
   bInterrupted = True
   cmdPrev.Enabled = True
   cmdNext.Enabled = True
   cmdStart.Enabled = True
   
End Sub

Private Sub Form_Activate()
  GetSelectedDbxFiles
  GetExclusionList (App.Path & "\" & EXCLUSION_FILE)
  
End Sub

'***
' Sub to start main search loop
' Parameters: None
' Returns   : None
'*****
Private Sub StartSearch()
  Dim nLoop, nMax As Integer
                
                 ' clear stats display
  lblFilesProcessed.Caption = "0"
  lblEmailsAdded.Caption = "0"
  lblDuplicates.Caption = "0"
  prgFile.Value = 0
  prgTotal.Value = 0
  lstEmails.Clear
  nLoop = 0
  nEmailsAdded = 0
  nDuplicates = 0
  
  Set collEmails = New Collection
  
  nMax = UBound(arrSelectedDbxFiles)
  Do While Not bInterrupted And nLoop < nMax
  
     If Len(Trim(arrSelectedDbxFiles(nLoop))) > 5 Then
        SearchEachFile (arrSelectedDbxFiles(nLoop))
     End If
            ' update stats
     lblFilesProcessed.Caption = nLoop + 1
     prgTotal.Value = Int(((nLoop + 1) * 100) / nMax)
     DoEvents
     Sleep (10)
     nLoop = nLoop + 1
  Loop
  
     
  
End Sub
'***
' Sub to search each file
' Parameters: strFile-> filename
' Returns   : None
'*****
Private Sub SearchEachFile(ByVal strfile As String)
  Dim strLine As String
  Dim arrBytes() As Byte
  Dim nSize, nCurrPos As Long
  Dim nHandle As Integer
  
  On Error GoTo ErrHandler
  nSize = FileLen(strfile)      ' get filesize
  
  nHandle = FreeFile
  Open strfile For Binary As #nHandle
  
  
  prgFile.Value = 0
  While Not EOF(nHandle) And Not bInterrupted
    DoEvents
    ReDim arrBytes(0 To 2048)
    Get #nHandle, , arrBytes
    strLine = StrConv(arrBytes, vbUnicode)
    DoScan strLine
                        ' update stats
    nCurrPos = Loc(nHandle)
    If nCurrPos > nSize Then
        nCurrPos = nSize
    End If
    prgFile.Value = Int((nCurrPos * 100) / nSize)
    
  Wend
  
  Close #nHandle
  
  Exit Sub
  
ErrHandler:
     MsgBox Err.Number & ", " & Err.Description, vbExclamation + vbOKOnly, _
           App.Title
     Close #1
End Sub


'***
' Sub to search for emails in file data
' Parameters: strData-> file data
' Returns   : None
'*********
Private Sub DoScan(ByVal strData As String)
    Dim objRegExp As New RegExp
    Dim objMatches, objMatch As Object
    Dim strMatch As String
    Dim nLoop As Integer
    
    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
                            ' regexp for email
    objRegExp.Pattern = "[a-zA-Z][\w\.\-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]"
    objRegExp.MultiLine = True
    
    On Error Resume Next
    Set objMatches = objRegExp.Execute(strData)
    If Err.Number = 0 Then
        For nLoop = 0 To objMatches.Count - 1
            strMatch = objMatches(nLoop)
                        ' update collection if not added already
                Err.Clear
                collEmails.Add strMatch, strMatch
                If Err.Number = 0 Then  ' duplicate not found
                    If Not ToBeExcluded(strMatch) Then
                                ' update stats
                        nEmailsAdded = nEmailsAdded + 1
                        lblEmailsAdded.Caption = nEmailsAdded
                        
                        lstEmails.AddItem strMatch
                    End If
                Else
                    nDuplicates = nDuplicates + 1
                    lblDuplicates.Caption = nDuplicates
                End If
        Next
    End If
    
    On Error GoTo 0
    Set objMatch = Nothing
    Set objRegExp = Nothing

End Sub


'***
' Sub to create file for storing emails
' Parameters: strName-> filename
' Returns   : Filehandle
'****
Private Function CreateEmailFile(ByVal strName As String) As Integer
   Dim nHandle As Integer
   
   On Error GoTo ErrHandler
   nHandle = FreeFile
   Open strName For Output As #nHandle
   
   CreateEmailFile = nHandle
   
   Exit Function
   
ErrHandler:
      MsgBox Err.Number & ", " & Err.Description, vbExclamation + vbOKOnly, _
                App.Title
      Close #1
End Function

'***
' Sub to close email file
' Parameters: nHandle-> handle number
' Returns   : None
'****
Private Sub CloseEmailFile(ByVal nHandle As Integer)
    
    Close #nHandle
End Sub


'***
' Sub to add email to file
' Parameters: nHandle-> handle number
'             strLine->email
' Returns   : None
'****
Private Sub WriteToFile(ByVal nHandle As Integer, ByVal strLine As String)
    
    Print #nHandle, strLine
End Sub
'***
' Function to return email collection
' Parameters: None
' Returns   : collEmails
'****
Public Function GetEmails() As Collection

    Set GetEmails = collEmails
End Function


'***
' Sub to get exclusion list
' Parameters: strFile-> filename
' Returns   : None
'****
Private Sub GetExclusionList(ByVal strfile As String)
    Dim nFileHandle As Integer
    Dim strLine As String
    
    On Error GoTo ErrHandler
    Set collExclude = Nothing
    Set collExclude = New Collection
                    ' check if file exists
    If Dir(strfile) <> "" Then
       
        nFileHandle = FreeFile
        Open strfile For Input As #nFileHandle
        While Not EOF(nFileHandle)
            Input #nFileHandle, strLine
            collExclude.Add strLine, strLine
        Wend
        
        Close #nFileHandle
    End If
    
    Exit Sub
    
ErrHandler:
       MsgBox Err.Number & ": " & Err.Description, vbExclamation + vbOKOnly, _
            App.Title

End Sub


'***
' Function to check if email is in exclusion list
' Parameters: strEmail-> email to check
' Returns   : false/true
'****
Private Function ToBeExcluded(ByVal strEmail As String) As Boolean
    Dim Item As Variant
    
    strEmail = LCase(strEmail)
    For Each Item In collExclude
        If InStr(strEmail, LCase(Item)) > 0 Then
            ToBeExcluded = True
            Exit For
        End If
    Next
    
End Function


Private Sub Form_Load()
  ReDim arrSelectedDbxFiles(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Hide
  frmFinish.Show

End Sub
