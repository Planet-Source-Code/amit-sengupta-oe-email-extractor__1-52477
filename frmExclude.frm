VERSION 5.00
Begin VB.Form frmExclude 
   Caption         =   "Exclusion List"
   ClientHeight    =   3345
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdAbort 
         Caption         =   "Abort"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Save"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtAdd 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ListBox lstExclude 
         Height          =   2400
         Left            =   3360
         TabIndex        =   1
         ToolTipText     =   "Right clck for options"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Enter new exlusion entry"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblText1 
         Caption         =   "Label1"
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   ""
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmExclude"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objEmails As Collection

Private Sub cmdAbort_Click()
  Unload Me
End Sub

Private Sub cmdAdd_Click()
  If Len(Trim(txtAdd.Text)) > 0 Then
     If AddEntry(txtAdd.Text) <> 0 Then
        MsgBox "Entry already exists", vbExclamation + vbOKOnly, _
            App.Title
     Else
        txtAdd.Text = ""
     End If
  Else
    MsgBox "No entry to add", vbExclamation + vbOKOnly, _
        App.Title
  End If
End Sub

Private Sub cmdClose_Click()
  If objEmails.Count > 0 Then
     SaveList (App.Path & "\" & EXCLUSION_FILE)
  End If
  Unload Me
End Sub

Private Sub Form_Load()

    lblText1.Caption = "Emails added in this list will be excluded from the " & _
                "scanning process. You can enter full email ids like email@domain.com " & _
                " or partial ids like @domain.com"
                    
    Set objEmails = New Collection
    LoadList App.Path & "\" & EXCLUSION_FILE

    
End Sub
'***
' Function to add a new entry
' Parameters: strEntry-> entry
' Returns   : 0-success/error code
'*********
Private Function AddEntry(ByVal strEntry) As Integer
   
   On Error GoTo ErrHandler
   
   objEmails.Add LCase(strEntry), LCase(strEntry)
   lstExclude.AddItem LCase(strEntry)
   
   Exit Function
   
ErrHandler:
    AddEntry = Err.Number
    
    If Err.Number <> 457 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation + vbOKOnly, _
                App.Title
    End If
   
End Function
'***
' Sub to load exclusion list from file
' Parameters: strfile-> file name
' Returns   : None
'*********
Private Sub LoadList(ByVal strfile As String)
    Dim nFileHandle As Integer
    Dim strLine As String
    
    On Error GoTo ErrHandler
                    ' check if file exists
    If Dir(strfile) <> "" Then
        nFileHandle = FreeFile
        Open strfile For Input As #nFileHandle
        While Not EOF(nFileHandle)
            Input #nFileHandle, strLine
            AddEntry strLine
        Wend
        
        Close #nFileHandle
    End If
    
    Exit Sub
    
ErrHandler:
       MsgBox Err.Number & ": " & Err.Description, vbExclamation + vbOKOnly, _
            App.Title
End Sub
'***
' Sub to save exclusion list to file
' Parameters: strfile-> file name
' Returns   : None
'*********
Private Sub SaveList(ByVal strfile As String)
    Dim nFileHandle As Integer
    Dim strLine As String
    Dim Item As Variant
    
    On Error GoTo ErrHandler
                    ' delete if file exists
    If Dir(strfile) <> "" Then
        Kill (strfile)
    End If
        
    nFileHandle = FreeFile
    Open strfile For Output As #nFileHandle
    For Each Item In objEmails
        Print #nFileHandle, Item
    Next
    
    Close #nFileHandle
    
    Exit Sub
ErrHandler:
       MsgBox Err.Number & ": " & Err.Description, vbExclamation + vbOKOnly, _
            App.Title
    
End Sub

Private Sub lstExclude_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    If lstExclude.Text <> "" Then
        DeleteEntry lstExclude.Text
    End If
  End If
End Sub

Private Sub lstExclude_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbRightButton Then
    Me.PopupMenu mnuOpts
  End If
End Sub

Private Sub mnuDelete_Click()
    If lstExclude.Text <> "" Then
        DeleteEntry lstExclude.Text
    End If
End Sub

'***
' Sub to delete entry from list
' Parameters: strentry-> entry
' Returns   : None
'*********
Private Sub DeleteEntry(ByVal strEntry As String)
    Dim Item As Variant
    
    For Each Item In objEmails
        If Item = strEntry Then
            objEmails.Remove Item
            Exit For
        End If
    Next
                        'refresh list
    lstExclude.Clear
    For Each Item In objEmails
        lstExclude.AddItem Item
    Next
End Sub

