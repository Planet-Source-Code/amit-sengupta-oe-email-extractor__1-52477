VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStep1 
   Caption         =   "OE Email XTract (Step 1 of 3)"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8145
   Icon            =   "frmStep1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlgFile 
      Left            =   6840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<<Prev"
         Height          =   375
         Left            =   5040
         TabIndex        =   13
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next>>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "End"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Selected DBX files"
         ForeColor       =   &H00FF0000&
         Height          =   2415
         Left            =   120
         TabIndex        =   2
         Top             =   3720
         Width           =   7695
         Begin VB.CommandButton cmdClear2 
            Caption         =   "Clear All Entries.."
            Height          =   255
            Left            =   6000
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid grdFiles2 
            Height          =   1695
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   2990
            _Version        =   393216
            FocusRect       =   0
            OLEDropMode     =   1
         End
         Begin VB.Label lblCount2 
            Caption         =   "0 files"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "DBX File Selection"
         ForeColor       =   &H00FF0000&
         Height          =   2895
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   7695
         Begin VB.CommandButton cmdExclude 
            Caption         =   "Exclusion List.."
            Height          =   255
            Left            =   5160
            TabIndex        =   15
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdClear1 
            Caption         =   "Clear All Entries.."
            Height          =   255
            Left            =   5880
            TabIndex        =   9
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add DBX manually.."
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdIdentities 
            Caption         =   "Search OE Identities"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2175
         End
         Begin MSFlexGridLib.MSFlexGrid grdFiles1 
            Height          =   1575
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Right click for options"
            Top             =   1200
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   2778
            _Version        =   393216
            AllowBigSelection=   0   'False
            FocusRect       =   0
         End
         Begin VB.Label Label2 
            Caption         =   "Select files in the grid and right click to move them"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label lblCount1 
            Caption         =   "0 files"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Step 1 of 3 "
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
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Menu mnuOpts1 
      Caption         =   ""
      Begin VB.Menu mnuSelectOne 
         Caption         =   "&Select file(s)"
      End
   End
End
Attribute VB_Name = "frmStep1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

                        ' array of OE folder paths
Private arrOEPaths() As String
Private arrDBXfiles() As String
Private arrSelectedDbxFiles() As String

Private Sub cmdAdd_Click()
  Dim arrFiles() As String
  Dim strPath, strfile, strSelection As String
  Dim nLoop As Integer
  
  On Error GoTo ErrHandler
  
  cdlgFile.Filter = "DBX (*.dbx)|*.dbx"
  cdlgFile.Flags = cdlOFNAllowMultiselect + cdlOFNExplorer
  cdlgFile.MaxFileSize = 1000
  cdlgFile.ShowOpen
  
  If cdlgFile.FileName <> "" Then
     strSelection = cdlgFile.FileName
     If Right(strSelection, 1) <> Chr(0) Then
       strSelection = strSelection & Chr(0)
     End If
     
     arrFiles = Split(strSelection, Chr(0))
                        ' check if multiple files chosen
     If UBound(arrFiles) > 1 Then
        strPath = arrFiles(0) & "\"
     Else
        strPath = ""
     End If
     For nLoop = 0 To UBound(arrFiles) - 1
        If strPath <> "" And nLoop = 0 Then
        Else
            strfile = strPath & arrFiles(nLoop)
            If Not DBXAlreadyInList(strfile) Then
               AddDBXToSearchDBXGrid (strfile)
            Else
               MsgBox "This file is already in the list", vbOKOnly + vbExclamation, _
                           App.Title
            End If
        End If
     Next
  End If
  Exit Sub
  
ErrHandler:
    If Err.Number = 20476 Then
        MsgBox "Please select only " & cdlgFile.Max & " files at a time", _
            vbExclamation + vbOKOnly, App.Title
    Else
       MsgBox Err.Number & ":" & Err.Description, vbExclamation + vbOKOnly, _
            App.Title
    End If
End Sub

Private Sub cmdClear1_Click()
   Dim nSaveRow As Integer
   
   grdFiles1.Row = 1
   grdFiles1.Col = 1
   
   If grdFiles1.Text <> "" Then
      If MsgBox("This will clear all the files in the grid." & vbCrLf & _
                 " Proceed?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            ClearSearchedFilesGrid
      End If
   End If
                 
End Sub

Private Sub cmdClear2_Click()
   Dim nSaveRow As Integer
   
   grdFiles2.Row = 1
   grdFiles2.Col = 1
   
   If grdFiles2.Text <> "" Then
      If MsgBox("This will clear all the files in the selected grid." & vbCrLf & _
                 " Proceed?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            ClearSelectedFilesGrid
            cmdNext.Enabled = False
      Else
        grdFiles2.Row = nSaveRow
      End If
   End If

End Sub

Private Sub cmdEnd_Click()
 If MsgBox("Sure you want to stop now?", vbQuestion + vbYesNo, _
        App.Title) = vbYes Then
    Me.Hide
    frmFinish.Show
 End If
End Sub

Private Sub cmdExclude_Click()
  frmExclude.Show
End Sub

Private Sub cmdIdentities_Click()
 Dim strText As String
 Dim nSaveRow As Integer
 
 nSaveRow = grdFiles1.Row
 grdFiles1.Row = 1
 strText = grdFiles1.Text
 If strText <> "" Then
     If MsgBox("This will clear the files list.Re-search DBX files again?", vbYesNo + vbQuestion, _
                App.Title) = vbYes Then
        ClearSearchedFilesGrid
        SearchIdentities
        FillSearchedDBXGrid
     End If
 Else
        SearchIdentities
        FillSearchedDBXGrid
 End If
 
End Sub

Private Sub cmdNext_Click()
   frmStep2.Show
   Me.Hide
End Sub

Private Sub cmdPrev_Click()
  frmIntro.Show
  Me.Hide
End Sub

'****
' Sub to fill in searched files grid
' Parameters: None
' Returns   : None
'**********
Private Sub FillSearchedDBXGrid()
    Dim nLoop As Integer
                    ' clear and init grid
        frmStep1.grdFiles1.Clear
        With frmStep1.grdFiles1
            .Rows = 2
            .Cols = 2
            .FixedRows = 1
            
            .Row = 0
            .ColWidth(0) = DBX_GRID_COLWIDTH_0
            .ColWidth(1) = DBX_GRID_COLWIDTH_1
            .Col = 1
            .Text = "Searched DBX Files"
        
                        ' add rows
            For nLoop = 0 To UBound(arrDBXfiles)
                .Row = .Rows - 1
                .Col = 1
                If Len(arrDBXfiles(nLoop)) > GRID_CHARS_VISIBLE Then
                    .Text = "..." & Right(arrDBXfiles(nLoop), GRID_CHARS_VISIBLE)
                Else
                    .Text = arrDBXfiles(nLoop)
                End If
                If nLoop < UBound(arrDBXfiles) Then
                    .Rows = .Rows + 1
                End If
            Next
            
        End With
                                ' show count
        lblCount1.Caption = UBound(arrDBXfiles) & " files"
End Sub



'***
' Sub to clear searchedfiles grid
' Parameters: None
' Returns   : None
'********
Private Sub ClearSearchedFilesGrid()

    ReDim arrDBXfiles(0)
    
    grdFiles1.Rows = 2
    grdFiles1.Clear
    lblCount1.Caption = "0 files"
End Sub

'***
' Sub to clear selectedfiles grid
' Parameters: None
' Returns   : None
'********
Private Sub ClearSelectedFilesGrid()

    ReDim arrSelectedDbxFiles(0)
    
    grdFiles2.Rows = 2
    grdFiles2.Clear
    lblCount2.Caption = "0 files"
End Sub


'***
' Sub to check if a dbx file already exists in the searched dbx list
' Parameters: strFile-> filename
' Returns   : True/false
'*********
Private Function DBXAlreadyInList(ByVal strfile) As Boolean
    Dim nLoop As Integer
    
    On Error GoTo ErrHandler
    
    For nLoop = 0 To UBound(arrDBXfiles) - 1
        If LCase(arrDBXfiles(nLoop)) = LCase(strfile) Then
            DBXAlreadyInList = True
            Exit For
        End If
    Next
    
    Exit Function
    
ErrHandler:
       If Err.Number = 9 Then
             Exit Function
       Else
            Err.Raise Err.Number
       End If
    
    

End Function

'***
' Sub to add a manually added dbx to searched files grid
' Parameters: strFile-> filename
' Returns   : None
'*********
Private Sub AddDBXToSearchDBXGrid(ByVal strfile)
   Dim bInitFlag As Boolean
   Dim nSize As Integer
   
                ' check if grid is empty
    If grdFiles1.Rows = 2 Then
        grdFiles1.Row = 1
        grdFiles1.Col = 1
        If grdFiles1.Text = "" Then
          With grdFiles1
                        ' init grid
            .Rows = 2
            .Cols = 2
            .FixedRows = 1
            
            .Row = 0
            .ColWidth(0) = DBX_GRID_COLWIDTH_0
            .ColWidth(1) = DBX_GRID_COLWIDTH_1
            .Col = 1
            .Text = "Searched DBX Files"
          End With
          bInitFlag = True
        End If
    End If
                        ' add new entry
    With grdFiles1
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
    
    End With
                        ' add to array
    On Error Resume Next
    nSize = UBound(arrDBXfiles)
    If IsEmpty(nSize) Then
        nSize = 0
    End If
    On Error GoTo 0
    ReDim Preserve arrDBXfiles(nSize + 1)
    arrDBXfiles(nSize) = strfile
    
                        ' update count
    lblCount1.Caption = grdFiles1.Rows - 1 & " files"
    
End Sub


'***
' Sub to search identities for dbx files
' Parameters: None
' Returns   : None
'********
Private Sub SearchIdentities()

  Dim arrTemp() As String
  Dim nSize, nLoop, nLoop2 As Integer
  Dim objDBXFiles As New DBXFiles
  
  
  arrOEPaths = objDBXFiles.GetStoreFolder()
  If UBound(arrOEPaths) > 0 Then
                            ' get all files in each identity
     For nLoop = 0 To UBound(arrOEPaths) - 1
        Call objDBXFiles.GetDBXFilesInPath(arrOEPaths(nLoop), arrTemp)
                            ' copy to main array
        On Error Resume Next
        nSize = UBound(arrDBXfiles)
        On Error GoTo 0
        If IsEmpty(nSize) Then
            nSize = 0
        End If
        For nLoop2 = 0 To UBound(arrTemp)
          ReDim Preserve arrDBXfiles(nSize + nLoop2)
          arrDBXfiles(nSize + nLoop2) = arrTemp(nLoop2)
        Next
     Next
  End If
  
  Set objDBXFiles = Nothing

End Sub

Private Sub Form_Load()
 ReDim arrDBXfiles(0)
 ReDim arrSelectedDbxFiles(0)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Me.Hide
 frmFinish.Show
End Sub

Private Sub grdFiles1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (Button And vbRightButton) = vbRightButton Then
                        ' if grid is not empty
        If grdFiles1.Text <> "" Then
            Me.PopupMenu mnuOpts1
        End If
    End If
End Sub
'***
' Sub to add a dbx to selected files grid
' Parameters: strFile-> filename
' Returns   : None
'*********
Private Sub AddDBXToSelectedDBXGrid(ByVal strfile)
   Dim bInitFlag As Boolean
   Dim nSize As Integer
   
   If Not DBXAlreadyInSelectedList(GetFullPath(strfile)) Then
                    ' check if grid is empty
        If grdFiles2.Rows = 2 Then
            grdFiles2.Row = 1
            grdFiles2.Col = 1
            If grdFiles2.Text = "" Then
              With grdFiles2
                            ' init grid
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
            End If
        End If
                            ' add new entry
        With grdFiles2
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
        
        End With
                            
                            ' update array
        On Error Resume Next
        nSize = UBound(arrSelectedDbxFiles)
        On Error GoTo 0
        If IsEmpty(nSize) Then
            nSize = 0
        End If
        ReDim Preserve arrSelectedDbxFiles(nSize + 1)
        arrSelectedDbxFiles(nSize) = GetFullPath(strfile)
        
                            ' update count
        lblCount2.Caption = UBound(arrSelectedDbxFiles) - 1 & " files"
        
    End If
End Sub

Private Sub grdFiles1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (Button And vbLeftButton) = vbLeftButton And _
        (Shift = vbCtrlMask) Then
            grdFiles1.Drag
    End If

End Sub

Private Sub grdFiles1_OLEStartDrag(Data As MSFlexGridLib.DataObject, AllowedEffects As Long)
  
  AllowedEffects = vbDropEffectCopy
  
  Data.SetData , 1
  
End Sub

Private Sub mnuSelectOne_Click()
  
  XferFromSearchedToSelected
  cmdNext.Enabled = True
  
End Sub



'***
' Sub to transfer entries from searched grid to selected grid
' Parameters: None
' Returns   : None
'********
Private Sub XferFromSearchedToSelected()

    Dim nStart, nStop, nSaveRow, nLoop  As Integer
    Dim strText As String
    
    nSaveRow = grdFiles1.Row
    If grdFiles1.Row < grdFiles1.RowSel Then
        nStart = grdFiles1.Row
        nStop = grdFiles1.RowSel
    Else
        nStart = grdFiles1.RowSel
        nStop = grdFiles1.Row
    End If
                ' if single row selection
    If nStop = nStart Then
        strText = grdFiles1.Text
        AddDBXToSelectedDBXGrid (strText)
    Else
        For nLoop = nStart To nStop
            grdFiles1.Row = nLoop
            strText = grdFiles1.Text
            AddDBXToSelectedDBXGrid (strText)
        Next
    End If
    
    grdFiles1.Row = nSaveRow
    
End Sub

'***
' Function to check if an entry already exists in the searched dbx grid
' Parameters: strFile-> filename
' Returns   : True/false
'*********
Private Function DBXAlreadyInSelectedList(ByVal strfile) As Boolean
    Dim nLoop As Integer
    
    On Error GoTo ErrHandler
    
    For nLoop = 0 To UBound(arrSelectedDbxFiles) - 1
        If LCase(arrSelectedDbxFiles(nLoop)) = LCase(strfile) Then
            DBXAlreadyInSelectedList = True
            Exit For
        End If
    Next
    
    Exit Function
    
ErrHandler:
       If Err.Number = 9 Then
             Exit Function
       Else
            Err.Raise Err.Number
       End If
    
    

End Function

'***
' Function to return list of selected dbx files
' Parameters: None
' REturns   : arrSelectedDbxfiles-> array of selected dbx files
'*****
Public Function GetSelectedDbxFiles() As String()

    GetSelectedDbxFiles = arrSelectedDbxFiles
End Function


'***
' Function to return fullpath of file from array
' Parameters: strFile-> partial filename from grid
' REturns   : full path
'*****
Private Function GetFullPath(ByVal strfile) As String
   Dim nLoop As Integer
   
   For nLoop = 0 To UBound(arrDBXfiles) - 1
    If Right(strfile, GRID_CHARS_VISIBLE) = Right(arrDBXfiles(nLoop), GRID_CHARS_VISIBLE) Then
       GetFullPath = arrDBXfiles(nLoop)
       Exit For
    End If
   Next
End Function


