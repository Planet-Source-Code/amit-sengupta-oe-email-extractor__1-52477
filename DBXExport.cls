VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBXExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
                        ' file type constants
Const DATA_CSV = 1
Const DATA_TAB = 2
Const DATA_ROW = 3


Dim nFileHandle As Integer
Dim strFileName As String
Dim collEmails As Collection


'***
' Function to init object
' Parameters: collE-> collection of emails
'             strName->file name
' Returns   : None
'********
Public Sub Init(ByRef collE As Collection, ByVal strName As String)
    
    Set collEmails = collE
    strFileName = strName
    On Error Resume Next
    Kill strFileName        ' delete existing file if any
    On Error GoTo 0

End Sub

'***
' Function to export as csv file
' Parameters: None
' Returns   : 0->success/else error code
'********
Public Function ExportAsCSV() As Integer
  Dim nError As Integer
  
    ExportAsCSV = WriteData(DATA_CSV)
End Function
'***
' Function to export as tab delimited file
' Parameters: none
' Returns   : 0->success/else error code
'********
Public Function ExportAsTabDelimited() As Integer
  Dim nError As Integer
  
  ExportAsTabDelimited = WriteData(DATA_TAB)

End Function

'***
' Function to export as row delimited file
' Parameters: None
' Returns   : 0->success/else error code
'********
Public Function ExportAsRowDelimited() As Integer
  Dim nError As Integer
  
  ExportAsRowDelimited = WriteData(DATA_ROW)

End Function
'***
' Sub to create file
' Parameters: None
' Returns   : 0-success/else error code
'*****
Private Function CreateFile() As Integer

    On Error GoTo ErrHandler
    nFileHandle = FreeFile
    Open strFileName For Binary Access Write As #nFileHandle
    
    CreateFile = 0
    Exit Function

ErrHandler:
    
        CreateFile = Err.Number
End Function
'***
' Sub to close file
' Parameters: None
' Returns   : 0-success/else error code
'*****
Private Function CloseFile() As Integer

    On Error GoTo ErrHandler
    Close #nFileHandle
    
    CloseFile = 0
    Exit Function

ErrHandler:
    
        CloseFile = Err.Number
End Function
'***
' Function to write data to file
' Parameters: nType-> type of file
' Returns   : 0-success/else error code
'*****
Private Function WriteData(ByVal nType As Integer)
  Dim nLoop, nError As Integer
  Dim arrData() As Byte
  Dim strData, strDelimiter As String
  
  Select Case nType
    Case DATA_CSV
        strDelimiter = ","
    Case DATA_TAB
        strDelimiter = vbTab
    Case DATA_ROW
        strDelimiter = vbCrLf
  End Select
  
  nError = CreateFile()
  If nError = 0 Then
    For nLoop = 1 To collEmails.Count
      strData = collEmails.Item(nLoop)
      arrData = StrConv(strData, vbFromUnicode)
      
      Put #nFileHandle, , arrData
      If nLoop < collEmails.Count Then
           arrData = StrConv(strDelimiter, vbFromUnicode)
          Put #nFileHandle, , arrData
      End If
    Next
  Else
    WriteData = nError
  End If

End Function

