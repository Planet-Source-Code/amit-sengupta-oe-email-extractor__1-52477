VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBXFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
                        ' dbx folders to ignore
Const FOLDERS_DBX = "folders.dbx"
Const POP3UIDL_DBX = "pop3uidl.dbx"


'****
' Function to return paths of OE dbx files
'       Searches for all identity paths of current user
' Parameters: None
' Returns   : array of path strings
'**********
Public Function GetStoreFolder() As String()
    Const IDENTITY_KEY = "HKEY_CURRENT_USER\Identities"
    Const STORE_FOLDER_KEY = "Software\Microsoft\Outlook Express"
    Dim arrStoreFolderValue() As String
    Dim collOEKeys, collOEVersion, collOEValues As Collection
    Dim varSubkey As Variant
    Dim nCount, nSize, nResult As Integer
    Dim strUserProfile As String * 255
    Dim strProfilePath As String
    Dim nLoop As Integer
    
    Set collOEKeys = EnumRegistryKeys(IDENTITY_KEY)
                                ' get identities
    If collOEKeys.Count > 0 Then
        For Each varSubkey In collOEKeys
          If Not EmptyString(varSubkey) Then
            nCount = nCount + 1
            ReDim Preserve arrStoreFolderValue(nCount)
            arrStoreFolderValue(nCount - 1) = varSubkey & "\"
          End If
        Next
    End If
                                ' get OE version
    For nLoop = 0 To UBound(arrStoreFolderValue) - 1
        arrStoreFolderValue(nLoop) = IDENTITY_KEY & "\" & arrStoreFolderValue(nLoop) & _
                        STORE_FOLDER_KEY
        Set collOEVersion = EnumRegistryKeys(arrStoreFolderValue(nLoop))
        arrStoreFolderValue(nLoop) = arrStoreFolderValue(nLoop) & "\" & collOEVersion.Item(1)
    Next
                                ' get OE folder location
    For nLoop = 0 To UBound(arrStoreFolderValue) - 1
        Set collOEValues = EnumRegistryValues(arrStoreFolderValue(nLoop))
        arrStoreFolderValue(nLoop) = collOEValues("Store Root")
    Next
                                ' get userprofile value
    nResult = ExpandEnvironmentStrings("%UserProfile%", strUserProfile, 255)
    strProfilePath = Left(strUserProfile, nResult - 1)
    
    For nLoop = 0 To UBound(arrStoreFolderValue) - 1
        'arrStoreFolderValue(nLoop) = collOEValues("Store Root")
        arrStoreFolderValue(nLoop) = Replace(arrStoreFolderValue(nLoop), "%UserProfile%", strProfilePath)
    Next
    
    GetStoreFolder = arrStoreFolderValue
End Function
'****
' Function to return all dbx files in a path
' Parameters: strPath-> file path
' Returns   : arrFiles-> array of dbx files
'**********
Public Sub GetDBXFilesInPath(ByVal strPath As String, ByRef arrFiles() As String)
    Dim strFileName As String
    Dim nCount As Integer
    
    strFileName = Dir(strPath & "*.dbx")
    If strFileName <> "" Then
        If Not ExcludeDBX(strFileName) Then
            nCount = nCount + 1
            ReDim Preserve arrFiles(nCount)
            arrFiles(0) = strPath & strFileName
        End If
        Do While strFileName <> ""
            strFileName = Dir
            If strFileName <> "" And Not ExcludeDBX(strFileName) Then
                nCount = nCount + 1
                ReDim Preserve arrFiles(nCount)
                arrFiles(nCount - 1) = strPath & strFileName
            End If
        Loop
    End If

End Sub
'****
' Function to check if a dbx folder is to be excluded
' Parameters: strFolder-> folder name
' Returns   : True/false
'**********
Private Function ExcludeDBX(ByVal strFolder As String) As Boolean

    ExcludeDBX = LCase(strFolder) = FOLDERS_DBX Or _
       LCase(strFolder) = POP3UIDL_DBX
          
End Function


